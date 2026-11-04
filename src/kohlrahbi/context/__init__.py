"""
Export of the explanatory text around the AHB tables ("context").

The AHB scraping (kohlrahbi ahb/conditions/changehistory) only reads tables and skips all other text of the
AHB documents. This module collects exactly that skipped text: everything outside the AHB tables, the package table
and the change history. It is split into passages (one per section) which carry anchors to the keys of the
structured exports, so agents can find the explanation for a Prüfidentifikator, segment, condition, ... without the
text being mixed into the structured data. See kohlrahbi.context.models.
"""

import json
import re
from collections import defaultdict
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from pathlib import Path

import docx
from docx.document import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from efoli import EdifactFormat, EdifactFormatVersion, get_format_of_pruefidentifikator

from kohlrahbi.ahb import get_pruefi_to_file_mapping
from kohlrahbi.conditions import find_all_files_from_all_pruefis
from kohlrahbi.context.anchors import find_anchors, find_explicit_pruefis, merge_anchors
from kohlrahbi.context.models import Anchor, AnchorKind, AnchorSource, ContextDocument, Passage, PruefiContext
from kohlrahbi.logger import logger
from kohlrahbi.read_functions import (
    get_all_paragraphs_and_tables,
    get_style_name,
    is_item_header_of_change_history_section,
    is_item_table_with_pruefidentifikatoren,
    is_last_row_end_of_message,
)
from kohlrahbi.seed import Seed

_heading_style = re.compile(r"^(?:Heading|Überschrift)\s*(\d)$")
_INTERMEDIATE_HEADING_STYLE = "Zwischenüberschrift"
_LIST_STYLE_PREFIXES = ("List", "Liste")
_PACKAGE_SECTION_HEADING = "Übersicht der Pakete"
_whitespace = re.compile(r"[ \t ]+")

HeadingPath = tuple[str, ...]


def _normalize(text: str) -> str:
    """Replaces tabs and non-breaking spaces, collapses runs of spaces and strips every line."""
    lines = (_whitespace.sub(" ", line).strip() for line in text.splitlines())
    return "\n".join(line for line in lines if line)


def _heading_level(style_name: str) -> int | None:
    match = _heading_style.match(style_name)
    return int(match.group(1)) if match else None


def table_to_markdown(table: Table) -> str:
    """
    Renders a docx table as Markdown table. The first row is used as header row.
    Horizontally merged cells are padded with empty cells so that the columns stay aligned; the text of nested
    tables is included in the cell of the outer table.
    """
    rows: list[list[str]] = []
    for tr in table._tbl.tr_lst:  # pylint:disable=protected-access
        cells: list[str] = []
        for tc in tr.tc_lst:
            paragraphs = (_normalize(Paragraph(p, table).text) for p in tc.iter(qn("w:p")))
            text = "<br>".join(p.replace("\n", "<br>") for p in paragraphs if p)
            cells.append(text.replace("|", "\\|"))
            cells.extend([""] * (tc.grid_span - 1))
        if any(cells):
            rows.append(cells)
    if not rows:
        return ""
    width = max(len(row) for row in rows)
    rows = [row + [""] * (width - len(row)) for row in rows]
    lines = ["| " + " | ".join(rows[0]) + " |", "|" + "---|" * width]
    lines += ["| " + " | ".join(row) + " |" for row in rows[1:]]
    return "\n".join(lines)


@dataclass
class _Block:
    text: str
    is_list_item: bool = False
    is_table: bool = False


@dataclass
class _RawPassage:
    heading_path: HeadingPath
    blocks: list[_Block]

    def to_markdown(self) -> str:
        """Joins the blocks; consecutive list items are separated by a single line break only."""
        parts: list[str] = []
        previous: _Block | None = None
        for block in self.blocks:
            if previous is not None:
                parts.append("\n" if previous.is_list_item and block.is_list_item else "\n\n")
            parts.append(f"- {block.text}" if block.is_list_item else block.text)
            previous = block
        return "".join(parts)


@dataclass
class ExtractionResult:
    """The passages of a document (without ids and anchors yet) and the Prüfis of the AHB tables per section."""

    passages: list[_RawPassage] = field(default_factory=list)
    table_pruefis_by_section: dict[HeadingPath, list[str]] = field(default_factory=lambda: defaultdict(list))

    @property
    def table_pruefis(self) -> list[str]:
        """All Prüfidentifikatoren of the AHB tables of the document, in order of appearance."""
        return list(dict.fromkeys(p for pruefis in self.table_pruefis_by_section.values() for p in pruefis))


class _PassageCollector:
    """Keeps track of the current heading path and collects the blocks below it."""

    def __init__(self) -> None:
        self.result = ExtractionResult()
        self._headings: list[tuple[int, str]] = []  # (level, text); intermediate headings have level 99
        self._blocks: list[_Block] = []

    @property
    def heading_path(self) -> HeadingPath:
        """The texts of the current headings, outermost first."""
        return tuple(text for _, text in self._headings)

    def flush(self) -> None:
        """Closes the current passage (if it has content)."""
        if self._blocks:
            self.result.passages.append(_RawPassage(heading_path=self.heading_path, blocks=self._blocks))
            self._blocks = []

    def enter_heading(self, level: int, text: str) -> None:
        """Starts a new section. Intermediate headings (level 99) replace each other but no real heading."""
        self.flush()
        self._headings = [(lvl, txt) for lvl, txt in self._headings if lvl < level]
        if text:
            self._headings.append((level, text))

    def add(self, block: _Block) -> None:
        """Adds a paragraph, list item or table to the current passage."""
        self._blocks.append(block)


def extract_passages(document: Document) -> ExtractionResult:
    """
    Walks through the document and collects everything that is neither an AHB table, nor the package table, nor
    part of the change history. The front matter before the first heading (title page, disclaimer) is skipped, too.

    An AHB table starts with a table whose header contains Prüfidentifikatoren and ends with the table whose last row
    is the end of the message (UNT/UNZ); all tables in between are its headless continuations.
    """
    collector = _PassageCollector()
    started = False  # the front matter ends with the first heading
    in_ahb_table = False
    in_package_section = False

    for item in get_all_paragraphs_and_tables(document):
        if isinstance(item, Table):
            if not started:
                continue
            if is_item_table_with_pruefidentifikatoren(item):
                collector.flush()
                in_ahb_table = True
                section_pruefis = collector.result.table_pruefis_by_section[collector.heading_path]
                new_pruefis = Seed.from_table(item).pruefidentifikatoren
                section_pruefis.extend(p for p in new_pruefis if p not in section_pruefis)
            elif in_ahb_table or in_package_section:
                collector.flush()
            else:
                markdown = table_to_markdown(item)
                if markdown:
                    collector.add(_Block(text=markdown, is_table=True))
            if in_ahb_table and is_last_row_end_of_message(item):
                in_ahb_table = False
            continue

        style_name = get_style_name(item)
        level = _heading_level(style_name)
        text = _normalize(item.text)
        if level is not None:
            if is_item_header_of_change_history_section(item, style_name):
                break
            started = True
            in_package_section = _PACKAGE_SECTION_HEADING in text
            collector.enter_heading(level, text)
        elif not started or not text or style_name.upper().startswith("TOC"):
            continue
        elif style_name == _INTERMEDIATE_HEADING_STYLE:
            collector.enter_heading(99, text)
        else:
            collector.add(_Block(text=text, is_list_item=style_name.startswith(_LIST_STYLE_PREFIXES)))

    collector.flush()
    return collector.result


def build_passages(extraction: ExtractionResult, document_id: str, known_pruefis: Iterable[str]) -> list[Passage]:
    """
    Turns the raw passages into Passages with ids and anchors.
    Besides the anchors found in the text itself (see kohlrahbi.context.anchors), a passage gets
    - the Prüfis of the AHB tables in the same section (source: table_in_section) and
    - the Prüfis explicitly named in a passage of an enclosing (or the same) section (source: section_scope).
    """
    known = set(known_pruefis) | set(extraction.table_pruefis)
    texts = [raw.to_markdown() for raw in extraction.passages]
    explicit_by_section: dict[HeadingPath, list[str]] = defaultdict(list)
    for raw, text in zip(extraction.passages, texts, strict=True):
        explicit_by_section[raw.heading_path].extend(find_explicit_pruefis(text))

    passages: list[Passage] = []
    for order, (raw, text) in enumerate(zip(extraction.passages, texts, strict=True)):
        path = raw.heading_path
        anchors: list[Anchor] = find_anchors(text, known)
        anchors += [
            Anchor(kind=AnchorKind.PRUEFIDENTIFIKATOR, value=p, source=AnchorSource.TABLE_IN_SECTION)
            for p in extraction.table_pruefis_by_section.get(path, [])
        ]
        anchors += [
            Anchor(kind=AnchorKind.PRUEFIDENTIFIKATOR, value=p, source=AnchorSource.SECTION_SCOPE)
            for section, pruefis in explicit_by_section.items()
            if path[: len(section)] == section
            for p in pruefis
        ]
        passages.append(
            Passage(
                id=f"{document_id}:{order:04d}",
                order=order,
                heading_path=list(path),
                text=text,
                contains_table=any(block.is_table for block in raw.blocks),
                anchors=merge_anchors(anchors),
            )
        )
    return passages


def _is_of_format(pruefi: str, edifact_format: EdifactFormat) -> bool:
    try:
        return get_format_of_pruefidentifikator(pruefi) is edifact_format
    except ValueError:
        return False


def get_pruefi_contexts(context_document: ContextDocument) -> list[PruefiContext]:
    """
    Groups the passages of a document by Prüfidentifikator: one PruefiContext for each Prüfi of the document's
    format that has an AHB table in the document or is explicitly described by a passage.
    """
    explicit = [
        anchor.value
        for passage in context_document.passages
        for anchor in passage.anchors
        if anchor.kind is AnchorKind.PRUEFIDENTIFIKATOR and anchor.source is AnchorSource.EXPLICIT
    ]
    pruefis = [
        p
        for p in dict.fromkeys([*context_document.pruefidentifikatoren, *explicit])
        if _is_of_format(p, context_document.edifact_format)
    ]
    return [
        PruefiContext(
            pruefidentifikator=pruefi,
            edifact_format=context_document.edifact_format,
            format_version=context_document.format_version,
            source_file=context_document.source_file,
            passages=[
                passage
                for passage in context_document.passages
                if any(a.kind is AnchorKind.PRUEFIDENTIFIKATOR and a.value == pruefi for a in passage.anchors)
            ],
        )
        for pruefi in pruefis
    ]


def get_context_document(
    document: Document | ExtractionResult,
    source_file: str,
    edifact_format: EdifactFormat,
    format_version: EdifactFormatVersion,
    known_pruefis: Iterable[str] = (),
) -> ContextDocument:
    """
    Extracts the context (explanatory text with anchors) of one AHB document for the given format.
    An already extracted document can be passed to avoid walking through it again for each format it contains.
    """
    extraction = document if isinstance(document, ExtractionResult) else extract_passages(document)
    return ContextDocument(
        edifact_format=edifact_format,
        format_version=format_version,
        source_file=source_file,
        pruefidentifikatoren=[p for p in extraction.table_pruefis if _is_of_format(p, edifact_format)],
        passages=build_passages(extraction, document_id=Path(source_file).stem, known_pruefis=known_pruefis),
    )


def get_context_directory(output_path: Path, edifact_format: EdifactFormat) -> Path:
    """The directory of the context export of a format, next to flatahb/csv/xlsx."""
    return output_path / str(edifact_format) / "context"


def dump_context_document(context_document: ContextDocument, output_path: Path) -> list[Path]:
    """
    Writes <FORMAT>/context/<document>.json and one <FORMAT>/context/pruefi/<pruefi>.json per Prüfidentifikator.
    Returns the paths of the written files.
    """
    directory = get_context_directory(output_path, context_document.edifact_format)
    (directory / "pruefi").mkdir(parents=True, exist_ok=True)
    document_path = directory / f"{Path(context_document.source_file).stem}.json"
    document_path.write_text(context_document.model_dump_json(indent=2), encoding="utf-8")
    written = [document_path]
    for pruefi_context in get_pruefi_contexts(context_document):
        pruefi_path = directory / "pruefi" / f"{pruefi_context.pruefidentifikator}.json"
        pruefi_path.write_text(pruefi_context.model_dump_json(indent=2), encoding="utf-8")
        written.append(pruefi_path)
    return written


def dump_json_schemas(output_path: Path) -> list[Path]:
    """Writes the JSON schemas of the context export to <output>/context_schemas/."""
    directory = output_path / "context_schemas"
    directory.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []
    for name, model in (("context_document", ContextDocument), ("pruefi_context", PruefiContext)):
        path = directory / f"{name}.schema.json"
        path.write_text(json.dumps(model.model_json_schema(), indent=2, ensure_ascii=False), encoding="utf-8")
        written.append(path)
    return written


def scrape_context(
    basic_input_path: Path,
    output_path: Path,
    format_version: EdifactFormatVersion,
    *,
    on_start: Callable[[int], None] | None = None,
    on_file: Callable[[str], None] | None = None,
) -> list[ContextDocument]:
    """
    Exports the context of all AHB documents of the format version (the same documents ``kohlrahbi ahb`` uses).
    ``on_start`` is invoked once with the number of documents and ``on_file`` after each document.
    """
    pruefi_to_file_mapping = get_pruefi_to_file_mapping(basic_input_path, format_version)
    format_to_files = find_all_files_from_all_pruefis(pruefi_to_file_mapping)
    files = list(dict.fromkeys(file for files in format_to_files.values() for file in files))
    if on_start is not None:
        on_start(len(files))
    dump_json_schemas(output_path)

    context_documents: list[ContextDocument] = []
    for file in files:
        path = basic_input_path / "edi_energy_de" / format_version.value / file
        logger.info("Extracting the context of %s", file)
        document = docx.Document(str(path))
        extraction = extract_passages(document)
        for edifact_format in (f for f, fs in format_to_files.items() if file in fs):
            context_document = get_context_document(
                extraction, file, edifact_format, format_version, known_pruefis=pruefi_to_file_mapping
            )
            dump_context_document(context_document, output_path)
            context_documents.append(context_document)
        if on_file is not None:
            on_file(file)
    return context_documents
