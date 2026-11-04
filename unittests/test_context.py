import json
from pathlib import Path

import docx
import pytest
from docx.document import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.table import Table
from efoli import EdifactFormat, EdifactFormatVersion
from typer.testing import CliRunner

import kohlrahbi.context
from kohlrahbi import app
from kohlrahbi.context import (
    build_passages,
    dump_context_document,
    dump_json_schemas,
    extract_passages,
    get_context_document,
    get_pruefi_contexts,
    table_to_markdown,
)
from kohlrahbi.context.anchors import find_anchors, find_explicit_pruefis, merge_anchors
from kohlrahbi.context.models import Anchor, AnchorKind, AnchorSource, ContextDocument, PruefiContext
from unittests import path_to_test_files_fv2310, strip_ansi


def _anchor_values(anchors: list[Anchor], kind: AnchorKind) -> list[tuple[str, AnchorSource]]:
    return [(a.value, a.source) for a in anchors if a.kind is kind]


class TestAnchors:
    def test_explicit_pruefis(self) -> None:
        text = (
            "Tabellenspalte = Lastgang Messlokation, Netzkoppelpunkt, Netzlokation 13018\n"
            "Tabellenspalte = Werte nach Typ 2 13027\n"
            "Kommunikationspartner, Identifikationsangabe und Art der Werte für diesen Anwendungsfall "
            "Prüfidentifikator: 13025\n"
            "Der LF sendet die SDÄ (Prüfidentifikator 55230) an den NB."
        )
        assert find_explicit_pruefis(text) == ["13018", "13027", "13025"]

    def test_mentions(self) -> None:
        text = (
            "Bei der Übermittlung ist im BGM-Segment DE1001 der Qualifier Z41 zu verwenden. "
            "Die Angabe erfolgt über SG10 DTM+163 und SG 10 DTM + 164, siehe auch SG9 PIA+5 DE 7140 und "
            "Datenelement 6060. Es gelten [492] und das Paket [1P0..1], geprüft wird mit E_0623. "
            "Die Anmeldung (55001) und der Betrag 12345 EUR, der MSB und die NNE."
        )
        anchors = find_anchors(text, known_pruefis=["55001"])
        assert _anchor_values(anchors, AnchorKind.PRUEFIDENTIFIKATOR) == [("55001", AnchorSource.MENTION)]
        assert [v for v, _ in _anchor_values(anchors, AnchorKind.SEGMENT_GROUP)] == ["SG10", "SG9"]
        assert [v for v, _ in _anchor_values(anchors, AnchorKind.SEGMENT)] == ["DTM+163", "DTM+164", "PIA+5", "BGM"]
        assert [v for v, _ in _anchor_values(anchors, AnchorKind.DATA_ELEMENT)] == ["DE1001", "DE7140", "DE6060"]
        assert [v for v, _ in _anchor_values(anchors, AnchorKind.CODE)] == ["Z41"]
        assert [v for v, _ in _anchor_values(anchors, AnchorKind.CONDITION)] == ["492"]
        assert [v for v, _ in _anchor_values(anchors, AnchorKind.PACKAGE)] == ["1P"]
        assert [v for v, _ in _anchor_values(anchors, AnchorKind.EBD)] == ["E_0623"]

    def test_explicit_wins_over_mention(self) -> None:
        anchors = find_anchors("Tabellenspalte = Zählerstand (Strom) 13017\nSiehe auch 13017.", ["13017"])
        assert _anchor_values(anchors, AnchorKind.PRUEFIDENTIFIKATOR) == [("13017", AnchorSource.EXPLICIT)]

    def test_merge_keeps_strongest_source(self) -> None:
        merged = merge_anchors(
            [
                Anchor(kind=AnchorKind.SEGMENT, value="DTM", source=AnchorSource.MENTION),
                Anchor(kind=AnchorKind.PRUEFIDENTIFIKATOR, value="1", source=AnchorSource.MENTION),
                Anchor(kind=AnchorKind.PRUEFIDENTIFIKATOR, value="1", source=AnchorSource.SECTION_SCOPE),
                Anchor(kind=AnchorKind.PRUEFIDENTIFIKATOR, value="1", source=AnchorSource.TABLE_IN_SECTION),
            ]
        )
        assert merged == [
            Anchor(kind=AnchorKind.PRUEFIDENTIFIKATOR, value="1", source=AnchorSource.TABLE_IN_SECTION),
            Anchor(kind=AnchorKind.SEGMENT, value="DTM", source=AnchorSource.MENTION),
        ]


def test_table_to_markdown() -> None:
    table = docx.Document().add_table(rows=3, cols=3)
    for row_index, texts in enumerate([["Sparte", "Kommunikation von", ""], ["Strom", "MSB an NB", "a|b"]]):
        for col_index, text in enumerate(texts):
            table.cell(row_index, col_index).text = text
    table.cell(0, 1).merge(table.cell(0, 2))  # the header cell spans two columns
    table.cell(1, 1).paragraphs[0].add_run("\tmit  Tab")
    assert table_to_markdown(table) == (
        "| Sparte | Kommunikation von |  |\n|---|---|---|\n| Strom | MSB an NB mit Tab | a\\|b |"
    )


# ---- extraction from a synthetic document ----------------------------------------------------------------------
# The detection of AHB tables is covered by the tests of kohlrahbi.read_functions; here it is replaced by a simple
# convention: a table whose first cell is "EDIFACT Struktur" starts an AHB table (its Prüfis are in the second cell),
# a table whose last row starts with "UNT" ends it.


class _FakeSeed:
    def __init__(self, pruefidentifikatoren: list[str]) -> None:
        self.pruefidentifikatoren = pruefidentifikatoren

    @classmethod
    def from_table(cls, docx_table: Table) -> "_FakeSeed":
        return cls(docx_table.cell(0, 1).text.split())


@pytest.fixture
def fake_ahb_table_detection(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        kohlrahbi.context, "is_item_table_with_pruefidentifikatoren", lambda t: t.cell(0, 0).text == "EDIFACT Struktur"
    )
    monkeypatch.setattr(kohlrahbi.context, "is_last_row_end_of_message", lambda t: t.cell(-1, 0).text == "UNT")
    monkeypatch.setattr(kohlrahbi.context, "Seed", _FakeSeed)


def _add_table(document: Document, rows: list[list[str]]) -> None:
    table = document.add_table(rows=len(rows), cols=len(rows[0]))
    for row_index, row in enumerate(rows):
        for col_index, text in enumerate(row):
            table.cell(row_index, col_index).text = text


@pytest.fixture
def synthetic_ahb() -> Document:
    document = docx.Document()
    document.styles.add_style("Zwischenüberschrift", WD_STYLE_TYPE.PARAGRAPH)
    _add_table(document, [["Version:", "1.0"]])  # front matter
    document.add_paragraph("Die PDF-Datei ist das allein gültige Dokument.")  # front matter
    document.add_heading("Übersicht der Pakete in der MSCONS", level=1)
    _add_table(document, [["Paket", "Bedingungen"], ["[1P]", "[1] ∧ [2]"]])
    document.add_heading("Zählerstände", level=1)
    document.add_paragraph("In SG10 QTY DE6060 werden Zählerstände angegeben.")
    document.add_heading("Übertragung von Zählerständen Strom", level=2)
    document.add_paragraph("Tabellenspalte = Zählerstand (Strom) 13017")
    document.add_paragraph("Dieser Anwendungsfall dient zur Übertragung von:")
    document.add_paragraph("Turnuswerten,", style="List Bullet")
    document.add_paragraph("Zwischenablesungen.", style="List Bullet")
    _add_table(document, [["Sparte", "Kommunikation von"], ["Strom", "MSB an NB"]])
    document.add_paragraph("Werte aus dem SMGw.", style="Zwischenüberschrift")
    document.add_paragraph("Es ist die Konfigurations-ID anzugeben.")
    document.add_heading("Anwendungsübersicht Zählerstand Strom", level=2)
    document.add_paragraph("Die folgende Tabelle gilt nur für Strom.")
    _add_table(document, [["EDIFACT Struktur", "13017 13019"], ["UNH", "Muss"]])
    _add_table(document, [["SG10 QTY", "Muss"], ["UNT", "Muss"]])  # headless continuation
    document.add_paragraph("Hinweis nach der Tabelle.")
    _add_table(document, [["Gastag von", "Gastag bis"], ["24.10.", "25.10."]])  # side table after the AHB table
    document.add_heading("Änderungshistorie", level=1)
    document.add_paragraph("Änderung von 13017.")
    return document


@pytest.mark.usefixtures("fake_ahb_table_detection")
class TestExtraction:
    def test_passages(self, synthetic_ahb: Document) -> None:
        extraction = extract_passages(synthetic_ahb)
        passages = build_passages(extraction, "doc", known_pruefis=["13017", "13019"])
        assert [(p.heading_path, p.text) for p in passages] == [
            (["Zählerstände"], "In SG10 QTY DE6060 werden Zählerstände angegeben."),
            (
                ["Zählerstände", "Übertragung von Zählerständen Strom"],
                "Tabellenspalte = Zählerstand (Strom) 13017\n\n"
                "Dieser Anwendungsfall dient zur Übertragung von:\n\n"
                "- Turnuswerten,\n- Zwischenablesungen.\n\n"
                "| Sparte | Kommunikation von |\n|---|---|\n| Strom | MSB an NB |",
            ),
            (
                ["Zählerstände", "Übertragung von Zählerständen Strom", "Werte aus dem SMGw."],
                "Es ist die Konfigurations-ID anzugeben.",
            ),
            (["Zählerstände", "Anwendungsübersicht Zählerstand Strom"], "Die folgende Tabelle gilt nur für Strom."),
            (
                ["Zählerstände", "Anwendungsübersicht Zählerstand Strom"],
                "Hinweis nach der Tabelle.\n\n| Gastag von | Gastag bis |\n|---|---|\n| 24.10. | 25.10. |",
            ),
        ]
        assert [p.id for p in passages] == [f"doc:{i:04d}" for i in range(5)]
        assert [p.contains_table for p in passages] == [False, True, False, False, True]
        assert extraction.table_pruefis == ["13017", "13019"]

    def test_pruefi_anchors(self, synthetic_ahb: Document) -> None:
        passages = build_passages(extract_passages(synthetic_ahb), "doc", known_pruefis=[])
        pruefi_anchors = [_anchor_values(p.anchors, AnchorKind.PRUEFIDENTIFIKATOR) for p in passages]
        assert pruefi_anchors == [
            [],
            [("13017", AnchorSource.EXPLICIT)],
            [("13017", AnchorSource.SECTION_SCOPE)],
            [("13017", AnchorSource.TABLE_IN_SECTION), ("13019", AnchorSource.TABLE_IN_SECTION)],
            [("13017", AnchorSource.TABLE_IN_SECTION), ("13019", AnchorSource.TABLE_IN_SECTION)],
        ]

    def test_pruefi_contexts_and_dump(self, synthetic_ahb: Document, tmp_path: Path) -> None:
        context_document = get_context_document(
            synthetic_ahb, "AHB_MSCONS_test.docx", EdifactFormat.MSCONS, EdifactFormatVersion.FV2610
        )
        assert context_document.pruefidentifikatoren == ["13017", "13019"]
        contexts = {c.pruefidentifikator: c for c in get_pruefi_contexts(context_document)}
        assert [p.order for p in contexts["13017"].passages] == [1, 2, 3, 4]
        assert [p.order for p in contexts["13019"].passages] == [3, 4]

        written = dump_context_document(context_document, tmp_path)
        context_directory = tmp_path / "MSCONS" / "context"
        assert written == [
            context_directory / "AHB_MSCONS_test.json",
            context_directory / "pruefi" / "13017.json",
            context_directory / "pruefi" / "13019.json",
        ]
        assert ContextDocument.model_validate_json(written[0].read_text(encoding="utf-8")) == context_document
        pruefi_context = PruefiContext.model_validate_json(written[1].read_text(encoding="utf-8"))
        assert pruefi_context.layer == "context"
        assert pruefi_context == contexts["13017"]


def test_json_schemas(tmp_path: Path) -> None:
    written = dump_json_schemas(tmp_path)
    assert [path.name for path in written] == ["context_document.schema.json", "pruefi_context.schema.json"]
    schema = json.loads(written[0].read_text(encoding="utf-8"))
    assert schema["properties"]["layer"]["const"] == "context"
    assert "Structured facts" in schema["description"]


# ---- extraction from real (FV2310) AHB documents committed to the test mirror ---------------------------------


class TestRealDocuments:
    def test_ahb_tables_are_excluded(self) -> None:
        path = next(path_to_test_files_fv2310.glob("ORDERSORDRSPAHBMaBiS*.docx"))
        extraction = extract_passages(docx.Document(str(path)))
        passages = build_passages(extraction, path.stem, known_pruefis=[])
        assert extraction.table_pruefis == [*(f"172{i:02d}" for i in range(1, 9)), "19204", "17209", "17210", "17211"]
        assert not any("EDIFACT Struktur" in p.text or "Nachrichten-Endesegment" in p.text for p in passages)
        assert passages[0].heading_path == ["Vorwort"]  # the title page and the disclaimer are skipped

    def test_document_with_two_formats(self) -> None:
        path = next(path_to_test_files_fv2310.glob("ORDERSORDRSPAHBMaBiS*.docx"))
        document = docx.Document(str(path))
        ordrsp = get_context_document(document, path.name, EdifactFormat.ORDRSP, EdifactFormatVersion.FV2310)
        assert ordrsp.pruefidentifikatoren == ["19204"]
        (context,) = get_pruefi_contexts(ordrsp)
        assert [p.heading_path for p in context.passages] == [["Anwendungsübersichten", "Anwendungsfall ORDRSP"]]
        assert _anchor_values(context.passages[0].anchors, AnchorKind.PRUEFIDENTIFIKATOR) == [
            ("19204", AnchorSource.TABLE_IN_SECTION)
        ]

    def test_side_table_is_kept(self) -> None:
        path = next(path_to_test_files_fv2310.glob("PARTINAHB*.docx"))
        passages = build_passages(extract_passages(docx.Document(str(path))), path.stem, known_pruefis=[])
        side_table_passages = [p for p in passages if p.contains_table]
        assert len(side_table_passages) == 1
        assert "|---|" in side_table_passages[0].text


def test_cli_help() -> None:
    # a wide terminal, otherwise rich truncates the long option names in the help table
    result = CliRunner().invoke(app, ["context", "--help"], env={"COLUMNS": "200"})
    assert result.exit_code == 0
    assert "--edi-energy-mirror-path" in strip_ansi(result.output)
