"""
Models of the context export: the explanatory text around the AHB tables.

The context export is deliberately kept apart from the structured exports (flatahb, csv, xlsx, conditions, packages,
change history). A passage holds unstructured text only. It never repeats structured facts; it merely *points* to
them via anchors, i.e. the keys under which the structured exports can be looked up (Prüfidentifikator, segment
group, segment, data element, condition, package, EBD code).
"""

from enum import StrEnum
from typing import Literal

from efoli import EdifactFormat, EdifactFormatVersion
from pydantic import BaseModel, ConfigDict, Field

CONTEXT_LAYER_DESCRIPTION = (
    "Unstructured explanatory text from an AHB document (everything outside the AHB, package and change history "
    "tables). It is meant as additional context for humans and agents. Structured facts (AHB rows, conditions, "
    "packages) are NOT part of this layer; use the structured exports and follow the anchors to them."
)


class AnchorKind(StrEnum):
    """The kind of structured key an anchor points to."""

    PRUEFIDENTIFIKATOR = "pruefidentifikator"  #: e.g. "13017"
    SEGMENT_GROUP = "segment_group"  #: e.g. "SG10"
    SEGMENT = "segment"  #: e.g. "DTM" or, with qualifier, "DTM+163"
    DATA_ELEMENT = "data_element"  #: e.g. "DE1001"
    CODE = "code"  #: e.g. "Z41" (from "Qualifier Z41" / "Code Z41")
    CONDITION = "condition"  #: e.g. "492" (from "[492]")
    PACKAGE = "package"  #: e.g. "1P" (from "[1P0..1]")
    EBD = "ebd"  #: e.g. "E_0623"


class AnchorSource(StrEnum):
    """
    How an anchor was found. The sources are ordered from strongest to weakest evidence;
    if the same key is found in several ways, only the strongest source is kept.
    """

    #: the passage states that it describes this Prüfidentifikator,
    #: e.g. "Tabellenspalte = Zählerstand (Strom) 13017" or "... für diesen Anwendungsfall Prüfidentifikator: 13025"
    EXPLICIT = "explicit"
    #: the passage stands in the same section (same heading path) as the AHB table of this Prüfidentifikator
    TABLE_IN_SECTION = "table_in_section"
    #: a passage of an enclosing section states that it describes this Prüfidentifikator (see EXPLICIT)
    SECTION_SCOPE = "section_scope"
    #: the key is merely mentioned somewhere in the text
    MENTION = "mention"


ANCHOR_SOURCE_PRIORITY: dict[AnchorSource, int] = {source: index for index, source in enumerate(AnchorSource)}


class Anchor(BaseModel):
    """A pointer from a passage to a key of the structured exports."""

    model_config = ConfigDict(frozen=True)

    kind: AnchorKind
    value: str = Field(description="The key in its normalized form, e.g. '13017', 'SG10', 'DTM+163', 'DE1001'.")
    source: AnchorSource


class Passage(BaseModel):
    """
    A contiguous piece of explanatory text below one heading, e.g. all paragraphs, lists and small tables of a
    section up to the next heading or AHB table.
    """

    id: str = Field(description="Unique within the format version: '<source file stem>:<order>'.")
    order: int = Field(description="Position of the passage within its document (0-based).")
    heading_path: list[str] = Field(description="The headings the passage stands below, outermost first.")
    text: str = Field(description="The text as Markdown. Lists become '- ' items, tables become Markdown tables.")
    contains_table: bool = Field(description="True if the text contains at least one (non-AHB) table.")
    anchors: list[Anchor] = Field(default_factory=list)


class ContextDocument(BaseModel):
    """All passages of one AHB document."""

    model_config = ConfigDict(json_schema_extra={"description": CONTEXT_LAYER_DESCRIPTION})

    layer: Literal["context"] = "context"
    edifact_format: EdifactFormat
    format_version: EdifactFormatVersion
    source_file: str = Field(description="File name of the AHB docx the passages were taken from.")
    pruefidentifikatoren: list[str] = Field(
        description="The Prüfidentifikatoren of this format whose AHB tables are in the document."
    )
    passages: list[Passage]


class PruefiContext(BaseModel):
    """The passages of an AHB document that have an anchor to one Prüfidentifikator."""

    model_config = ConfigDict(json_schema_extra={"description": CONTEXT_LAYER_DESCRIPTION})

    layer: Literal["context"] = "context"
    pruefidentifikator: str
    edifact_format: EdifactFormat
    format_version: EdifactFormatVersion
    source_file: str = Field(description="File name of the AHB docx the passages were taken from.")
    passages: list[Passage]
