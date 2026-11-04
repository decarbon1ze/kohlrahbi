"""
Deterministic extraction of anchors (pointers to structured keys) from the explanatory text of AHB documents.
"""

import re
from collections.abc import Iterable

from kohlrahbi.context.models import ANCHOR_SOURCE_PRIORITY, Anchor, AnchorKind, AnchorSource

#: UN/EDIFACT segment tags used in the EDI@Energy messages; restricting to known tags avoids false positives
#: on other three letter abbreviations in the text (e.g. MSB, EEG, NNE).
SEGMENT_TAGS: frozenset[str] = frozenset(
    """
    ADR AGR AJT ALC ALI APR ARD ATT AUT BGM CAV CCI CNT COD COM CTA CUX DGS DLM DOC DTM EFI EQD ERC FII FTX
    GEI GIN GIR HAN IDE IMD INP IRQ LIN LOC MEA MKS MOA NAD PAC PAI PAT PCD PCI PER PGI PIA PRC PRI PRV PYT
    QTY QVR RCS RFF RNG RTE SCC SEQ SGP STS TAX TDT TOD UCD UCI UCM UCS UNA UNB UNE UNG UNH UNS UNT UNZ
    """.split()
)
_segment_tags = "|".join(sorted(SEGMENT_TAGS))

# "Tabellenspalte = Zählerstand (Strom) 13017": the Prüfi is the last five-digit number of the line
_explicit_tabellenspalte = re.compile(r"^Tabellenspalte\s*=[^\n]*\b([1-9]\d{4})\b\s*$", re.MULTILINE)
# "Kommunikationspartner, ... für diesen Anwendungsfall Prüfidentifikator: 13025" (the line ends with the Prüfi)
_explicit_line_ending_with_pruefi = re.compile(r"Prüfidentifikator(?:en)?\s*:\s*([1-9]\d{4})\s*$", re.MULTILINE)
_five_digits = re.compile(r"(?<![\d.,])\b([1-9]\d{4})\b(?![\d.,]\d)")
_segment_group = re.compile(r"\bSG\s?(\d{1,2})\b")
_segment_with_qualifier = re.compile(rf"\b({_segment_tags})\s?\+\s?([A-Z0-9]{{1,3}})\b")
_segment = re.compile(rf"\b({_segment_tags})\b(?!\s?\+)")
_data_element = re.compile(r"\b(?:DE\s?|Datenelements?\s+)(\d{4})\b")
_code = re.compile(r"\b(?:Qualifier|Code)\s+[„\"]?([A-Z0-9]{1,3})\b")
_condition = re.compile(r"\[(\d{1,4})\]")
_package = re.compile(r"\[(\d{1,4})P\d*\.\.\d*\]")
_ebd = re.compile(r"\bE_\d{4}\b")


def find_explicit_pruefis(text: str) -> list[str]:
    """Returns the Prüfidentifikatoren the text explicitly states to describe."""
    found = [match.group(1) for match in _explicit_tabellenspalte.finditer(text)]
    found += [match.group(1) for match in _explicit_line_ending_with_pruefi.finditer(text)]
    return list(dict.fromkeys(found))


def find_anchors(text: str, known_pruefis: Iterable[str]) -> list[Anchor]:
    """
    Finds all anchors in the given text.
    Five-digit numbers only count as Prüfidentifikator if they are in ``known_pruefis``, because the text also
    contains other five-digit numbers (amounts, postal codes, ...).
    """
    known = set(known_pruefis)
    anchors = [
        Anchor(kind=AnchorKind.PRUEFIDENTIFIKATOR, value=p, source=AnchorSource.EXPLICIT)
        for p in find_explicit_pruefis(text)
    ]
    anchors += [
        Anchor(kind=AnchorKind.PRUEFIDENTIFIKATOR, value=m.group(1), source=AnchorSource.MENTION)
        for m in _five_digits.finditer(text)
        if m.group(1) in known
    ]
    mentions: list[tuple[AnchorKind, str]] = [
        *((AnchorKind.SEGMENT_GROUP, f"SG{int(m.group(1))}") for m in _segment_group.finditer(text)),
        *((AnchorKind.SEGMENT, f"{m.group(1)}+{m.group(2)}") for m in _segment_with_qualifier.finditer(text)),
        *((AnchorKind.SEGMENT, m.group(1)) for m in _segment.finditer(text)),
        *((AnchorKind.DATA_ELEMENT, f"DE{m.group(1)}") for m in _data_element.finditer(text)),
        *((AnchorKind.CODE, m.group(1)) for m in _code.finditer(text)),
        *((AnchorKind.CONDITION, m.group(1)) for m in _condition.finditer(text)),
        *((AnchorKind.PACKAGE, f"{m.group(1)}P") for m in _package.finditer(text)),
        *((AnchorKind.EBD, m.group(0)) for m in _ebd.finditer(text)),
    ]
    anchors += [Anchor(kind=kind, value=value, source=AnchorSource.MENTION) for kind, value in mentions]
    return merge_anchors(anchors)


def merge_anchors(anchors: Iterable[Anchor]) -> list[Anchor]:
    """
    Removes duplicates: for each (kind, value) only the anchor with the strongest source is kept.
    The result is sorted by kind (in the order of AnchorKind) and keeps the order of first occurrence within a kind.
    """
    best: dict[tuple[AnchorKind, str], Anchor] = {}
    for anchor in anchors:
        key = (anchor.kind, anchor.value)
        current = best.get(key)
        if current is None or ANCHOR_SOURCE_PRIORITY[anchor.source] < ANCHOR_SOURCE_PRIORITY[current.source]:
            best[key] = anchor
    kind_order = {kind: index for index, kind in enumerate(AnchorKind)}
    first_seen = {key: index for index, key in enumerate(best)}
    return sorted(best.values(), key=lambda a: (kind_order[a.kind], first_seen[(a.kind, a.value)]))
