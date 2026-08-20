"""ICAI Division I – Non-Ind AS Schedule III reference analysis.

The Guidance Note PDF is the primary source. Line items, section headings,
and terminology are taken from that document rather than from general
accounting knowledge.
"""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

import pdfplumber
from pypdf import PdfReader

from app.config import settings
from app.mapping.field_mapping import build_field_mappings, slugify

HEADER_NOISE_RE = re.compile(
    r"Guidance Note on Division I[^\n]*\n(?:\d+\n)?",
    re.IGNORECASE,
)
FOOTNOTE_RE = re.compile(
    r"(?:Amended|Inserted|Omitted) pursuant to MCA Notification[^\n]*",
    re.IGNORECASE,
)
SKIP_TITLES = {"particulars", "note", "note no", "1", "2", "3", "4", "xxx", "total"}
SHAREHOLDERS_RE = re.compile(r"(Shareholders[’'] funds)", re.IGNORECASE)

MARKER_RE = re.compile(
    r"(?:^|\s)(?:"
    r"(?P<roman>I{1,3}|IV|VI{0,3}|IX|XI{0,3}|XIV|XV|XVI)\.\s+"
    r"|(?P<group>\(\d+\))\s+"
    r"|(?P<sub>\([ivx]+\))\s+"
    r"|(?P<letter>\([a-z]\))\s+"
    r"|(?P<cap>\([A-Z]\))\s+"
    r")"
)

P_AND_L_ROMAN_RE = re.compile(
    r"^(I{1,3}|IV|VI{0,3}|IX|X|XI{0,3}|XIV|XV|XVI)\.?\s+(.*)$",
    re.IGNORECASE,
)

NOTE_HEADING_RE = re.compile(
    r"^([A-Z]{1,2})(?:\.\d+)?\.?\s*\d*\s+(.+)$"
)

RATIO_HEADING_RE = re.compile(
    r"^(\d+)\.\s+(.+?)\s*$"
)

# Headings taken from the Guidance Note index / body. Used only as search
# keys against extracted PDF text.
SECTION_QUERIES: tuple[tuple[str, str, tuple[str, ...]], ...] = (
    (
        "general_instructions",
        "General instructions",
        (
            "General Instructions to Division I to The Schedule III",
            "GENERAL INSTRUCTIONS FOR PREPARATION OF BALANCE SHEET AND STATEMENT OF PROFIT AND LOSS",
        ),
    ),
    (
        "balance_sheet",
        "Balance Sheet",
        (
            "8. Part I: Form of Balance Sheet",
            "Part I: Form of Balance Sheet and Note 6",
            "PART I – Form of BALANCE SHEET",
        ),
    ),
    (
        "profit_and_loss",
        "Statement of Profit and Loss",
        (
            "9. Part II – Statement of Profit and Loss",
            "Part II – Statement of Profit and Loss",
            "PART II – Form of STATEMENT OF PROFIT AND LOSS",
        ),
    ),
    (
        "notes_to_accounts",
        "Notes to Accounts",
        (
            "Notes to accounts shall contain information",
            "GENERAL INSTRUCTIONS FOR PREPARATION OF BALANCE SHEET",
        ),
    ),
    (
        "current_non_current",
        "Current/non-current classification",
        (
            "An asset shall be classified as current when it satisfies any of the",
            "A liability shall be classified as current when it satisfies any of the",
        ),
    ),
    (
        "shareholders_funds",
        "Shareholders' funds",
        ("Shareholders’ Funds", "Shareholders' Funds", "Shareholders’ funds"),
    ),
    (
        "non_current_liabilities",
        "Non-current liabilities",
        ("Non-current liabilities",),
    ),
    (
        "current_liabilities",
        "Current liabilities",
        ("Current liabilities",),
    ),
    (
        "non_current_assets",
        "Non-current assets",
        ("Non-current assets", "Non Current Assets"),
    ),
    (
        "current_assets",
        "Current assets",
        ("Current assets",),
    ),
    (
        "revenue_from_operations",
        "Revenue from operations",
        ("Revenue from operations",),
    ),
    (
        "other_income",
        "Other income",
        ("Other income", "Other income shall be classified as"),
    ),
    (
        "expenses",
        "Expenses",
        (
            "The aggregate of the following expenses are to be disclosed on the face",
            "IV. Expenses",
        ),
    ),
    (
        "tax_expense",
        "Tax expense",
        ("Tax expense:", "X Tax expense"),
    ),
    (
        "earnings_per_share",
        "Earnings per share",
        ("Earnings per equity share",),
    ),
    (
        "other_disclosures",
        "Other disclosures",
        (
            "11. Other Disclosures",
            "Y. Additional Regulatory Information",
        ),
    ),
    (
        "ratios",
        "Ratios",
        (
            "Following Ratios to be disclosed",
            "Annexure B",
            "Analytical Ratios",
        ),
    ),
)


@dataclass(slots=True)
class HierarchyNode:
    """One caption in the Schedule III hierarchy."""

    code: str
    label: str
    kind: str
    level: int
    source_pages: list[int] = field(default_factory=list)
    excerpt: str = ""
    children: list["HierarchyNode"] = field(default_factory=list)
    note_classifications: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "label": self.label,
            "kind": self.kind,
            "level": self.level,
            "source_pages": self.source_pages,
            "excerpt": self.excerpt,
            "note_classifications": self.note_classifications,
            "children": [child.to_dict() for child in self.children],
        }


@dataclass(slots=True)
class IdentifiedSection:
    """A Guidance Note / Schedule III section located in the PDF."""

    section_id: str
    title: str
    start_page: int
    excerpt: str
    matched_heading: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "section_id": self.section_id,
            "title": self.title,
            "start_page": self.start_page,
            "matched_heading": self.matched_heading,
            "excerpt": self.excerpt,
        }


@dataclass(slots=True)
class ScheduleIIIReference:
    """Structured representation of Division I Schedule III from the ICAI PDF."""

    source_path: str
    source_title: str
    page_count: int
    sections: list[IdentifiedSection]
    balance_sheet: HierarchyNode
    profit_and_loss: HierarchyNode
    notes_to_accounts: HierarchyNode
    ratios: list[HierarchyNode]
    field_mappings: dict[str, dict[str, Any]]

    def to_dict(self) -> dict[str, Any]:
        return {
            "source": {
                "path": self.source_path,
                "title": self.source_title,
                "page_count": self.page_count,
            },
            "sections": [section.to_dict() for section in self.sections],
            "hierarchy": {
                "balance_sheet": self.balance_sheet.to_dict(),
                "statement_of_profit_and_loss": self.profit_and_loss.to_dict(),
                "notes_to_accounts": self.notes_to_accounts.to_dict(),
                "ratios": [node.to_dict() for node in self.ratios],
            },
            "field_mappings": self.field_mappings,
        }


@dataclass(slots=True)
class _RawItem:
    kind: str
    marker: str
    title: str
    page: int
    excerpt: str


class ICAIReferenceReader:
    """Read page-level text from the ICAI Guidance Note PDF."""

    def __init__(self, pdf_path: Path) -> None:
        self.pdf_path = Path(pdf_path)
        if not self.pdf_path.exists():
            raise FileNotFoundError(f"ICAI reference PDF not found: {self.pdf_path}")
        self._pages: list[str] | None = None
        self._title: str | None = None
        self._index_pages: set[int] | None = None

    @property
    def pages(self) -> list[str]:
        if self._pages is None:
            reader = PdfReader(str(self.pdf_path))
            extracted: list[str] = []
            for page in reader.pages:
                text = page.extract_text() or ""
                extracted.append(self.normalize_text(text))
            if len(extracted) < 20:
                raise ValueError(
                    f"{self.pdf_path} does not look like the ICAI Guidance Note "
                    f"(found {len(extracted)} page(s))."
                )
            joined = "\n".join(extracted[:5]).lower()
            if "schedule iii" not in joined and "division i" not in joined:
                raise ValueError(
                    f"{self.pdf_path} does not contain the ICAI Division I Schedule III Guidance Note."
                )
            self._pages = extracted
        return self._pages

    @property
    def page_count(self) -> int:
        return len(self.pages)

    @property
    def title(self) -> str:
        if self._title is None:
            head = self.pages[0] if self.pages else ""
            if "GUIDANCE NOTE ON" in head.upper():
                self._title = (
                    "Guidance Note on Division I – Non Ind AS Schedule III "
                    "to the Companies Act, 2013 (Revised January, 2022 Edition)"
                )
            else:
                self._title = "ICAI Guidance Note on Division I – Non-Ind AS Schedule III"
        return self._title

    @staticmethod
    def normalize_text(text: str) -> str:
        cleaned = text.replace("\xa0", " ")
        cleaned = cleaned.replace("’", "'").replace("‘", "'")
        cleaned = HEADER_NOISE_RE.sub("\n", cleaned)
        cleaned = FOOTNOTE_RE.sub("", cleaned)
        cleaned = re.sub(r"[ \t]+", " ", cleaned)
        cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
        return cleaned.strip()

    def page_text(self, page_number: int) -> str:
        return self.pages[page_number - 1]

    @property
    def index_pages(self) -> set[int]:
        """Pages that only list contents, not the Schedule III forms themselves."""
        if self._index_pages is None:
            ignored: set[int] = set()
            for page_number, text in enumerate(self.pages, start=1):
                head = text[:500]
                if "Index" in head and "Introduction" in head and "...." in text:
                    ignored.add(page_number)
            self._index_pages = ignored
        return self._index_pages

    def find_first(
        self,
        needles: Iterable[str],
        *,
        start_page: int = 1,
        ignore_pages: Iterable[int] | None = None,
    ) -> tuple[int, str] | None:
        skipped = set(ignore_pages) if ignore_pages is not None else set()
        compact_needles = [(original, self._compact(original)) for original in needles]
        for page_number, text in enumerate(self.pages, start=1):
            if page_number < start_page or page_number in skipped:
                continue
            loose_haystack = self._loose(text)
            compact_haystack = self._compact(text)
            for original, compact_needle in compact_needles:
                if not compact_needle:
                    continue
                matched = (
                    compact_needle in compact_haystack
                    if len(compact_needle) >= 20
                    else self._loose(original) in loose_haystack
                )
                if matched:
                    excerpt = self._excerpt_for(text, original)
                    return page_number, excerpt
        return None

    def slice_text(self, start_page: int, end_page: int) -> str:
        return "\n".join(self.pages[start_page - 1 : end_page])

    @staticmethod
    def _loose(value: str) -> str:
        return re.sub(r"\s+", " ", value).strip().lower()

    @staticmethod
    def _compact(value: str) -> str:
        return re.sub(r"[^a-z0-9]+", "", value.lower())

    def _excerpt_for(self, text: str, needle: str, radius: int = 180) -> str:
        loose_text = re.sub(r"\s+", " ", text)
        loose_needle = re.sub(r"\s+", " ", needle)
        idx = loose_text.lower().find(loose_needle.lower())
        if idx < 0:
            return loose_text[: radius * 2].strip()
        start = max(0, idx - radius)
        end = min(len(loose_text), idx + len(loose_needle) + radius)
        excerpt = loose_text[start:end].strip()
        return excerpt


class ScheduleIIIMapper:
    """Load the ICAI PDF and expose a JSON-serializable Schedule III model."""

    def __init__(self, pdf_path: Path | None = None) -> None:
        self.pdf_path = Path(pdf_path) if pdf_path else settings.reference_path
        self.reader = ICAIReferenceReader(self.pdf_path)
        self._reference: ScheduleIIIReference | None = None

    def load_reference(self) -> ScheduleIIIReference:
        if self._reference is None:
            self._reference = self._build_reference()
        return self._reference

    def get_hierarchy(self) -> dict[str, Any]:
        return self.load_reference().to_dict()["hierarchy"]

    def get_sections(self) -> list[dict[str, Any]]:
        return [section.to_dict() for section in self.load_reference().sections]

    def get_field_mappings(self) -> dict[str, dict[str, Any]]:
        return self.load_reference().field_mappings

    def to_json(self, *, indent: int = 2) -> str:
        return json.dumps(self.load_reference().to_dict(), indent=indent, ensure_ascii=False)

    def format_hierarchy(self) -> str:
        reference = self.load_reference()
        blocks = [
            _format_tree(reference.balance_sheet),
            "",
            _format_tree(reference.profit_and_loss),
            "",
            _format_tree(reference.notes_to_accounts),
        ]
        if reference.ratios:
            ratio_root = HierarchyNode(
                code="ratios",
                label="Analytical Ratios",
                kind="root",
                level=0,
                source_pages=sorted({page for node in reference.ratios for page in node.source_pages}),
                children=reference.ratios,
            )
            blocks.extend(["", _format_tree(ratio_root)])
        return "\n".join(blocks)

    def map(
        self,
        classified_data: dict | None = None,
        *,
        template_metadata: dict | None = None,
        inspect_template: bool = True,
    ) -> dict[str, Any]:
        """Map classified source data onto Schedule III categories and Excel cells."""
        payload = self.load_reference().to_dict()
        if classified_data is None:
            return payload
        mapped = self.map_classified(
            classified_data,
            template_metadata=template_metadata,
            inspect_template=inspect_template,
        )
        payload.update(mapped)
        return payload

    def map_classified(
        self,
        classified_data: dict | Any,
        *,
        template_metadata: dict | None = None,
        inspect_template: bool = True,
        field_map: Any | None = None,
    ) -> dict[str, Any]:
        """Build Excel placements and a mapping report from classified values."""
        from app.mapping.field_mapping import ExcelMappingEngine, load_excel_field_map

        metadata = template_metadata
        if metadata is None and inspect_template:
            metadata = self._inspect_template()
        engine = ExcelMappingEngine(
            field_map=field_map or load_excel_field_map(),
            icai_mappings=self.get_field_mappings(),
            template_metadata=metadata or {},
        )
        return engine.map(classified_data)

    def _inspect_template(self) -> dict[str, Any]:
        try:
            from app.excel.template_manager import TemplateManager

            return TemplateManager().inspect().to_dict()
        except Exception:
            return {}

    def _build_reference(self) -> ScheduleIIIReference:
        sections = self._identify_sections()
        balance_sheet = self._parse_balance_sheet()
        profit_and_loss = self._parse_profit_and_loss()
        notes = self._parse_notes_to_accounts()
        _attach_note_classifications(balance_sheet, notes)
        _attach_note_classifications(profit_and_loss, notes)
        ratios = self._parse_ratios()
        reference = ScheduleIIIReference(
            source_path=str(self.pdf_path),
            source_title=self.reader.title,
            page_count=self.reader.page_count,
            sections=sections,
            balance_sheet=balance_sheet,
            profit_and_loss=profit_and_loss,
            notes_to_accounts=notes,
            ratios=ratios,
            field_mappings={},
        )
        reference.field_mappings = build_field_mappings(reference)
        return reference

    def _identify_sections(self) -> list[IdentifiedSection]:
        identified: list[IdentifiedSection] = []
        for section_id, title, needles in SECTION_QUERIES:
            found = self.reader.find_first(needles, ignore_pages=self.reader.index_pages)
            if found is None:
                continue
            page_number, excerpt = found
            identified.append(
                IdentifiedSection(
                    section_id=section_id,
                    title=title,
                    start_page=page_number,
                    excerpt=excerpt,
                    matched_heading=needles[0],
                )
            )
        return identified

    def _parse_balance_sheet(self) -> HierarchyNode:
        start = self._require_page(
            ("PART I – Form of BALANCE SHEET", "PART I - Form of BALANCE SHEET"),
            after="SCHEDULE III (See section 129)",
        )
        chunks = self._table_chunks(range(start - 1, min(start + 2, self.reader.page_count)))
        items = _explode_chunks(chunks)
        items = _merge_continuations(items)
        items = [item for item in items if item.title.lower() not in SKIP_TITLES]
        items = _split_shareholders_heading(items)

        root = HierarchyNode(
            code="balance_sheet",
            label="Balance Sheet",
            kind="statement",
            level=0,
            source_pages=[start],
            excerpt="PART I – Form of BALANCE SHEET",
        )
        _build_tree(root, items, code_prefix="balance_sheet")
        if not root.children:
            raise ValueError("Could not parse the Balance Sheet form from Annexure A.")
        return root

    def _parse_profit_and_loss(self) -> HierarchyNode:
        start = self._require_page(
            (
                "PART II – Form of STATEMENT OF PROFIT AND LOSS",
                "PART II - Form of STATEMENT OF PROFIT AND LOSS",
            ),
            after="SCHEDULE III (See section 129)",
        )
        chunks = self._table_chunks(range(start - 1, min(start + 2, self.reader.page_count)))
        items = _explode_chunks(chunks, allow_plain_under_expenses=True)
        items = _merge_continuations(items)
        items = [item for item in items if item.title.lower() not in SKIP_TITLES]
        items = _reconcile_face_expenses(items, self.reader)

        root = HierarchyNode(
            code="profit_and_loss",
            label="Statement of Profit and Loss",
            kind="statement",
            level=0,
            source_pages=[start],
            excerpt="PART II – Form of STATEMENT OF PROFIT AND LOSS",
        )
        _build_tree(root, items, code_prefix="profit_and_loss")
        if not root.children:
            raise ValueError("Could not parse the Statement of Profit and Loss form from Annexure A.")
        return root

    def _parse_notes_to_accounts(self) -> HierarchyNode:
        start = self._require_page(
            ("A. Share Capital",),
            after="SCHEDULE III (See section 129)",
        )
        end = self._require_page(
            (
                "PART II – Form of STATEMENT OF PROFIT AND LOSS",
                "PART II - Form of STATEMENT OF PROFIT AND LOSS",
            ),
            after="SCHEDULE III (See section 129)",
        )
        text = self.reader.slice_text(start, end - 1)
        notes_root = HierarchyNode(
            code="notes_to_accounts",
            label="Notes to Accounts",
            kind="notes",
            level=0,
            source_pages=[start],
            excerpt="GENERAL INSTRUCTIONS FOR PREPARATION OF BALANCE SHEET",
        )

        current: HierarchyNode | None = None
        skip_letters_before_notes = {"I", "II", "III", "IV"}
        raw_lines = text.splitlines()
        index = 0
        while index < len(raw_lines):
            line = clean_title(raw_lines[index])
            index += 1
            if not line:
                continue
            match = NOTE_HEADING_RE.match(line)
            if not match:
                continue
            letter, heading = match.groups()
            heading = clean_title(re.sub(r"\s+\d{1,2}$", "", heading))
            dangling = {
                "the", "and", "of", "for", "not", "from", "at", "to", "in", "or",
                "as", "a", "an", "with", "by", "other", "than", "held", "per",
            }
            while index < len(raw_lines) and len(heading) < 100:
                last = heading.rstrip(".;:,").split()[-1].lower() if heading.split() else ""
                if last not in dangling and not heading.endswith("-"):
                    break
                nxt = clean_title(raw_lines[index])
                if not nxt or NOTE_HEADING_RE.match(nxt):
                    break
                heading = clean_title(f"{heading} {nxt}")
                index += 1
            if "(" in heading:
                heading = heading.split(" (i)", 1)[0]
                heading = heading.split(" (a)", 1)[0]
                heading = heading.split(" (i)", 1)[0]
            lowered = heading.lower()
            if letter in skip_letters_before_notes and lowered.startswith(("equity", "assets", "revenue", "other income", "total income", "expenses")):
                continue
            if lowered.startswith("omitted") or lowered.startswith("no promoter"):
                continue
            if len(heading) < 8:
                continue
            if lowered.startswith("raw materials") or lowered.startswith("export of goods"):
                continue
            node = HierarchyNode(
                code=f"notes_to_accounts.{slugify(letter)}_{slugify(heading)[:60]}",
                label=f"{letter}. {heading}",
                kind="note",
                level=1,
                source_pages=[start],
                excerpt=heading[:240],
            )
            notes_root.children.append(node)
            current = node

        additional = self._parse_additional_regulatory_information()
        if additional is not None:
            if current and "additional regulatory information" in current.label.lower():
                current.children.extend(additional.children)
                current.source_pages = sorted(set(current.source_pages + additional.source_pages))
            else:
                notes_root.children.append(additional)
        if not notes_root.children:
            raise ValueError("Could not parse Notes to Accounts headings from Annexure A.")
        return notes_root

    def _parse_additional_regulatory_information(self) -> HierarchyNode | None:
        found = self.reader.find_first(
            ("Additional Regulatory Information",),
            start_page=150,
            ignore_pages=self.reader.index_pages,
        )
        if found is None:
            return None
        start_page, excerpt = found
        text = self.reader.slice_text(start_page, min(start_page + 8, self.reader.page_count))
        collapsed = re.sub(r"\s+", " ", text)
        marker = re.search(r"Additional Regulatory Information", collapsed, flags=re.IGNORECASE)
        if marker:
            collapsed = collapsed[marker.end() :]
        root = HierarchyNode(
            code="notes_to_accounts.additional_regulatory_information",
            label="Y. Additional Regulatory Information",
            kind="note",
            level=1,
            source_pages=[start_page],
            excerpt=excerpt[:240],
        )
        item_re = re.compile(r"\(([ivx]+)\)\s+")
        expected = [
            "i",
            "ii",
            "iii",
            "iv",
            "v",
            "vi",
            "vii",
            "viii",
            "ix",
            "x",
            "xi",
            "xii",
            "xiii",
            "xiv",
        ]
        seen: set[str] = set()
        matches = list(item_re.finditer(collapsed))
        for index, match in enumerate(matches):
            roman = match.group(1).lower()
            if roman not in expected or roman in seen:
                continue
            start = match.end()
            end = matches[index + 1].start() if index + 1 < len(matches) else min(len(collapsed), start + 180)
            title = clean_title(collapsed[start:end])
            title = re.split(r"(?<=[a-z)*])\.\s", title, maxsplit=1)[0]
            if ":" in title:
                head = title.split(":", 1)[0].strip()
                if len(head) >= 12:
                    title = head
            title = title.split(" (a)", 1)[0]
            title = title.split(" (A)", 1)[0]
            title = title.split(" The company", 1)[0]
            title = title.split(" Where a company", 1)[0]
            title = title.split(" Where the company", 1)[0]
            title = title.split(" Where any", 1)[0]
            title = title.rstrip(" ,;:-")
            if len(title) < 8:
                continue
            if len(title) > 90:
                title = title[:90].rsplit(" ", 1)[0]
            seen.add(roman)
            root.children.append(
                HierarchyNode(
                    code=f"{root.code}.{slugify(title)[:70]}",
                    label=title,
                    kind="disclosure",
                    level=2,
                    source_pages=[start_page],
                    excerpt=title,
                )
            )
        return root

    def _parse_ratios(self) -> list[HierarchyNode]:
        required = self._parse_required_ratio_list()
        explained = self._parse_annexure_b_ratios()
        by_slug = {slugify(node.label): node for node in explained}
        merged: list[HierarchyNode] = []
        seen: set[str] = set()
        for node in required + explained:
            key = slugify(node.label)
            if key in seen:
                continue
            seen.add(key)
            if key in by_slug:
                extra = by_slug[key]
                node.excerpt = extra.excerpt or node.excerpt
                node.source_pages = sorted(set(node.source_pages + extra.source_pages))
            merged.append(node)
        return merged

    def _parse_required_ratio_list(self) -> list[HierarchyNode]:
        found = self.reader.find_first(
            ("Following Ratios to be disclosed",),
            start_page=150,
            ignore_pages=self.reader.index_pages,
        )
        if found is None:
            return []
        page_number, _ = found
        text = self.reader.slice_text(page_number, min(page_number + 2, self.reader.page_count))
        block = text.split("Following Ratios to be disclosed", 1)[-1]
        block = block.split("The company shall explain", 1)[0]
        nodes: list[HierarchyNode] = []
        for match in re.finditer(r"\(([a-k])\)\s*([^,\n]+)", block, flags=re.IGNORECASE):
            title = clean_title(match.group(2)).rstrip(",;")
            if not title:
                continue
            nodes.append(
                HierarchyNode(
                    code=f"ratios.{slugify(title)}",
                    label=title,
                    kind="ratio",
                    level=1,
                    source_pages=[page_number],
                    excerpt=title,
                )
            )
        return nodes

    def _parse_annexure_b_ratios(self) -> list[HierarchyNode]:
        found = self.reader.find_first(
            ("Annexure B", "Analytical Ratios"),
            ignore_pages=self.reader.index_pages,
        )
        if found is None:
            return []
        start_page, _ = found
        end_page = start_page
        for page_number in range(start_page, min(start_page + 8, self.reader.page_count + 1)):
            if page_number > start_page and "Annexure C" in self.reader.page_text(page_number):
                break
            end_page = page_number
        text = self.reader.slice_text(start_page, end_page)
        nodes: list[HierarchyNode] = []
        for match in RATIO_HEADING_RE.finditer(text):
            title = clean_title(match.group(2)).rstrip(":")
            if title.lower().startswith("section ") or len(title) < 4:
                continue
            if title.lower().startswith("where a company"):
                continue
            nodes.append(
                HierarchyNode(
                    code=f"ratios.{slugify(title)}",
                    label=title,
                    kind="ratio",
                    level=1,
                    source_pages=[start_page],
                    excerpt=title,
                )
            )
        return nodes

    def _require_page(self, needles: tuple[str, ...], *, after: str | None = None) -> int:
        start_page = 1
        ignore = self.reader.index_pages
        if after:
            found_after = self.reader.find_first((after,), ignore_pages=ignore)
            if found_after:
                start_page = found_after[0]
        found = self.reader.find_first(needles, start_page=start_page, ignore_pages=ignore)
        if found is None:
            raise ValueError(f"Could not locate {needles[0]!r} in {self.pdf_path}")
        return found[0]

    def _table_chunks(self, page_indexes: range) -> list[tuple[int, str]]:
        chunks: list[tuple[int, str]] = []
        with pdfplumber.open(str(self.pdf_path)) as pdf:
            for index in page_indexes:
                page_number = index + 1
                for table in pdf.pages[index].extract_tables() or []:
                    for row in table:
                        chunk = _row_particular(row)
                        if chunk:
                            chunks.append((page_number, chunk))
        return chunks


PDF_WORD_REPAIRS = (
    ("p rogress", "progress"),
    ("e xtent", "extent"),
    ("fo r ", "for "),
    ("o r ", "or "),
    ("Plan t ", "Plant "),
    ("fina ncial", "financial"),
    ("Arran gements", "Arrangements"),
    ("prem ium", "premium"),
    ("R egistrar", "Registrar"),
    ("rela ted", "related"),
    ("classifie d", "classified"),
    ("SHE ET", "SHEET"),
)


def clean_title(text: str) -> str:
    value = text.replace("\n", " ")
    value = value.replace("–", "-").replace("—", "-")
    value = value.replace("’", "'").replace("‘", "'")
    value = value.replace("[", " ").replace("]", " ")
    for broken, repaired in PDF_WORD_REPAIRS:
        value = value.replace(broken, repaired)
    value = re.sub(r"(?<=[A-Za-z])\d{1,2}\b", "", value)
    value = re.sub(r"-\s+", "-", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip(" :-")


def _row_particular(row: list[str | None]) -> str | None:
    cells = [(cell or "").strip() for cell in row]
    if not cells:
        return None
    left = cells[0]
    mid = cells[1] if len(cells) > 1 else ""
    combined = " ".join(part for part in (left, mid) if part)
    combined = clean_title(combined)
    if not combined or combined.lower() in SKIP_TITLES or combined == "1 2 3 4":
        return None
    if combined.lower() in {"particulars", "note no"}:
        return None
    return combined


def _explode_chunks(
    chunks: list[tuple[int, str]],
    *,
    allow_plain_under_expenses: bool = False,
) -> list[_RawItem]:
    items: list[_RawItem] = []
    for page_number, chunk in chunks:
        items.extend(_explode_text(chunk, page_number, allow_plain_under_expenses=allow_plain_under_expenses))
    return items


def _explode_text(text: str, page_number: int, *, allow_plain_under_expenses: bool) -> list[_RawItem]:
    value = clean_title(text)
    if not value:
        return []

    matches = list(MARKER_RE.finditer(f" {value}" if not value.startswith(" ") else value))
    pnl_match = P_AND_L_ROMAN_RE.match(value)
    if not matches and pnl_match:
        roman, title = pnl_match.groups()
        return [
            _RawItem(
                kind="roman",
                marker=f"{roman.upper()}.",
                title=clean_title(title),
                page=page_number,
                excerpt=value,
            )
        ]

    if not matches:
        if allow_plain_under_expenses:
            return [
                _RawItem(
                    kind="plain",
                    marker="",
                    title=value,
                    page=page_number,
                    excerpt=value,
                )
            ]
        return [
            _RawItem(
                kind="plain",
                marker="",
                title=value,
                page=page_number,
                excerpt=value,
            )
        ]

    items: list[_RawItem] = []
    search_text = f" {value}"
    matches = list(MARKER_RE.finditer(search_text))
    for index, match in enumerate(matches):
        kind = next(name for name, captured in match.groupdict().items() if captured)
        marker = (match.group(kind) or "").strip()
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(search_text)
        title = clean_title(search_text[start:end])
        if not title:
            continue
        items.append(
            _RawItem(
                kind={"roman": "roman", "group": "group", "sub": "sub", "letter": "letter", "cap": "cap"}[kind],
                marker=marker if marker.endswith(".") or marker.startswith("(") else f"{marker}.",
                title=title,
                page=page_number,
                excerpt=value,
            )
        )
    return items


def _split_shareholders_heading(items: list[_RawItem]) -> list[_RawItem]:
    expanded: list[_RawItem] = []
    for item in items:
        match = SHAREHOLDERS_RE.search(item.title)
        if item.kind == "roman" and match and match.start() > 0:
            head = clean_title(item.title[: match.start()])
            tail = match.group(1)
            expanded.append(
                _RawItem(
                    kind=item.kind,
                    marker=item.marker,
                    title=head,
                    page=item.page,
                    excerpt=item.excerpt,
                )
            )
            expanded.append(
                _RawItem(
                    kind="heading",
                    marker="",
                    title=tail,
                    page=item.page,
                    excerpt=item.excerpt,
                )
            )
        else:
            expanded.append(item)
    return expanded


def _merge_continuations(items: list[_RawItem]) -> list[_RawItem]:
    if not items:
        return []
    merged: list[_RawItem] = [items[0]]
    for item in items[1:]:
        previous = merged[-1]
        continuation = (
            item.kind == "plain"
            and not item.marker
            and not P_AND_L_ROMAN_RE.match(item.title)
            and not item.title.lower().startswith(("cost of", "purchases of", "employee", "depreciation", "other expenses", "total expenses", "finance"))
        )
        if continuation and previous.kind in {"letter", "sub", "plain", "roman"}:
            previous.title = clean_title(f"{previous.title} {item.title}")
            previous.excerpt = f"{previous.excerpt} {item.title}".strip()
            continue
        merged.append(item)
    for item in merged:
        if "finished goods" in item.title.lower() and "work-in-progress" in item.title.lower():
            item.title = "Changes in inventories of finished goods, work-in-progress and Stock-in-Trade"
    return merged


def _face_expense_labels(reader: ICAIReferenceReader) -> list[tuple[int, str]]:
    """Read the face-of-P&L expense captions from Guidance Note paragraph 9.5."""
    found = reader.find_first(
        ("The aggregate of the following expenses are to be disclosed on the face",),
        ignore_pages=reader.index_pages,
    )
    if found is None:
        return []
    page_number, _ = found
    text = reader.page_text(page_number)
    start = text.find("➢")
    if start < 0:
        start = text.find("•")
    if start < 0:
        return []
    block = text[start:]
    block = re.split(r"9\.5\.1", block, maxsplit=1)[0]
    labels: list[tuple[int, str]] = []
    for raw in re.split(r"[➢•]", block):
        label = clean_title(raw)
        if not label or len(label) < 8 or len(label) > 140:
            continue
        labels.append((page_number, label))
    return labels


def _reconcile_face_expenses(items: list[_RawItem], reader: ICAIReferenceReader) -> list[_RawItem]:
    """Keep Annexure A items and insert any face captions listed in paragraph 9.5."""
    reference_expenses = _face_expense_labels(reader)
    if not reference_expenses:
        return items

    start = next(
        (
            index
            for index, item in enumerate(items)
            if item.kind == "roman" and item.title.lower().startswith("expense")
        ),
        None,
    )
    if start is None:
        return items
    end = start + 1
    while end < len(items) and items[end].kind != "roman":
        end += 1
    current = items[start + 1 : end]

    def match_existing(label: str) -> _RawItem | None:
        key = "_".join(slugify(label).split("_")[:6])
        for item in current:
            item_key = "_".join(slugify(item.title).split("_")[:6])
            if key == item_key:
                return item
        return None

    rebuilt: list[_RawItem] = []
    used: set[int] = set()
    for page_number, label in reference_expenses:
        existing = match_existing(label)
        if existing is not None:
            rebuilt.append(existing)
            used.add(id(existing))
        else:
            rebuilt.append(
                _RawItem(kind="plain", marker="", title=label, page=page_number, excerpt=label)
            )
    for item in current:
        if id(item) not in used:
            rebuilt.append(item)
    return items[: start + 1] + rebuilt + items[end:]


def _kind_level(kind: str) -> int:
    return {
        "roman": 1,
        "heading": 2,
        "group": 2,
        "letter": 3,
        "plain": 3,
        "sub": 4,
        "cap": 4,
    }.get(kind, 3)


def _build_tree(root: HierarchyNode, items: list[_RawItem], *, code_prefix: str) -> None:
    stack: list[HierarchyNode] = [root]
    for item in items:
        if item.title.lower() in SKIP_TITLES:
            continue
        level = _kind_level(item.kind)
        node = HierarchyNode(
            code="",
            label=item.title,
            kind=item.kind,
            level=level,
            source_pages=[item.page],
            excerpt=item.excerpt[:240],
        )
        while len(stack) > 1 and stack[-1].level >= level:
            stack.pop()
        parent = stack[-1]
        node.code = f"{parent.code}.{slugify(node.label)}" if parent.code else slugify(node.label)
        if not node.code.startswith(code_prefix):
            node.code = f"{code_prefix}.{node.code}" if node.code else code_prefix
        parent.children.append(node)
        stack.append(node)


def _attach_note_classifications(statement: HierarchyNode, notes: HierarchyNode) -> None:
    note_by_slug: dict[str, str] = {}
    for note in notes.children:
        heading = re.sub(r"^[A-Z]{1,2}\.\s*", "", note.label)
        note_by_slug[slugify(heading)] = note.label
        note_by_slug[slugify(note.label)] = note.label

    def walk(node: HierarchyNode) -> None:
        key = slugify(node.label)
        if key in note_by_slug:
            node.note_classifications.append(note_by_slug[key])
        for child in node.children:
            walk(child)

    walk(statement)


def _format_tree(node: HierarchyNode) -> str:
    lines = [node.label]
    for index, child in enumerate(node.children):
        lines.extend(_format_branch(child, "", index == len(node.children) - 1))
    return "\n".join(lines)


def _format_branch(node: HierarchyNode, prefix: str, is_last: bool) -> list[str]:
    connector = "└── " if is_last else "├── "
    lines = [f"{prefix}{connector}{node.label}"]
    extension = "    " if is_last else "│   "
    for index, child in enumerate(node.children):
        lines.extend(_format_branch(child, prefix + extension, index == len(node.children) - 1))
    return lines


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Print the ICAI Division I Schedule III hierarchy extracted from the Guidance Note PDF."
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print the full JSON-serializable mapping instead of the tree.",
    )
    parser.add_argument(
        "--pdf",
        type=Path,
        default=None,
        help="Override path to the ICAI Guidance Note PDF.",
    )
    args = parser.parse_args(argv)
    mapper = ScheduleIIIMapper(pdf_path=args.pdf)
    if args.json:
        print(mapper.to_json())
    else:
        reference = mapper.load_reference()
        print(reference.source_title)
        print(f"Source: {reference.source_path} ({reference.page_count} pages)")
        print()
        print(mapper.format_hierarchy())
        print()
        print("Identified sections:")
        for section in reference.sections:
            print(f"  - {section.title} (PDF page {section.start_page})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
