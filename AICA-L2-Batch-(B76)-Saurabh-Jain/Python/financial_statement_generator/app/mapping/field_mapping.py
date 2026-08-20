"""JSON-serializable field mappings from ICAI Schedule III to Excel destinations.

ICAI captions come from ``reference/ICAI_GN_Div_I_Sch_III.pdf``. Excel sheet
and cell destinations come from ``excel_field_map.json`` so they stay
configurable instead of being hard-coded in Python.
"""

from __future__ import annotations

import json
import re
import unicodedata
from dataclasses import asdict, dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Iterable

from app.config import settings

DEFAULT_MAPPING_CONFIG = Path(__file__).resolve().parent / "excel_field_map.json"
COORD_RE = re.compile(r"^\$?([A-Za-z]+)\$?(\d+)$")
DEFINED_NAME_RE = re.compile(r"(?:'([^']+)'|([^!]+))!\$?([A-Za-z]+)\$?(\d+)")
FUZZY_LABEL_THRESHOLD = 0.92

SOURCE_BLOCK_KEYS = (
    "balance_sheet",
    "profit_and_loss",
    "cash_flow",
    "notes_to_accounts",
    "share_capital",
    "reserves_and_surplus",
    "borrowings",
    "ppe",
    "cwip",
    "investments",
    "inventory",
    "trade_receivables",
    "trade_payables",
    "cash_and_cash_equivalents",
    "revenue",
    "expenses",
    "tax_expense",
    "eps",
    "other_disclosures",
    "ratios",
)


def slugify(label: str) -> str:
    """Create a stable identifier from a Schedule III label."""
    normalized = unicodedata.normalize("NFKD", label)
    normalized = normalized.replace("’", "'").replace("‘", "'").replace("“", '"').replace("”", '"')
    normalized = re.sub(r"[^A-Za-z0-9]+", "_", normalized)
    return normalized.strip("_").lower()


def normalize_label(value: str) -> str:
    cleaned = (value or "").replace("’", "'").replace("‘", "'")
    cleaned = cleaned.replace("–", "-").replace("—", "-")
    cleaned = re.sub(r"[^a-z0-9']+", " ", cleaned.lower())
    return " ".join(cleaned.split())


@dataclass(slots=True)
class FieldMapping:
    """One Schedule III caption that downstream modules can look up."""

    code: str
    label: str
    statement: str
    parent_code: str | None
    path: list[str]
    level: int
    kind: str
    source_pages: list[int] = field(default_factory=list)
    excerpt: str = ""
    aliases: list[str] = field(default_factory=list)
    note_classifications: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


# Populated when the ICAI reference is loaded. Kept as a dict so existing
# imports of FIELD_MAPPINGS continue to work.
FIELD_MAPPINGS: dict[str, dict[str, Any]] = {}


def aliases_for(label: str) -> list[str]:
    """Build lookup aliases that still originate from the extracted label."""
    compact = re.sub(r"\s+", " ", label).strip()
    variants = {
        compact,
        compact.lower(),
        compact.replace("’", "'"),
        re.sub(r"\s*\([^)]*\)\s*", " ", compact).strip(),
        slugify(compact).replace("_", " "),
    }
    return [item for item in variants if item]


def mapping_from_node(
    node: Any,
    *,
    statement: str,
    parent_code: str | None,
    path: list[str],
) -> FieldMapping:
    """Convert a hierarchy node into a serializable field mapping."""
    full_path = [*path, node.label]
    return FieldMapping(
        code=node.code,
        label=node.label,
        statement=statement,
        parent_code=parent_code,
        path=full_path,
        level=node.level,
        kind=node.kind,
        source_pages=list(node.source_pages),
        excerpt=node.excerpt,
        aliases=aliases_for(node.label),
        note_classifications=list(getattr(node, "note_classifications", []) or []),
    )


def flatten_hierarchy(
    nodes: Iterable[Any],
    *,
    statement: str,
    parent_code: str | None = None,
    path: list[str] | None = None,
) -> dict[str, FieldMapping]:
    """Walk a hierarchy tree and return ``code -> FieldMapping``."""
    mappings: dict[str, FieldMapping] = {}
    current_path = path or []
    for node in nodes:
        mapping = mapping_from_node(
            node,
            statement=statement,
            parent_code=parent_code,
            path=current_path,
        )
        mappings[mapping.code] = mapping
        mappings.update(
            flatten_hierarchy(
                node.children,
                statement=statement,
                parent_code=node.code,
                path=mapping.path,
            )
        )
    return mappings


def build_field_mappings(reference: Any) -> dict[str, dict[str, Any]]:
    """Create JSON-ready mappings from a loaded Schedule III reference."""
    mappings: dict[str, FieldMapping] = {}
    mappings.update(
        flatten_hierarchy(
            [reference.balance_sheet],
            statement="balance_sheet",
        )
    )
    mappings.update(
        flatten_hierarchy(
            [reference.profit_and_loss],
            statement="profit_and_loss",
        )
    )
    if reference.notes_to_accounts is not None:
        mappings.update(
            flatten_hierarchy(
                [reference.notes_to_accounts],
                statement="notes_to_accounts",
            )
        )
    if reference.ratios:
        mappings.update(
            flatten_hierarchy(
                reference.ratios,
                statement="ratios",
            )
        )

    serialized = {code: mapping.to_dict() for code, mapping in mappings.items()}
    FIELD_MAPPINGS.clear()
    FIELD_MAPPINGS.update(serialized)
    return serialized


def lookup_mapping(label: str, mappings: dict[str, dict[str, Any]] | None = None) -> dict[str, Any] | None:
    """Find a field mapping by code, exact label, or alias."""
    catalog = mappings if mappings is not None else FIELD_MAPPINGS
    needle = label.strip()
    if needle in catalog:
        return catalog[needle]

    needle_lower = needle.lower()
    for item in catalog.values():
        if item["label"].lower() == needle_lower:
            return item
        if needle_lower in {alias.lower() for alias in item.get("aliases", [])}:
            return item
        if slugify(item["label"]) == slugify(needle):
            return item
    return None


@dataclass(slots=True)
class PeriodTargets:
    current: str | None = None
    previous: str | None = None
    note: str | None = None

    def as_dict(self) -> dict[str, str | None]:
        return {"current": self.current, "previous": self.previous, "note": self.note}


@dataclass(slots=True)
class ExcelFieldConfig:
    """One configurable mapping from source captions to an Excel destination."""

    key: str
    schedule_iii_section: str
    schedule_iii_label: str
    statement: str
    excel_sheet: str
    source_labels: list[str] = field(default_factory=list)
    label_match: list[str] = field(default_factory=list)
    target_cells: PeriodTargets = field(default_factory=PeriodTargets)
    defined_names: PeriodTargets = field(default_factory=PeriodTargets)
    note_sheet: str | None = None
    kind: str = "exact"
    addends: list[str] = field(default_factory=list)
    overwrite_formula: bool = False
    value_role: str = "amount"
    write_note_number: bool = True

    def match_labels(self) -> list[str]:
        labels = [self.schedule_iii_label, self.key.replace("_", " "), *self.source_labels, *self.label_match]
        seen: set[str] = set()
        ordered: list[str] = []
        for label in labels:
            key = normalize_label(label)
            if not key or key in seen:
                continue
            seen.add(key)
            ordered.append(label)
        return ordered

    def to_dict(self) -> dict[str, Any]:
        payload = {
            "key": self.key,
            "schedule_iii_section": self.schedule_iii_section,
            "schedule_iii_label": self.schedule_iii_label,
            "statement": self.statement,
            "excel_sheet": self.excel_sheet,
            "note_sheet": self.note_sheet,
            "source_labels": list(self.source_labels),
            "label_match": list(self.label_match),
            "target_cells": self.target_cells.as_dict(),
            "defined_names": self.defined_names.as_dict(),
            "kind": self.kind,
            "addends": list(self.addends),
            "overwrite_formula": self.overwrite_formula,
            "value_role": self.value_role,
            "write_note_number": self.write_note_number,
        }
        return payload


@dataclass(slots=True)
class SourceItem:
    label: str
    current_value: float | str | int | None = None
    previous_value: float | str | int | None = None
    note_no: str | int | float | None = None
    confidence: float | None = None
    source_page: int | None = None
    source_text: str | None = None
    origin: str = ""


@dataclass(slots=True)
class MappingPlacement:
    field_key: str
    source_label: str
    extracted_value: float | str | int | None
    schedule_iii_category: str
    schedule_iii_label: str
    excel_sheet: str | None
    excel_cell: str | None
    note_sheet: str | None
    period: str
    confidence: float | None
    action: str
    resolution: str
    warnings: list[str] = field(default_factory=list)
    formula: str | None = None
    source_page: int | None = None

    def to_report_row(self) -> dict[str, Any]:
        return {
            "source_label": self.source_label,
            "extracted_value": self.extracted_value,
            "source_page": self.source_page,
            "schedule_iii_category": self.schedule_iii_category,
            "excel_destination": _format_destination(self.excel_sheet, self.excel_cell),
            "excel_sheet": self.excel_sheet,
            "excel_cell": self.excel_cell,
            "note_sheet": self.note_sheet,
            "period": self.period,
            "confidence": self.confidence,
            "action": self.action,
            "resolution": self.resolution,
            "warnings": list(self.warnings),
        }

    def to_dict(self) -> dict[str, Any]:
        payload = self.to_report_row()
        payload.update(
            {
                "field_key": self.field_key,
                "schedule_iii_label": self.schedule_iii_label,
                "formula": self.formula,
            }
        )
        return payload


class ExcelFieldMap:
    """Configuration catalog: source labels → Schedule III → Excel."""

    def __init__(self, payload: dict[str, Any], *, source_path: Path | None = None) -> None:
        self.source_path = source_path
        self.version = payload.get("version", 1)
        self.sheet_aliases = {
            key: [normalize_label(alias) for alias in aliases]
            for key, aliases in (payload.get("sheet_aliases") or {}).items()
        }
        self.sheets = payload.get("sheets") or {}
        self.synonyms = {
            normalize_label(canonical): [normalize_label(item) for item in variants]
            for canonical, variants in (payload.get("synonyms") or {}).items()
        }
        self.fields = {
            key: _field_from_config(key, spec) for key, spec in (payload.get("fields") or {}).items()
        }

    @classmethod
    def load(cls, path: str | Path | None = None) -> ExcelFieldMap:
        config_path = Path(path) if path else Path(getattr(settings, "mapping_config_path", DEFAULT_MAPPING_CONFIG))
        if not config_path.exists():
            config_path = DEFAULT_MAPPING_CONFIG
        payload = json.loads(config_path.read_text(encoding="utf-8"))
        return cls(payload, source_path=config_path)

    def resolve_source(self, label: str, icai_mappings: dict[str, dict[str, Any]] | None = None) -> tuple[ExcelFieldConfig, str, float] | None:
        """Match a source caption using exact labels, aliases, synonyms, then ICAI."""
        needle = normalize_label(label)
        if not needle:
            return None
        slug = slugify(label)
        if slug in self.fields:
            return self.fields[slug], "exact", 1.0

        for config in self.fields.values():
            for candidate in config.match_labels():
                if normalize_label(candidate) == needle:
                    return config, "exact", 1.0
                if slugify(candidate) == slug:
                    return config, "alias", 0.96

        synonym_canonical = self._canonical_from_synonym(needle)
        if synonym_canonical:
            for config in self.fields.values():
                if normalize_label(config.schedule_iii_label) == synonym_canonical:
                    return config, "synonym", 0.9
                if synonym_canonical in {normalize_label(item) for item in config.match_labels()}:
                    return config, "synonym", 0.9

        icai = lookup_mapping(label, icai_mappings)
        if icai:
            icai_slug = slugify(icai.get("label") or "")
            if icai_slug in self.fields:
                return self.fields[icai_slug], "icai_alias", 0.93
            for config in self.fields.values():
                if normalize_label(config.schedule_iii_label) == normalize_label(icai.get("label") or ""):
                    return config, "icai_alias", 0.93

        best: tuple[ExcelFieldConfig, str, float] | None = None
        for config in self.fields.values():
            for candidate in config.match_labels():
                score = SequenceMatcher(None, needle, normalize_label(candidate)).ratio()
                if score < FUZZY_LABEL_THRESHOLD:
                    continue
                if best is None or score > best[2]:
                    best = (config, "fuzzy", round(score, 4))
        return best

    def _canonical_from_synonym(self, needle: str) -> str | None:
        for canonical, variants in self.synonyms.items():
            if needle == canonical or needle in variants:
                return canonical
        return None

    def resolve_sheet_name(self, preferred: str, available: Iterable[str]) -> str | None:
        names = [name for name in available if name]
        if preferred in names:
            return preferred
        wanted = {normalize_label(preferred), *self.sheet_aliases.get(preferred, [])}
        for name in names:
            if normalize_label(name) in wanted:
                return name
        preferred_stripped = preferred.strip()
        for name in names:
            if name.strip() == preferred_stripped:
                return name
        for name in names:
            if normalize_label(preferred) in normalize_label(name):
                return name
        return None

    def sheet_config(self, sheet_name: str | None) -> dict[str, Any]:
        if not sheet_name:
            return {}
        if sheet_name in self.sheets:
            return self.sheets[sheet_name]
        stripped = sheet_name.strip()
        if stripped in self.sheets:
            return self.sheets[stripped]
        needle = normalize_label(sheet_name)
        for key, config in self.sheets.items():
            aliases = {normalize_label(key), *self.sheet_aliases.get(key, [])}
            if needle in aliases:
                return config
        return {}


class ExcelMappingEngine:
    """Map classified source values onto Schedule III categories and Excel cells."""

    def __init__(
        self,
        *,
        field_map: ExcelFieldMap | None = None,
        icai_mappings: dict[str, dict[str, Any]] | None = None,
        template_metadata: dict[str, Any] | None = None,
    ) -> None:
        self.field_map = field_map or ExcelFieldMap.load()
        self.icai_mappings = icai_mappings if icai_mappings is not None else FIELD_MAPPINGS
        self.template_metadata = template_metadata or {}

    def map(self, classified_data: Any) -> dict[str, Any]:
        sources = collect_source_items(classified_data)
        placements: list[MappingPlacement] = []
        warnings: list[str] = []
        matched_keys: set[str] = set()
        unmapped: list[dict[str, Any]] = []

        for source in sources:
            resolved = self.field_map.resolve_source(source.label, self.icai_mappings)
            if resolved is None:
                unmapped.append(
                    {
                        "source_label": source.label,
                        "extracted_value": source.current_value,
                        "origin": source.origin,
                        "warning": "No Schedule III / Excel mapping matched this source label.",
                    }
                )
                continue
            config, match_kind, match_score = resolved
            matched_keys.add(config.key)
            confidence = _combine_confidence(source.confidence, match_score)
            placements.extend(
                self._placements_for_source(
                    config, source, match_kind, confidence, source_page=source.source_page
                )
            )

        for config in self.field_map.fields.values():
            if config.kind == "derived_total":
                placements.extend(self._derived_total_placements(config, matched_keys, sources))

        placements = _dedupe_placements(placements)

        derived_keys = {key for key, spec in self.field_map.fields.items() if spec.kind == "derived_total"}
        report_rows = [item.to_report_row() for item in placements]
        summary = {
            "mapped": sum(1 for item in placements if item.action == "write"),
            "skipped_formulas": sum(1 for item in placements if item.action == "skip_formula"),
            "missing_values": sum(1 for item in placements if item.action == "missing_value"),
            "unmapped_sources": len(unmapped),
            "derived_totals": sum(1 for item in placements if item.field_key in derived_keys),
        }
        if unmapped:
            warnings.append(f"{len(unmapped)} source label(s) could not be mapped.")
        return {
            "placements": [item.to_dict() for item in placements],
            "report": {"rows": report_rows, "summary": summary},
            "unmapped_sources": unmapped,
            "warnings": warnings + [msg for item in placements for msg in item.warnings],
            "config_path": str(self.field_map.source_path) if self.field_map.source_path else None,
        }

    def _placements_for_source(
        self,
        config: ExcelFieldConfig,
        source: SourceItem,
        match_kind: str,
        confidence: float | None,
        source_page: int | None = None,
    ) -> list[MappingPlacement]:
        periods: list[tuple[str, float | str | int | None]] = [
            ("current", source.current_value),
        ]
        if config.value_role != "text":
            periods.append(("previous", source.previous_value))
            if config.write_note_number:
                periods.append(("note", source.note_no))
        placements: list[MappingPlacement] = []
        for period, value in periods:
            if period == "note" and value is None:
                continue
            if period != "note" and value is None:
                placements.append(
                    self._empty_placement(
                        config,
                        period=period,
                        warning=f"{period.replace('_', ' ').title()} value was not extracted.",
                        action="missing_value",
                        source_label=source.label,
                        confidence=confidence,
                        source_page=source_page,
                    )
                )
                continue
            placements.append(
                self._placement(
                    config,
                    period=period,
                    value=value,
                    source_label=source.label,
                    confidence=confidence,
                    match_kind=match_kind,
                    source_page=source_page,
                )
            )
        return placements

    def _derived_total_placements(
        self,
        config: ExcelFieldConfig,
        matched_keys: set[str],
        sources: list[SourceItem],
    ) -> list[MappingPlacement]:
        by_key: dict[str, SourceItem] = {}
        for source in sources:
            resolved = self.field_map.resolve_source(source.label, self.icai_mappings)
            if resolved is None:
                continue
            by_key[resolved[0].key] = source
        missing = [key for key in config.addends if key not in by_key or by_key[key].current_value is None]
        if missing:
            return [
                self._empty_placement(
                    config,
                    period="current",
                    warning=(
                        "Derived total was not computed because one or more addends "
                        f"are missing: {', '.join(missing)}."
                    ),
                    action="missing_value",
                    resolution="derived_total",
                )
            ]
        current_total = 0.0
        previous_values: list[float] = []
        confidences: list[float] = []
        for key in config.addends:
            item = by_key[key]
            current_total += float(item.current_value)  # type: ignore[arg-type]
            if isinstance(item.previous_value, (int, float)):
                previous_values.append(float(item.previous_value))
            if item.confidence is not None:
                confidences.append(item.confidence)
        confidence = min(confidences) if confidences else None
        placements = [
            self._placement(
                config,
                period="current",
                value=current_total,
                source_label=" + ".join(config.addends),
                confidence=confidence,
                match_kind="derived_total",
            )
        ]
        if len(previous_values) == len(config.addends):
            placements.append(
                self._placement(
                    config,
                    period="previous",
                    value=sum(previous_values),
                    source_label=" + ".join(config.addends),
                    confidence=confidence,
                    match_kind="derived_total",
                )
            )
        return placements

    def _placement(
        self,
        config: ExcelFieldConfig,
        *,
        period: str,
        value: float | str | int | None,
        source_label: str,
        confidence: float | None,
        match_kind: str,
        source_page: int | None = None,
    ) -> MappingPlacement:
        sheet, cell, resolution, formula, locate_warnings = self._locate_destination(config, period)
        action = "write"
        warnings = list(locate_warnings)
        if match_kind == "fuzzy":
            warnings.append(f"Matched '{source_label}' to '{config.schedule_iii_label}' with a fuzzy alias.")
        if cell and formula and not config.overwrite_formula:
            action = "skip_formula"
            warnings.append(
                f"{sheet}!{cell} contains a formula and will not be overwritten."
            )
        elif cell is None:
            action = "unmapped_destination"
            warnings.append("Excel destination cell could not be resolved from labels, names, or config.")
        return MappingPlacement(
            field_key=config.key,
            source_label=source_label,
            extracted_value=value,
            schedule_iii_category=config.schedule_iii_section,
            schedule_iii_label=config.schedule_iii_label,
            excel_sheet=sheet or config.excel_sheet,
            excel_cell=cell,
            note_sheet=config.note_sheet,
            period=period,
            confidence=confidence,
            action=action,
            resolution=resolution or match_kind,
            warnings=warnings,
            formula=formula,
            source_page=source_page,
        )

    def _empty_placement(
        self,
        config: ExcelFieldConfig,
        *,
        period: str,
        warning: str,
        action: str,
        source_label: str | None = None,
        confidence: float | None = None,
        resolution: str = "config",
        source_page: int | None = None,
    ) -> MappingPlacement:
        sheet, cell, located, formula, locate_warnings = self._locate_destination(config, period)
        return MappingPlacement(
            field_key=config.key,
            source_label=source_label or config.schedule_iii_label,
            extracted_value=None,
            schedule_iii_category=config.schedule_iii_section,
            schedule_iii_label=config.schedule_iii_label,
            excel_sheet=sheet or config.excel_sheet,
            excel_cell=cell,
            note_sheet=config.note_sheet,
            period=period,
            confidence=confidence,
            action=action,
            resolution=located or resolution,
            warnings=[warning, *locate_warnings],
            formula=formula,
            source_page=source_page,
        )

    def _locate_destination(
        self,
        config: ExcelFieldConfig,
        period: str,
    ) -> tuple[str | None, str | None, str, str | None, list[str]]:
        warnings: list[str] = []
        available = self.template_metadata.get("sheet_names") or [
            sheet.get("name") for sheet in self.template_metadata.get("sheets") or []
        ]
        sheet_name = self.field_map.resolve_sheet_name(config.excel_sheet, [name for name in available if name])
        if sheet_name is None:
            sheet_name = config.excel_sheet
            if self.template_metadata:
                warnings.append(
                    f"Sheet '{config.excel_sheet}' was not found in the template; "
                    "configured destination was kept without label verification."
                )

        named = getattr(config.defined_names, period, None)
        if named:
            resolved = _resolve_defined_name(self.template_metadata, named)
            if resolved:
                named_sheet, named_cell = resolved
                return named_sheet, named_cell, "defined_name", _cell_formula(self._sheet_meta(named_sheet), named_cell), warnings

        label_cell = self._find_label_cell(sheet_name, config)
        if label_cell is not None:
            _label_coord, row = label_cell
            column = self._column_for_period(sheet_name, period, row)
            if column:
                coord = f"{column}{row}"
                return (
                    sheet_name,
                    coord,
                    "template_label",
                    _cell_formula(self._sheet_meta(sheet_name), coord),
                    warnings,
                )
            warnings.append(
                f"Matched label '{config.schedule_iii_label}' on {sheet_name} row {row} "
                f"but could not resolve the {period} column."
            )

        configured = getattr(config.target_cells, period, None)
        if configured:
            if label_cell is None and self._sheet_meta(sheet_name):
                warnings.append(
                    f"Used configured cell {configured} because the Schedule III label "
                    f"was not found on '{sheet_name}'."
                )
            return (
                sheet_name,
                configured.upper(),
                "configured_cell",
                _cell_formula(self._sheet_meta(sheet_name), configured.upper()),
                warnings,
            )
        return sheet_name, None, "unresolved", None, warnings

    def _find_label_cell(self, sheet_name: str | None, config: ExcelFieldConfig) -> tuple[str, int] | None:
        meta = self._sheet_meta(sheet_name)
        if not meta:
            return None
        sheet_cfg = self.field_map.sheet_config(config.excel_sheet) or self.field_map.sheet_config(sheet_name)
        label_columns = {col.upper() for col in sheet_cfg.get("label_columns") or ["A", "B"]}
        needles = [normalize_label(item) for item in config.match_labels()]
        best: tuple[str, int, float] | None = None
        for coord, cell in (meta.get("cells") or {}).items():
            parsed = split_coord(coord)
            if parsed is None:
                continue
            column, row = parsed
            if label_columns and column not in label_columns:
                continue
            value = cell.get("value")
            if not isinstance(value, str) or not value.strip():
                continue
            score = _best_label_score(normalize_label(value), needles)
            if score < 0.92:
                continue
            if best is None or score > best[2]:
                best = (coord.upper(), row, score)
        if best is None:
            return None
        return best[0], best[1]

    def _column_for_period(self, sheet_name: str | None, period: str, row: int) -> str | None:
        meta = self._sheet_meta(sheet_name)
        sheet_cfg = self.field_map.sheet_config(sheet_name)
        discovered = self._discover_columns(meta, sheet_cfg)
        if period in discovered:
            return discovered[period]
        configured = (sheet_cfg.get("columns") or {}).get(period) or []
        if configured:
            return str(configured[0]).upper()
        return None

    def _discover_columns(self, sheet_meta: dict[str, Any] | None, sheet_cfg: dict[str, Any]) -> dict[str, str]:
        if not sheet_meta:
            return {}
        keywords = sheet_cfg.get("header_keywords") or {
            "note": ["note"],
            "current": ["current", "as at", "year ended"],
            "previous": ["previous", "prior"],
        }
        found: dict[str, str] = {}
        for coord, cell in (sheet_meta.get("cells") or {}).items():
            parsed = split_coord(coord)
            if parsed is None:
                continue
            column, row = parsed
            if row > 20:
                continue
            if column in {col.upper() for col in sheet_cfg.get("label_columns") or ["A", "B"]}:
                continue
            value = cell.get("value")
            if not isinstance(value, str):
                continue
            blob = normalize_label(value)
            for role, needles in keywords.items():
                if not any(normalize_label(needle) in blob for needle in needles):
                    continue
                if role not in found:
                    found[role] = column
                elif (
                    role == "current"
                    and "previous" not in found
                    and column != found[role]
                ):
                    found["previous"] = column
        return found

    def _sheet_meta(self, sheet_name: str | None) -> dict[str, Any] | None:
        if not sheet_name:
            return None
        for sheet in self.template_metadata.get("sheets") or []:
            if sheet.get("name") == sheet_name:
                return sheet
        wanted = sheet_name.strip()
        for sheet in self.template_metadata.get("sheets") or []:
            if (sheet.get("name") or "").strip() == wanted:
                return sheet
        return None


def collect_source_items(classified_data: Any) -> list[SourceItem]:
    """Flatten classified models/dicts into mappable source items."""
    payload = classified_data
    if hasattr(classified_data, "model_dump"):
        payload = classified_data.model_dump(mode="json")
    if not isinstance(payload, dict):
        return []
    items: list[SourceItem] = []
    company = payload.get("company") or {}
    if isinstance(company, dict):
        for field_name, sourced in company.items():
            parsed = _sourced_dict(sourced)
            if parsed is None or parsed["value"] is None:
                continue
            items.append(
                SourceItem(
                    label=field_name.replace("_", " "),
                    current_value=parsed["value"],
                    confidence=parsed["confidence"],
                    source_page=parsed["source_page"],
                    source_text=parsed["source_text"],
                    origin="company",
                )
            )
    for block_key in SOURCE_BLOCK_KEYS:
        block = payload.get(block_key) or {}
        if not isinstance(block, dict):
            continue
        for line in block.get("line_items") or []:
            item = _line_item_to_source(line, origin=block_key)
            if item is not None:
                items.append(item)
        for field_name, sourced in block.items():
            if field_name in {"line_items", "notes", "identified", "start_page", "end_page", "pages", "confidence", "excerpt"}:
                continue
            parsed = _sourced_dict(sourced)
            if parsed is None or parsed["value"] is None:
                continue
            items.append(
                SourceItem(
                    label=field_name.replace("_", " "),
                    current_value=parsed["value"],
                    confidence=parsed["confidence"],
                    source_page=parsed["source_page"],
                    source_text=parsed["source_text"],
                    origin=block_key,
                )
            )
        for note in block.get("notes") or []:
            if not isinstance(note, dict):
                continue
            for line in note.get("line_items") or []:
                item = _line_item_to_source(line, origin=f"{block_key}.notes")
                if item is not None:
                    items.append(item)
    return items


def _dedupe_placements(placements: list[MappingPlacement]) -> list[MappingPlacement]:
    unique: dict[tuple[str, str], MappingPlacement] = {}
    order: list[tuple[str, str]] = []
    for item in placements:
        key = (item.field_key, item.period)
        existing = unique.get(key)
        if existing is None:
            unique[key] = item
            order.append(key)
            continue
        if (item.confidence or 0) > (existing.confidence or 0):
            unique[key] = item
    return [unique[key] for key in order]


def split_coord(coord: str) -> tuple[str, int] | None:
    match = COORD_RE.match((coord or "").replace("$", "").strip().upper())
    if not match:
        return None
    return match.group(1).upper(), int(match.group(2))


def load_excel_field_map(path: str | Path | None = None) -> ExcelFieldMap:
    return ExcelFieldMap.load(path)


def _field_from_config(key: str, spec: dict[str, Any]) -> ExcelFieldConfig:
    return ExcelFieldConfig(
        key=key,
        schedule_iii_section=spec.get("schedule_iii_section") or "",
        schedule_iii_label=spec.get("schedule_iii_label") or key.replace("_", " "),
        statement=spec.get("statement") or "",
        excel_sheet=spec.get("excel_sheet") or "",
        source_labels=list(spec.get("source_labels") or []),
        label_match=list(spec.get("label_match") or []),
        target_cells=_period_targets(spec.get("target_cells")),
        defined_names=_period_targets(spec.get("defined_names")),
        note_sheet=spec.get("note_sheet"),
        kind=spec.get("kind") or "exact",
        addends=list(spec.get("addends") or []),
        overwrite_formula=bool(spec.get("overwrite_formula", False)),
        value_role=spec.get("value_role") or "amount",
        write_note_number=bool(spec.get("write_note_number", True)),
    )


def _period_targets(raw: Any) -> PeriodTargets:
    if raw is None:
        return PeriodTargets()
    if isinstance(raw, list):
        current = raw[0] if len(raw) > 0 else None
        previous = raw[1] if len(raw) > 1 else None
        note = raw[2] if len(raw) > 2 else None
        return PeriodTargets(current=current, previous=previous, note=note)
    if isinstance(raw, dict):
        return PeriodTargets(
            current=raw.get("current"),
            previous=raw.get("previous"),
            note=raw.get("note"),
        )
    if isinstance(raw, str):
        return PeriodTargets(current=raw)
    return PeriodTargets()


def _line_item_to_source(line: Any, *, origin: str) -> SourceItem | None:
    if not isinstance(line, dict):
        return None
    label = str(line.get("label") or "").strip()
    if not label:
        return None
    current = _sourced_dict(line.get("current_period"))
    previous = _sourced_dict(line.get("previous_period"))
    note = _sourced_dict(line.get("note_no"))
    current_value = None if current is None else current["value"]
    previous_value = None if previous is None else previous["value"]
    note_value = None if note is None else note["value"]
    if current_value is None and previous_value is None and note_value is None:
        return None
    confidence = None
    for parsed in (current, previous):
        if parsed and parsed["confidence"] is not None:
            confidence = parsed["confidence"]
            break
    source_page = (current or previous or note or {}).get("source_page")
    source_text = (current or previous or note or {}).get("source_text")
    return SourceItem(
        label=label,
        current_value=current_value,
        previous_value=previous_value,
        note_no=note_value,
        confidence=confidence,
        source_page=source_page,
        source_text=source_text,
        origin=origin,
    )


def _sourced_dict(value: Any) -> dict[str, Any] | None:
    if value is None:
        return None
    if hasattr(value, "model_dump"):
        value = value.model_dump(mode="json")
    if not isinstance(value, dict) or "value" not in value:
        return None
    return {
        "value": value.get("value"),
        "source_page": value.get("source_page"),
        "source_text": value.get("source_text"),
        "confidence": value.get("confidence"),
    }


def _combine_confidence(extracted: float | None, match_score: float) -> float:
    if extracted is None:
        return round(match_score, 4)
    return round(min(float(extracted), match_score), 4)


def _best_label_score(value: str, needles: list[str]) -> float:
    best = 0.0
    for needle in needles:
        if not needle:
            continue
        if value == needle:
            return 1.0
        if needle in value or value in needle:
            ratio = min(len(value), len(needle)) / max(len(value), len(needle))
            best = max(best, 0.95 if ratio >= 0.7 else 0.9)
        else:
            best = max(best, SequenceMatcher(None, value, needle).ratio())
    return best


def _format_destination(sheet: str | None, cell: str | None) -> str | None:
    if sheet and cell:
        return f"{sheet}!{cell}"
    if sheet:
        return sheet
    return cell


def _cell_formula(sheet_meta: dict[str, Any] | None, coord: str | None) -> str | None:
    if not sheet_meta or not coord:
        return None
    formulas = sheet_meta.get("formulas") or {}
    if coord in formulas:
        return formulas[coord]
    cell = (sheet_meta.get("cells") or {}).get(coord) or (sheet_meta.get("cells") or {}).get(coord.upper())
    if not cell:
        return None
    return cell.get("formula")


def _resolve_defined_name(metadata: dict[str, Any], name: str) -> tuple[str, str] | None:
    for item in metadata.get("defined_names") or []:
        if (item.get("name") or "").lower() != name.lower():
            continue
        attr = item.get("attr_text") or item.get("value") or ""
        match = DEFINED_NAME_RE.search(str(attr))
        if not match:
            continue
        sheet = match.group(1) or match.group(2)
        cell = f"{match.group(3).upper()}{match.group(4)}"
        return sheet, cell
    return None
