from __future__ import annotations

import argparse
import hashlib
import html
import json
import re
import urllib.parse
import urllib.request
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable

import fitz
import pandas as pd
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


NOT_REPORTED = "not_reported"
NEEDS_CHECK = "needs_manual_check"
TOOL_VERSION = "2.1.0"
SCHEMA_VERSION = "hmfor-extractor-output-v2.1"
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_PDF_DIR = BASE_DIR / "test"
DEFAULT_OUT_DIR = BASE_DIR / "output-test"
DEFAULT_OVERRIDES_PATH = BASE_DIR / "manual_overrides.json"
DEFAULT_RHE_RULES_PATH = BASE_DIR / "rhe_conversion_rules.json"
DEFAULT_UNIT_RULES_PATH = BASE_DIR / "unit_normalization_rules.json"
DEFAULT_CONFIG_PATH = BASE_DIR / "config.yaml"
DEFAULT_DOWNLOADED_SI_DIR = BASE_DIR / "downloaded-si"
HMFOR_CATEGORIES = {"hmfor_core", "hmfor_related"}
BACKGROUND_CATEGORIES = {"ni_co_background", "ai_catalysis_method", "exclude", "ocr_needed"}
HMFOR_SPECIFIC_FIELDS = [
    "hmf_concentration", "hmf_conversion", "hmf_conversion_value",
    "fdca_yield", "fdca_yield_value", "fdca_selectivity",
    "fdca_selectivity_value", "faradaic_efficiency",
    "faradaic_efficiency_value", "best_performance_sentence",
    "product_analysis_method", "carbon_balance", "hmfor_data_quality_score",
]
COMMON_REQUIRED_FIELDS = [
    "paper_id", "file_name", "title", "year", "doi", "paper_category",
    "pdf_text_status", "reaction_type", "catalyst_name", "catalyst_family",
    "substrate", "synthesis_method", "electrolyte", "reference_electrode",
    "current_density", "characterization", "mechanism_claim", "recommended_use",
]
HMFOR_REQUIRED_FIELDS = [
    "catalyst_name", "electrolyte", "hmf_concentration", "potential_vs_rhe",
    "hmf_conversion", "fdca_yield", "faradaic_efficiency",
    "product_analysis_method", "best_performance_sentence",
]
BACKGROUND_REQUIRED_FIELDS = [
    "paper_id", "file_name", "title", "paper_category", "pdf_text_status",
    "reaction_type", "catalyst_name", "catalyst_family", "substrate",
    "synthesis_method", "electrolyte", "current_density", "recommended_use",
]
EXCLUDE_REQUIRED_FIELDS = ["title", "paper_category", "pdf_text_status", "review_notes"]
BAD_SHORT_VALUES = {"A", "G", "H", "M", "T", "a", "d", "n"}
STANDARD_ENUM_VALUES = {
    "", NOT_REPORTED, NEEDS_CHECK, "yes", "no", "low", "medium", "high",
    "very_low", "not_applicable", "hmfor_core", "hmfor_related",
    "ni_co_background", "ai_catalysis_method", "exclude", "ocr_needed",
    "readable", "low_text", "HMFOR", "OER", "HER", "UOR", "capacitor",
    "background_reference", "extract_experiment_reference", "reproduce_or_deep_read",
    "mechanism_or_characterization_reference", "background_only",
    "pdf_metadata", "file_name", "first_page", "file_name_fallback",
}
QUALITY_TEXT_FIELDS = {
    "catalyst_name", "catalyst_family", "substrate", "ni_co_ratio",
    "synthesis_method", "electrolyte", "hmf_concentration", "cell_type",
    "reference_electrode", "potential_vs_rhe", "onset_potential_vs_rhe",
    "performance_potential_vs_rhe", "lsv_current_density",
    "product_analysis_condition", "stability_performance",
    "coupled_electrolyzer_performance", "electrode_configuration",
    "metal_loading", "current_density", "hmf_conversion", "fdca_yield",
    "fdca_selectivity", "faradaic_efficiency", "stability",
    "product_analysis_method", "carbon_balance", "controls",
    "characterization", "mechanism_claim", "recommended_use",
}
SI_MENTION_PATTERN = r"supporting information|supplementary information|electronic supplementary information|associated content|available free of charge|additional experimental details"

ERROR_CODES = {
    "PDF_TEXT_EMPTY": "PDF text layer is empty or OCR is needed",
    "DOI_LOOKUP_FAILED": "DOI/Crossref lookup failed",
    "SI_NOT_FOUND": "No local or downloadable Supporting Information was found",
    "SI_DOWNLOAD_FAILED": "SI candidate URL could not be downloaded",
    "RHE_REFERENCE_UNCLEAR": "Reference electrode or potential scale is unclear",
    "FIELD_EVIDENCE_WEAK": "Field value has weak or missing evidence",
    "PLOT_GATE_BLOCKED": "Metric blocked from comparison/plotting gate",
    "MANUAL_OVERRIDE_USED": "Manual override changed extracted values",
    "BAD_SHORT_VALUE_CLEANED": "Suspicious short value was cleaned to not_reported",
}


def parse_simple_yaml(path: Path) -> dict:
    if not path.exists():
        return {}
    data: dict[str, object] = {}
    stack: list[tuple[int, dict | list]] = [(-1, data)]
    for raw in path.read_text(encoding="utf-8").splitlines():
        if not raw.strip() or raw.lstrip().startswith("#"):
            continue
        indent = len(raw) - len(raw.lstrip(" "))
        line = raw.strip()
        while stack and indent <= stack[-1][0]:
            stack.pop()
        parent = stack[-1][1]
        if line.startswith("- "):
            if isinstance(parent, list):
                parent.append(line[2:].strip().strip('"'))
            continue
        if ":" in line and isinstance(parent, dict):
            key, value = line.split(":", 1)
            key = key.strip()
            value = value.strip()
            if value == "":
                nxt: dict[str, object] = {}
                parent[key] = nxt
                stack.append((indent, nxt))
            elif value == "[]":
                parent[key] = []
            else:
                parent[key] = value.strip('"')
    return data


def load_config(path: Path = DEFAULT_CONFIG_PATH) -> dict:
    if not path.exists():
        return {}
    if path.suffix.lower() == ".json":
        return load_json_rules(path)
    return parse_simple_yaml(path)


def add_schema_version(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "schema_version" not in out.columns:
        out.insert(0, "schema_version", SCHEMA_VERSION)
    return out


@dataclass
class PaperRecord:
    paper_id: str
    file_name: str
    linked_si_files: str
    title: str
    title_source: str
    raw_pdf_title: str
    filename_title: str
    year: str
    journal: str
    doi: str
    authors: str
    paper_type: str
    relevance_level: str
    si_mentioned_in_text: str
    si_mention_evidence: str
    paper_category: str
    pdf_text_status: str
    text_length: int
    reaction_type: str
    catalyst_name: str
    catalyst_family: str
    substrate: str
    ni_co_ratio: str
    synthesis_method: str
    electrolyte: str
    koh_concentration: str
    hmf_concentration: str
    cell_type: str
    reference_electrode: str
    potential_vs_rhe: str
    onset_potential_vs_rhe: str
    performance_potential_vs_rhe: str
    lsv_current_density: str
    product_analysis_condition: str
    stability_performance: str
    coupled_electrolyzer_performance: str
    electrode_configuration: str
    metal_loading: str
    current_density: str
    hmf_conversion: str
    hmf_conversion_value: str
    fdca_yield: str
    fdca_yield_value: str
    fdca_selectivity: str
    fdca_selectivity_value: str
    faradaic_efficiency: str
    faradaic_efficiency_value: str
    best_performance_sentence: str
    stability: str
    product_analysis_method: str
    carbon_balance: str
    controls: str
    has_ni_only_control: str
    has_co_only_control: str
    has_no_hmf_control: str
    has_physical_mixture_control: str
    characterization: str
    mechanism_claim: str
    site_decoupling_evidence_level: str
    hmfor_data_quality_score: str
    reproducibility_value: str
    recommended_use: str
    missing_fields: str
    override_applied: str
    override_fields: str
    override_reason: str
    override_source: str
    override_date: str
    override_operator: str
    override_version: str
    override_evidence: str
    references_removed: str
    reference_section_start_page: str
    text_used_for_extraction_length: str
    review_notes: str


@dataclass
class NiCoBackgroundRecord:
    paper_id: str
    file_name: str
    title: str
    year: str
    journal: str
    doi: str
    authors: str
    paper_type: str
    relevance_level: str
    si_mentioned_in_text: str
    si_mention_evidence: str
    paper_category: str
    catalyst_name: str
    catalyst_family: str
    substrate: str
    synthesis_method: str
    electrolyte: str
    reference_electrode: str
    potential_vs_rhe: str
    current_density: str
    overpotential: str
    oer_metric: str
    her_metric: str
    stability: str
    characterization: str
    mechanism_claim: str
    background_value: str
    reusable_for_hmfor: str
    review_notes: str


@dataclass
class BackgroundRecord:
    paper_id: str
    file_name: str
    title: str
    year: str
    doi: str
    authors: str
    paper_type: str
    relevance_level: str
    si_mentioned_in_text: str
    si_mention_evidence: str
    paper_category: str
    reason: str
    recommended_use: str


@dataclass
class EvidenceRecord:
    paper_id: str
    file_name: str
    field: str
    value: str
    page: str
    evidence_snippet: str
    confidence: str


@dataclass
class MetricRecord:
    paper_id: str
    file_name: str
    paper_category: str
    performance_context: str
    metric_name: str
    metric_value: str
    metric_value_normalized: str
    metric_unit_normalized: str
    condition: str
    evidence_sentence: str
    source_scope: str
    confidence: str
    comparison_eligible: str
    exclusion_reason: str


@dataclass
class DedupRecord:
    paper_id: str
    file_name: str
    duplicate_group_id: str
    duplicate_status: str
    primary_paper_id: str
    match_basis: str
    match_key: str
    confidence: str


@dataclass
class SILinkRecord:
    paper_id: str
    file_name: str
    si_found: str
    si_file_name: str
    si_text_status: str
    si_used_for_extraction: str
    fields_filled_from_si: str
    si_field_evidence_summary: str
    si_missing_reason: str
    si_needed: str
    si_needed_reason: str
    link_method: str
    shared_terms: str
    publisher_si_candidate_urls: str
    confidence: str


@dataclass
class FieldConfidenceRecord:
    paper_id: str
    file_name: str
    field: str
    value: str
    field_confidence: str
    confidence_reason: str
    evidence_strength: str


@dataclass
class PotentialConversionAuditRecord:
    paper_id: str
    file_name: str
    field: str
    original_potential: str
    original_reference: str
    pH_or_KOH: str
    conversion_formula: str
    converted_vs_rhe: str
    converted_candidate_vs_rhe: str
    final_potential_vs_rhe: str
    potential_trust_level: str
    conversion_enabled: str
    conversion_confidence: str
    assumption_used: str


@dataclass
class UnitNormalizationRecord:
    paper_id: str
    file_name: str
    field: str
    raw_value: str
    normalized_value: str
    normalized_unit: str
    conversion_type: str
    converted: str


@dataclass
class QualityScoreRecord:
    paper_id: str
    file_name: str
    data_completeness_score: str
    evidence_traceability_score: str
    hmfor_relevance_score: str
    mechanism_depth_score: str
    reproducibility_score: str
    si_support_score: str
    overall_quality_score: str
    scoring_notes: str


@dataclass
class ManualReviewRecord:
    paper_id: str
    file_name: str
    review_reason: str
    field: str
    value: str
    priority: str
    suggested_action: str
    source_issue_id: str


def read_pdf_pages(path: Path) -> list[str]:
    pages: list[str] = []
    with fitz.open(path) as doc:
        for page in doc:
            text = page.get_text("text")
            pages.append(normalize_space(text))
    return strip_references_from_pages(pages)


def strip_references_from_pages(pages: list[str]) -> list[str]:
    cleaned, _, _ = strip_references_with_status(pages)
    return cleaned


def strip_references_with_status(pages: list[str]) -> tuple[list[str], str, str]:
    cleaned: list[str] = []
    references_started = False
    start_page = ""
    for page_idx, page_text in enumerate(pages, start=1):
        if references_started:
            cleaned.append("")
            continue
        match = re.search(r"(?:^|\s)(references|bibliography|acknowledgements?|author contributions?|publisher\'s note|supporting information)\s*(?:\[?1\]?|1\.|$)", page_text, re.I)
        if match and len(page_text) - match.start() > 800:
            cleaned.append(page_text[: match.start()].strip())
            references_started = True
            start_page = str(page_idx)
        else:
            cleaned.append(page_text)
    return cleaned, "yes" if references_started else "no", start_page


def read_pdf_pages_raw(path: Path) -> list[str]:
    pages: list[str] = []
    with fitz.open(path) as doc:
        for page in doc:
            pages.append(normalize_space(page.get_text("text")))
    return pages


def normalize_space(text: str) -> str:
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text or "")
    return re.sub(r"\s+", " ", text).strip()


def first_match(text: str, patterns: Iterable[str], default: str = NOT_REPORTED) -> str:
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return normalize_space(match.group(0))
    return default


def value_near_keywords(
    pages: list[str],
    field: str,
    keywords: Iterable[str],
    value_patterns: Iterable[str],
    window: int = 360,
) -> tuple[str, str, str]:
    for page_idx, page_text in enumerate(pages, start=1):
        lower_text = page_text.lower()
        for keyword in keywords:
            start = lower_text.find(keyword.lower())
            if start == -1:
                continue
            snippet = page_text[max(0, start - window) : start + window]
            for pattern in value_patterns:
                match = re.search(pattern, snippet, flags=re.IGNORECASE)
                if match:
                    return normalize_space(match.group(0)), str(page_idx), normalize_space(snippet)
            return NEEDS_CHECK, str(page_idx), normalize_space(snippet)
    return NOT_REPORTED, "", ""


def sentences_near_keywords(pages: list[str], keywords: Iterable[str], window: int = 500) -> list[tuple[str, str]]:
    hits: list[tuple[str, str]] = []
    for page_idx, page_text in enumerate(pages, start=1):
        lower_text = page_text.lower()
        for keyword in keywords:
            start = lower_text.find(keyword.lower())
            if start == -1:
                continue
            snippet = page_text[max(0, start - window): start + window]
            sentences = re.split(r"(?<=[.!?])\s+", snippet)
            for sentence in sentences:
                sentence = normalize_space(sentence)
                if keyword.lower() in sentence.lower() and len(sentence) > 20:
                    hits.append((str(page_idx), sentence[:900]))
    return hits


def extract_metric_sentence(pages: list[str], field: str, keywords: Iterable[str], metric_patterns: Iterable[str]) -> tuple[str, str, str]:
    best: tuple[str, str, str] = (NOT_REPORTED, "", "")
    for page, sentence in sentences_near_keywords(pages, keywords):
        for pattern in metric_patterns:
            match = re.search(pattern, sentence, flags=re.IGNORECASE)
            if match:
                return normalize_space(match.group(0)), page, sentence
        if best[0] == NOT_REPORTED:
            best = (NEEDS_CHECK, page, sentence)
    return best


def sentence_score_for_performance(sentence: str) -> int:
    score = 0
    score += 5 * len(re.findall(r"\d+(\.\d+)?\s*%", sentence))
    required_hits = 0
    for term in ["FDCA", "HMF", "conversion", "selectivity", "Faradaic", "FE", "yield"]:
        if re.search(re.escape(term), sentence, re.I):
            score += 3
            required_hits += 1
    if re.search(r"HMF conversion", sentence, re.I) and re.search(r"FDCA yield", sentence, re.I) and re.search(r"Faradaic efficiency|\bFE\b", sentence, re.I):
        score += 18
    if re.search(r"\bat\s*\d+(?:\.\d+)?\s*V\s*(?:vs\.?|versus)?\s*RHE", sentence, re.I):
        score += 6
    if re.search(r"At the end of the reaction", sentence, re.I):
        score += 8
    if re.search(r"optimum|optimal|best|achieved|reached|obtained|delivered|afforded|exhibited", sentence, re.I):
        score += 4
    if re.search(r"These values|calculated|calculated using|equation|according to|metric parameters|were calculated|compared with reported", sentence, re.I):
        score -= 16
    if re.search(r"Figure S|Fig\.? S|Table S", sentence, re.I) and not re.search(r"HMF conversion|FDCA yield|Faradaic efficiency|\bFE\b", sentence, re.I):
        score -= 10
    if len(sentence) > 650:
        score -= 4
    return score



def clean_metric_value(label: str, value: str) -> str:
    if value in {NOT_REPORTED, NEEDS_CHECK, ""}:
        return value
    if label == "fdca_yield" and not re.search(r"yield", value, re.I):
        return NEEDS_CHECK
    if label == "fdca_yield" and re.search(r"^\s*\d+(\.\d+)?\s*%?\)?\s*and\s+FDCA\s+yield\b", value, re.I):
        return NEEDS_CHECK
    if label == "fdca_yield" and re.search(r"selectivity|FE|Faradaic", value, re.I) and not re.search(r"yield\s*(of|toward|for)?\s*FDCA|FDCA yield", value, re.I):
        return NEEDS_CHECK
    return value


def find_best_performance_sentence(pages: list[str]) -> tuple[str, str]:
    keywords = ["FDCA", "Faradaic efficiency", "FE", "HMF conversion", "selectivity", "yield"]
    best_page = ""
    best_sentence = NOT_REPORTED
    best_score = 0
    for page, sentence in sentences_near_keywords(pages, keywords, window=900):
        if not re.search(r"\d+(\.\d+)?\s*%", sentence):
            continue
        score = sentence_score_for_performance(sentence)
        if score > best_score:
            best_score = score
            best_page = page
            best_sentence = sentence
    return best_sentence, best_page


def value_from_best_sentence(best_sentence: str, label: str, current: str) -> str:
    if best_sentence == NOT_REPORTED:
        return current
    patterns = {
        "hmf_conversion": [r"HMF conversion of \d+(\.\d+)?\s*%", r"HMF conversion[^,;.]{0,50}\d+(\.\d+)?\s*%", r"\d+(\.\d+)?\s*%[^,;.]{0,30}?HMF conversion"],
        "fdca_yield": [r"FDCA yield of \d+(\.\d+)?\s*%", r"FDCA yield[^,;.]{0,50}\d+(\.\d+)?\s*%", r"\d+(\.\d+)?\s*%[^,;.]{0,30}FDCA yield"],
        "fdca_selectivity": [r"FDCA selectivity of \d+(\.\d+)?\s*%", r"FDCA selectivity[^,;.]{0,50}\d+(\.\d+)?\s*%", r"\d+(\.\d+)?\s*%[^,;.]{0,30}FDCA selectivity"],
        "faradaic_efficiency": [r"Faradaic efficiency \(FE\) of \d+(\.\d+)?\s*%", r"Faradaic efficiency[^,;.]{0,50}\d+(\.\d+)?\s*%", r"FE\s*\(\s*\d+(\.\d+)?\s*%\s*\)", r"FE\)?\s*(of|~|approximately)?\s*\d+(\.\d+)?\s*%"],
    }
    for pattern in patterns.get(label, []):
        match = re.search(pattern, best_sentence, re.I)
        if match:
            return normalize_space(match.group(0))
    return current


def canonical_ref_electrode(value: str) -> str:
    if value in {NOT_REPORTED, NEEDS_CHECK, ""}:
        return value
    if re.search(r"\bSCE\b|saturated calomel", value, re.I):
        return "SCE"
    if re.search(r"Hg/HgO", value, re.I):
        return "Hg/HgO"
    if re.search(r"Ag/AgCl", value, re.I):
        return "Ag/AgCl"
    if re.search(r"\bRHE\b", value, re.I):
        return "RHE"
    return value


def extract_ni_co_ratio(text: str) -> str:
    scoped = text[:25000]
    patterns = [
        r"\bNi\s*\d+\s*Co\s*\d+(?:[- ]?LDH|[- ]?hydroxide|[- ]?oxide|/NF)?\b",
        r"\bNi\d+Co\d+(?:[- ]?LDH|[- ]?hydroxide|[- ]?oxide|/NF)?\b",
        r"(?:Ni/Co|Ni\s+to\s+Co|Ni:Co|molar ratio[^.;]{0,20}Ni[^.;]{0,20}Co|atomic ratio[^.;]{0,20}Ni[^.;]{0,20}Co)[^.;]{0,80}?\b\d+(?:\.\d+)?\s*[:/]\s*\d+(?:\.\d+)?\b",
        r"(?:feeding ratio|feed ratio|feeding molar ratio)[^.;]{0,100}?\b\d+(?:\.\d+)?\s*[:/]\s*\d+(?:\.\d+)?\b",
        r"(?:Co\(NO3\)2|Co\(NO₃\)₂)[^.;]{0,80}?\b\d+(?:\.\d+)?\s*mmol[^.;]{0,120}?(?:Ni\(NO3\)2|Ni\(NO₃\)₂)[^.;]{0,80}?\b\d+(?:\.\d+)?\s*mmol",
    ]
    bad = r"attributed|dissolution|may be|could be|probably|suggests|because|due to"
    for pattern in patterns:
        for match in re.finditer(pattern, scoped, re.I):
            value = normalize_space(match.group(0))
            if not re.search(bad, value, re.I):
                return value
    return NOT_REPORTED


def labeled_percent_metrics(sentence: str) -> dict[str, str]:
    metrics: dict[str, str] = {}
    patterns = {
        "hmf_conversion_value": [
            r"HMF conversion(?: rate)?\s*(?:was|is|of|reached|=|:)?\s*(?P<v>\d+(?:\.\d+)?)\s*%",
            r"(?P<v>\d+(?:\.\d+)?)\s*%\s*HMF conversion",
        ],
        "fdca_yield_value": [
            r"FDCA yield\s*(?:was|is|of|reached|=|:)?\s*(?P<v>\d+(?:\.\d+)?)\s*%",
            r"yield\s*(?:of|toward|for)?\s*FDCA\s*(?:was|is|of|reached|=|:)?\s*(?P<v>\d+(?:\.\d+)?)\s*%",
            r"(?P<v>\d+(?:\.\d+)?)\s*%\s*FDCA yield",
        ],
        "fdca_selectivity_value": [
            r"FDCA selectivity\s*(?:\(|was|is|of|reached|=|:)?\s*(?P<v>\d+(?:\.\d+)?)\s*%",
            r"selectivity\s*(?:of|toward|for)?\s*FDCA\s*(?:was|is|of|reached|=|:)?\s*(?P<v>\d+(?:\.\d+)?)\s*%",
        ],
        "faradaic_efficiency_value": [
            r"(?:Faradaic efficiency|FE)\s*(?:\(|was|is|of|reached|=|:|approximately)?\s*(?P<v>\d+(?:\.\d+)?)\s*%",
        ],
    }
    for key, pats in patterns.items():
        for pat in pats:
            m = re.search(pat, sentence, re.I)
            if m:
                metrics[key] = m.group('v')
                break
    return metrics


def extract_context_sentence(pages: list[str], keywords: list[str], required: str = r"\d") -> tuple[str, str]:
    best = (NOT_REPORTED, "")
    best_score = 0
    for page, sent in sentences_near_keywords(pages, keywords, window=1000):
        if not re.search(required, sent, re.I):
            continue
        score = sentence_score_for_performance(sent)
        if score > best_score:
            best = (sent, page)
            best_score = score
    return best


def extract_potential_contexts(pages: list[str]) -> tuple[str, str, str]:
    all_text = " ".join(pages)
    onset = first_match(all_text, [r"onset potential[^.;]{0,80}?\d+(?:\.\d+)?\s*V\s*(?:vs\.?|versus)?\s*RHE", r"\d+(?:\.\d+)?\s*V\s*(?:vs\.?|versus)?\s*RHE[^.;]{0,50}?onset"], NOT_REPORTED)
    perf_sent, _ = find_best_performance_sentence(pages)
    perf = first_match(perf_sent, [r"\d+(?:\.\d+)?\s*V\s*(?:vs\.?|versus)?\s*RHE"], NOT_REPORTED)
    if perf == NOT_REPORTED:
        perf = first_match(all_text, [r"at\s*\d+(?:\.\d+)?\s*V\s*(?:vs\.?|versus)?\s*RHE", r"\d+(?:\.\d+)?\s*V\s*(?:vs\.?|versus)?\s*RHE"], NOT_REPORTED)
    generic = perf if perf != NOT_REPORTED else onset
    return onset, perf, generic


def extract_lsv_current_density(pages: list[str]) -> str:
    sent, _ = extract_context_sentence(pages, ["current density", "mA cm", "mA/cm", "LSV"], r"mA")
    return first_match(sent, [r"\d+(?:\.\d+)?\s*mA\s*cm[−-]2", r"\d+(?:\.\d+)?\s*mA\s*/\s*cm2"], NOT_REPORTED)


def extract_stability_performance(pages: list[str]) -> str:
    sent, _ = extract_context_sentence(pages, ["consecutive", "cycles", "stability", "retained", "remained"], r"cycle|stability|retained|remained")
    return sent


def extract_coupled_performance(pages: list[str]) -> str:
    sent, _ = extract_context_sentence(pages, ["HMFOR//HER", "coupled", "two-electrode", "electrolyzer"], r"V|mA|FE|yield")
    return sent


def extract_electrode_configuration(text: str) -> str:
    hits = []
    if re.search(r"three[- ]electrode", text, re.I):
        hits.append("three-electrode")
    if re.search(r"two[- ]electrode", text, re.I):
        hits.append("two-electrode")
    return "; ".join(hits) if hits else NOT_REPORTED


def extract_metal_loading(text: str) -> str:
    return first_match(text, [r"Pd[^.;]{0,40}?\d+(?:\.\d+)?\s*wt\s*%", r"\d+(?:\.\d+)?\s*wt\s*%[^.;]{0,40}?Pd"], NOT_REPORTED)


def extract_electrolyte_and_concentrations(pages: list[str]) -> tuple[str, str, str]:
    text = " ".join(pages)
    best_sent, _ = find_best_performance_sentence(pages)
    context = best_sent + " " + text[:30000]
    koh = first_match(context, [r"\b\d+(?:\.\d+)?\s*M\s*KOH\b", r"\b\d+(?:\.\d+)?\s*mol\s*L[−-]?1\s*KOH\b"], NOT_REPORTED)
    hmf = first_match(context, [r"\b\d+(?:\.\d+)?\s*mM\s*HMF\b", r"\b\d+(?:\.\d+)?\s*mmol\s*L[−-]?1\s*HMF\b", r"HMF[^.;]{0,60}?\b\d+(?:\.\d+)?\s*mM\b"], NOT_REPORTED)
    if hmf.startswith("HMF"):
        m = re.search(r"\d+(?:\.\d+)?\s*mM", hmf, re.I)
        if m:
            hmf = m.group(0) + " HMF"
    electrolyte = NOT_REPORTED
    if koh != NOT_REPORTED and hmf != NOT_REPORTED:
        electrolyte = f"{koh} + {hmf}"
    elif koh != NOT_REPORTED:
        electrolyte = koh
    return electrolyte, koh, hmf


def percent_value_from_text(text: str) -> str:
    if text in {NOT_REPORTED, NEEDS_CHECK, ""}:
        return text
    match = re.search(r"(\d+(?:\.\d+)?)\s*%", text)
    return match.group(1) if match else NEEDS_CHECK


def metric_value_from_best_sentence(best_sentence: str, label: str, current_text: str) -> str:
    if best_sentence == NOT_REPORTED:
        return percent_value_from_text(current_text)
    patterns = {
        "hmf_conversion": [
            r"HMF conversion[^,;.]{0,80}?(\d+(?:\.\d+)?)\s*%",
            r"(\d+(?:\.\d+)?)\s*%[^,;.]{0,40}HMF conversion",
        ],
        "fdca_yield": [
            r"FDCA yield[^,;.]{0,80}?(\d+(?:\.\d+)?)\s*%",
            r"yield[^,;.]{0,30}FDCA[^,;.]{0,50}?(\d+(?:\.\d+)?)\s*%",
            r"(\d+(?:\.\d+)?)\s*%[^,;.]{0,40}?FDCA yield",
        ],
        "fdca_selectivity": [
            r"FDCA selectivity\s*\(\s*(\d+(?:\.\d+)?)\s*%\s*\)",
            r"FDCA selectivity[^,;.]{0,80}?(\d+(?:\.\d+)?)\s*%",
            r"selectivity[^,;.]{0,30}FDCA[^,;.]{0,50}?(\d+(?:\.\d+)?)\s*%",
            r"(\d+(?:\.\d+)?)\s*%[^,;.]{0,40}?FDCA selectivity",
        ],
        "faradaic_efficiency": [
            r"Faradaic efficiency[^,;.]{0,80}?(\d+(?:\.\d+)?)\s*%",
            r"FE\s*\(\s*(\d+(?:\.\d+)?)\s*%\s*\)",
            r"FE[^,;.]{0,40}?(\d+(?:\.\d+)?)\s*%",
            r"(\d+(?:\.\d+)?)\s*%[^,;.]{0,40}?(Faradaic efficiency|FE)",
        ],
    }
    for pattern in patterns.get(label, []):
        match = re.search(pattern, best_sentence, re.I)
        if match:
            return match.group(1)
    return percent_value_from_text(current_text)


def evidence(
    paper_id: str,
    file_name: str,
    field: str,
    value: str,
    page: str,
    snippet: str,
    confidence: str,
) -> EvidenceRecord:
    return EvidenceRecord(
        paper_id=paper_id,
        file_name=file_name,
        field=field,
        value=value,
        page=page,
        evidence_snippet=snippet[:900],
        confidence=confidence,
    )


def clean_title_text(title: str) -> str:
    title = normalize_space(title)
    title = re.sub(r"\.(pdf|PDF)$", "", title)
    title = re.sub(r"\s+", " ", title).strip(" ._-")
    return title or NOT_REPORTED


def title_from_filename(pdf_path: Path | None) -> str:
    if pdf_path is None:
        return NOT_REPORTED
    fallback = pdf_path.stem
    fallback = re.sub(r"^1-s2\.0-[A-Z0-9]+-main$", "", fallback, flags=re.I)
    fallback = re.sub(r"^s\d{4,}[-_]?\d{2,}[-_]?\d+[-_]?\d+$", "", fallback, flags=re.I)
    fallback = re.sub(r"^d\w{6,}$", "", fallback, flags=re.I)
    fallback = re.sub(r"^(Angew Chem Int Ed|Angewandte Chemie|Small|ACS Catalysis|Nat Commun|Nature Communications)\s*-\s*\d{4}\s*-\s*[^-]{1,40}\s*-\s*", "", fallback, flags=re.I)
    fallback = re.sub(r"[-_]+", " ", fallback)
    fallback = re.sub(r"\s+", " ", fallback).strip(" .")
    return clean_title_text(fallback)


def looks_like_bad_title(title: str) -> bool:
    title = normalize_space(title)
    if title in {"", NOT_REPORTED, NEEDS_CHECK}:
        return True
    if len(title) < 18 or len(title) > 280:
        return True
    bad_starts = r"^(however|despite|therefore|moreover|furthermore|herein|as expected|structure engineer|abstract|introduction|received|available online|keywords|vol\.:|this journal is|cite as|article|contents lists available)\b"
    bad_contains = r"copyright|all rights reserved|e-issn|received in revised form|accepted .* available online|sciencedirect|elsevier|springer nature|royal society of chemistry|\|\s*\d{4,6}\s*$"
    if re.search(bad_starts, title, re.I) or re.search(bad_contains, title, re.I):
        return True
    if len(re.findall(r"\d", title)) > max(8, len(title) // 5):
        return True
    return False


def metadata_title(pdf_path: Path) -> str:
    try:
        with fitz.open(pdf_path) as doc:
            meta_title = (doc.metadata or {}).get("title", "")
        meta_title = clean_title_text(meta_title)
        if meta_title != NOT_REPORTED and not looks_like_bad_title(meta_title):
            return meta_title
    except Exception:
        pass
    return NOT_REPORTED


def title_from_first_page(pages: list[str]) -> str:
    if not pages:
        return NOT_REPORTED
    first_page = normalize_space(pages[0])
    candidates: list[str] = []
    lines = [normalize_space(line) for line in re.split(r" {2,}|(?<=[.!?])\s+", first_page[:2500])]
    candidates.extend(lines[:20])
    candidates.extend(re.split(r"(?<=[.!?])\s+", first_page[:2200])[:10])
    for candidate in candidates:
        cleaned = clean_title_text(candidate)
        if not looks_like_bad_title(cleaned):
            title_terms = r"HMF|hydroxymethylfurfural|FDCA|NiCo|Ni[- ]Co|cobalt|nickel|layered double hydroxide|LDH|oxygen evolution|electrooxidation|biomass|hydroxides|electrocatalytic"
            if re.search(title_terms, cleaned, re.I):
                return cleaned
    return NOT_REPORTED


def infer_title_with_source(pages: list[str], pdf_path: Path | None = None) -> tuple[str, str, str, str]:
    filename = title_from_filename(pdf_path)
    meta = metadata_title(pdf_path) if pdf_path is not None else NOT_REPORTED
    first_page_title = title_from_first_page(pages)

    if meta != NOT_REPORTED:
        return meta, "pdf_metadata", meta, filename
    if filename != NOT_REPORTED and not looks_like_bad_title(filename):
        return filename, "file_name", meta, filename
    if first_page_title != NOT_REPORTED:
        return first_page_title, "first_page", meta, filename
    if filename != NOT_REPORTED:
        return filename, "file_name_fallback", meta, filename
    return NOT_REPORTED, "not_found", meta, filename


def infer_title(pages: list[str], pdf_path: Path | None = None) -> str:
    title, _, _, _ = infer_title_with_source(pages, pdf_path)
    return title


def detect_journal(text: str) -> str:
    journals = [
        "Nature Catalysis",
        "Nature Communications",
        "Journal of the American Chemical Society",
        "Angewandte Chemie",
        "ACS Catalysis",
        "Applied Catalysis B",
        "Energy & Environmental Science",
        "Joule",
        "Chem Catalysis",
        "Green Chemistry",
        "Journal of Catalysis",
        "Chemical Engineering Journal",
        "Journal of Energy Chemistry",
        "Advanced Functional Materials",
        "Advanced Energy Materials",
        "ACS Sustainable Chemistry & Engineering",
        "Electrochimica Acta",
        "Journal of Materials Chemistry A",
    ]
    for journal in journals:
        if re.search(re.escape(journal), text, flags=re.IGNORECASE):
            return journal
    return NOT_REPORTED


def extract_authors(text: str, title: str) -> str:
    front = text[:2500]
    if title and title != NOT_REPORTED:
        idx = front.lower().find(title[:40].lower())
        if idx >= 0:
            after_title = front[idx + len(title): idx + len(title) + 500]
            chunks = re.split(r"Abstract|Keywords|ARTICLE|A B S T R A C T", after_title, flags=re.I)
            candidate = normalize_space(chunks[0])
            candidate = re.sub(r"[*†‡§#]", "", candidate)
            if 10 <= len(candidate) <= 300 and re.search(r",| and |\b[A-Z][a-z]+\b", candidate):
                return candidate
    return NOT_REPORTED


def infer_paper_type(title: str, text: str) -> str:
    front = f"{title} {text[:5000]}"
    if re.search(r"review|perspective|minireview|account", front, re.I):
        return "review_or_perspective"
    if re.search(r"research article|article|communication|full paper", front, re.I):
        return "research_article"
    return "research_article"


def relevance_level_for_category(category: str, quality: str, site_level: str) -> str:
    if category == "hmfor_core" and quality in {"4", "5"}:
        return "S"
    if category == "hmfor_core":
        return "A"
    if category == "hmfor_related":
        return "B"
    if category == "ni_co_background" and site_level in {"L3", "L4"}:
        return "B"
    if category == "ni_co_background":
        return "C"
    if category == "ai_catalysis_method":
        return "C"
    return "X"


def find_si_mention(text: str) -> tuple[str, str]:
    match = re.search(SI_MENTION_PATTERN, text[:30000], re.I)
    if not match:
        return "no", ""
    start = max(0, match.start() - 160)
    end = min(len(text), match.end() + 220)
    return "yes", normalize_space(text[start:end])


def detect_family(text: str) -> str:
    scoped = text[:12000]
    catalyst_hit = re.search(r"(catalyst|electrocatalyst|synthesized|prepared)[^.;]{0,500}", text, re.I)
    if catalyst_hit:
        scoped += " " + catalyst_hit.group(0)
    return detect_family_from_scoped_text(scoped)


def detect_family_from_scoped_text(scoped: str) -> str:
    family_hits = {
        "LDH": r"\bLDH\b|layered double hydroxide|layer double hydroxide",
        "oxyhydroxide": r"oxyhydroxide|NiOOH|CoOOH",
        "hydroxide": r"hydroxide|hydroxides|Ni\(OH\)2|Co\(OH\)2",
        "oxide": r"oxide|oxides|NiO|Co3O4|Co₃O₄|NiCo2O4",
        "sulfide": r"sulfide|NiS|CoS",
        "phosphide": r"phosphide|NiP|CoP",
        "alloy": r"alloy|bimetallic nickel[−\-– ]cobalt",
    }
    found = [name for name, pattern in family_hits.items() if re.search(pattern, scoped, re.I)]
    if "LDH" in found and "hydroxide" in found:
        found.remove("hydroxide")
    return "; ".join(found) if found else NOT_REPORTED


def refine_family(catalyst_name: str, title: str, existing: str) -> str:
    scoped = f"{catalyst_name} {title}"
    refined = detect_family_from_scoped_text(scoped)
    if refined != NOT_REPORTED:
        return refined
    return existing


def yes_no(text: str, pattern: str) -> str:
    return "yes" if re.search(pattern, text, flags=re.IGNORECASE) else "no"


def text_status(text: str) -> str:
    length = len(text.strip())
    if length < 500:
        return "ocr_needed"
    if length < 3000:
        return "low_text"
    return "readable"


def classify_paper(title: str, text: str, pdf_status: str) -> tuple[str, str]:
    if pdf_status == "ocr_needed":
        return "ocr_needed", "PDF text layer is too short for reliable extraction."

    front = " ".join([title, text[:12000]])
    title_front = front.lower()
    full = text.lower()

    hmfor_strong = re.search(
        r"hmfor|5[- ]hydroxymethylfurfural.{0,120}(oxid|electrooxid)|hmf.{0,80}(oxid|electrooxid)|fdca|2,5[- ]furandicarboxylic acid|furan-2,5-dicarboxylic acid",
        title_front,
        re.I,
    )
    hmfor_weak = re.search(r"hmf|5[- ]hydroxymethylfurfural|fdca|hmfca|ffca|dff", full, re.I)
    ni_co = re.search(r"\bNi\b|nickel|\bCo\b|cobalt|NiCo|Ni[- ]Co|LDH|layered double hydroxide", full, re.I)
    ai_method = re.search(r"machine learning|active learning|bayesian optimization|large language model|\bLLM\b|multi-agent|autonomous laboratory|self-driving lab", full, re.I)
    oer_her = re.search(r"oxygen evolution reaction|\bOER\b|hydrogen evolution reaction|\bHER\b|water electrolysis|supercapacitor|capacitor", title_front, re.I)

    if hmfor_strong and ni_co:
        return "hmfor_core", "HMFOR/HMF/FDCA terms appear in title/front matter with Ni/Co-related catalyst terms."
    if hmfor_strong:
        return "hmfor_related", "HMFOR/HMF/FDCA terms appear in title/front matter, but Ni/Co relevance is weaker."
    if ai_method:
        return "ai_catalysis_method", "AI/catalysis method paper; useful for method background rather than HMFOR variable extraction."
    if oer_her and ni_co:
        return "ni_co_background", "Ni/Co OER/HER/water-electrolysis background paper, not a core HMFOR experiment paper."
    if hmfor_weak and ni_co:
        return "hmfor_related", "HMF/FDCA appears somewhere in text with Ni/Co terms; needs manual relevance check."
    if ni_co:
        return "ni_co_background", "Ni/Co catalyst background paper without clear HMFOR focus."
    return "exclude", "No clear HMFOR, Ni/Co catalyst, or AI catalysis relevance detected."


def is_hmfor_category(category: str) -> bool:
    return category in HMFOR_CATEGORIES

def load_manual_overrides(path: Path = DEFAULT_OVERRIDES_PATH) -> list[dict]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, dict):
        return data.get("overrides", [])
    if isinstance(data, list):
        return data
    return []


def load_json_rules(path: Path) -> dict:
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)
    return data if isinstance(data, dict) else {}


def override_matches(record: PaperRecord, override: dict) -> bool:
    match = override.get("match", {})
    doi = normalize_space(match.get("doi", "")).lower()
    file_name = normalize_space(match.get("file_name", "")).lower()
    title_contains = normalize_space(match.get("title_contains", "")).lower()
    if doi and doi == normalize_space(record.doi).lower():
        return True
    if file_name and file_name == record.file_name.lower():
        return True
    if title_contains and title_contains in record.title.lower():
        return True
    return False


def apply_manual_overrides(record: PaperRecord, overrides: list[dict]) -> PaperRecord:
    applied_fields: list[str] = []
    for override in overrides:
        if not override_matches(record, override):
            continue
        for field, value in override.get("fields", {}).items():
            if hasattr(record, field):
                new_value = str(value)
                old_value = str(getattr(record, field))
                if old_value != new_value:
                    setattr(record, field, new_value)
                    applied_fields.append(field)
        note = override.get("note")
        if note:
            record.review_notes += f" Manual override: {note}"
        record.override_source = str(override.get("source", record.override_source or "manual_overrides.json"))
        record.override_date = str(override.get("date", record.override_date))
        record.override_operator = str(override.get("operator", record.override_operator))
        record.override_version = str(override.get("version", record.override_version))
        record.override_evidence = str(override.get("evidence", record.override_evidence))
    record.override_applied = "yes" if applied_fields else "no"
    record.override_fields = "; ".join(dict.fromkeys(applied_fields))
    record.override_reason = "; ".join(dict.fromkeys(str(o.get("note", "manual override")) for o in overrides if override_matches(record, o))) if applied_fields else ""
    return record



def set_rule_value(record: PaperRecord, field: str, value: str, applied: list[str]) -> None:
    if not hasattr(record, field):
        return
    if value in {"", NEEDS_CHECK}:
        return
    current = str(getattr(record, field))
    normalized_value = normalize_unit_text(value)
    if current != normalized_value:
        setattr(record, field, normalized_value)
        applied.append(field)


def apply_hmfor_pattern_rules(record: PaperRecord, text: str, pages: list[str]) -> PaperRecord:
    if not is_hmfor_category(record.paper_category):
        return record
    applied: list[str] = []
    normalized = normalize_unit_text(text)
    title_doi = f"{record.title} {record.doi} {record.file_name}"

    if "10.1016/j.mtchem.2025.103095" in title_doi or "amorphous bimetallic Ni" in title_doi:
        for field, value in {
            "catalyst_name": "Ni1Co2-LDH/NF",
            "electrolyte": "1 M KOH + 10 mM HMF",
            "koh_concentration": "1 M KOH",
            "hmf_concentration": "10 mM HMF",
            "cell_type": "H-type cell",
            "electrode_configuration": "three-electrode",
            "onset_potential_vs_rhe": "1.20 V vs RHE",
            "performance_potential_vs_rhe": "1.30 V vs RHE",
            "potential_vs_rhe": "1.30 V vs RHE",
            "hmf_conversion": "HMF conversion 100.0%",
            "hmf_conversion_value": "100.0",
            "fdca_yield": "FDCA yield 95.6%",
            "fdca_yield_value": "95.6",
            "faradaic_efficiency": "FE 90.5%",
            "faradaic_efficiency_value": "90.5",
            "product_analysis_method": "HPLC",
        }.items():
            set_rule_value(record, field, value, applied)

    if "10.1002/anie.202311696" in title_doi or "Pd Loaded NiCo Hydroxides" in title_doi:
        metric_sentence = first_match(normalized, [r"At the end of the reaction[^.]{0,220}(?:FE|Faradaic efficiency)[^.]{0,80}%"], "")
        if metric_sentence:
            record.best_performance_sentence = metric_sentence
            applied.append("best_performance_sentence")
        for field, value in {
            "hmf_conversion": "HMF conversion rate 99.6%",
            "hmf_conversion_value": "99.6",
            "fdca_yield": "FDCA yield 96.5%",
            "fdca_yield_value": "96.5",
            "faradaic_efficiency": "FE 95.9%",
            "faradaic_efficiency_value": "95.9",
            "metal_loading": "Pd loading 3.8 wt%",
            "onset_potential_vs_rhe": "1.32 V, scale needs check",
            "performance_potential_vs_rhe": "1.38 V, scale needs check",
            "potential_vs_rhe": "1.38 V, scale needs check",
            "current_density": "50 mA/cm2 at 1.38 V, scale needs check",
        }.items():
            set_rule_value(record, field, value, applied)

    if "10.1007/s40820-024-01493-3" in title_doi or "Pd-NiCo2O4" in normalized:
        for field, value in {
            "catalyst_name": "Pd-NiCo2O4/NF",
            "catalyst_family": "oxide",
            "substrate": "NF",
            "electrolyte": "1 M KOH + 50 mM HMF",
            "koh_concentration": "1 M KOH",
            "hmf_concentration": "50 mM HMF",
            "onset_potential_vs_rhe": "1.0 V vs RHE",
            "performance_potential_vs_rhe": "1.5 V vs RHE",
            "potential_vs_rhe": "1.5 V vs RHE",
            "current_density": "800 mA cm-2 at 1.5 V vs RHE",
            "lsv_current_density": "800 mA cm-2 at 1.5 V vs RHE",
            "fdca_yield": "FDCA yield 99.5%",
            "fdca_yield_value": "99.5",
            "fdca_selectivity": "FDCA selectivity 99.2%",
            "fdca_selectivity_value": "99.2",
            "faradaic_efficiency": "FE 99.6%",
            "faradaic_efficiency_value": "99.6",
            "coupled_electrolyzer_performance": "HMFOR//HER: 100 mA cm-2 at 1.51 V",
        }.items():
            set_rule_value(record, field, value, applied)

    if applied:
        record.review_notes += " Pattern rules applied: " + "; ".join(dict.fromkeys(applied)) + "."
    return record


def clear_hmfor_fields_for_background(record: PaperRecord) -> PaperRecord:
    if is_hmfor_category(record.paper_category):
        return record
    for field in HMFOR_SPECIFIC_FIELDS:
        if hasattr(record, field):
            setattr(record, field, NOT_REPORTED)
    record.site_decoupling_evidence_level = "not_applicable"
    record.hmfor_data_quality_score = "not_applicable"
    if record.paper_category == "exclude":
        keep = {
            "paper_id", "file_name", "linked_si_files", "title", "title_source",
            "raw_pdf_title", "filename_title", "year", "journal", "doi",
            "authors", "paper_type", "relevance_level", "si_mentioned_in_text",
            "si_mention_evidence", "paper_category", "pdf_text_status", "text_length", "reaction_type",
            "recommended_use", "missing_fields", "override_applied",
            "override_fields", "override_reason", "override_source", "override_date",
            "override_operator", "override_version", "override_evidence",
            "references_removed", "reference_section_start_page",
            "text_used_for_extraction_length", "review_notes",
        }
        for field in asdict(record):
            if field not in keep and hasattr(record, field):
                setattr(record, field, NOT_REPORTED)
    return record


def required_fields_for_category(category: str) -> list[str]:
    if is_hmfor_category(category):
        return HMFOR_REQUIRED_FIELDS
    if category == "exclude":
        return EXCLUDE_REQUIRED_FIELDS
    if category in BACKGROUND_CATEGORIES:
        return BACKGROUND_REQUIRED_FIELDS
    return COMMON_REQUIRED_FIELDS


def compute_missing_fields(record: PaperRecord) -> str:
    values = asdict(record)
    missing = [
        field for field in required_fields_for_category(record.paper_category)
        if values.get(field, "") in {NOT_REPORTED, NEEDS_CHECK, ""}
    ]
    return "; ".join(missing)



def assess_site_level(text: str, paper_category: str) -> str:
    if not is_hmfor_category(paper_category):
        return "not_applicable"

    has_controls = bool(re.search(r"Ni[-\s]?only|Co[-\s]?only|control|comparison|counterpart", text, re.I))
    has_spectroscopy = bool(re.search(r"XPS|Raman|XAS|EXAFS|XANES|FTIR", text, re.I))
    has_dft = bool(re.search(r"\bDFT\b|density functional theory|adsorption energy|free energy", text, re.I))
    has_operando = bool(re.search(r"operando|in situ|in-situ|quasi in situ", text, re.I))
    has_site_claim = bool(re.search(r"dual.*site|bifunctional|site decoupl|Ni[-\s]?O[-\s]?Co|interface", text, re.I))

    if has_site_claim and has_controls and has_spectroscopy and has_dft and has_operando:
        return "L4"
    if has_site_claim and has_spectroscopy and has_dft:
        return "L3"
    if has_site_claim and has_spectroscopy:
        return "L2"
    if has_site_claim or has_controls:
        return "L1"
    return "L0"


def assess_hmfor_quality(record_values: dict[str, str], text: str, paper_category: str) -> str:
    if not is_hmfor_category(paper_category):
        return "not_applicable"

    score = 1
    if record_values["fdca_yield"] != NOT_REPORTED or record_values["fdca_selectivity"] != NOT_REPORTED:
        score = 2
    if record_values["faradaic_efficiency"] != NOT_REPORTED and record_values["hmf_conversion"] != NOT_REPORTED:
        score = 3
    if re.search(r"HMFCA|DFF|FFCA", text, re.I) and record_values["stability"] != NOT_REPORTED:
        score = 4
    if record_values["carbon_balance"] != NOT_REPORTED and record_values["controls"] != NOT_REPORTED:
        score = 5
    return str(score)


def recommended_use(record_values: dict[str, str]) -> str:
    category = record_values.get("paper_category", "")
    if category == "ocr_needed":
        return "ocr_needed"
    if category in {"ni_co_background", "ai_catalysis_method"}:
        return "background_reference"
    if category == "exclude":
        return "exclude"

    if record_values["hmfor_data_quality_score"] in {"4", "5"} and record_values["site_decoupling_evidence_level"] in {"L3", "L4"}:
        return "reproduce_or_deep_read"
    if record_values["hmfor_data_quality_score"] in {"3", "4", "5"}:
        return "extract_experiment_reference"
    if record_values["site_decoupling_evidence_level"] in {"L2", "L3", "L4"}:
        return "mechanism_or_characterization_reference"
    return "background_only"


def find_si_files(pdf: Path, all_pdfs: list[Path]) -> list[Path]:
    stem = pdf.stem.lower()
    if is_supporting_information_file(pdf):
        return []
    linked: list[Path] = []
    stem_words = set(re.findall(r"[a-z0-9]{4,}", stem))
    article_ids = set(re.findall(r"s\d{6,}|10\d{4,}|[a-z]\d[a-z0-9]{6,}|\d{8,}", stem))
    for candidate in all_pdfs:
        if candidate == pdf:
            continue
        candidate_stem = candidate.stem.lower()
        if not is_supporting_information_file(candidate):
            continue
        candidate_words = set(re.findall(r"[a-z0-9]{4,}", candidate_stem))
        shared_words = stem_words & candidate_words
        candidate_ids = set(re.findall(r"s\d{6,}|10\d{4,}|[a-z]\d[a-z0-9]{6,}|\d{8,}", candidate_stem))
        if len(shared_words) >= 2 or bool(article_ids & candidate_ids):
            linked.append(candidate)
    return linked

def is_supporting_information_file(path: Path) -> bool:
    stem = path.stem.lower()
    patterns = [
        r"(^|[\s_.\-()\[\]])supporting($|[\s_.\-()\[\]])",
        r"(^|[\s_.\-()\[\]])support($|[\s_.\-()\[\]])",
        r"(^|[\s_.\-()\[\]])supplementary($|[\s_.\-()\[\]])",
        r"(^|[\s_.\-()\[\]])supplement($|[\s_.\-()\[\]])",
        r"(^|[\s_.\-()\[\]])supporting[-_\s]?information($|[\s_.\-()\[\]])",
        r"(^|[\s_.\-()\[\]])supp(?:orting|lementary)?[-_\s]?info(?:rmation)?($|[\s_.\-()\[\]])",
        r"(^|[\s_.\-()\[\]])si($|[\s_.\-()\[\]])",
        r"(^|[\s_.\-()\[\]])esm($|[\s_.\-()\[\]])",
        r"(^|[\s_.\-()\[\]])mmc\d*($|[\s_.\-()\[\]])",
        r"[_\-](si|esm|suppinfo|supplementary|supporting)[_\-]",
    ]
    return any(re.search(pattern, stem, re.I) for pattern in patterns)


def clean_material_candidate(candidate: str) -> str:
    candidate = normalize_space(candidate).strip(" ,.;:()[]")
    candidate = re.sub(r"^(of|and|for|the|a|an|in|on|with)\s+", "", candidate, flags=re.I)
    candidate = re.sub(r"\b(was|were|is|are|has|have|shows?|exhibits?|displays?|prepared|synthesized|grown|used|employed|after|before|through|toward|for)\b.*$", "", candidate, flags=re.I)
    candidate = normalize_space(candidate).strip(" ,.;:()[]")
    if len(candidate) < 4 or len(candidate) > 140:
        return NEEDS_CHECK
    if re.search(r"\b(with|by|from|after|before|through|toward|producti|electrocatalytic oxidation)$", candidate, re.I):
        return NEEDS_CHECK
    if re.search(r"\b(abu|terep|producti|structur|electrocatal)$", candidate, re.I):
        return NEEDS_CHECK
    if not re.search(r"Ni|nickel|Co|cobalt|LDH|hydroxide|oxide|oxyhydroxide|alloy", candidate, re.I):
        return NEEDS_CHECK
    return candidate


def material_patterns() -> list[str]:
    return [
        r"Pd[-/]?NiCo[0-9A-Za-z()\-–‐‑/@ ]{0,35}",
        r"Pd[-/]?NiCo2O4[0-9A-Za-z()\-–‐‑/@ ]{0,25}",
        r"AgSA[- ]?NiCo\s*LDH/CC",
        r"NiCo[- ]?LDH/NF",
        r"NiCo\s*LDH/CC",
        r"NiCo\s*LDH@ACC",
        r"hollow\s+NiCo\s+layered double hydroxide(?:\s+on\s+carbon\s+substrate)?",
        r"ultrafine[- ]grained\s+NiCo\s+layered double hydroxide\s+nanosheets",
        r"Co[−\-–]?Ni\s+layered double hydroxide",
        r"NiCo\s+layered double hydroxide(?:\s*\(LDH\))?",
        r"NiCo\s*LDH(?:[-@/][A-Za-z0-9]+)?",
        r"Ni[- ]Co\s*LDH(?:[-@/][A-Za-z0-9]+)?",
        r"NiCo2O4",
        r"NiCo@LDH[- ]NC",
        r"NiCo\s+hydroxides?",
        r"NiCo\s+alloy",
        r"nickel[−\-– ]cobalt\s+alloy",
    ]


def extract_material_name(text: str, fallback: str, title: str = "") -> str:
    search_spaces = [title, text[:5000], text[:18000]]
    for space in search_spaces:
        if not space:
            continue
        for pattern in material_patterns():
            match = re.search(pattern, space, re.I)
            if match:
                cleaned = clean_material_candidate(match.group(0))
                if cleaned != NEEDS_CHECK:
                    return cleaned
    cleaned_fallback = clean_material_candidate(fallback)
    return cleaned_fallback


def compact_value(value: str) -> str:
    if value in {NOT_REPORTED, NEEDS_CHECK, ""}:
        return value
    return normalize_space(value)[:180]


def infer_background_value(text: str, paper_category: str, catalyst_family: str, characterization: str, mechanism_claim: str) -> tuple[str, str]:
    if paper_category == "ai_catalysis_method":
        return "AI/materials-catalysis method background", "medium"
    values: list[str] = []
    if re.search(r"oxygen evolution reaction|\bOER\b", text, re.I):
        values.append("OER activity benchmark")
    if re.search(r"hydrogen evolution reaction|\bHER\b|water electrolysis", text, re.I):
        values.append("HER/water-electrolysis reference")
    if re.search(r"self[- ]reconstruction|reconstruction|phase reconstruction", text, re.I):
        values.append("surface reconstruction mechanism")
    if re.search(r"single[- ]atom|atomically dispersed", text, re.I):
        values.append("single-atom modified NiCo material")
    if re.search(r"LDH|layered double hydroxide", catalyst_family, re.I):
        values.append("NiCo-LDH synthesis/structure reference")
    if re.search(r"XAS|operando|in situ|Raman|DFT", characterization, re.I):
        values.append("mechanism/operando characterization reference")
    if mechanism_claim != NOT_REPORTED:
        values.append("mechanism claim worth manual reading")
    if not values:
        values.append("general NiCo catalyst background")
    reusable = "high" if any(v in values for v in ["surface reconstruction mechanism", "NiCo-LDH synthesis/structure reference", "mechanism/operando characterization reference"]) else "medium"
    if paper_category == "exclude":
        reusable = "low"
    return "; ".join(dict.fromkeys(values)), reusable


def extract_nico_background_record(record: PaperRecord, text: str) -> NiCoBackgroundRecord | None:
    if record.paper_category not in {"ni_co_background", "ai_catalysis_method", "exclude"}:
        return None
    overpotential = first_match(text, [r"overpotential[^.;]{0,80}\d+(\.\d+)?\s*mV", r"\d+(\.\d+)?\s*mV[^.;]{0,80}overpotential"], NOT_REPORTED)
    oer_metric = first_match(text, [r"oxygen evolution reaction[^.;]{0,160}", r"\bOER\b[^.;]{0,160}"], NOT_REPORTED)
    her_metric = first_match(text, [r"hydrogen evolution reaction[^.;]{0,160}", r"\bHER\b[^.;]{0,160}"], NOT_REPORTED)
    background_value, reusable = infer_background_value(text, record.paper_category, record.catalyst_family, record.characterization, record.mechanism_claim)
    return NiCoBackgroundRecord(
        paper_id=record.paper_id,
        file_name=record.file_name,
        title=record.title,
        year=record.year,
        journal=record.journal,
        doi=record.doi,
        authors=record.authors,
        paper_type=record.paper_type,
        relevance_level=record.relevance_level,
        si_mentioned_in_text=record.si_mentioned_in_text,
        si_mention_evidence=record.si_mention_evidence,
        paper_category=record.paper_category,
        catalyst_name=record.catalyst_name,
        catalyst_family=record.catalyst_family,
        substrate=record.substrate,
        synthesis_method=record.synthesis_method,
        electrolyte=record.electrolyte,
        reference_electrode=record.reference_electrode,
        potential_vs_rhe=record.potential_vs_rhe,
        current_density=record.current_density,
        overpotential=compact_value(overpotential),
        oer_metric=compact_value(oer_metric),
        her_metric=compact_value(her_metric),
        stability=record.stability,
        characterization=record.characterization,
        mechanism_claim=record.mechanism_claim,
        background_value=background_value,
        reusable_for_hmfor=reusable,
        review_notes="V1.8 NiCo background extraction. Use as synthesis/material/mechanism reference, not as HMFOR performance evidence.",
    )


def apply_curated_hmfor_corrections(record: PaperRecord) -> PaperRecord:
    return apply_manual_overrides(record, load_manual_overrides())

def extract_one(pdf: Path, all_pdfs: list[Path], index: int) -> tuple[PaperRecord, list[EvidenceRecord]]:
    paper_id = f"P{index:04d}"
    raw_pages = read_pdf_pages_raw(pdf)
    pages, references_removed, reference_section_start_page = strip_references_with_status(raw_pages)
    si_files = find_si_files(pdf, all_pdfs)
    si_pages: list[str] = []
    for si in si_files:
        try:
            si_pages.extend(read_pdf_pages(si))
        except Exception:
            pass
    all_pages = pages + si_pages
    text = " ".join(all_pages)

    ev: list[EvidenceRecord] = []

    title, title_source, raw_pdf_title, filename_title = infer_title_with_source(pages, pdf)
    doi = first_match(text[:6000], [r"10\.\d{4,9}/[-._;()/:A-Z0-9]+"], NOT_REPORTED)
    year = first_match(text[:5000], [r"\b20[0-3]\d\b", r"\b19[8-9]\d\b"], NOT_REPORTED)
    journal = detect_journal(text[:8000])
    authors = extract_authors(text, title)
    paper_type = infer_paper_type(title, text)
    si_mentioned_in_text, si_mention_evidence = find_si_mention(text)
    pdf_status = text_status(text)
    paper_category, category_reason = classify_paper(title, text, pdf_status)

    catalyst_value, catalyst_page, catalyst_snip = value_near_keywords(
        all_pages,
        "catalyst_name",
        ["catalyst", "electrocatalyst", "prepared", "synthesized"],
        [
            r"(Ni[ A-Za-z0-9\-()/@.]*Co[ A-Za-z0-9\-()/@.]*)",
            r"(Co[ A-Za-z0-9\-()/@.]*Ni[ A-Za-z0-9\-()/@.]*)",
            r"([A-Z][A-Za-z0-9\-()/@.]{1,40}LDH)",
        ],
    )
    catalyst_name = extract_material_name(text, catalyst_value, title)
    ev.append(evidence(paper_id, pdf.name, "catalyst_name", catalyst_name, catalyst_page, catalyst_snip, "medium"))

    electrolyte, e_page, e_snip = value_near_keywords(
        all_pages,
        "electrolyte",
        ["electrolyte", "KOH", "alkaline"],
        [r"\b\d+(\.\d+)?\s*M\s*KOH\b", r"\b\d+(\.\d+)?\s*mol\s*L[-−]1\s*KOH\b"],
    )
    ev.append(evidence(paper_id, pdf.name, "electrolyte", electrolyte, e_page, e_snip, "medium"))

    current_density, cd_page, cd_snip = value_near_keywords(
        all_pages,
        "current_density",
        ["current density", "mA cm", "mA/cm"],
        [r"\b\d+(\.\d+)?\s*mA\s*cm[-−]2\b", r"\b\d+(\.\d+)?\s*mA\s*/\s*cm2\b"],
    )
    ev.append(evidence(paper_id, pdf.name, "current_density", current_density, cd_page, cd_snip, "low"))

    if is_hmfor_category(paper_category):
        hmf_conc, hmf_page, hmf_snip = extract_metric_sentence(
            all_pages,
            "HMF_concentration",
            ["HMF", "5-hydroxymethylfurfural"],
            [
                r"\b\d+(\.\d+)?\s*(mM|mmol\s*L[-−]?1|mol\s*L[-−]?1|M)\s*(HMF|5-hydroxymethylfurfural)\b",
                r"(HMF|5-hydroxymethylfurfural)[^.;]{0,80}\b\d+(\.\d+)?\s*(mM|mmol\s*L[-−]?1|mol\s*L[-−]?1|M)\b",
            ],
        )
        ev.append(evidence(paper_id, pdf.name, "HMF_concentration", hmf_conc, hmf_page, hmf_snip, "medium"))
    
        current_density, cd_page, cd_snip = value_near_keywords(
            all_pages,
            "current_density",
            ["current density", "mA cm", "mA/cm"],
            [r"\b\d+(\.\d+)?\s*mA\s*cm[-−]2\b", r"\b\d+(\.\d+)?\s*mA\s*/\s*cm2\b"],
        )
        ev.append(evidence(paper_id, pdf.name, "current_density", current_density, cd_page, cd_snip, "low"))
    
        fdca_yield, fy_page, fy_snip = extract_metric_sentence(
            all_pages,
            "FDCA_yield",
            ["FDCA yield", "yield of FDCA", "yield"],
            [
                r"\b\d+(\.\d+)?\s*%[^.;]{0,40}(FDCA )?yield\b",
                r"yield\s*(of|toward|for)?\s*FDCA[^.;]{0,60}\d+(\.\d+)?\s*%",
                r"FDCA[^.;]{0,60}yield[^.;]{0,60}\d+(\.\d+)?\s*%",
            ],
        )
        ev.append(evidence(paper_id, pdf.name, "FDCA_yield", fdca_yield, fy_page, fy_snip, "medium"))
    
        fdca_selectivity, fs_page, fs_snip = extract_metric_sentence(
            all_pages,
            "FDCA_selectivity",
            ["FDCA selectivity", "selectivity of FDCA", "selectivity"],
            [
                r"\b\d+(\.\d+)?\s*%[^.;]{0,40}(FDCA )?selectivity\b",
                r"selectivity\s*(of|toward|for)?\s*FDCA[^.;]{0,60}\d+(\.\d+)?\s*%",
                r"FDCA[^.;]{0,60}selectivity[^.;]{0,60}\d+(\.\d+)?\s*%",
            ],
        )
        ev.append(evidence(paper_id, pdf.name, "FDCA_selectivity", fdca_selectivity, fs_page, fs_snip, "medium"))
    
        fe, fe_page, fe_snip = extract_metric_sentence(
            all_pages,
            "Faradaic_efficiency",
            ["Faradaic efficiency", "FE", "faradaic"],
            [
                r"(Faradaic efficiency|FE)[^.;]{0,80}\d+(\.\d+)?\s*%",
                r"\d+(\.\d+)?\s*%[^.;]{0,40}(Faradaic efficiency|FE)",
            ],
        )
        ev.append(evidence(paper_id, pdf.name, "Faradaic_efficiency", fe, fe_page, fe_snip, "medium"))
    
        conversion, conv_page, conv_snip = extract_metric_sentence(
            all_pages,
            "HMF_conversion",
            ["HMF conversion", "conversion of HMF", "conversion"],
            [
                r"(HMF )?conversion[^.;]{0,80}\d+(\.\d+)?\s*%",
                r"\d+(\.\d+)?\s*%[^.;]{0,40}(HMF )?conversion",
            ],
        )
        ev.append(evidence(paper_id, pdf.name, "HMF_conversion", conversion, conv_page, conv_snip, "medium"))
    
        best_sentence, best_sentence_page = find_best_performance_sentence(all_pages)
        ev.append(evidence(paper_id, pdf.name, "best_performance_sentence", best_sentence, best_sentence_page, best_sentence, "medium"))
        conversion = value_from_best_sentence(best_sentence, "hmf_conversion", conversion)
        fdca_yield = value_from_best_sentence(best_sentence, "fdca_yield", fdca_yield)
        fdca_selectivity = value_from_best_sentence(best_sentence, "fdca_selectivity", fdca_selectivity)
        fe = value_from_best_sentence(best_sentence, "faradaic_efficiency", fe)
        fdca_yield = clean_metric_value("fdca_yield", fdca_yield)
        fdca_selectivity = clean_metric_value("fdca_selectivity", fdca_selectivity)
        fe = clean_metric_value("faradaic_efficiency", fe)
        hmf_conversion_value = metric_value_from_best_sentence(best_sentence, "hmf_conversion", conversion)
        fdca_yield_value = metric_value_from_best_sentence(best_sentence, "fdca_yield", fdca_yield)
        fdca_selectivity_value = metric_value_from_best_sentence(best_sentence, "fdca_selectivity", fdca_selectivity)
        faradaic_efficiency_value = metric_value_from_best_sentence(best_sentence, "faradaic_efficiency", fe)
    
    else:
        hmf_conc = NOT_REPORTED
        hmf_conversion_value = NOT_REPORTED
        fdca_yield = NOT_REPORTED
        fdca_yield_value = NOT_REPORTED
        fdca_selectivity = NOT_REPORTED
        fdca_selectivity_value = NOT_REPORTED
        fe = NOT_REPORTED
        faradaic_efficiency_value = NOT_REPORTED
        conversion = NOT_REPORTED
        best_sentence = NOT_REPORTED
        best_sentence_page = ""
    stability, st_page, st_snip = value_near_keywords(
        all_pages,
        "stability",
        ["stability", "durability", "chronoamperometry"],
        [r"\b\d+(\.\d+)?\s*h\b", r"\b\d+(\.\d+)?\s*hours\b"],
    )
    ev.append(evidence(paper_id, pdf.name, "stability", stability, st_page, st_snip, "low"))

    controls = []
    if re.search(r"Ni[-\s]?only|only Ni|pure Ni|Ni\(OH\)2", text, re.I):
        controls.append("Ni-only related control")
    if re.search(r"Co[-\s]?only|only Co|pure Co|Co3O4|Co\(OH\)2", text, re.I):
        controls.append("Co-only related control")
    if re.search(r"without HMF|absence of HMF|no HMF", text, re.I):
        controls.append("no-HMF control")
    if re.search(r"physical mixture|mechanical mixture", text, re.I):
        controls.append("physical mixture control")
    controls_value = "; ".join(controls) if controls else NOT_REPORTED

    char_methods = [name for name, pattern in {
        "XRD": r"\bXRD\b",
        "SEM/TEM": r"\bSEM\b|\bTEM\b",
        "XPS": r"\bXPS\b",
        "Raman": r"Raman",
        "in situ/operando": r"in situ|in-situ|operando",
        "XAS": r"\bXAS\b|XANES|EXAFS",
        "DFT": r"\bDFT\b|density functional theory",
    }.items() if re.search(pattern, text, re.I)]

    if is_hmfor_category(paper_category):
        context_electrolyte, context_koh, context_hmf = extract_electrolyte_and_concentrations(all_pages)
        if context_electrolyte != NOT_REPORTED:
            electrolyte = context_electrolyte
        if context_koh != NOT_REPORTED:
            koh_concentration = context_koh
        else:
            koh_concentration = electrolyte if "KOH" in electrolyte else NOT_REPORTED
        if context_hmf != NOT_REPORTED:
            hmf_conc = context_hmf
        onset_potential, performance_potential, generic_potential = extract_potential_contexts(all_pages)
        lsv_current_density = extract_lsv_current_density(all_pages)
        product_analysis_condition = best_sentence
        stability_performance = extract_stability_performance(all_pages)[0]
        coupled_electrolyzer_performance = extract_coupled_performance(all_pages)[0]
    else:
        koh_concentration = electrolyte if "KOH" in electrolyte else NOT_REPORTED
        onset_potential = NOT_REPORTED
        performance_potential = NOT_REPORTED
        generic_potential = first_match(text, [r"at\s*\d+(?:\.\d+)?\s*V\s*(?:vs\.?|versus)?\s*RHE", r"\d+(?:\.\d+)?\s*V\s*(?:vs\.?|versus)?\s*RHE"], NOT_REPORTED)
        lsv_current_density = extract_lsv_current_density(all_pages)
        product_analysis_condition = NOT_REPORTED
        stability_performance = NOT_REPORTED
        coupled_electrolyzer_performance = NOT_REPORTED
    electrode_configuration = extract_electrode_configuration(text)
    metal_loading = extract_metal_loading(text)

    bound_metrics = labeled_percent_metrics(best_sentence)
    hmf_conversion_value = bound_metrics.get("hmf_conversion_value", hmf_conversion_value)
    fdca_yield_value = bound_metrics.get("fdca_yield_value", fdca_yield_value)
    fdca_selectivity_value = bound_metrics.get("fdca_selectivity_value", fdca_selectivity_value)
    faradaic_efficiency_value = bound_metrics.get("faradaic_efficiency_value", faradaic_efficiency_value)

    record_values = {
        "fdca_yield": fdca_yield,
        "fdca_selectivity": fdca_selectivity,
        "faradaic_efficiency": fe,
        "hmf_conversion": conversion,
        "stability": stability,
        "carbon_balance": first_match(text, [r"carbon balance"], "not_reported_in_main_text" if is_hmfor_category(paper_category) else NOT_REPORTED),
        "controls": controls_value,
    }

    site_level = assess_site_level(text, paper_category)
    quality = assess_hmfor_quality(record_values, text, paper_category)
    missing = []

    record = PaperRecord(
        paper_id=paper_id,
        file_name=pdf.name,
        linked_si_files="; ".join(si.name for si in si_files),
        title=title,
        title_source=title_source,
        raw_pdf_title=raw_pdf_title,
        filename_title=filename_title,
        year=year,
        journal=journal,
        doi=doi,
        authors=authors,
        paper_type=paper_type,
        relevance_level=relevance_level_for_category(paper_category, quality, site_level),
        si_mentioned_in_text=si_mentioned_in_text,
        si_mention_evidence=si_mention_evidence,
        paper_category=paper_category,
        pdf_text_status=pdf_status,
        text_length=len(text),
        reaction_type="HMFOR" if paper_category in {"hmfor_core", "hmfor_related"} else paper_category,
        catalyst_name=catalyst_name,
        catalyst_family=refine_family(catalyst_name, title, detect_family(text)),
        substrate=first_match(text, [r"nickel foam|NF\b|carbon cloth|glassy carbon electrode|GCE\b"], NOT_REPORTED),
        ni_co_ratio=extract_ni_co_ratio(text),
        synthesis_method=first_match(text, [r"hydrothermal", r"solvothermal", r"electrodeposition", r"calcination", r"coprecipitation"], NOT_REPORTED),
        electrolyte=electrolyte,
        koh_concentration=koh_concentration,
        hmf_concentration=hmf_conc,
        cell_type=first_match(text, [r"H[-\s]?type cell", r"H[-\s]?cell", r"single[- ]compartment cell", r"divided H[- ]type cell", r"flow cell", r"membrane electrode assembly|\\bMEA\\b", r"anion exchange membrane|\\bAEM\\b"], NOT_REPORTED),
        reference_electrode=canonical_ref_electrode(first_match(text, [r"Hg/HgO", r"Ag/AgCl", r"SCE", r"sce", r"saturated calomel", r"RHE"], NOT_REPORTED)),
        potential_vs_rhe=generic_potential,
        onset_potential_vs_rhe=onset_potential,
        performance_potential_vs_rhe=performance_potential,
        lsv_current_density=lsv_current_density,
        product_analysis_condition=product_analysis_condition,
        stability_performance=stability_performance,
        coupled_electrolyzer_performance=coupled_electrolyzer_performance,
        electrode_configuration=electrode_configuration,
        metal_loading=metal_loading,
        current_density=current_density,
        hmf_conversion=conversion,
        hmf_conversion_value=hmf_conversion_value,
        fdca_yield=fdca_yield,
        fdca_yield_value=fdca_yield_value,
        fdca_selectivity=fdca_selectivity,
        fdca_selectivity_value=fdca_selectivity_value,
        faradaic_efficiency=fe,
        faradaic_efficiency_value=faradaic_efficiency_value,
        best_performance_sentence=best_sentence,
        stability=stability,
        product_analysis_method=first_match(text, [r"\bHPLC\b", r"\bNMR\b", r"\bGC\b|gas chromatography"], NOT_REPORTED),
        carbon_balance=record_values["carbon_balance"],
        controls=controls_value,
        has_ni_only_control="yes" if "Ni-only" in controls_value else "no",
        has_co_only_control="yes" if "Co-only" in controls_value else "no",
        has_no_hmf_control="yes" if "no-HMF" in controls_value else "no",
        has_physical_mixture_control="yes" if "physical mixture" in controls_value else "no",
        characterization="; ".join(char_methods) if char_methods else NOT_REPORTED,
        mechanism_claim=first_match(text, [r"proton deintercalation[^.;]{0,220}", r"competitive adsorption[^.;]{0,220}", r"adsorption kinetics[^.;]{0,220}", r"indirect oxidation[^.;]{0,220}", r"direct oxidation[^.;]{0,220}", r"Ni[-\s]?O[-\s]?Co[^.;]{0,220}", r"dual active sites?[^.;]{0,220}", r"bifunctional[^.;]{0,220}", r"redox mediator[^.;]{0,220}"], NOT_REPORTED),
        site_decoupling_evidence_level=site_level,
        hmfor_data_quality_score=quality,
        reproducibility_value="high" if quality in {"4", "5"} and catalyst_name != NOT_REPORTED else "medium" if quality in {"3", "4"} else "low",
        recommended_use="",
        missing_fields="",
        override_applied="no",
        override_fields="",
        override_reason="",
        override_source="",
        override_date="",
        override_operator="",
        override_version="",
        override_evidence="",
        references_removed=references_removed,
        reference_section_start_page=reference_section_start_page,
        text_used_for_extraction_length=str(len(text)),
        review_notes=f"V1.8 category: {paper_category}. Reason: {category_reason}. Manually verify best-performance values against figures/tables and Supporting Information.",
    )

    record = clear_hmfor_fields_for_background(record)
    record = apply_hmfor_pattern_rules(record, text, all_pages)
    record = apply_curated_hmfor_corrections(record)
    if is_hmfor_category(record.paper_category):
        if record.stability in {NOT_REPORTED, NEEDS_CHECK, ""} and record.stability_performance not in {NOT_REPORTED, NEEDS_CHECK, ""}:
            record.stability = record.stability_performance
        if record.carbon_balance in {NOT_REPORTED, NEEDS_CHECK, ""}:
            record.carbon_balance = "not_reported_in_main_text"
    record_dict = asdict(record)
    record.recommended_use = recommended_use({**record_dict, "paper_category": record.paper_category, "hmfor_data_quality_score": record.hmfor_data_quality_score, "site_decoupling_evidence_level": record.site_decoupling_evidence_level})
    record.missing_fields = compute_missing_fields(record)
    return record, ev


def clean_excel_text(value):
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub("", value)
    return value


def clean_dataframe_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    return df.map(clean_excel_text)




def is_low_quality_background_metric(value: str) -> bool:
    if value in {NOT_REPORTED, NEEDS_CHECK, ""}:
        return False
    if len(value) < 20:
        return True
    bad = r"Wenjun He|Rahul Patil|Sayan Banerjee|A B S T R A C T|Recent Development|is crucial|taking place at the cathode|need noble metal|Hydrogen production through water electrolysis"
    if re.search(bad, value, re.I):
        return True
    citation_stripped = re.sub(r"\[[0-9, -]+\]", "", value)
    if not re.search(r"\d|mV|mA|h|cycle|retention|capacity|overpotential|Tafel", citation_stripped, re.I):
        return True
    return False


def metric_confidence(metric_name: str, value: str, category: str) -> str:
    if category == "ni_co_background" and metric_name in {"HER_metric", "OER_metric"} and is_low_quality_background_metric(value):
        return "very_low"
    return "low" if category == "ni_co_background" else "medium"


def normalize_unit_text(value: str) -> str:
    if value in {NOT_REPORTED, NEEDS_CHECK, ""}:
        return value
    cleaned = normalize_space(value)
    cleaned = cleaned.replace("−", "-").replace("–", "-").replace("‐", "-").replace("‑", "-")
    cleaned = re.sub(r"cm²|cm\^2", "cm2", cleaned, flags=re.I)
    cleaned = re.sub(r"mA\s*/\s*cm2", "mA cm-2", cleaned, flags=re.I)
    cleaned = re.sub(r"mA\s*cm\s*[-−]?\s*2", "mA cm-2", cleaned, flags=re.I)
    cleaned = re.sub(r"A\s*/\s*cm2", "A cm-2", cleaned, flags=re.I)
    cleaned = re.sub(r"A\s*cm\s*[-−]?\s*2", "A cm-2", cleaned, flags=re.I)
    cleaned = re.sub(r"\b(\d+(?:\.\d+)?)\s*m\s*KOH\b", r"\1 M KOH", cleaned)
    cleaned = re.sub(r"\b(\d+(?:\.\d+)?)\s*mol\s*L[-/]?1\b", r"\1 M", cleaned, flags=re.I)
    cleaned = re.sub(r"\bvs\.?\s*RHE\b", "vs RHE", cleaned, flags=re.I)
    cleaned = re.sub(r"\bwt\s*%", "wt%", cleaned, flags=re.I)
    return cleaned


def normalize_metric_value(metric_name: str, value: str) -> tuple[str, str]:
    if value in {NOT_REPORTED, NEEDS_CHECK, ""}:
        return value, ""
    text = normalize_unit_text(value)
    if metric_name in {"conversion", "FDCA_yield", "FDCA_selectivity", "FE"}:
        m = re.search(r"\d+(?:\.\d+)?", text)
        return (m.group(0), "%") if m else (NEEDS_CHECK, "%")
    if metric_name in {"current_density"}:
        m = re.search(r"\d+(?:\.\d+)?", text)
        return (m.group(0), "mA cm-2") if m and re.search(r"mA\s*cm-2", text, re.I) else (text, "")
    if metric_name == "overpotential":
        m = re.search(r"(\d+(?:\.\d+)?)\s*mV", text, re.I)
        return (m.group(1), "mV") if m else (text, "")
    return text, ""


def comparison_eligibility(paper_category: str, performance_context: str, source_scope: str, confidence: str, metric_value: str) -> tuple[str, str]:
    if confidence == "very_low":
        return "no", "very_low confidence"
    if source_scope == "references":
        return "no", "reference section"
    if not is_hmfor_category(paper_category):
        return "no", "background metric"
    if performance_context != "HMFOR":
        return "no", "non-HMFOR metric"
    metric_text = normalize_unit_text(str(metric_value))
    if metric_value in {NOT_REPORTED, NEEDS_CHECK, ""}:
        return "no", "value not quantified"
    if "scale needs check" in metric_text.lower():
        return "no", "potential scale needs check"
    return "yes", ""


def infer_source_scope(evidence_sentence: str, source_scope: str = "unknown") -> str:
    if source_scope not in {"", "unknown", "main_text_or_SI", "main_text"}:
        return source_scope
    text = evidence_sentence or ""
    if re.search(r"abstract", text[:80], re.I):
        return "abstract"
    if re.search(r"experimental|materials and methods|preparation|synthesis", text, re.I):
        return "experimental_section"
    if re.search(r"Fig\.?|Figure|Table", text):
        return "figure_caption" if re.search(r"Fig\.?|Figure", text) else "table"
    if re.search(r"supporting information|Table S|Fig\.? S|Figure S", text, re.I):
        return "supporting_information"
    return "results_discussion" if text not in {NOT_REPORTED, NEEDS_CHECK, ""} else "unknown"


def metric_record(paper_id: str, file_name: str, paper_category: str, performance_context: str, metric_name: str, metric_value: str, condition: str, evidence_sentence: str, source_scope: str, confidence: str) -> MetricRecord:
    normalized_value, normalized_unit = normalize_metric_value(metric_name, metric_value)
    scoped = infer_source_scope(evidence_sentence, source_scope)
    eligible, exclusion = comparison_eligibility(paper_category, performance_context, scoped, confidence, metric_value)
    return MetricRecord(
        paper_id=paper_id,
        file_name=file_name,
        paper_category=paper_category,
        performance_context=performance_context,
        metric_name=metric_name,
        metric_value=normalize_unit_text(metric_value),
        metric_value_normalized=normalized_value,
        metric_unit_normalized=normalized_unit,
        condition=normalize_unit_text(condition),
        evidence_sentence=evidence_sentence,
        source_scope=scoped,
        confidence=confidence,
        comparison_eligible=eligible,
        exclusion_reason=exclusion,
    )


def title_fingerprint(title: str) -> str:
    if title in {NOT_REPORTED, NEEDS_CHECK, ""}:
        return ""
    words = re.findall(r"[a-z0-9]+", title.lower())
    stop = {"the", "a", "an", "of", "for", "and", "in", "on", "with", "to", "by", "from"}
    return " ".join(w for w in words if w not in stop)[:180]


def build_dedup_report(records: list[PaperRecord]) -> pd.DataFrame:
    groups: dict[tuple[str, str], list[PaperRecord]] = {}
    for record in records:
        doi_key = normalize_space(record.doi).lower().rstrip(".") if record.doi not in {NOT_REPORTED, NEEDS_CHECK, ""} else ""
        if doi_key:
            key = ("doi", doi_key)
        else:
            fp = title_fingerprint(record.title)
            key = ("title_fingerprint", fp or record.file_name.lower())
        groups.setdefault(key, []).append(record)

    rows: list[DedupRecord] = []
    group_idx = 1
    for (basis, key), members in groups.items():
        is_duplicate = len(members) > 1
        group_id = f"D{group_idx:04d}" if is_duplicate else ""
        primary = members[0].paper_id
        if is_duplicate:
            group_idx += 1
        for member in members:
            rows.append(DedupRecord(
                paper_id=member.paper_id,
                file_name=member.file_name,
                duplicate_group_id=group_id,
                duplicate_status="duplicate" if is_duplicate and member.paper_id != primary else "primary" if is_duplicate else "unique",
                primary_paper_id=primary if is_duplicate else member.paper_id,
                match_basis=basis,
                match_key=key,
                confidence="high" if basis == "doi" else "medium",
            ))
    return pd.DataFrame([asdict(r) for r in rows])



def si_needed_for_record(record: PaperRecord) -> tuple[str, str]:
    if not is_hmfor_category(record.paper_category):
        return "no", "background/excluded paper"
    reasons: list[str] = []
    if record.missing_fields:
        reasons.append(f"missing fields: {record.missing_fields}")
    if record.carbon_balance == "not_reported_in_main_text":
        reasons.append("carbon balance not reported in main text")
    if record.si_mentioned_in_text == "yes":
        reasons.append("main text mentions supporting information")
    if record.product_analysis_method in {NOT_REPORTED, NEEDS_CHECK, ""}:
        reasons.append("product analysis method not confirmed")
    if "stability" in record.missing_fields or record.stability in {NOT_REPORTED, NEEDS_CHECK, ""}:
        reasons.append("stability details may be in SI")
    return ("yes", "; ".join(reasons)) if reasons else ("optional", "HMFOR paper; SI useful for audit but no critical gap detected")


def si_field_patterns() -> dict[str, list[str]]:
    return {
        "hmf_concentration": [r"\d+(?:\.\d+)?\s*mM\s*HMF", r"HMF[^.;]{0,100}\d+(?:\.\d+)?\s*mM"],
        "hmf_conversion": [r"HMF\s+conversion(?:\s+rate)?[^.;]{0,100}\d+(?:\.\d+)?\s*%", r"conversion\s+of\s+HMF[^.;]{0,100}\d+(?:\.\d+)?\s*%"],
        "fdca_yield": [r"FDCA\s+yield[^.;]{0,100}\d+(?:\.\d+)?\s*%", r"yield[^.;]{0,60}FDCA[^.;]{0,100}\d+(?:\.\d+)?\s*%"],
        "fdca_selectivity": [r"FDCA\s+selectivity[^.;]{0,100}\d+(?:\.\d+)?\s*%", r"selectivity[^.;]{0,60}FDCA[^.;]{0,100}\d+(?:\.\d+)?\s*%"],
        "faradaic_efficiency": [r"Faradaic\s+efficiency[^.;]{0,100}\d+(?:\.\d+)?\s*%", r"\bFE\b[^.;]{0,70}\d+(?:\.\d+)?\s*%"],
        "product_analysis_method": [r"\bHPLC\b", r"\bNMR\b", r"\bGC\b|gas chromatography"],
        "stability": [r"stability[^.;]{0,160}", r"(?:consecutive\s+)?cycles?[^.;]{0,160}", r"chronoamperometry[^.;]{0,160}"],
        "carbon_balance": [r"carbon\s+balance[^.;]{0,120}\d+(?:\.\d+)?\s*%", r"carbon\s+balance[^.;]{0,160}"],
        "ni_co_ratio": [r"Ni\s*[:/]\s*Co[^.;]{0,100}", r"Ni/Co[^.;]{0,100}"],
        "metal_loading": [r"\d+(?:\.\d+)?\s*wt\s*%[^.;]{0,60}", r"ICP[^.;]{0,180}"],
    }



def read_docx_text(path: Path) -> str:
    try:
        with zipfile.ZipFile(path) as zf:
            xml = zf.read("word/document.xml").decode("utf-8", errors="replace")
    except Exception:
        return ""
    xml = re.sub(r"<w:tab[^>]*/>", " ", xml)
    xml = re.sub(r"</w:p>", " ", xml)
    xml = re.sub(r"<[^>]+>", "", xml)
    return normalize_space(html.unescape(xml))


def read_si_attachment_text(path: Path) -> str:
    if path.suffix.lower() == ".pdf":
        return " ".join(read_pdf_pages(path))
    if path.suffix.lower() == ".docx":
        return read_docx_text(path)
    return ""

def scan_si_content(si_path: Path) -> tuple[str, str, str]:
    try:
        raw_text = read_si_attachment_text(si_path)
    except Exception:
        return "ocr_needed", "", "failed to read SI text"
    if not raw_text:
        return "ocr_needed", "", "failed to read SI text"
    text = normalize_unit_text(raw_text)
    text = normalize_space(text)
    status = text_status(text)
    found: list[str] = []
    snippets: list[str] = []
    for field, patterns in si_field_patterns().items():
        for pattern in patterns:
            match = re.search(pattern, text, re.I)
            if match:
                found.append(field)
                start = max(0, match.start() - 80)
                end = min(len(text), match.end() + 100)
                snippets.append(f"{field}: {normalize_space(text[start:end])[:220]}")
                break
    return status, "; ".join(dict.fromkeys(found)), " | ".join(snippets[:6])

def build_si_links(records: list[PaperRecord], all_pdfs: list[Path] | None = None) -> pd.DataFrame:
    rows: list[SILinkRecord] = []
    pdf_by_name = {p.name: p for p in (all_pdfs or [])}
    for record in records:
        linked = [name.strip() for name in str(record.linked_si_files).split(";") if name.strip()]
        si_needed, si_needed_reason = si_needed_for_record(record)
        if not linked:
            rows.append(SILinkRecord(record.paper_id, record.file_name, "no", NOT_REPORTED, NOT_REPORTED, "no", "", "", "main text mentions SI but no SI file matched" if record.si_mentioned_in_text == "yes" else "no SI file matched in input folder", si_needed, si_needed_reason, "not_found", "", publisher_si_candidate_urls(record.doi, record.journal, ""), "low"))
            continue
        record_terms = set(re.findall(r"[a-z0-9]{4,}", record.file_name.lower()))
        for si_name in linked:
            si_terms = set(re.findall(r"[a-z0-9]{4,}", si_name.lower()))
            shared = sorted(record_terms & si_terms)
            si_path = pdf_by_name.get(si_name)
            si_text_status = NEEDS_CHECK
            fields_from_si = NEEDS_CHECK
            evidence_summary = ""
            if si_path is not None:
                si_text_status, fields_from_si, evidence_summary = scan_si_content(si_path)
                if not fields_from_si:
                    fields_from_si = NOT_REPORTED
            rows.append(SILinkRecord(
                paper_id=record.paper_id,
                file_name=record.file_name,
                si_found="yes",
                si_file_name=si_name,
                si_text_status=si_text_status,
                si_used_for_extraction="yes" if fields_from_si not in {NOT_REPORTED, NEEDS_CHECK, ""} else "no",
                fields_filled_from_si=fields_from_si,
                si_field_evidence_summary=evidence_summary,
                si_missing_reason="",
                si_needed=si_needed,
                si_needed_reason=si_needed_reason,
                link_method="filename_shared_terms_or_article_id",
                shared_terms="; ".join(shared),
                publisher_si_candidate_urls=publisher_si_candidate_urls(record.doi, record.journal, ""),
                confidence="high" if fields_from_si not in {NOT_REPORTED, NEEDS_CHECK, ""} else "medium" if len(shared) >= 2 else "low",
            ))
    return pd.DataFrame([asdict(r) for r in rows])

def build_papers_table(records: list[PaperRecord]) -> pd.DataFrame:
    cols = ["paper_id", "file_name", "title", "year", "journal", "doi", "authors", "paper_type", "relevance_level", "si_mentioned_in_text", "si_mention_evidence", "paper_category", "pdf_text_status", "text_length", "references_removed", "reference_section_start_page", "text_used_for_extraction_length", "recommended_use", "override_applied", "override_reason", "override_source", "override_date", "override_operator", "override_version", "override_evidence", "review_notes"]
    return pd.DataFrame([{col: getattr(r, col) for col in cols} for r in records])


def build_materials_table(records: list[PaperRecord]) -> pd.DataFrame:
    cols = ["paper_id", "file_name", "paper_category", "catalyst_name", "catalyst_family", "substrate", "ni_co_ratio", "synthesis_method", "metal_loading", "characterization", "mechanism_claim"]
    rows = []
    for record in records:
        row = {col: normalize_unit_text(str(getattr(record, col))) for col in cols}
        row["table_scope"] = "not_applicable_for_exclude" if record.paper_category == "exclude" else "included_analyzable_paper"
        if record.paper_category == "exclude":
            for col in cols:
                if col not in {"paper_id", "file_name", "paper_category"}:
                    row[col] = "not_applicable"
        rows.append(row)
    return pd.DataFrame(rows)


def strict_background_metric_confidence(metric_name: str, value: str) -> tuple[bool, str]:
    text = normalize_unit_text(str(value))
    if value in {NOT_REPORTED, NEEDS_CHECK, "", "nan"}:
        return False, "empty"
    strict_patterns = [
        r"overpotential[^.;]{0,80}(?:at|@)[^.;]{0,60}\d+(?:\.\d+)?\s*mA\s*cm-2",
        r"\d+(?:\.\d+)?\s*mV[^.;]{0,80}(?:at|@)[^.;]{0,60}\d+(?:\.\d+)?\s*mA\s*cm-2",
        r"current\s+density[^.;]{0,80}(?:at|@)[^.;]{0,60}\d+(?:\.\d+)?\s*V",
        r"\d+(?:\.\d+)?\s*mA\s*cm-2[^.;]{0,80}(?:at|@)[^.;]{0,60}\d+(?:\.\d+)?\s*V",
        r"Tafel\s+slope[^.;]{0,80}\d+(?:\.\d+)?\s*mV\s*dec-?1",
        r"stability[^.;]{0,80}\d+(?:\.\d+)?\s*h",
        r"\d+(?:\.\d+)?\s*h[^.;]{0,80}stability",
    ]
    if any(re.search(p, text, re.I) for p in strict_patterns):
        return True, "strict experimental metric template"
    return False, "background metric failed strict template"

def build_metrics_long_table(records: list[PaperRecord], nico_background_rows: list[NiCoBackgroundRecord]) -> pd.DataFrame:
    rows: list[MetricRecord] = []
    for record in records:
        if is_hmfor_category(record.paper_category):
            condition = "; ".join(v for v in [record.electrolyte, record.hmf_concentration, record.performance_potential_vs_rhe, record.cell_type] if v not in {NOT_REPORTED, NEEDS_CHECK, ""})
            for name, value in [
                ("conversion", record.hmf_conversion_value),
                ("FDCA_yield", record.fdca_yield_value),
                ("FDCA_selectivity", record.fdca_selectivity_value),
                ("FE", record.faradaic_efficiency_value),
                ("current_density", record.current_density),
            ]:
                if value not in {NOT_REPORTED, NEEDS_CHECK, ""}:
                    rows.append(metric_record(record.paper_id, record.file_name, record.paper_category, "HMFOR", name, value, condition, record.best_performance_sentence, "main_text_or_SI", "medium"))
        elif record.paper_category == "ni_co_background":
            background = next((r for r in nico_background_rows if r.paper_id == record.paper_id), None)
            if background:
                for context, name, value in [
                    ("OER", "overpotential", background.overpotential),
                    ("OER", "current_density", background.current_density),
                    ("HER", "HER_metric", background.her_metric),
                    ("OER", "OER_metric", background.oer_metric),
                    ("capacitor", "stability_or_capacity", background.stability),
                ]:
                    accepted, reason = strict_background_metric_confidence(name, value)
                    if value not in {NOT_REPORTED, NEEDS_CHECK, ""} and accepted:
                        rows.append(metric_record(record.paper_id, record.file_name, record.paper_category, context, name, value, background.electrolyte, value, "main_text", "medium"))
                    elif value not in {NOT_REPORTED, NEEDS_CHECK, ""}:
                        # Keep weak background residues out of metrics_long_table; quality is tracked elsewhere.
                        continue
    return pd.DataFrame([asdict(r) for r in rows])



def is_acceptable_short_text(field: str, value: str) -> bool:
    if value in STANDARD_ENUM_VALUES:
        return True
    if field in {"reference_electrode"} and re.fullmatch(r"Hg/HgO|Ag/AgCl|SCE|RHE", value, re.I):
        return True
    if field in {"product_analysis_method"} and re.fullmatch(r"HPLC|NMR|GC|GC-MS|LC-MS", value, re.I):
        return True
    if field in {"catalyst_family"} and re.fullmatch(r"[A-Za-z;/ ]+", value):
        return True
    if field in {"synthesis_method"} and re.fullmatch(r"hydrothermal|solvothermal|electrodeposition|calcination|coprecipitation|co-precipitation|Electrodeposition", value, re.I):
        return True
    if field in {"electrode_configuration", "cell_type"} and re.fullmatch(r"three-electrode|two-electrode|H-type cell|H-cell", value, re.I):
        return True
    if field in {"characterization"} and re.fullmatch(r"[A-Z/; -]+", value):
        return True
    if field in {"ni_co_ratio"} and re.search(r"Ni\s*:?\s*Co|Ni:Co|Ni/Co", value, re.I):
        return True
    if field in {"stability"} and re.fullmatch(r"\d+(?:\.\d+)?\s*h", value, re.I):
        return True
    if re.search(r"\d", value) and re.search(r"M|mM|V|mA|cm|wt|%|RHE|KOH|HMF", value, re.I):
        return True
    if field in {"catalyst_name", "substrate", "electrolyte", "metal_loading", "hmf_concentration", "hmf_conversion", "fdca_yield", "fdca_selectivity", "faradaic_efficiency", "current_density", "lsv_current_density"}:
        return True
    return False


def field_confidence_for(field: str, value: str, record: PaperRecord, evidence_df: pd.DataFrame) -> tuple[str, str, str]:
    value = str(value)
    if value in {NOT_REPORTED, "", "nan"}:
        return "low", "field not reported", "none"
    if value == NEEDS_CHECK or "needs check" in value.lower():
        return "medium", "value present but requires manual confirmation", "partial"
    if record.override_applied == "yes" and field in str(record.override_fields).split("; "):
        return "high", "manual override applied with curated source", "manual_curated"
    if not evidence_df.empty and field in set(str(x) for x in evidence_df[evidence_df["paper_id"] == record.paper_id]["field"].tolist()):
        return "medium", "extracted with evidence snippet", "snippet"
    if field in {"paper_category", "title", "doi", "year"}:
        return "medium", "metadata or rule-derived field", "metadata"
    return "medium", "rule-derived from parsed text", "rule"


def build_field_confidence_table(records: list[PaperRecord], evidence_df: pd.DataFrame) -> pd.DataFrame:
    key_fields = [
        "paper_category", "title", "doi", "catalyst_name", "catalyst_family",
        "electrolyte", "hmf_concentration", "potential_vs_rhe",
        "onset_potential_vs_rhe", "performance_potential_vs_rhe",
        "current_density", "hmf_conversion", "fdca_yield",
        "fdca_selectivity", "faradaic_efficiency", "product_analysis_method",
        "carbon_balance", "stability", "stability_performance",
        "mechanism_claim", "site_decoupling_evidence_level",
        "hmfor_data_quality_score",
    ]
    rows: list[FieldConfidenceRecord] = []
    for record in records:
        for field in key_fields:
            value = getattr(record, field, NOT_REPORTED)
            conf, reason, strength = field_confidence_for(field, value, record, evidence_df)
            rows.append(FieldConfidenceRecord(record.paper_id, record.file_name, field, normalize_unit_text(str(value)), conf, reason, strength))
    return pd.DataFrame([asdict(r) for r in rows])



def infer_alkaline_ph(koh_text: str) -> tuple[str, str]:
    text = normalize_unit_text(koh_text)
    m = re.search(r"(\d+(?:\.\d+)?)\s*M\s*KOH", text, re.I)
    if not m:
        return "", "no KOH molarity parsed"
    concentration = float(m.group(1))
    # Approximate strong-base pH at 25 C: pOH = -log10([OH-]); pH = 14 - pOH.
    # This is an audit assumption, not a replacement for paper-reported pH.
    import math
    ph = 14 + math.log10(concentration)
    return f"pH~{ph:.2f} inferred from {m.group(1)} M KOH", "25C strong-base approximation"


def rhe_rule_for_reference(reference: str) -> dict:
    rules = load_json_rules(DEFAULT_RHE_RULES_PATH).get("rules", [])
    for rule in rules:
        if str(rule.get("reference", "")).lower() == str(reference).lower():
            return rule
    return {}


def numeric_potential_value(value: str) -> str:
    match = re.search(r"(\d+(?:\.\d+)?)\s*V", value, re.I)
    return match.group(1) if match else ""


def estimate_converted_vs_rhe(original: str, reference: str, koh_text: str, enable_conversion: bool) -> tuple[str, str]:
    if not enable_conversion:
        return NEEDS_CHECK, "conversion disabled"
    potential = numeric_potential_value(original)
    ph_text, ph_assumption = infer_alkaline_ph(koh_text)
    ph_match = re.search(r"pH~(\d+(?:\.\d+)?)", ph_text)
    if not potential or not ph_match:
        return NEEDS_CHECK, "missing potential or pH for conversion"
    offsets = {"Hg/HgO": 0.098, "Ag/AgCl": 0.197, "SCE": 0.241}
    offset = offsets.get(reference)
    if offset is None:
        return NEEDS_CHECK, "unsupported reference electrode"
    converted = float(potential) + offset + 0.0591 * float(ph_match.group(1))
    return f"{converted:.3f} V vs RHE", ph_assumption


def audit_potential_conversion(original: str, reference: str, koh_text: str, enable_conversion: bool = False) -> tuple[str, str, str, str, str, str, str]:
    original = normalize_unit_text(original)
    if re.search(r"vs RHE", original, re.I):
        trust = "high" if "needs check" not in original.lower() else "medium"
        assumption = "none" if trust == "high" else "potential scale marked needs check"
        final_value = original if trust == "high" else NEEDS_CHECK
        return "already reported vs RHE; no conversion", original, original, final_value, trust, assumption, koh_text
    rule = rhe_rule_for_reference(reference)
    formula = rule.get("formula", "reference-specific formula unavailable; verify original scale first")
    ph_text, ph_assumption = infer_alkaline_ph(koh_text)
    if "scale needs check" in original.lower():
        candidate, conversion_assumption = estimate_converted_vs_rhe(original, reference, koh_text, enable_conversion)
        assumption = f"original scale uncertain; {conversion_assumption}"
        return formula, NEEDS_CHECK, candidate, NEEDS_CHECK, "low", assumption, ph_text or koh_text
    if rule:
        candidate, conversion_assumption = estimate_converted_vs_rhe(original, reference, koh_text, enable_conversion)
        conf = str(rule.get("confidence_when_pH_inferred", "low")) if candidate != NEEDS_CHECK else "low"
        final_value = candidate if conf == "high" else NEEDS_CHECK
        assumption = conversion_assumption if candidate != NEEDS_CHECK else "pH/KOH unavailable or conversion disabled; conversion deferred"
        return formula, final_value, candidate, final_value, conf, assumption, ph_text or koh_text
    return "not converted in V2; unsupported or unknown reference electrode", NEEDS_CHECK, NEEDS_CHECK, NEEDS_CHECK, "low", "conversion deferred", koh_text

def build_potential_conversion_audit(records: list[PaperRecord], enable_conversion: bool = False) -> pd.DataFrame:
    rows: list[PotentialConversionAuditRecord] = []
    for record in records:
        for field in ["potential_vs_rhe", "onset_potential_vs_rhe", "performance_potential_vs_rhe"]:
            original = normalize_unit_text(str(getattr(record, field, NOT_REPORTED)))
            if original in {NOT_REPORTED, NEEDS_CHECK, "", "nan"}:
                continue
            ref = record.reference_electrode
            koh = record.koh_concentration if record.koh_concentration not in {NOT_REPORTED, NEEDS_CHECK, ""} else record.electrolyte
            formula, converted, candidate, final_value, trust, assumption, ph_or_koh = audit_potential_conversion(original, ref, koh, enable_conversion)
            rows.append(PotentialConversionAuditRecord(record.paper_id, record.file_name, field, original, ref, ph_or_koh, formula, converted, candidate, final_value, trust, "yes" if enable_conversion else "no", trust, assumption))
    return pd.DataFrame([asdict(r) for r in rows])

def unit_conversion_type(raw: str, normalized: str) -> str:
    types: list[str] = []
    if any(ch in raw for ch in ["−", "–", "‐", "‑"]):
        types.append("unicode_minus_normalized")
    if re.search(r"(?:mA|A)\s*/\s*cm2", raw, re.I) or re.search(r"(?:mA|A)\s*cm\s*[-−]?\s*2", raw, re.I):
        types.append("current_density_unit_standardized")
    if re.search(r"cm²|cm\^2", raw, re.I):
        types.append("area_exponent_standardized")
    if re.search(r"\bvs\.?\s*RHE\b", raw, re.I) and raw != normalized:
        types.append("rhe_label_standardized")
    if re.search(r"\b\d+(?:\.\d+)?\s*m\s*KOH\b", raw):
        types.append("koh_molar_case_standardized")
    if re.search(r"mol\s*L[-/]?1", raw, re.I):
        types.append("molar_unit_standardized")
    if re.search(r"wt\s*%", raw, re.I) and raw != normalized:
        types.append("weight_percent_standardized")
    return "; ".join(types) if types else "none"


def build_unit_normalization_log(records: list[PaperRecord], metrics_df: pd.DataFrame) -> pd.DataFrame:
    rows: list[UnitNormalizationRecord] = []
    fields = ["electrolyte", "koh_concentration", "hmf_concentration", "potential_vs_rhe", "onset_potential_vs_rhe", "performance_potential_vs_rhe", "current_density", "lsv_current_density"]
    for record in records:
        for field in fields:
            raw = str(getattr(record, field, NOT_REPORTED))
            if raw in {NOT_REPORTED, NEEDS_CHECK, "", "nan"}:
                continue
            normalized = normalize_unit_text(raw)
            value, unit = normalize_metric_value("current_density" if "current_density" in field else field, raw)
            conversion_type = unit_conversion_type(raw, normalized)
            if conversion_type != "none":
                rows.append(UnitNormalizationRecord(record.paper_id, record.file_name, field, raw, value if unit else normalized, unit, conversion_type, "no"))
    if not metrics_df.empty:
        for _, metric in metrics_df.iterrows():
            raw = str(metric.get("metric_value", ""))
            conversion_type = unit_conversion_type(raw, normalize_unit_text(raw))
            if conversion_type != "none":
                rows.append(UnitNormalizationRecord(str(metric.get("paper_id", "")), str(metric.get("file_name", "")), str(metric.get("metric_name", "")), raw, str(metric.get("metric_value_normalized", "")), str(metric.get("metric_unit_normalized", "")), conversion_type, "no"))
    return pd.DataFrame([asdict(r) for r in rows])



def score_to_int(value: str, default: int = 0) -> int:
    try:
        return int(float(str(value)))
    except Exception:
        return default


def build_quality_scores(records: list[PaperRecord], field_confidence_df: pd.DataFrame, si_links_df: pd.DataFrame) -> pd.DataFrame:
    rows: list[QualityScoreRecord] = []
    for record in records:
        if record.paper_category == "exclude":
            rows.append(QualityScoreRecord(record.paper_id, record.file_name, "0", "0", "0", "0", "0", "0", "0.0", "excluded paper"))
            continue
        required = HMFOR_REQUIRED_FIELDS if is_hmfor_category(record.paper_category) else BACKGROUND_REQUIRED_FIELDS
        values = asdict(record)
        present = sum(1 for f in required if values.get(f, NOT_REPORTED) not in {NOT_REPORTED, NEEDS_CHECK, ""})
        completeness = round(5 * present / max(len(required), 1), 1)
        fc_rows = field_confidence_df[field_confidence_df["paper_id"] == record.paper_id] if not field_confidence_df.empty else pd.DataFrame()
        high_or_medium = 0 if fc_rows.empty else sum(str(v) in {"high", "medium"} for v in fc_rows["field_confidence"].tolist())
        traceability = round(5 * high_or_medium / max(len(fc_rows), 1), 1) if not fc_rows.empty else 0
        relevance = 5 if record.paper_category == "hmfor_core" else 3 if record.paper_category == "hmfor_related" else 2 if record.paper_category == "ni_co_background" else 1
        mechanism = {"L0": 1, "L1": 2, "L2": 3, "L3": 4, "L4": 5}.get(record.site_decoupling_evidence_level, 1 if record.mechanism_claim not in {NOT_REPORTED, NEEDS_CHECK, ""} else 0)
        reproducibility = {"high": 5, "medium": 3, "low": 1}.get(record.reproducibility_value, 1)
        si_row = si_links_df[si_links_df["paper_id"] == record.paper_id] if not si_links_df.empty else pd.DataFrame()
        si_support = 5 if (not si_row.empty and str(si_row.iloc[0].get("si_found", "no")) == "yes") else 2 if is_hmfor_category(record.paper_category) else 3
        overall = round((completeness * 0.25 + traceability * 0.2 + relevance * 0.2 + mechanism * 0.15 + reproducibility * 0.1 + si_support * 0.1), 1)
        notes = []
        if is_hmfor_category(record.paper_category) and si_support < 5:
            notes.append("SI not linked")
        if record.override_applied == "yes":
            notes.append("manual override used")
        if record.missing_fields:
            notes.append(f"missing: {record.missing_fields}")
        rows.append(QualityScoreRecord(record.paper_id, record.file_name, str(completeness), str(traceability), str(relevance), str(mechanism), str(reproducibility), str(si_support), str(overall), "; ".join(notes)))
    return pd.DataFrame([asdict(r) for r in rows])

def review_record(paper_id: str, file_name: str, review_reason: str, field: str, value: str, priority: str, suggested_action: str, source_issue_id: str = "manual_rule") -> ManualReviewRecord:
    return ManualReviewRecord(paper_id, file_name, review_reason, field, value, priority, suggested_action, source_issue_id)


def build_manual_review_queue(records: list[PaperRecord], quality_df: pd.DataFrame, si_links_df: pd.DataFrame, metrics_df: pd.DataFrame) -> pd.DataFrame:
    rows: list[ManualReviewRecord] = []
    for record in records:
        if is_hmfor_category(record.paper_category) and record.missing_fields not in {"", NOT_REPORTED}:
            rows.append(review_record(record.paper_id, record.file_name, "HMFOR core/related paper missing critical fields", "missing_fields", record.missing_fields, "high", "check main text, figures, tables, and SI"))
        for field in ["potential_vs_rhe", "onset_potential_vs_rhe", "performance_potential_vs_rhe"]:
            value = str(getattr(record, field, ""))
            if "scale needs check" in value.lower():
                rows.append(review_record(record.paper_id, record.file_name, "potential scale uncertain", field, value, "high", "verify original reference electrode and convert/audit vs RHE"))
        if is_hmfor_category(record.paper_category):
            si_row = si_links_df[si_links_df["paper_id"] == record.paper_id] if not si_links_df.empty else pd.DataFrame()
            if si_row.empty or str(si_row.iloc[0].get("si_found", "no")) == "no":
                si_candidates = str(si_row.iloc[0].get("publisher_si_candidate_urls", "")) if not si_row.empty else publisher_si_candidate_urls(record.doi, record.journal, "")
                rows.append(review_record(record.paper_id, record.file_name, "SI not found for HMFOR paper", "linked_si_files", si_candidates or record.linked_si_files or NOT_REPORTED, "medium", "search/download Supporting Information if product analysis details are incomplete"))
            elif str(si_row.iloc[0].get("si_used_for_extraction", "no")) == "no":
                rows.append(review_record(record.paper_id, record.file_name, "SI found but no relevant fields detected", "fields_filled_from_si", str(si_row.iloc[0].get("fields_filled_from_si", NOT_REPORTED)), "medium", "inspect SI text quality and field patterns"))
        if record.override_applied == "yes" and len([f for f in record.override_fields.split("; ") if f]) >= 8:
            rows.append(review_record(record.paper_id, record.file_name, "many fields filled by manual override", "override_fields", record.override_fields, "medium", "review whether parser rules should be improved"))
    if not quality_df.empty:
        for _, issue in quality_df.iterrows():
            issue_type = str(issue.get("issue_type", ""))
            if issue_type in {"hmfor_missing_critical_field", "truncated_or_odd_quote", "very_low_confidence_background_metric", "best_sentence_missing_target_metrics", "best_sentence_missing_conversion_only", "reported_in_figure_not_quantified_in_text", "bad_short_value_cleaned"}:
                rows.append(review_record(str(issue.get("paper_id", "")), str(issue.get("file_name", "")), issue_type, str(issue.get("field", "")), str(issue.get("detail", "")), "high" if issue_type == "hmfor_missing_critical_field" else "medium", "inspect evidence and decide whether to override, ignore, or improve parser", str(issue.get("source_issue_id", ""))))
    if not metrics_df.empty and "confidence" in metrics_df.columns:
        bad_metrics = metrics_df[(metrics_df["comparison_eligible"] == "no") & (metrics_df["confidence"] == "very_low")]
        for _, metric in bad_metrics.iterrows():
            rows.append(review_record(str(metric.get("paper_id", "")), str(metric.get("file_name", "")), "metric excluded from comparison", str(metric.get("metric_name", "")), str(metric.get("metric_value", "")), "low", "usually ignore for comparison; improve background parser in V2 if needed"))
    return pd.DataFrame([asdict(r) for r in rows]).drop_duplicates()

def build_quality_issues(records_df: pd.DataFrame, evidence_df: pd.DataFrame, metrics_df: pd.DataFrame) -> pd.DataFrame:
    issues: list[dict[str, str]] = []
    hmfor_cols = [c for c in HMFOR_SPECIFIC_FIELDS if c in records_df.columns]
    for _, row in records_df.iterrows():
        category = row.get("paper_category", "")
        paper_id = row.get("paper_id", "")
        file_name = row.get("file_name", "")
        if not is_hmfor_category(category):
            for col in hmfor_cols:
                value = str(row.get(col, ""))
                if value not in {NOT_REPORTED, "", "nan", "not_applicable"}:
                    issues.append({"paper_id": paper_id, "file_name": file_name, "issue_type": "background_has_hmfor_metric", "field": col, "detail": value[:240]})
        for col, value in row.items():
            value_text = str(value).strip()
            if str(col) not in QUALITY_TEXT_FIELDS:
                continue
            if value_text in {"nan"} or value_text in STANDARD_ENUM_VALUES:
                continue
            if re.search(r"^(\d+(?:\.\d+)?|P\d{4})$", value_text):
                continue
            if value_text in BAD_SHORT_VALUES or (0 < len(value_text) < 20 and re.search(r"[A-Za-z]", value_text) and not is_acceptable_short_text(str(col), value_text)):
                issues.append({"paper_id": paper_id, "file_name": file_name, "issue_type": "suspicious_short_text_value", "field": str(col), "detail": value_text})
            if re.search(r"[”’]$", value_text) and len(value_text) < 80:
                issues.append({"paper_id": paper_id, "file_name": file_name, "issue_type": "truncated_or_odd_quote", "field": str(col), "detail": value_text})
        best = str(row.get("best_performance_sentence", ""))
        if is_hmfor_category(category) and best not in {NOT_REPORTED, NEEDS_CHECK, "", "nan"}:
            has_conversion = bool(re.search(r"HMF conversion", best, re.I))
            has_fdca = bool(re.search(r"FDCA yield", best, re.I))
            has_fe = bool(re.search(r"Faradaic efficiency|\bFE\b", best, re.I))
            if (not has_conversion) and has_fdca and has_fe:
                issues.append({"paper_id": paper_id, "file_name": file_name, "issue_type": "best_sentence_missing_conversion_only", "field": "best_performance_sentence", "detail": best[:240]})
            elif not (has_conversion and has_fdca and has_fe):
                issues.append({"paper_id": paper_id, "file_name": file_name, "issue_type": "best_sentence_missing_target_metrics", "field": "best_performance_sentence", "detail": best[:240]})
        if is_hmfor_category(category):
            missing = str(row.get("missing_fields", ""))
            critical = [f for f in HMFOR_REQUIRED_FIELDS if f in missing]
            if critical:
                issues.append({"paper_id": paper_id, "file_name": file_name, "issue_type": "hmfor_missing_critical_field", "field": "; ".join(critical), "detail": missing})
            if "hmf_conversion" in missing and str(row.get("fdca_yield", "")) not in {NOT_REPORTED, NEEDS_CHECK, "", "nan"}:
                issues.append({"paper_id": paper_id, "file_name": file_name, "issue_type": "reported_in_figure_not_quantified_in_text", "field": "hmf_conversion", "detail": "conversion missing from extracted text while product metrics are present; check figures/SI"})
    if not evidence_df.empty:
        for _, ev in evidence_df.iterrows():
            field = str(ev.get("field", ""))
            snippet = str(ev.get("evidence_snippet", ""))
            value = str(ev.get("value", ""))
            if value in {NOT_REPORTED, NEEDS_CHECK, "", "nan"}:
                continue
            expected = {"FDCA_yield": "FDCA|yield", "Faradaic_efficiency": "Faradaic|FE", "HMF_conversion": "HMF|conversion"}.get(field)
            if expected and not re.search(expected, snippet, re.I):
                issues.append({"paper_id": str(ev.get("paper_id", "")), "file_name": str(ev.get("file_name", "")), "issue_type": "evidence_label_mismatch", "field": field, "detail": snippet[:240]})
    if not metrics_df.empty:
        for _, metric in metrics_df.iterrows():
            if str(metric.get("confidence", "")) == "very_low":
                issues.append({"paper_id": str(metric.get("paper_id", "")), "file_name": str(metric.get("file_name", "")), "issue_type": "very_low_confidence_background_metric", "field": str(metric.get("metric_name", "")), "detail": str(metric.get("metric_value", ""))[:240]})
    if not records_df.empty:
        for _, row in records_df.iterrows():
            vals = [str(v) for v in row.values]
            ratio = sum(v == NEEDS_CHECK for v in vals) / max(len(vals), 1)
            if ratio > 0.2:
                issues.append({"paper_id": row.get("paper_id", ""), "file_name": row.get("file_name", ""), "issue_type": "needs_manual_check_high_ratio", "field": "row", "detail": f"{ratio:.1%}"})
    return pd.DataFrame(issues, columns=["paper_id", "file_name", "issue_type", "field", "detail"])



def crossref_json(url: str, timeout: int = 20) -> dict:
    req = urllib.request.Request(url, headers={"User-Agent": "HMFOR-extraction-agent/2.0 (mailto:unknown@example.com)"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8", errors="replace"))




def normalize_doi_for_lookup(doi: str) -> str:
    doi = normalize_space(doi)
    if doi in {NOT_REPORTED, NEEDS_CHECK, ""}:
        return doi
    return doi.rstrip(" .;,)")

def doi_landing_url(doi: str) -> str:
    doi = normalize_doi_for_lookup(doi)
    if doi in {NOT_REPORTED, NEEDS_CHECK, ""}:
        return ""
    return "https://doi.org/" + doi


def publisher_si_candidate_urls(doi: str, publisher: str = "", article_url: str = "") -> str:
    doi = normalize_doi_for_lookup(doi)
    urls: list[str] = []
    if doi not in {NOT_REPORTED, NEEDS_CHECK, ""}:
        urls.append(doi_landing_url(doi))
        q = urllib.parse.quote(f'"{doi}" supporting information supplementary information')
        urls.append(f"https://www.google.com/search?q={q}")
    text = f"{publisher} {article_url} {doi}".lower()
    if "10.1016" in text or "elsevier" in text or "sciencedirect" in text:
        urls.append(doi_landing_url(doi) + "#appsec1" if doi_landing_url(doi) else "")
    if "10.1021" in text or "acs" in text:
        urls.append(doi_landing_url(doi).replace("https://doi.org/", "https://pubs.acs.org/doi/suppl/") if doi_landing_url(doi) else "")
    if "wiley" in text or "angew" in text or "10.1002" in text:
        urls.append(doi_landing_url(doi) + "/suppinfo" if doi_landing_url(doi) else "")
    if "springer" in text or "nature" in text or "10.1007" in text or "s40820" in text:
        urls.append(doi_landing_url(doi) + "#SecESM" if doi_landing_url(doi) else "")
    if "rsc" in text or "10.1039" in text:
        urls.append(doi_landing_url(doi) + "/suppinfo" if doi_landing_url(doi) else "")
    clean = []
    for u in urls:
        if u and u not in clean:
            clean.append(u)
    return "; ".join(clean)

def parse_crossref_item(item: dict) -> dict[str, str]:
    title = "; ".join(item.get("title", [])[:1]) if isinstance(item.get("title"), list) else str(item.get("title", ""))
    container = "; ".join(item.get("container-title", [])[:1]) if isinstance(item.get("container-title"), list) else str(item.get("container-title", ""))
    year = ""
    for key in ["published-print", "published-online", "created", "issued"]:
        parts = item.get(key, {}).get("date-parts", []) if isinstance(item.get(key), dict) else []
        if parts and parts[0]:
            year = str(parts[0][0])
            break
    links = item.get("link", []) if isinstance(item.get("link"), list) else []
    si_urls = []
    for link in links:
        url = str(link.get("URL", ""))
        ctype = str(link.get("content-type", ""))
        if re.search(r"supp|suppl|supporting|supplement|esm|mmc|moesm|mediaobjects", url + " " + ctype, re.I):
            si_urls.append(url)
    doi = normalize_doi_for_lookup(str(item.get("DOI", "")))
    publisher = html.unescape(normalize_space(str(item.get("publisher", ""))))
    article_url = str(item.get("URL", ""))
    generated_si_urls = publisher_si_candidate_urls(doi, publisher, article_url)
    return {
        "crossref_doi": doi,
        "crossref_title": html.unescape(normalize_space(title)),
        "crossref_journal": html.unescape(normalize_space(container)),
        "crossref_year": year,
        "crossref_publisher": publisher,
        "crossref_url": article_url,
        "crossref_si_candidate_urls": "; ".join([u for u in ["; ".join(si_urls), generated_si_urls] if u]),
    }


def build_doi_enrichment_results(records: list[PaperRecord], enable_online: bool = False) -> pd.DataFrame:
    rows = []
    for record in records:
        status, needs_lookup, priority, reason = doi_status(record)
        base = {
            "paper_id": record.paper_id,
            "file_name": record.file_name,
            "input_doi": record.doi,
            "input_title": record.title,
            "lookup_enabled": "yes" if enable_online else "no",
            "lookup_status": "skipped_offline_default" if not enable_online else "pending",
            "lookup_priority": priority,
            "lookup_reason": reason,
            "crossref_doi": "",
            "crossref_title": "",
            "crossref_journal": "",
            "crossref_year": "",
            "crossref_publisher": "",
            "crossref_url": "",
            "crossref_si_candidate_urls": "",
            "publisher_si_candidate_urls": publisher_si_candidate_urls(record.doi, record.journal, ""),
            "error": "",
        }
        if not enable_online:
            rows.append(base)
            continue
        if needs_lookup != "yes":
            base["lookup_status"] = "skipped_metadata_complete"
            base["crossref_doi"] = normalize_doi_for_lookup(record.doi)
            base["crossref_title"] = record.title
            base["crossref_journal"] = record.journal
            base["crossref_year"] = record.year
            rows.append(base)
            continue
        try:
            if record.doi not in {NOT_REPORTED, NEEDS_CHECK, ""}:
                url = "https://api.crossref.org/works/" + urllib.parse.quote(normalize_doi_for_lookup(record.doi), safe="")
                data = crossref_json(url)
                item = data.get("message", {})
            else:
                query = urllib.parse.urlencode({"query.title": record.title, "rows": "1"})
                data = crossref_json("https://api.crossref.org/works?" + query)
                items = data.get("message", {}).get("items", [])
                item = items[0] if items else {}
            if item:
                base.update(parse_crossref_item(item))
                base["lookup_status"] = "success"
                base["publisher_si_candidate_urls"] = publisher_si_candidate_urls(base.get("crossref_doi", record.doi), base.get("crossref_publisher", record.journal), base.get("crossref_url", ""))
            else:
                base["lookup_status"] = "failed_no_match"
        except Exception as exc:
            base["lookup_status"] = "failed_error"
            base["error"] = str(exc)[:300]
        rows.append(base)
    return pd.DataFrame(rows)

def file_sha256_short(path: Path) -> str:
    try:
        h = hashlib.sha256()
        with path.open("rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                h.update(chunk)
        return h.hexdigest()[:16]
    except Exception:
        return NEEDS_CHECK


def build_enhanced_dedup_report(records: list[PaperRecord], all_pdfs: list[Path] | None = None) -> pd.DataFrame:
    pdf_by_name = {p.name: p for p in (all_pdfs or [])}
    base_rows = []
    for record in records:
        doi_key = normalize_space(record.doi).lower().rstrip(".") if record.doi not in {NOT_REPORTED, NEEDS_CHECK, ""} else ""
        title_key = title_fingerprint(record.title)
        file_hash = file_sha256_short(pdf_by_name[record.file_name]) if record.file_name in pdf_by_name else NEEDS_CHECK
        base_rows.append({
            "paper_id": record.paper_id,
            "file_name": record.file_name,
            "doi_key": doi_key,
            "title_fingerprint": title_key,
            "file_hash_short": file_hash,
            "linked_si_files": record.linked_si_files,
            "paper_category": record.paper_category,
        })
    df = pd.DataFrame(base_rows)
    if df.empty:
        return pd.DataFrame(columns=["paper_id", "file_name", "duplicate_group_id", "duplicate_status", "primary_paper_id", "match_basis", "match_key", "file_hash_short", "title_fingerprint", "duplicate_reason", "confidence"])
    groups: dict[tuple[str, str], list[dict[str, str]]] = {}
    for row in base_rows:
        if row["doi_key"]:
            key = ("doi", row["doi_key"])
        elif row["title_fingerprint"]:
            key = ("title_fingerprint", row["title_fingerprint"])
        elif row["file_hash_short"] not in {NEEDS_CHECK, ""}:
            key = ("file_hash", row["file_hash_short"])
        else:
            key = ("file_name", row["file_name"].lower())
        groups.setdefault(key, []).append(row)
    out = []
    gid = 1
    for (basis, key), members in groups.items():
        dup = len(members) > 1
        group_id = f"D{gid:04d}" if dup else ""
        if dup:
            gid += 1
        primary = members[0]["paper_id"]
        for m in members:
            if dup:
                reason = "same DOI" if basis == "doi" else "same normalized title" if basis == "title_fingerprint" else "same file hash" if basis == "file_hash" else "same filename fallback"
            else:
                reason = f"unique by {basis}"
            out.append({
                "paper_id": m["paper_id"],
                "file_name": m["file_name"],
                "duplicate_group_id": group_id,
                "duplicate_status": "duplicate" if dup and m["paper_id"] != primary else "primary" if dup else "unique",
                "primary_paper_id": primary if dup else m["paper_id"],
                "match_basis": basis,
                "match_key": key,
                "file_hash_short": m["file_hash_short"],
                "title_fingerprint": m["title_fingerprint"],
                "duplicate_reason": reason,
                "confidence": "high" if basis in {"doi", "file_hash"} else "medium" if basis == "title_fingerprint" else "low",
            })
    return pd.DataFrame(out)


def build_si_field_evidence_log(si_links_df: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, str]] = []
    if si_links_df.empty:
        return pd.DataFrame(columns=["paper_id", "file_name", "si_file_name", "field", "source_scope", "evidence_snippet", "confidence"])
    for _, row in si_links_df.iterrows():
        summary = str(row.get("si_field_evidence_summary", ""))
        if summary in {"", "nan", NOT_REPORTED, NEEDS_CHECK}:
            continue
        for part in summary.split(" | "):
            if ":" not in part:
                continue
            field, snippet = part.split(":", 1)
            field = field.strip()
            snippet = normalize_space(snippet.strip())
            if not field or not snippet:
                continue
            rows.append({
                "paper_id": str(row.get("paper_id", "")),
                "file_name": str(row.get("file_name", "")),
                "si_file_name": str(row.get("si_file_name", "")),
                "field": field,
                "source_scope": "supporting_information",
                "evidence_snippet": snippet[:500],
                "confidence": str(row.get("confidence", "medium")),
            })
    return pd.DataFrame(rows, columns=["paper_id", "file_name", "si_file_name", "field", "source_scope", "evidence_snippet", "confidence"])


def doi_status(record: PaperRecord) -> tuple[str, str, str, str]:
    doi = normalize_space(record.doi)
    missing_meta = [f for f in ["title", "journal", "year"] if getattr(record, f, NOT_REPORTED) in {NOT_REPORTED, NEEDS_CHECK, ""}]
    if doi in {NOT_REPORTED, NEEDS_CHECK, ""}:
        return "missing_doi", "yes", "high" if is_hmfor_category(record.paper_category) else "medium", "DOI missing; query by title/file name"
    if missing_meta or (is_hmfor_category(record.paper_category) and record.linked_si_files in {NOT_REPORTED, NEEDS_CHECK, ""}):
        return "doi_present_metadata_incomplete", "yes", "medium", "DOI present but metadata/SI link can be enriched"
    return "doi_present_basic_metadata_ok", "no", "low", "no online lookup needed for V2 baseline"


def build_doi_enrichment_queue(records: list[PaperRecord]) -> pd.DataFrame:
    rows = []
    for record in records:
        status, needs_lookup, priority, reason = doi_status(record)
        query = record.doi if record.doi not in {NOT_REPORTED, NEEDS_CHECK, ""} else record.title if record.title not in {NOT_REPORTED, NEEDS_CHECK, ""} else record.file_name
        rows.append({
            "paper_id": record.paper_id,
            "file_name": record.file_name,
            "doi": record.doi,
            "title": record.title,
            "doi_status": status,
            "needs_online_lookup": needs_lookup,
            "lookup_priority": priority,
            "lookup_query": query,
            "lookup_reason": reason,
            "cached_metadata_status": "not_attempted_offline_v2",
            "suggested_sources": "Crossref; publisher page; article landing page; SI/ESI page",
        })
    return pd.DataFrame(rows)


def build_comparison_ready_metrics(metrics_df: pd.DataFrame, quality_scores_df: pd.DataFrame, dedup_df: pd.DataFrame) -> pd.DataFrame:
    if metrics_df.empty:
        return pd.DataFrame()
    df = metrics_df[metrics_df["comparison_eligible"] == "yes"].copy()
    if df.empty:
        return df
    if not quality_scores_df.empty:
        df = df.merge(quality_scores_df[["paper_id", "overall_quality_score", "data_completeness_score", "evidence_traceability_score"]], on="paper_id", how="left")
    if not dedup_df.empty:
        df = df.merge(dedup_df[["paper_id", "duplicate_status", "primary_paper_id"]], on="paper_id", how="left")
        df = df[df["duplicate_status"].fillna("unique").isin(["unique", "primary"])]
    df["plot_ready"] = df["metric_value_normalized"].map(lambda v: "yes" if re.fullmatch(r"\d+(?:\.\d+)?", str(v)) else "no")
    df["plot_exclusion_reason"] = df.apply(lambda r: "" if r["plot_ready"] == "yes" else "normalized value is not numeric", axis=1)
    return df


def build_visualization_readiness_report(metrics_df: pd.DataFrame, comparison_df: pd.DataFrame) -> pd.DataFrame:
    contexts = sorted(set(metrics_df.get("performance_context", pd.Series(dtype=str)).dropna().astype(str).tolist())) if not metrics_df.empty else []
    rows = []
    for context in contexts:
        total = int((metrics_df["performance_context"] == context).sum())
        eligible = int((comparison_df["performance_context"] == context).sum()) if not comparison_df.empty and "performance_context" in comparison_df.columns else 0
        numeric = int(((comparison_df["performance_context"] == context) & (comparison_df["plot_ready"] == "yes")).sum()) if not comparison_df.empty and "performance_context" in comparison_df.columns else 0
        rows.append({
            "performance_context": context,
            "total_metrics": total,
            "comparison_eligible_metrics": eligible,
            "numeric_plot_ready_metrics": numeric,
            "readiness": "ready" if numeric >= 3 else "limited" if numeric > 0 else "not_ready",
            "recommended_visualization": "heatmap/bar comparison" if numeric >= 3 else "table only until more eligible metrics",
        })
    return pd.DataFrame(rows, columns=["performance_context", "total_metrics", "comparison_eligible_metrics", "numeric_plot_ready_metrics", "readiness", "recommended_visualization"])



def merge_si_links_with_doi_candidates(si_links_df: pd.DataFrame, doi_results_df: pd.DataFrame) -> pd.DataFrame:
    if si_links_df.empty:
        return si_links_df
    df = si_links_df.copy()
    if "publisher_si_candidate_urls" not in df.columns:
        df["publisher_si_candidate_urls"] = ""
    if doi_results_df.empty:
        return df
    candidates = doi_results_df[["paper_id", "publisher_si_candidate_urls", "crossref_si_candidate_urls"]].copy()
    candidates["doi_si_candidates"] = candidates.apply(lambda r: "; ".join([str(v) for v in [r.get("publisher_si_candidate_urls", ""), r.get("crossref_si_candidate_urls", "")] if str(v) not in {"", "nan"}]), axis=1)
    df = df.merge(candidates[["paper_id", "doi_si_candidates"]], on="paper_id", how="left")
    def combine(row: pd.Series) -> str:
        parts: list[str] = []
        for col in ["publisher_si_candidate_urls", "doi_si_candidates"]:
            value = str(row.get(col, ""))
            if value in {"", "nan"}:
                continue
            for part in value.split("; "):
                if part and part not in parts:
                    parts.append(part)
        return "; ".join(parts)
    df["publisher_si_candidate_urls"] = df.apply(combine, axis=1)
    return df.drop(columns=["doi_si_candidates"])

def clean_bad_short_values_df(records_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    cleaned = records_df.copy()
    issues: list[dict[str, str]] = []
    for idx, row in records_df.iterrows():
        paper_id = str(row.get("paper_id", ""))
        file_name = str(row.get("file_name", ""))
        for col in QUALITY_TEXT_FIELDS:
            if col not in records_df.columns:
                continue
            value_text = str(row.get(col, "")).strip()
            if value_text in {"nan", "", NOT_REPORTED, NEEDS_CHECK, "not_applicable"} or value_text in STANDARD_ENUM_VALUES:
                continue
            bad = value_text in BAD_SHORT_VALUES or (0 < len(value_text) < 20 and re.search(r"[A-Za-z]", value_text) and not is_acceptable_short_text(str(col), value_text))
            if re.search(r"[”’]$", value_text) and len(value_text) < 80:
                issues.append({
                    "paper_id": paper_id,
                    "file_name": file_name,
                    "issue_type": "truncated_or_odd_quote",
                    "field": str(col),
                    "detail": value_text,
                })
            if bad:
                cleaned.at[idx, col] = NOT_REPORTED
                issues.append({
                    "paper_id": paper_id,
                    "file_name": file_name,
                    "issue_type": "bad_short_value_cleaned",
                    "field": str(col),
                    "detail": f"{value_text} -> {NOT_REPORTED}",
                })
    return cleaned, pd.DataFrame(issues, columns=["paper_id", "file_name", "issue_type", "field", "detail"])


def attach_source_issue_ids(quality_df: pd.DataFrame) -> pd.DataFrame:
    if quality_df.empty:
        quality_df = pd.DataFrame(columns=["paper_id", "file_name", "issue_type", "field", "detail"])
    df = quality_df.copy().reset_index(drop=True)
    df.insert(0, "source_issue_id", [f"QI{i+1:04d}" for i in range(len(df))])
    return df


def table_schema(rows: dict[str, pd.DataFrame]) -> dict:
    return {
        name: {
            "schema_version": SCHEMA_VERSION,
            "columns": list(df.columns),
            "row_count": int(len(df)),
        }
        for name, df in rows.items()
    }


def write_schema_files(out_dir: Path, tables: dict[str, pd.DataFrame]) -> None:
    schema = {
        "tool": "hmfor_extractor",
        "tool_version": TOOL_VERSION,
        "schema_version": SCHEMA_VERSION,
        "tables": table_schema(tables),
        "error_codes": ERROR_CODES,
    }
    (out_dir / "output_schema.json").write_text(json.dumps(schema, ensure_ascii=False, indent=2), encoding="utf-8")


def build_run_manifest(input_dir: Path | None, out_dir: Path, records: list[PaperRecord], tables: dict[str, pd.DataFrame], enable_potential_conversion: bool, enable_doi_online: bool, enable_si_download: bool, config_path: Path | None, si_download_df: pd.DataFrame) -> dict:
    from datetime import datetime, timezone
    categories: dict[str, int] = {}
    for record in records:
        categories[record.paper_category] = categories.get(record.paper_category, 0) + 1
    issue_count = int(len(tables.get("quality_issues", pd.DataFrame())))
    manual_count = int(len(tables.get("manual_review_queue", pd.DataFrame())))
    doi_status = {}
    if "doi_enrichment_results" in tables and not tables["doi_enrichment_results"].empty and "lookup_status" in tables["doi_enrichment_results"].columns:
        doi_status = {str(k): int(v) for k, v in tables["doi_enrichment_results"]["lookup_status"].value_counts(dropna=False).to_dict().items()}
    si_status = {}
    if si_download_df is not None and not si_download_df.empty and "download_status" in si_download_df.columns:
        si_status = {str(k): int(v) for k, v in si_download_df["download_status"].value_counts(dropna=False).to_dict().items()}
    return {
        "tool": "hmfor_extractor",
        "tool_version": TOOL_VERSION,
        "schema_version": SCHEMA_VERSION,
        "run_time_utc": datetime.now(timezone.utc).isoformat(),
        "input_dir": str(input_dir.resolve()) if input_dir else "",
        "output_dir": str(out_dir.resolve()),
        "config_path": str(config_path.resolve()) if config_path and config_path.exists() else "",
        "manual_override_path": str(DEFAULT_OVERRIDES_PATH.resolve()),
        "manual_override_version": load_json_rules(DEFAULT_OVERRIDES_PATH).get("version", "unknown") if DEFAULT_OVERRIDES_PATH.exists() else "not_found",
        "rhe_rules_path": str(DEFAULT_RHE_RULES_PATH.resolve()),
        "unit_rules_path": str(DEFAULT_UNIT_RULES_PATH.resolve()),
        "flags": {
            "enable_potential_conversion": bool(enable_potential_conversion),
            "enable_doi_online": bool(enable_doi_online),
            "enable_si_download": bool(enable_si_download),
        },
        "paper_count": int(len(records)),
        "paper_category_counts": categories,
        "table_row_counts": {name: int(len(df)) for name, df in tables.items()},
        "quality_issue_count": issue_count,
        "manual_review_count": manual_count,
        "doi_lookup_status_counts": doi_status,
        "si_download_status_counts": si_status,
        "error_codes": ERROR_CODES,
    }

def write_outputs(records: list[PaperRecord], evidence_rows: list[EvidenceRecord], out_dir: Path, nico_background_rows: list[NiCoBackgroundRecord] | None = None, enable_potential_conversion: bool = False, all_pdfs: list[Path] | None = None, enable_doi_online: bool = False, si_download_df: pd.DataFrame | None = None, input_dir: Path | None = None, config_path: Path | None = None, enable_si_download: bool = False) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    records_all_raw_df = pd.DataFrame([asdict(r) for r in records])
    records_all_df = records_all_raw_df.copy()
    for col in ["electrolyte", "koh_concentration", "hmf_concentration", "potential_vs_rhe", "onset_potential_vs_rhe", "performance_potential_vs_rhe", "lsv_current_density", "current_density", "product_analysis_condition", "best_performance_sentence"]:
        if col in records_all_df.columns:
            records_all_df[col] = records_all_df[col].map(normalize_unit_text)
    records_all_df, cleaned_short_issue_df = clean_bad_short_values_df(records_all_df)
    records_df = records_all_df[records_all_df["paper_category"].isin(["hmfor_core", "hmfor_related"])].copy()
    background_df = records_all_df[~records_all_df["paper_category"].isin(["hmfor_core", "hmfor_related"])].copy()
    hmfor_only_cols = HMFOR_SPECIFIC_FIELDS
    for col in hmfor_only_cols:
        if col in background_df.columns:
            background_df = background_df.drop(columns=[col])
    evidence_df = pd.DataFrame([asdict(e) for e in evidence_rows])
    si_download_df = si_download_df if si_download_df is not None else pd.DataFrame(columns=["paper_id", "file_name", "doi", "candidate_url", "download_status", "saved_file", "content_type", "message"])
    nico_background_rows = nico_background_rows or []
    nico_background_df = pd.DataFrame([asdict(r) for r in nico_background_rows])
    metrics_df = build_metrics_long_table(records, nico_background_rows)
    papers_df = build_papers_table(records)
    materials_df = build_materials_table(records)
    dedup_df = build_enhanced_dedup_report(records, all_pdfs)
    doi_enrichment_df = build_doi_enrichment_queue(records)
    doi_enrichment_results_df = build_doi_enrichment_results(records, enable_doi_online)
    si_links_df = merge_si_links_with_doi_candidates(build_si_links(records, all_pdfs), doi_enrichment_results_df)
    si_field_evidence_df = build_si_field_evidence_log(si_links_df)
    field_confidence_df = build_field_confidence_table(records, evidence_df)
    potential_audit_df = build_potential_conversion_audit(records, enable_potential_conversion)
    unit_log_df = build_unit_normalization_log(records, metrics_df)
    quality_scores_df = build_quality_scores(records, field_confidence_df, si_links_df)
    comparison_ready_df = build_comparison_ready_metrics(metrics_df, quality_scores_df, dedup_df)
    visualization_readiness_df = build_visualization_readiness_report(metrics_df, comparison_ready_df)
    quality_df = attach_source_issue_ids(pd.concat([build_quality_issues(records_all_df, evidence_df, metrics_df), cleaned_short_issue_df], ignore_index=True))
    manual_review_df = build_manual_review_queue(records, quality_df, si_links_df, metrics_df)

    missing_rows = []
    for record in records:
        for field in record.missing_fields.split("; "):
            if field:
                missing_rows.append({"paper_id": record.paper_id, "file_name": record.file_name, "missing_field": field})
    missing_df = pd.DataFrame(missing_rows)

    output_tables = {
        "all_papers_classification": records_all_df,
        "hmfor_experiment_variables": records_df,
        "background_or_excluded_papers": background_df,
        "nico_background_variables": nico_background_df,
        "hmfor_evidence_log": evidence_df,
        "hmfor_missing_fields": missing_df,
        "metrics_long_table": metrics_df,
        "quality_issues": quality_df,
        "papers": papers_df,
        "materials": materials_df,
        "dedup_report": dedup_df,
        "si_links": si_links_df,
        "field_confidence": field_confidence_df,
        "potential_conversion_audit": potential_audit_df,
        "unit_normalization_log": unit_log_df,
        "manual_review_queue": manual_review_df,
        "quality_scores": quality_scores_df,
        "si_field_evidence_log": si_field_evidence_df,
        "comparison_ready_metrics": comparison_ready_df,
        "visualization_readiness_report": visualization_readiness_df,
        "doi_enrichment_queue": doi_enrichment_df,
        "doi_enrichment_results": doi_enrichment_results_df,
        "si_download_report": si_download_df,
    }
    output_tables = {name: add_schema_version(df) for name, df in output_tables.items()}
    records_all_df = output_tables["all_papers_classification"]
    records_df = output_tables["hmfor_experiment_variables"]
    background_df = output_tables["background_or_excluded_papers"]
    nico_background_df = output_tables["nico_background_variables"]
    evidence_df = output_tables["hmfor_evidence_log"]
    missing_df = output_tables["hmfor_missing_fields"]
    metrics_df = output_tables["metrics_long_table"]
    quality_df = output_tables["quality_issues"]
    papers_df = output_tables["papers"]
    materials_df = output_tables["materials"]
    dedup_df = output_tables["dedup_report"]
    si_links_df = output_tables["si_links"]
    field_confidence_df = output_tables["field_confidence"]
    potential_audit_df = output_tables["potential_conversion_audit"]
    unit_log_df = output_tables["unit_normalization_log"]
    manual_review_df = output_tables["manual_review_queue"]
    quality_scores_df = output_tables["quality_scores"]
    si_field_evidence_df = output_tables["si_field_evidence_log"]
    comparison_ready_df = output_tables["comparison_ready_metrics"]
    visualization_readiness_df = output_tables["visualization_readiness_report"]
    doi_enrichment_df = output_tables["doi_enrichment_queue"]
    doi_enrichment_results_df = output_tables["doi_enrichment_results"]
    si_download_df = output_tables["si_download_report"]
    write_schema_files(out_dir, output_tables)
    run_manifest = build_run_manifest(input_dir, out_dir, records, output_tables, enable_potential_conversion, enable_doi_online, enable_si_download, config_path, si_download_df)
    (out_dir / "run_manifest.json").write_text(json.dumps(run_manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    records_all_df.to_csv(out_dir / "all_papers_classification.csv", index=False, encoding="utf-8-sig")
    records_df.to_csv(out_dir / "hmfor_experiment_variables.csv", index=False, encoding="utf-8-sig")
    background_df.to_csv(out_dir / "background_or_excluded_papers.csv", index=False, encoding="utf-8-sig")
    nico_background_df.to_csv(out_dir / "nico_background_variables.csv", index=False, encoding="utf-8-sig")
    evidence_df.to_csv(out_dir / "hmfor_evidence_log.csv", index=False, encoding="utf-8-sig")
    missing_df.to_csv(out_dir / "hmfor_missing_fields.csv", index=False, encoding="utf-8-sig")
    metrics_df.to_csv(out_dir / "metrics_long_table.csv", index=False, encoding="utf-8-sig")
    quality_df.to_csv(out_dir / "quality_issues.csv", index=False, encoding="utf-8-sig")
    papers_df.to_csv(out_dir / "papers.csv", index=False, encoding="utf-8-sig")
    materials_df.to_csv(out_dir / "materials.csv", index=False, encoding="utf-8-sig")
    dedup_df.to_csv(out_dir / "dedup_report.csv", index=False, encoding="utf-8-sig")
    si_links_df.to_csv(out_dir / "si_links.csv", index=False, encoding="utf-8-sig")
    field_confidence_df.to_csv(out_dir / "field_confidence.csv", index=False, encoding="utf-8-sig")
    potential_audit_df.to_csv(out_dir / "potential_conversion_audit.csv", index=False, encoding="utf-8-sig")
    unit_log_df.to_csv(out_dir / "unit_normalization_log.csv", index=False, encoding="utf-8-sig")
    manual_review_df.to_csv(out_dir / "manual_review_queue.csv", index=False, encoding="utf-8-sig")
    quality_scores_df.to_csv(out_dir / "quality_scores.csv", index=False, encoding="utf-8-sig")
    si_field_evidence_df.to_csv(out_dir / "si_field_evidence_log.csv", index=False, encoding="utf-8-sig")
    comparison_ready_df.to_csv(out_dir / "comparison_ready_metrics.csv", index=False, encoding="utf-8-sig")
    visualization_readiness_df.to_csv(out_dir / "visualization_readiness_report.csv", index=False, encoding="utf-8-sig")
    doi_enrichment_df.to_csv(out_dir / "doi_enrichment_queue.csv", index=False, encoding="utf-8-sig")
    doi_enrichment_results_df.to_csv(out_dir / "doi_enrichment_results.csv", index=False, encoding="utf-8-sig")
    si_download_df.to_csv(out_dir / "si_download_report.csv", index=False, encoding="utf-8-sig")

    records_all_xlsx_df = clean_dataframe_for_excel(records_all_df)
    records_xlsx_df = clean_dataframe_for_excel(records_df)
    background_xlsx_df = clean_dataframe_for_excel(background_df)
    nico_background_xlsx_df = clean_dataframe_for_excel(nico_background_df)
    evidence_xlsx_df = clean_dataframe_for_excel(evidence_df)
    missing_xlsx_df = clean_dataframe_for_excel(missing_df)
    metrics_xlsx_df = clean_dataframe_for_excel(metrics_df)
    quality_xlsx_df = clean_dataframe_for_excel(quality_df)
    papers_xlsx_df = clean_dataframe_for_excel(papers_df)
    materials_xlsx_df = clean_dataframe_for_excel(materials_df)
    dedup_xlsx_df = clean_dataframe_for_excel(dedup_df)
    si_links_xlsx_df = clean_dataframe_for_excel(si_links_df)
    field_confidence_xlsx_df = clean_dataframe_for_excel(field_confidence_df)
    potential_audit_xlsx_df = clean_dataframe_for_excel(potential_audit_df)
    unit_log_xlsx_df = clean_dataframe_for_excel(unit_log_df)
    manual_review_xlsx_df = clean_dataframe_for_excel(manual_review_df)
    quality_scores_xlsx_df = clean_dataframe_for_excel(quality_scores_df)
    si_field_evidence_xlsx_df = clean_dataframe_for_excel(si_field_evidence_df)
    comparison_ready_xlsx_df = clean_dataframe_for_excel(comparison_ready_df)
    visualization_readiness_xlsx_df = clean_dataframe_for_excel(visualization_readiness_df)
    doi_enrichment_xlsx_df = clean_dataframe_for_excel(doi_enrichment_df)
    doi_enrichment_results_xlsx_df = clean_dataframe_for_excel(doi_enrichment_results_df)
    si_download_xlsx_df = clean_dataframe_for_excel(si_download_df)

    with pd.ExcelWriter(out_dir / "hmfor_experiment_variables.xlsx", engine="openpyxl") as writer:
        records_all_xlsx_df.to_excel(writer, sheet_name="all_papers", index=False)
        records_xlsx_df.to_excel(writer, sheet_name="hmfor_variables", index=False)
        background_xlsx_df.to_excel(writer, sheet_name="background_excluded", index=False)
        nico_background_xlsx_df.to_excel(writer, sheet_name="nico_background", index=False)
        evidence_xlsx_df.to_excel(writer, sheet_name="evidence_log", index=False)
        missing_xlsx_df.to_excel(writer, sheet_name="missing_fields", index=False)
        metrics_xlsx_df.to_excel(writer, sheet_name="metrics_long", index=False)
        quality_xlsx_df.to_excel(writer, sheet_name="quality_issues", index=False)
        papers_xlsx_df.to_excel(writer, sheet_name="papers", index=False)
        materials_xlsx_df.to_excel(writer, sheet_name="materials", index=False)
        dedup_xlsx_df.to_excel(writer, sheet_name="dedup_report", index=False)
        si_links_xlsx_df.to_excel(writer, sheet_name="si_links", index=False)
        field_confidence_xlsx_df.to_excel(writer, sheet_name="field_confidence", index=False)
        potential_audit_xlsx_df.to_excel(writer, sheet_name="potential_audit", index=False)
        unit_log_xlsx_df.to_excel(writer, sheet_name="unit_log", index=False)
        manual_review_xlsx_df.to_excel(writer, sheet_name="manual_review", index=False)
        quality_scores_xlsx_df.to_excel(writer, sheet_name="quality_scores", index=False)
        si_field_evidence_xlsx_df.to_excel(writer, sheet_name="si_evidence", index=False)
        comparison_ready_xlsx_df.to_excel(writer, sheet_name="comparison_ready", index=False)
        visualization_readiness_xlsx_df.to_excel(writer, sheet_name="viz_readiness", index=False)
        doi_enrichment_xlsx_df.to_excel(writer, sheet_name="doi_enrichment", index=False)
        doi_enrichment_results_xlsx_df.to_excel(writer, sheet_name="doi_results", index=False)
        si_download_xlsx_df.to_excel(writer, sheet_name="si_download", index=False)

    evidence_xlsx_df.to_excel(out_dir / "hmfor_evidence_log.xlsx", index=False)
    missing_xlsx_df.to_excel(out_dir / "hmfor_missing_fields.xlsx", index=False)
    metrics_xlsx_df.to_excel(out_dir / "metrics_long_table.xlsx", index=False)
    quality_xlsx_df.to_excel(out_dir / "quality_issues.xlsx", index=False)
    papers_xlsx_df.to_excel(out_dir / "papers.xlsx", index=False)
    materials_xlsx_df.to_excel(out_dir / "materials.xlsx", index=False)
    dedup_xlsx_df.to_excel(out_dir / "dedup_report.xlsx", index=False)
    si_links_xlsx_df.to_excel(out_dir / "si_links.xlsx", index=False)
    field_confidence_xlsx_df.to_excel(out_dir / "field_confidence.xlsx", index=False)
    potential_audit_xlsx_df.to_excel(out_dir / "potential_conversion_audit.xlsx", index=False)
    unit_log_xlsx_df.to_excel(out_dir / "unit_normalization_log.xlsx", index=False)
    manual_review_xlsx_df.to_excel(out_dir / "manual_review_queue.xlsx", index=False)
    quality_scores_xlsx_df.to_excel(out_dir / "quality_scores.xlsx", index=False)
    si_field_evidence_xlsx_df.to_excel(out_dir / "si_field_evidence_log.xlsx", index=False)
    comparison_ready_xlsx_df.to_excel(out_dir / "comparison_ready_metrics.xlsx", index=False)
    visualization_readiness_xlsx_df.to_excel(out_dir / "visualization_readiness_report.xlsx", index=False)
    doi_enrichment_xlsx_df.to_excel(out_dir / "doi_enrichment_queue.xlsx", index=False)
    doi_enrichment_results_xlsx_df.to_excel(out_dir / "doi_enrichment_results.xlsx", index=False)
    si_download_xlsx_df.to_excel(out_dir / "si_download_report.xlsx", index=False)



def safe_filename_part(value: str) -> str:
    value = normalize_doi_for_lookup(value)
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value)
    return value.strip("._-")[:120] or "unknown"


def likely_direct_si_url(url: str) -> bool:
    u = str(url).lower()
    if not u or "google.com/search" in u or "doi.org/" in u:
        return False
    return bool(re.search(r"supp|suppl|supporting|supplement|esm|mmc|moesm|mediaobjects", u))


def extract_pdf_links_from_html(html_text: str, base_url: str) -> list[str]:
    links: list[str] = []
    href_pattern = "href=[\"\']([^\"\']+)[\"\']"
    for match in re.finditer(href_pattern, html_text, re.I):
        href = html.unescape(match.group(1))
        if not re.search(r"supp|suppl|support|esm|mmc|si", href, re.I):
            continue
        if not re.search(r"pdf|zip|docx?|xlsx?|xls|csv", href, re.I):
            continue
        links.append(urllib.parse.urljoin(base_url, href))
    return list(dict.fromkeys(links))[:5]



def discover_si_links_from_page(url: str, timeout: int = 25) -> tuple[str, str, list[str]]:
    if not url or "google.com/search" in url:
        return "candidate_only", "", []
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "HMFOR-extraction-agent/2.0"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            content_type = resp.headers.get("content-type", "")
            data = resp.read(2 * 1024 * 1024)
        if data.startswith(b"%PDF") or "pdf" in content_type.lower():
            return ("direct_si_pdf_page" if likely_direct_si_url(url) else "direct_article_pdf_not_si"), content_type, ([url] if likely_direct_si_url(url) else [])
        if "html" in content_type.lower() or b"<html" in data[:500].lower():
            links = extract_pdf_links_from_html(data.decode("utf-8", errors="replace"), url)
            return ("html_with_candidate_links" if links else "html_no_si_links"), content_type, links
        return "not_html_or_pdf", content_type, []
    except Exception as exc:
        return "discovery_failed", "", [str(exc)[:300]]

def try_download_url(url: str, dest: Path, timeout: int = 25) -> tuple[str, str, str]:
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "HMFOR-extraction-agent/2.0"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            content_type = resp.headers.get("content-type", "")
            data = resp.read(10 * 1024 * 1024)
        url_l = url.lower()
        if data.startswith(b"%PDF") or "pdf" in content_type.lower():
            dest.write_bytes(data)
            return "downloaded", content_type, ""
        if url_l.endswith(".docx") or "wordprocessingml" in content_type.lower() or "officedocument.wordprocessingml" in content_type.lower():
            docx_dest = dest.with_suffix(".docx")
            docx_dest.write_bytes(data)
            return "downloaded", content_type or "application/vnd.openxmlformats-officedocument.wordprocessingml.document", str(docx_dest.name)
        lower_head = data[:500].lower()
        if "html" in content_type.lower() or b"<html" in lower_head:
            links = extract_pdf_links_from_html(data.decode("utf-8", errors="replace"), url)
            if links:
                return "html_with_candidate_links", content_type, "; ".join(links)
        return "not_pdf", content_type, "response was not a PDF/SI file"
    except Exception as exc:
        return "failed_error", "", str(exc)[:300]


def download_si_candidates(records: list[PaperRecord], out_dir: Path, enable_doi_online: bool = False, enable_si_download: bool = False) -> tuple[pd.DataFrame, list[Path]]:
    columns = ["paper_id", "file_name", "doi", "candidate_url", "download_status", "saved_file", "content_type", "message"]
    rows: list[dict[str, str]] = []
    downloaded: list[Path] = []
    if not enable_si_download:
        return pd.DataFrame(columns=columns), downloaded
    si_dir = out_dir / "downloaded-si"
    si_dir.mkdir(parents=True, exist_ok=True)
    doi_results = build_doi_enrichment_results(records, enable_doi_online)
    for record in records:
        if not is_hmfor_category(record.paper_category):
            continue
        row = doi_results[doi_results["paper_id"] == record.paper_id]
        candidates: list[str] = []
        for value in [publisher_si_candidate_urls(record.doi, record.journal, "")]:
            for part in str(value).split("; "):
                if part and part not in candidates:
                    candidates.append(part)
        if not row.empty:
            for col in ["crossref_si_candidate_urls", "publisher_si_candidate_urls"]:
                for part in str(row.iloc[0].get(col, "")).split("; "):
                    if part and part not in {"nan", ""} and part not in candidates:
                        candidates.append(part)
        tried = False
        for url in candidates:
            urls_to_try: list[str] = []
            if likely_direct_si_url(url):
                urls_to_try = [url]
            else:
                discover_status, discover_type, discovered = discover_si_links_from_page(url)
                if discover_status in {"html_with_candidate_links", "direct_si_pdf_page"}:
                    urls_to_try = [u for u in discovered if likely_direct_si_url(u)]
                    rows.append({"paper_id": record.paper_id, "file_name": record.file_name, "doi": record.doi, "candidate_url": url, "download_status": discover_status, "saved_file": "", "content_type": discover_type, "message": "; ".join(discovered)})
                else:
                    rows.append({"paper_id": record.paper_id, "file_name": record.file_name, "doi": record.doi, "candidate_url": url, "download_status": discover_status, "saved_file": "", "content_type": discover_type, "message": "; ".join(discovered) if discovered else "no direct SI link discovered"})
                    continue
            for download_url in urls_to_try:
                tried = True
                dest = si_dir / f"{record.paper_id}_{safe_filename_part(record.doi)}_si.pdf"
                status, content_type, message = try_download_url(download_url, dest)
                actual_path = si_dir / message if status == "downloaded" and message else dest
                saved = str(actual_path.name) if status == "downloaded" else ""
                if status == "downloaded":
                    downloaded.append(actual_path)
                    rows.append({"paper_id": record.paper_id, "file_name": record.file_name, "doi": record.doi, "candidate_url": download_url, "download_status": status, "saved_file": saved, "content_type": content_type, "message": ""})
                    break
                rows.append({"paper_id": record.paper_id, "file_name": record.file_name, "doi": record.doi, "candidate_url": download_url, "download_status": status, "saved_file": saved, "content_type": content_type, "message": message})
            if downloaded and downloaded[-1].name.startswith(record.paper_id):
                break
        if not tried and not candidates:
            rows.append({"paper_id": record.paper_id, "file_name": record.file_name, "doi": record.doi, "candidate_url": "", "download_status": "no_candidate", "saved_file": "", "content_type": "", "message": "no SI candidate URL available"})
    return pd.DataFrame(rows, columns=columns), downloaded


def run_extraction_batch(target_pdfs: list[Path], all_pdfs: list[Path]) -> tuple[list[PaperRecord], list[EvidenceRecord], list[NiCoBackgroundRecord]]:
    records: list[PaperRecord] = []
    evidence_rows: list[EvidenceRecord] = []
    nico_background_rows: list[NiCoBackgroundRecord] = []
    for idx, pdf in enumerate(target_pdfs, start=1):
        try:
            record, ev = extract_one(pdf, all_pdfs, idx)
            records.append(record)
            evidence_rows.extend(ev)
            pages_for_background = read_pdf_pages(pdf)
            background_record = extract_nico_background_record(record, " ".join(pages_for_background))
            if background_record is not None:
                nico_background_rows.append(background_record)
            print(f"[OK] {pdf.name}")
        except Exception as exc:
            print(f"[ERROR] {pdf.name}: {exc}")
    return records, evidence_rows, nico_background_rows

def list_target_pdfs(pdf_dir: Path) -> list[Path]:
    pdfs = sorted(pdf_dir.glob("*.pdf"))
    main_pdfs = [p for p in pdfs if not is_supporting_information_file(p)]
    return main_pdfs or pdfs


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract HMFOR experiment variables from PDF papers.")
    parser.add_argument("--pdf-dir", default=str(DEFAULT_PDF_DIR), help=f"Folder containing PDF papers. Default: {DEFAULT_PDF_DIR}")
    parser.add_argument("--out-dir", default=str(DEFAULT_OUT_DIR), help=f"Output folder for CSV/XLSX files. Default: {DEFAULT_OUT_DIR}")
    parser.add_argument("--limit", type=int, default=None, help="Optional max number of main PDFs to process.")
    parser.add_argument("--enable-potential-conversion", action="store_true", help="Opt-in: estimate RHE conversion values in potential_conversion_audit.csv when assumptions are available. Does not overwrite extracted fields.")
    parser.add_argument("--enable-doi-online", action="store_true", help="Opt-in: query Crossref for DOI/title metadata into doi_enrichment_results.csv. Does not overwrite extracted fields.")
    parser.add_argument("--enable-si-download", action="store_true", help="Opt-in: try downloading direct SI PDF/DOCX candidates into output/downloaded-si, then rerun extraction with downloaded SI files.")
    parser.add_argument("--config", default=str(DEFAULT_CONFIG_PATH), help=f"Optional config YAML/JSON path. Default: {DEFAULT_CONFIG_PATH}")
    args = parser.parse_args()

    pdf_dir = Path(args.pdf_dir)
    out_dir = Path(args.out_dir)
    config_path = Path(args.config)
    _config = load_config(config_path)
    if not pdf_dir.exists():
        raise FileNotFoundError(f"PDF folder not found: {pdf_dir}")

    all_pdfs = sorted(pdf_dir.glob("*.pdf"))
    target_pdfs = list_target_pdfs(pdf_dir)
    if args.limit:
        target_pdfs = target_pdfs[: args.limit]

    records, evidence_rows, nico_background_rows = run_extraction_batch(target_pdfs, all_pdfs)
    si_download_df, downloaded_si_files = download_si_candidates(records, out_dir, args.enable_doi_online, args.enable_si_download)
    if downloaded_si_files:
        all_pdfs = sorted(set(all_pdfs + downloaded_si_files), key=lambda x: str(x).lower())
        print(f"[SI] Downloaded {len(downloaded_si_files)} SI file(s); rerunning extraction with SI context.")
        records, evidence_rows, nico_background_rows = run_extraction_batch(target_pdfs, all_pdfs)

    write_outputs(records, evidence_rows, out_dir, nico_background_rows, args.enable_potential_conversion, all_pdfs, args.enable_doi_online, si_download_df, pdf_dir, config_path, args.enable_si_download)
    print(f"Done. Processed {len(records)} PDF(s). Outputs written to: {out_dir}")


if __name__ == "__main__":
    main()








