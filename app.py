import json
import os
import re
import threading
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_COLOR_INDEX
from docx.opc.part import Part
from anthropic import Anthropic
from google import genai
from google.genai import types as genai_types
from openpyxl import Workbook
import yaml

APP_TITLE = "Indeed Proto QA Reviewer"
DEFAULT_ENV_CANDIDATES = [
    r"C:\Internal Apps\MosAIQ LQA\MosAIQ-LQA\.env",
    r"C:\Internal Apps\MosAIQ LQA\MosAIQ-LQA-GUI\.env",
    r"C:\Internal Apps\MosAIQ LQA\MosAIQ-LQA-GUI-RECURSIVE\.env",
    r"C:\Internal Apps\Terminology-Suite\.env",
]
WORD_COST_USD = 0.005
DEFAULT_CLAUDE_MODEL_AGENT1 = "claude-opus-4-5"
DEFAULT_CLAUDE_MODEL_AGENT2 = "claude-haiku-4-5"
DEFAULT_MOSAIQ_CONFIG = Path(r"C:\Internal Apps\MosAIQ LQA\MosAIQ-LQA\config.yaml")

# ── Colour palette ────────────────────────────────────────────────────────────
BG_DARK        = "#1e1f2e"
BG_PANEL       = "#252637"
BG_CARD        = "#2d2f45"
BG_INPUT       = "#1a1b2a"
ACCENT         = "#7c6af7"
ACCENT_HOVER   = "#9b8df9"
ACCENT_DIM     = "#4a4580"
SUCCESS        = "#4caf82"
WARNING        = "#f0a843"
DANGER         = "#e05c6e"
TEXT_PRIMARY   = "#e8e9f3"
TEXT_SECONDARY = "#9091a8"
TEXT_MUTED     = "#5c5d7a"
BORDER         = "#3a3b52"


# ═══════════════════════════════════════════════════════════════════════════════
#  Data classes
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class Issue:
    article: str
    severity: str
    rule_id: str
    rule: str
    evidence: str
    recommendation: str
    source_refs: List[str] = field(default_factory=list)
    source_files: List[str] = field(default_factory=list)


def normalize_rule_id(value: str, fallback_idx: int) -> str:
    rid = str(value or "").strip().upper()
    if re.fullmatch(r"R\d+", rid):
        return rid
    return f"R{fallback_idx}"


def coerce_str_list(value) -> List[str]:
    if isinstance(value, list):
        return [str(v).strip() for v in value if str(v).strip()]
    if isinstance(value, str):
        return [value.strip()] if value.strip() else []
    return []


def build_source_catalog(
    references: List[Tuple[str, str]],
    spotchecks: List[Tuple[str, str]],
) -> Dict[str, dict]:
    catalog: Dict[str, dict] = {}
    for idx, (name, _text) in enumerate(references, 1):
        sid = f"REF{idx}"
        catalog[sid] = {
            "source_id": sid,
            "source_type": "reference",
            "file": name,
        }
    for idx, (name, _text) in enumerate(spotchecks, 1):
        sid = f"SPOT{idx}"
        catalog[sid] = {
            "source_id": sid,
            "source_type": "spotcheck",
            "file": name,
        }
    return catalog


def normalize_rules_bundle(
    rules_bundle: dict,
    references: List[Tuple[str, str]],
    spotchecks: List[Tuple[str, str]],
) -> dict:
    source_catalog = build_source_catalog(references, spotchecks)
    file_to_source = {
        meta["file"].lower(): sid for sid, meta in source_catalog.items()
    }
    normalized_rules = []
    raw_rules = rules_bundle.get("qa_rules", []) if isinstance(rules_bundle, dict) else []
    for idx, rule in enumerate(raw_rules, 1):
        if not isinstance(rule, dict):
            continue
        rule_id = normalize_rule_id(rule.get("rule_id", ""), idx)
        source_refs = [s.upper() for s in coerce_str_list(rule.get("source_refs"))]
        if not source_refs:
            for doc_name in coerce_str_list(rule.get("source_docs")):
                sid = file_to_source.get(doc_name.lower())
                if sid:
                    source_refs.append(sid)
        source_refs = [s for s in source_refs if s in source_catalog]
        source_files = [source_catalog[s]["file"] for s in source_refs]
        normalized_rules.append({
            "rule_id": rule_id,
            "rule": str(rule.get("rule", "")).strip(),
            "severity": str(rule.get("severity", "low")).strip().lower(),
            "source": str(rule.get("source", "reference")).strip().lower(),
            "source_refs": source_refs,
            "source_files": source_files,
        })
    return {
        "qa_rules": normalized_rules,
        "summary": str((rules_bundle or {}).get("summary", "")),
        "source_catalog": sorted(source_catalog.values(), key=lambda x: x["source_id"]),
    }


@dataclass
class ArticleAssessment:
    article: str
    summary: str
    readiness_score: float
    issues: List[Issue]


# ═══════════════════════════════════════════════════════════════════════════════
#  Generic helpers
# ═══════════════════════════════════════════════════════════════════════════════

def safe_json_loads(text: str) -> Optional[dict]:
    try:
        return json.loads(text)
    except Exception:
        pass
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(0))
        except Exception:
            return None
    return None


def load_env_file(path: Path) -> Dict[str, str]:
    env = {}
    if not path.exists():
        return env
    for line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            continue
        key, value = stripped.split("=", 1)
        env[key.strip()] = value.strip().strip('"').strip("'")
    return env


def first_existing_env(candidates: List[str]) -> Optional[Path]:
    for candidate in candidates:
        p = Path(candidate)
        if p.exists():
            return p
    return None


def list_docx(folder: Path) -> List[Path]:
    if not folder.exists():
        return []
    return sorted(p for p in folder.rglob("*.docx") if not p.name.startswith("~$"))


def docx_to_text(path: Path) -> str:
    doc = Document(str(path))
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return "\n".join(lines)


def word_count(text: str) -> int:
    return len(re.findall(r"\b\w+\b", text))


def save_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def load_agent_settings(config_path: Path) -> dict:
    settings = {
        "step1_model": DEFAULT_CLAUDE_MODEL_AGENT1,
        "step1_temp":  0.2,
        "step2_model": DEFAULT_CLAUDE_MODEL_AGENT2,
        "step2_temp":  0.2,
    }
    if not config_path.exists():
        return settings
    try:
        data  = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
        lqa   = data.get("lqa", {}) or {}
        step1 = lqa.get("step1", {}) or {}
        step2 = lqa.get("step2", {}) or {}
        settings["step1_model"] = str(step1.get("model") or settings["step1_model"])
        settings["step2_model"] = str(step2.get("model") or settings["step2_model"])
        settings["step1_temp"]  = float(step1.get("temperature", settings["step1_temp"]))
        settings["step2_temp"]  = float(step2.get("temperature", settings["step2_temp"]))
    except Exception:
        pass
    return settings


# ═══════════════════════════════════════════════════════════════════════════════
#  Agent API calls
# ═══════════════════════════════════════════════════════════════════════════════

def call_claude_extract_rules(
    api_key: str,
    references: List[Tuple[str, str]],
    spotchecks: List[Tuple[str, str]],
    model: str,
    temperature: float,
) -> dict:
    joined_refs = "\n\n".join([f"[REF] {n}\n{t[:15000]}" for n, t in references])
    joined_spot = "\n\n".join([f"[SPOT] {n}\n{t[:15000]}" for n, t in spotchecks])

    prompt = f"""
You are Agent 1 for QA rules extraction.

Task:
1) Read references and internal spot checks.
2) Generate a concise list of QA rules/checkpoints.
3) Return JSON only, no extra text.

JSON schema:
{{
  "qa_rules": [
    {{"rule_id": "R1", "rule": "...", "severity": "high|medium|low", "source": "reference|spotcheck|both"}}
  ],
  "summary": "short summary"
}}

References:\n{joined_refs}\n\nInternal Spot Checks:\n{joined_spot}
"""
    client = Anthropic(api_key=api_key)
    msg    = client.messages.create(
        model=model,
        max_tokens=8192,
        temperature=temperature,
        messages=[{"role": "user", "content": prompt}],
    )
    text_blocks = [
        getattr(item, "text", "")
        for item in msg.content
        if getattr(item, "type", "") == "text"
    ]
    text   = "\n".join(text_blocks)
    parsed = safe_json_loads(text)
    if not parsed:
        raise ValueError("Claude Agent 1 response could not be parsed as JSON")
    return parsed


def call_claude_assess_article(
    api_key: str,
    article_name: str,
    article_text: str,
    qa_rules: List[dict],
    model: str,
    temperature: float,
) -> "ArticleAssessment":
    rules_text = json.dumps(qa_rules, ensure_ascii=False, indent=2)
    rules_by_id = {
        str(r.get("rule_id", "")).strip().upper(): r
        for r in qa_rules if isinstance(r, dict)
    }
    rules_by_text = {
        str(r.get("rule", "")).strip().lower(): r
        for r in qa_rules if isinstance(r, dict)
    }
    prompt = f"""
You are Agent 2 for article QA assessment.

Assess the article against the QA rules below. Return JSON only.

Schema:
{{
  "article": "{article_name}",
  "summary": "...",
  "readiness_score": 0,
  "issues": [
    {{
      "rule_id": "R1",
      "severity": "high|medium|low",
      "rule": "...",
      "source_refs": ["REF1"],
      "evidence": "exact sentence or phrase from the article that illustrates the problem",
      "recommendation": "..."
    }}
  ]
}}

Important: the "evidence" field must contain an exact quote (or very close paraphrase) from
the article text so the annotation tool can locate and highlight it.
Keep issue.rule_id and source_refs aligned with the supplied Rules JSON.

Rules:\n{rules_text}\n\nArticle:\n{article_text[:40000]}
"""
    client = Anthropic(api_key=api_key)
    msg    = client.messages.create(
        model=model,
        max_tokens=8192,
        temperature=temperature,
        messages=[{"role": "user", "content": prompt}],
    )
    text_blocks = [
        getattr(item, "text", "")
        for item in msg.content
        if getattr(item, "type", "") == "text"
    ]
    text   = "\n".join(text_blocks)
    parsed = safe_json_loads(text)
    if not parsed:
        raise ValueError(f"Claude response could not be parsed as JSON for {article_name}")

    issues = []
    for i in parsed.get("issues", []):
        if not isinstance(i, dict):
            continue
        issue_rule_id = str(i.get("rule_id", "")).strip().upper()
        issue_rule_text = str(i.get("rule", "")).strip()
        matched_rule = rules_by_id.get(issue_rule_id) or rules_by_text.get(issue_rule_text.lower())
        if matched_rule:
            if not issue_rule_id:
                issue_rule_id = str(matched_rule.get("rule_id", "")).strip().upper()
            if not issue_rule_text:
                issue_rule_text = str(matched_rule.get("rule", "")).strip()
        source_refs = [s.upper() for s in coerce_str_list(i.get("source_refs"))]
        if not source_refs and matched_rule:
            source_refs = [s.upper() for s in coerce_str_list(matched_rule.get("source_refs"))]
        source_files = []
        if matched_rule:
            source_files = coerce_str_list(matched_rule.get("source_files"))
        issues.append(Issue(
            article=article_name,
            severity=str(i.get("severity", "low")),
            rule_id=issue_rule_id,
            rule=issue_rule_text,
            evidence=str(i.get("evidence", "")),
            recommendation=str(i.get("recommendation", "")),
            source_refs=source_refs,
            source_files=source_files,
        ))
    return ArticleAssessment(
        article=parsed.get("article", article_name),
        summary=parsed.get("summary", ""),
        readiness_score=float(parsed.get("readiness_score", 0)),
        issues=issues,
    )


# ═══════════════════════════════════════════════════════════════════════════════
#  Fallbacks (no API key / API failure)
# ═══════════════════════════════════════════════════════════════════════════════

def fallback_rules(
    references: List[Tuple[str, str]],
    spotchecks: List[Tuple[str, str]],
) -> dict:
    spot1 = ["SPOT1"] if spotchecks else []
    ref1 = ["REF1"] if references else []
    return {
        "qa_rules": [
            {"rule_id": "R1", "rule": "Brand names must match official casing.",
             "severity": "high",   "source": "spotcheck", "source_refs": spot1},
            {"rule_id": "R2", "rule": "Terminology must align with reference materials.",
             "severity": "high",   "source": "reference", "source_refs": ref1},
            {"rule_id": "R3", "rule": "Tone and style must be consistent with internal guidance.",
             "severity": "medium", "source": "both", "source_refs": ref1 + spot1},
        ],
        "summary": "Fallback rules generated because API step failed.",
    }


def fallback_assessment(
    article_name: str,
    article_text: str,
    qa_rules: List[dict],
) -> ArticleAssessment:
    issues: List[Issue] = []
    if article_text and article_text.isupper():
        issues.append(Issue(
            article=article_name, severity="high",
            rule_id="",
            rule="Excessive uppercase",
            evidence=article_text[:120],
            recommendation="Use sentence case unless brand-standard uppercase is required.",
        ))
    if len(article_text) < 80:
        issues.append(Issue(
            article=article_name, severity="medium",
            rule_id="",
            rule="Insufficient content",
            evidence=article_text[:120] or "(empty)",
            recommendation="Ensure full article content is provided.",
        ))
    if not issues:
        issues.append(Issue(
            article=article_name, severity="low",
            rule_id="",
            rule="Manual review recommended",
            evidence=article_text[:120] or "(empty)",
            recommendation="Run with valid API keys for full agentic assessment.",
        ))
    score = max(
        0,
        100
        - len([i for i in issues if i.severity == "high"])   * 25
        - len([i for i in issues if i.severity == "medium"]) * 10,
    )
    return ArticleAssessment(
        article=article_name,
        summary="Fallback assessment used.",
        readiness_score=score,
        issues=issues,
    )


# ═══════════════════════════════════════════════════════════════════════════════
#  DOCX annotation with real Word comments
# ═══════════════════════════════════════════════════════════════════════════════

def _ensure_comments_part(doc: Document):
    """
    Return the comments XML root element (<w:comments>).
    Uses XmlPart so the lxml element is managed natively by python-docx
    and serialised correctly on doc.save() without any blob patching.
    """
    import lxml.etree as etree
    from docx.opc.part import XmlPart
    from docx.opc.packuri import PackURI

    REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    CT  = ("application/vnd.openxmlformats-officedocument"
           ".wordprocessingml.comments+xml")
    PN  = PackURI("/word/comments.xml")

    # ── return existing part if already related ───────────────────────────────
    for rel in doc.part.rels.values():
        if rel.reltype == REL:
            return rel._target._element

    # ── build the root element directly in lxml ───────────────────────────────
    WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    root = etree.Element(
        f"{{{WNS}}}comments",
        nsmap={"w": WNS, "r": RNS},
    )

    # ── create an XmlPart so python-docx owns and serialises the element ──────
    part_obj = XmlPart(
        partname=PN,
        content_type=CT,
        element=root,
        package=doc.part.package,
    )

    doc.part.relate_to(part_obj, REL)
    return root


def _add_word_comment(
    doc: Document,
    paragraph,
    comment_text: str,
    author: str = "QA Reviewer",
    comment_id: int = 1,
) -> None:
    """
    Attach a proper Word review comment to *paragraph*.
    """
    from datetime import timezone

    comments_root = _ensure_comments_part(doc)

    # ── build <w:comment> ─────────────────────────────────────────────────────
    comment_elem = OxmlElement("w:comment")
    comment_elem.set(qn("w:id"),       str(comment_id))
    comment_elem.set(qn("w:author"),   author)
    comment_elem.set(
        qn("w:date"),
        datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    )
    comment_elem.set(qn("w:initials"), "QA")

    # paragraph inside the comment balloon
    cp  = OxmlElement("w:p")
    cpr = OxmlElement("w:pPr")
    cps = OxmlElement("w:pStyle")
    cps.set(qn("w:val"), "CommentText")
    cpr.append(cps)
    cp.append(cpr)

    # split multi-line comment into runs separated by line breaks
    lines = comment_text.split("\n")
    for line_idx, line in enumerate(lines):
        cr = OxmlElement("w:r")
        if line_idx == 0:
            rpr = OxmlElement("w:rPr")
            rs  = OxmlElement("w:rStyle")
            rs.set(qn("w:val"), "CommentReference")
            rpr.append(rs)
            cr.append(rpr)
        ct = OxmlElement("w:t")
        ct.text = line
        ct.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        cr.append(ct)
        cp.append(cr)
        if line_idx < len(lines) - 1:
            br_run = OxmlElement("w:r")
            br_run.append(OxmlElement("w:br"))
            cp.append(br_run)

    comment_elem.append(cp)
    comments_root.append(comment_elem)

    # ── wrap paragraph runs with commentRangeStart / End ─────────────────────
    p_elem = paragraph._p

    range_start = OxmlElement("w:commentRangeStart")
    range_start.set(qn("w:id"), str(comment_id))

    range_end = OxmlElement("w:commentRangeEnd")
    range_end.set(qn("w:id"), str(comment_id))

    runs = p_elem.findall(qn("w:r"))
    if runs:
        runs[0].addprevious(range_start)
        runs[-1].addnext(range_end)
    else:
        p_elem.insert(0, range_start)
        p_elem.append(range_end)

    # ── commentReference run (draws the balloon anchor) ───────────────────────
    ref_run = OxmlElement("w:r")
    ref_rpr = OxmlElement("w:rPr")
    ref_rs  = OxmlElement("w:rStyle")
    ref_rs.set(qn("w:val"), "CommentReference")
    ref_rpr.append(ref_rs)
    ref_run.append(ref_rpr)

    comment_ref = OxmlElement("w:commentReference")
    comment_ref.set(qn("w:id"), str(comment_id))
    ref_run.append(comment_ref)

    range_end.addnext(ref_run)


def _highlight_paragraph(paragraph) -> None:
    """Yellow-highlight every run in a paragraph."""
    for run in paragraph.runs:
        run.font.highlight_color = WD_COLOR_INDEX.YELLOW


def _paragraph_matches_issue(ptxt_lower: str, issue: Issue) -> bool:
    """
    Return True when the paragraph text is likely the location of *issue*.
    Checks:
      1. An exact substring of the evidence (first 80 chars) appears in the paragraph.
      2. The rule keyword appears in the paragraph.
    """
    evidence_snippet = issue.evidence.strip().lower()[:80]
    rule_snippet     = issue.rule.strip().lower()
    return (
        (evidence_snippet and evidence_snippet in ptxt_lower)
        or (rule_snippet   and rule_snippet   in ptxt_lower)
    )


def issue_sources_text(issue: Issue) -> str:
    if issue.source_files:
        return ", ".join(issue.source_files)
    if issue.source_refs:
        return ", ".join(issue.source_refs)
    return "N/A"


def issue_heading(issue: Issue) -> str:
    if issue.rule_id:
        return f"[{issue.severity.upper()}] {issue.rule_id} - {issue.rule}"
    return f"[{issue.severity.upper()}] {issue.rule}"


def create_annotated_docx(
    source_path: Path,
    out_path: Path,
    issues: List[Issue],
) -> None:
    """
    Saves an annotated copy of *source_path* to *out_path*.

    For every issue whose evidence can be located in a paragraph:
      • The paragraph text is highlighted yellow.
      • A proper Word review comment (visible in the Review pane / comment
        balloons) is attached to that paragraph.

    Issues that could not be matched to a specific paragraph are collected
    into a single summary comment on the first paragraph so they are never
    silently lost.
    """
    doc = Document(str(source_path))

    if not issues:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(out_path))
        return

    comment_id    = 1
    unmatched     = list(issues)   # shrinks as we match issues to paragraphs

    for paragraph in doc.paragraphs:
        ptxt = paragraph.text.strip()
        if not ptxt:
            continue

        ptxt_lower = ptxt.lower()
        matched    = [iss for iss in issues if _paragraph_matches_issue(ptxt_lower, iss)]

        if not matched:
            continue

        # Remove matched issues from the unmatched pool
        for iss in matched:
            if iss in unmatched:
                unmatched.remove(iss)

        _highlight_paragraph(paragraph)

        comment_lines: List[str] = []
        for iss in matched:
            comment_lines.append(
                f"{issue_heading(iss)}\n"
                f"Source Material: {issue_sources_text(iss)}\n"
                f"Evidence: {iss.evidence}\n"
                f"Recommendation: {iss.recommendation}"
            )

        _add_word_comment(
            doc, paragraph,
            "\n\n".join(comment_lines),
            author="QA Reviewer",
            comment_id=comment_id,
        )
        comment_id += 1

    # ── attach unmatched issues to the first non-empty paragraph ─────────────
    if unmatched:
        first_para = next(
            (p for p in doc.paragraphs if p.text.strip()), None
        )
        if first_para:
            _highlight_paragraph(first_para)
            comment_lines = ["[UNMATCHED ISSUES — could not be located in text]\n"]
            for iss in unmatched:
                comment_lines.append(
                    f"{issue_heading(iss)}\n"
                    f"Source Material: {issue_sources_text(iss)}\n"
                    f"Evidence: {iss.evidence}\n"
                    f"Recommendation: {iss.recommendation}"
                )
            _add_word_comment(
                doc, first_para,
                "\n\n".join(comment_lines),
                author="QA Reviewer",
                comment_id=comment_id,
            )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))


# ═══════════════════════════════════════════════════════════════════════════════
#  Report helpers
# ═══════════════════════════════════════════════════════════════════════════════

def save_summary_docx(path: Path, assessments: List[ArticleAssessment]) -> None:
    doc = Document()
    doc.add_heading("QA Issues Summary", level=1)
    for a in assessments:
        doc.add_heading(a.article, level=2)
        doc.add_paragraph(f"Summary: {a.summary}")
        doc.add_paragraph(f"Readiness Score: {a.readiness_score}")
        if a.issues:
            doc.add_paragraph("Issues:")
            for idx, iss in enumerate(a.issues, 1):
                doc.add_paragraph(
                    f"  {idx}. {issue_heading(iss)}\n"
                    f"     Source Material: {issue_sources_text(iss)}\n"
                    f"     Evidence: {iss.evidence}\n"
                    f"     Recommendation: {iss.recommendation}",
                    style="List Bullet",
                )
        doc.add_paragraph("")
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


def write_sheet(ws, rows: List[dict], headers: List[str]) -> None:
    ws.append(headers)
    for row in rows:
        ws.append([row.get(h, "") for h in headers])


def build_excel_report(
    path: Path,
    assessments: List[ArticleAssessment],
    word_counts: List[dict],
) -> None:
    rows = []
    for a in assessments:
        high = len([i for i in a.issues if i.severity.lower() == "high"])
        med  = len([i for i in a.issues if i.severity.lower() == "medium"])
        low  = len([i for i in a.issues if i.severity.lower() == "low"])
        rows.append({
            "Article":       a.article,
            "Summary":       a.summary,
            "ReadinessScore":a.readiness_score,
            "HighIssues":    high,
            "MediumIssues":  med,
            "LowIssues":     low,
            "TotalIssues":   len(a.issues),
        })

    path.parent.mkdir(parents=True, exist_ok=True)
    wb    = Workbook()
    ws_qa = wb.active
    ws_qa.title = "QA_Readiness"
    write_sheet(ws_qa, rows,
                ["Article", "Summary", "ReadinessScore",
                 "HighIssues", "MediumIssues", "LowIssues", "TotalIssues"])

    scores  = [float(r.get("ReadinessScore", 0) or 0) for r in rows]
    overall = (sum(scores) / len(scores)) if scores else 0.0
    ws_overall = wb.create_sheet("Overall")
    write_sheet(ws_overall, [{"OverallReadinessScore": overall}], ["OverallReadinessScore"])

    ws_wc = wb.create_sheet("WordCount")
    write_sheet(ws_wc, word_counts, ["File", "Category", "WordCount"])

    total_words = sum(int(r.get("WordCount", 0) or 0) for r in word_counts)
    ws_cost = wb.create_sheet("CostEstimate")
    write_sheet(
        ws_cost,
        [{
            "TotalWords":        total_words,
            "RateUSDPerWord":    WORD_COST_USD,
            "EstimatedCostUSD":  round(total_words * WORD_COST_USD, 2),
        }],
        ["TotalWords", "RateUSDPerWord", "EstimatedCostUSD"],
    )
    wb.save(path)


# ═══════════════════════════════════════════════════════════════════════════════
#  Custom Tkinter Widgets
# ═══════════════════════════════════════════════════════════════════════════════

class FlatButton:
    """
    Flat rounded button — pure composition (no Tkinter subclassing).
    Pack/grid the instance directly; .pack()/.grid()/.place() delegate
    to the internal Canvas widget.
    """

    def __init__(self, parent, text="", command=None,
                 bg=ACCENT, hover_bg=ACCENT_HOVER,
                 fg=TEXT_PRIMARY, font=("Segoe UI", 9, "bold"),
                 width=130, height=32, radius=8):
        try:
            parent_bg = parent.cget("bg")
        except Exception:
            parent_bg = BG_DARK

        self._text    = text
        self._command = command
        self._bg      = bg
        self._hbg     = hover_bg
        self._fg      = fg
        self._font    = font
        self._radius  = radius
        self._w       = width
        self._h       = height

        self.widget = tk.Canvas(
            parent,
            width=width, height=height,
            bg=parent_bg,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )

        # Defer first draw until the canvas is fully registered in Tcl/Tk
        self.widget.after(0, lambda: self._draw(self._bg))

        self.widget.bind("<Enter>",    lambda _: self._draw(self._hbg))
        self.widget.bind("<Leave>",    lambda _: self._draw(self._bg))
        self.widget.bind("<Button-1>", lambda _: self._on_click())

    # ── geometry / config shims ───────────────────────────────────────────────
    def pack(self, **kwargs):   self.widget.pack(**kwargs)
    def grid(self, **kwargs):   self.widget.grid(**kwargs)
    def place(self, **kwargs):  self.widget.place(**kwargs)
    def config(self, **kwargs): self.widget.config(**kwargs)

    # ── drawing ───────────────────────────────────────────────────────────────
    def _round_rect(self, x1, y1, x2, y2, r, **kw):
        pts = [
            x1+r, y1,    x2-r, y1,
            x2,   y1,    x2,   y1+r,
            x2,   y2-r,  x2,   y2,
            x2-r, y2,    x1+r, y2,
            x1,   y2,    x1,   y2-r,
            x1,   y1+r,  x1,   y1,
            x1+r, y1,
        ]
        return self.widget.create_polygon(pts, smooth=True, **kw)

    def _draw(self, color):
        self.widget.delete("all")
        self._round_rect(0, 0, self._w, self._h, self._radius,
                         fill=color, outline="")
        self.widget.create_text(
            self._w // 2, self._h // 2,
            text=self._text,
            fill=self._fg,
            font=self._font,
            anchor="center",
        )

    def _on_click(self):
        self._draw(self._bg)
        if self._command:
            self._command()


class PathRow(tk.Frame):
    """
    Self-contained path-picker row.
    - Remembers the last used directory per-row across sessions via a small JSON cache
    - Falls back to the last globally used directory from any other row
    - Updates initdir live as the user types or pastes a path
    """

    _STATE_FILE = Path.home() / ".qa_reviewer_dirs.json"
    _state: dict = {}          # loaded once at class level
    _state_loaded: bool = False

    @classmethod
    def _load_state(cls) -> None:
        if cls._state_loaded:
            return
        cls._state_loaded = True
        try:
            if cls._STATE_FILE.exists():
                raw = json.loads(cls._STATE_FILE.read_text(encoding="utf-8"))
                # Only keep entries that still exist on disk
                cls._state = {
                    k: v for k, v in raw.items()
                    if Path(v).exists()
                }
        except Exception:
            cls._state = {}

    @classmethod
    def _save_state(cls) -> None:
        try:
            cls._STATE_FILE.write_text(
                json.dumps(cls._state, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
        except Exception:
            pass

    @classmethod
    def _get_global_last(cls) -> str:
        """Most recently used directory across ALL rows."""
        return cls._state.get("__last_global__", "")

    @classmethod
    def _set_global_last(cls, directory: str) -> None:
        cls._state["__last_global__"] = directory
        cls._save_state()

    def __init__(self, parent, label: str, icon: str,
                 pick_mode: str = "folder",
                 file_types=None,
                 row_key: str = "",
                 **kwargs):
        super().__init__(parent, bg=BG_CARD, **kwargs)

        PathRow._load_state()

        self._mode      = pick_mode
        self._filetypes = file_types or [("All", "*.*")]
        # row_key lets each row remember its OWN last directory independently
        self._row_key   = row_key or label
        self._initdir   = self._resolve_initdir()

        # ── icon + label ──────────────────────────────────────────────────────
        lbl_frame = tk.Frame(self, bg=BG_CARD)
        lbl_frame.pack(fill=tk.X, padx=12, pady=(10, 2))

        tk.Label(lbl_frame, text=icon, bg=BG_CARD,
                 fg=ACCENT, font=("Segoe UI", 11)).pack(side=tk.LEFT)
        tk.Label(lbl_frame, text=f"  {label}", bg=BG_CARD,
                 fg=TEXT_SECONDARY,
                 font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT)

        # ── entry + button row ────────────────────────────────────────────────
        row = tk.Frame(self, bg=BG_CARD)
        row.pack(fill=tk.X, padx=12, pady=(0, 10))

        entry_wrap = tk.Frame(row, bg=BORDER, bd=0)
        entry_wrap.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        inner = tk.Frame(entry_wrap, bg=BG_INPUT, bd=0)
        inner.pack(fill=tk.BOTH, padx=1, pady=1)

        self.var    = tk.StringVar()
        self._entry = tk.Entry(
            inner,
            textvariable=self.var,
            bg=BG_INPUT, fg=TEXT_PRIMARY,
            insertbackground=TEXT_PRIMARY,
            relief=tk.FLAT, bd=6,
            font=("Segoe UI", 9),
            disabledbackground=BG_INPUT,
            disabledforeground=TEXT_SECONDARY,
        )
        self._entry.pack(fill=tk.X)

        self._entry.bind("<FocusIn>",
                         lambda _: entry_wrap.config(bg=ACCENT))
        self._entry.bind("<FocusOut>",
                         lambda _: entry_wrap.config(bg=BORDER))

        # Update initdir live as the user types / pastes
        self.var.trace_add("write", self._on_entry_change)

        FlatButton(
            row,
            text="Browse",
            bg=ACCENT_DIM, hover_bg=ACCENT,
            width=90, height=30,
            command=self._pick,
        ).pack(side=tk.LEFT)

    # ── directory resolution ──────────────────────────────────────────────────
    def _resolve_initdir(self) -> str:
        """
        Priority order for the starting directory:
        1. This row's own last-used directory (saved from a previous session)
        2. The most recently used directory from any other row (same session or saved)
        3. The user's Desktop
        4. The user's Documents folder
        5. Home directory
        """
        # 1. Row-specific saved dir
        row_saved = PathRow._state.get(self._row_key, "")
        if row_saved and Path(row_saved).exists():
            return row_saved

        # 2. Global last used
        global_last = self._get_global_last()
        if global_last and Path(global_last).exists():
            return global_last

        # 3. Desktop
        desktop = Path.home() / "Desktop"
        if desktop.exists():
            return str(desktop)

        # 4. Documents (including OneDrive variants)
        for candidate in [
            Path.home() / "Documents",
            *sorted(Path.home().glob("OneDrive*/Documents")),
        ]:
            if candidate.exists():
                return str(candidate)

        # 5. Home
        return str(Path.home())

    def _remember(self, directory: str) -> None:
        """Persist this directory for this row and update the global last."""
        if not directory or not Path(directory).exists():
            return
        PathRow._state[self._row_key] = directory
        self._set_global_last(directory)   # also saves state

    def _on_entry_change(self, *_) -> None:
        """Track valid paths typed or pasted into the entry."""
        typed = self.var.get().strip()
        if not typed:
            return
        p = Path(typed)
        # If it looks like a file path use the parent; otherwise use as-is
        candidate = (
            p.parent
            if (p.is_file() or (not p.exists() and p.suffix))
            else p
        )
        if candidate.exists():
            self._initdir = str(candidate)

    # ── public ────────────────────────────────────────────────────────────────
    def get(self) -> str:
        return self.var.get().strip()

    def set(self, value: str) -> None:
        self.var.set(value)
        if value:
            p       = Path(value)
            new_dir = str(p.parent if p.is_file() else p)
            self._initdir = new_dir
            self._remember(new_dir)

    # ── private ───────────────────────────────────────────────────────────────
    def _pick(self) -> None:
        # Honour whatever the user has typed before opening the dialog
        current = self.var.get().strip()
        if current:
            p         = Path(current)
            candidate = p.parent if p.is_file() else p
            if candidate.exists():
                self._initdir = str(candidate)

        if self._mode == "folder":
            selected = filedialog.askdirectory(
                initialdir=self._initdir,
                title=f"Select folder",
            )
        else:
            selected = filedialog.askopenfilename(
                initialdir=self._initdir,
                filetypes=self._filetypes,
                title=f"Select file",
            )

        if selected:
            self.set(selected)


class StatusBar(tk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=BG_DARK, height=28, **kwargs)
        self.pack_propagate(False)

        self._dot = tk.Label(self, text="●", bg=BG_DARK,
                             fg=TEXT_MUTED, font=("Segoe UI", 10))
        self._dot.pack(side=tk.LEFT, padx=(12, 4))

        self._lbl = tk.Label(self, text="Ready", bg=BG_DARK,
                             fg=TEXT_SECONDARY, font=("Segoe UI", 8))
        self._lbl.pack(side=tk.LEFT)

        self._time_lbl = tk.Label(self, text="", bg=BG_DARK,
                                  fg=TEXT_MUTED, font=("Segoe UI", 8))
        self._time_lbl.pack(side=tk.RIGHT, padx=12)

    def set(self, text: str, state: str = "idle") -> None:
        colours = {"idle": TEXT_MUTED, "running": WARNING,
                   "ok": SUCCESS, "error": DANGER}
        self._dot.config(fg=colours.get(state, TEXT_MUTED))
        self._lbl.config(
            text=text,
            fg=TEXT_PRIMARY if state != "idle" else TEXT_SECONDARY,
        )
        self._time_lbl.config(text=datetime.now().strftime("%H:%M:%S"))


def list_available_gemini_models(api_key: str) -> List[str]:
    """Return names of Gemini models that support generateContent."""
    try:
        client = genai.Client(api_key=api_key)
        models = client.models.list()
        return sorted([
            m.name for m in models
            if "generateContent" in (m.supported_actions or [])
            and "gemini" in m.name.lower()
        ])
    except Exception as e:
        return [f"Error listing models: {e}"]

# ═══════════════════════════════════════════════════════════════════════════════
#  Main Application
# ═══════════════════════════════════════════════════════════════════════════════

class App(tk.Tk):
    def _check_gemini_models(self) -> None:
        env_file = Path(self.row_env.get())
        if not env_file.exists():
            messagebox.showerror("Error", "Please select a valid .env file first.")
            return

        env        = load_env_file(env_file)
        google_key = env.get("GOOGLE_API_KEY", "")
        if not google_key:
            messagebox.showerror("Error", "GOOGLE_API_KEY not found in the .env file.")
            return

        self._log("Fetching available Gemini models for your API key…", "agent1")
        self._status_bar.set("Fetching Gemini models…", "running")

        def _fetch():
            models = list_available_gemini_models(google_key)
            self._log(f"Available Gemini models ({len(models)}):", "agent1")
            for m in models:
                self._log(f"  • {m}", "agent1")
            self._status_bar.set("Model list complete", "ok")

            # Pick the best available pro model and suggest it
            preferred = [
                "models/gemini-2.5-pro",
                "models/gemini-2.5-flash",
                "models/gemini-1.5-pro",
                "models/gemini-1.5-flash",
            ]
            suggestion = next(
                (p for p in preferred if any(m.startswith(p) for m in models)),
                None,
            )
            if suggestion:
                matched = next(m for m in models if m.startswith(suggestion))
                self._log(
                    f"Suggested model for your key → update DEFAULT_GEMINI_MODEL "
                    f"to: \"{matched}\"",
                    "success",
                )
                # Ask user if they want to use it for this session
                use_it = messagebox.askyesno(
                    "Model suggestion",
                    f"Best available model for your API key:\n\n{matched}\n\n"
                    f"Use this model for the current session?",
                )
                if use_it:
                    global DEFAULT_GEMINI_MODEL
                    DEFAULT_GEMINI_MODEL = matched
                    self._log(f"Session model updated → {matched}", "success")

        threading.Thread(target=_fetch, daemon=True).start()
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1100x860")
        self.minsize(900, 700)
        self.configure(bg=BG_DARK)

        try:
            self.tk.call("tk", "scaling", 1.25)
        except Exception:
            pass

        self._build_ui()

        default_env = first_existing_env(DEFAULT_ENV_CANDIDATES)
        if default_env:
            self.row_env.set(str(default_env))
        self.row_output.set(str(Path.cwd() / "output"))

    # ── UI construction ───────────────────────────────────────────────────────
    def _build_ui(self) -> None:

        # ── header ────────────────────────────────────────────────────────────
        header = tk.Frame(self, bg=BG_PANEL, height=56)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        tk.Label(
            header,
            text="⬡  Indeed Proto QA Reviewer",
            bg=BG_PANEL, fg=TEXT_PRIMARY,
            font=("Segoe UI", 13, "bold"),
        ).pack(side=tk.LEFT, padx=18, pady=14)

        tk.Label(
            header,
            text="2-Agent Agentic Workflow  |  Claude Opus + Claude Haiku",
            bg=BG_PANEL, fg=TEXT_MUTED,
            font=("Segoe UI", 8),
        ).pack(side=tk.RIGHT, padx=18)

        tk.Frame(self, bg=ACCENT, height=2).pack(fill=tk.X)

        # ── body ──────────────────────────────────────────────────────────────
        body = tk.Frame(self, bg=BG_DARK)
        body.pack(fill=tk.BOTH, expand=True, padx=16, pady=12)

        left = tk.Frame(body, bg=BG_DARK, width=440)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 8))
        left.pack_propagate(False)

        right = tk.Frame(body, bg=BG_DARK)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # ── section label helper ──────────────────────────────────────────────
        def section_label(parent, text):
            f = tk.Frame(parent, bg=BG_DARK)
            f.pack(fill=tk.X, pady=(10, 4))
            tk.Label(f, text=text, bg=BG_DARK,
                     fg=ACCENT, font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT)
            tk.Frame(f, bg=BORDER, height=1).pack(
                side=tk.LEFT, fill=tk.X, expand=True, padx=8, pady=6)

        # ── input folder rows ─────────────────────────────────────────────────
        section_label(left, "INPUT FOLDERS")

        self.row_articles = PathRow(
            left, "Article Folder (.docx)", "📄",
            row_key="articles",
        )
        self.row_articles.pack(fill=tk.X, pady=2)

        self.row_refs = PathRow(
            left, "Reference Materials Folder (.docx)", "📚",
            row_key="references",
        )
        self.row_refs.pack(fill=tk.X, pady=2)

        self.row_spots = PathRow(
            left, "Internal Spot Check Folder (.docx)", "🔍",
            row_key="spotchecks",
        )
        self.row_spots.pack(fill=tk.X, pady=2)

        self.row_output = PathRow(
            left, "Output Folder", "📁",
            row_key="output",
        )
        self.row_output.pack(fill=tk.X, pady=2)

        self.row_env = PathRow(
            left,
            ".env File  (ANTHROPIC_API_KEY)", "🔑",
            pick_mode="file",
            file_types=[("ENV file", ".env"), ("All files", "*.*")],
            row_key="env",
        )
        self.row_env.pack(fill=tk.X, pady=2)

        # ── run button ────────────────────────────────────────────────────────
        btn_frame = tk.Frame(left, bg=BG_DARK)
        btn_frame.pack(fill=tk.X, pady=(18, 4))

        self._run_btn = FlatButton(
            btn_frame,
            text="▶  Run QA Workflow",
            command=self.run_workflow,
            bg=ACCENT, hover_bg=ACCENT_HOVER,
            width=200, height=38,
            font=("Segoe UI", 10, "bold"),
        )
        self._run_btn.pack(side=tk.RIGHT)

        # ── agent info cards ──────────────────────────────────────────────────
        section_label(left, "AGENT PIPELINE")
        self._build_agent_cards(left)

        # ── log panel ─────────────────────────────────────────────────────────
        log_header = tk.Frame(right, bg=BG_PANEL)
        log_header.pack(fill=tk.X)

        tk.Label(log_header, text="  ◈  Activity Log", bg=BG_PANEL,
                 fg=TEXT_PRIMARY, font=("Segoe UI", 9, "bold")).pack(
            side=tk.LEFT, pady=8, padx=4)

        self._clear_btn = tk.Label(
            log_header, text="Clear  ✕", bg=BG_PANEL,
            fg=TEXT_MUTED, font=("Segoe UI", 8), cursor="hand2")
        self._clear_btn.pack(side=tk.RIGHT, padx=10)
        self._clear_btn.bind("<Button-1>", lambda _: self._clear_log())
        self._clear_btn.bind("<Enter>", lambda _: self._clear_btn.config(fg=TEXT_PRIMARY))
        self._clear_btn.bind("<Leave>", lambda _: self._clear_btn.config(fg=TEXT_MUTED))

        tk.Frame(right, bg=BORDER, height=1).pack(fill=tk.X)

        # Log takes all remaining vertical space
        log_wrap = tk.Frame(right, bg=BG_INPUT)
        log_wrap.pack(fill=tk.BOTH, expand=True)

        self.log_box = tk.Text(
            log_wrap,
            bg=BG_INPUT, fg=TEXT_PRIMARY,
            insertbackground=TEXT_PRIMARY,
            font=("Consolas", 9),        # slightly larger font
            relief=tk.FLAT, bd=10,
            wrap=tk.WORD,
            state=tk.DISABLED,
            selectbackground=ACCENT_DIM,
        )
        self.log_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.log_box.tag_config("time",    foreground=TEXT_MUTED)
        self.log_box.tag_config("info",    foreground=TEXT_PRIMARY)
        self.log_box.tag_config("success", foreground=SUCCESS)
        self.log_box.tag_config("warning", foreground=WARNING)
        self.log_box.tag_config("error",   foreground=DANGER)
        self.log_box.tag_config("agent1",  foreground="#c97bf7")
        self.log_box.tag_config("agent2",  foreground="#c97bf7")

        scrollbar = ttk.Scrollbar(log_wrap, orient=tk.VERTICAL,
                                  command=self.log_box.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_box.config(yscrollcommand=scrollbar.set)

        # ── progress bar ──────────────────────────────────────────────────────
        prog_frame = tk.Frame(right, bg=BG_DARK)
        prog_frame.pack(fill=tk.X, pady=(4, 0))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "QA.Horizontal.TProgressbar",
            troughcolor=BG_PANEL,
            background=ACCENT,
            bordercolor=BG_PANEL,
            lightcolor=ACCENT,
            darkcolor=ACCENT,
        )
        self._progress = ttk.Progressbar(
            prog_frame,
            style="QA.Horizontal.TProgressbar",
            orient=tk.HORIZONTAL,
            mode="indeterminate",
            length=200,
        )
        self._progress.pack(fill=tk.X, padx=0, pady=2)

        # ── status bar ────────────────────────────────────────────────────────
        self._status_bar = StatusBar(self)
        self._status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def _build_agent_cards(self, parent) -> None:
        cards_frame = tk.Frame(parent, bg=BG_DARK)
        cards_frame.pack(fill=tk.X, pady=2)

        for title, model, color, icon, side_pad in [
            ("Agent 1", "Claude Opus — Rules", "#c97bf7", "🧠", (0, 4)),
            ("Agent 2", "Claude Haiku — QA",   "#c97bf7", "🧠", (4, 0)),
        ]:
            card = tk.Frame(cards_frame, bg=BG_CARD)
            card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True,
                      padx=side_pad)

            strip = tk.Frame(card, bg=color, width=3)
            strip.pack(side=tk.LEFT, fill=tk.Y)
            strip.pack_propagate(False)

            inner = tk.Frame(card, bg=BG_CARD, padx=10, pady=8)
            inner.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            tk.Label(inner, text=f"{icon} {title}", bg=BG_CARD,
                     fg=color, font=("Segoe UI", 8, "bold")).pack(anchor="w")
            tk.Label(inner, text=model, bg=BG_CARD,
                     fg=TEXT_SECONDARY, font=("Segoe UI", 8)).pack(anchor="w")

    # ── logging ───────────────────────────────────────────────────────────────
    def _log(self, text: str, level: str = "info") -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        if level == "info":
            tl = text.lower()
            if "error" in tl or "fail" in tl:
                level = "error"
            elif "gemini" in tl or "agent 1" in tl:
                level = "agent1"
            elif "claude" in tl or "agent 2" in tl:
                level = "agent2"
            elif "complet" in tl or "finish" in tl or "done" in tl:
                level = "success"
            elif "warning" in tl or "fallback" in tl or "missing" in tl:
                level = "warning"

        self.log_box.config(state=tk.NORMAL)
        self.log_box.insert(tk.END, f"[{stamp}]  ", "time")
        self.log_box.insert(tk.END, f"{text}\n", level)
        self.log_box.see(tk.END)
        self.log_box.config(state=tk.DISABLED)
        self.update_idletasks()

    def _clear_log(self) -> None:
        self.log_box.config(state=tk.NORMAL)
        self.log_box.delete("1.0", tk.END)
        self.log_box.config(state=tk.DISABLED)

    # ── workflow ──────────────────────────────────────────────────────────────
    def run_workflow(self) -> None:
        threading.Thread(target=self._run_workflow_inner, daemon=True).start()

    def _run_workflow_inner(self) -> None:
        try:
            self._status_bar.set("Running workflow…", "running")
            self._progress.start(12)
            self._log("Starting 2-agent QA workflow…")

            article_dir = Path(self.row_articles.get())
            ref_dir     = Path(self.row_refs.get())
            spot_dir    = Path(self.row_spots.get())
            out_dir     = Path(self.row_output.get())
            env_file    = Path(self.row_env.get())

            if not article_dir.exists() or not ref_dir.exists() or not spot_dir.exists():
                raise ValueError("Please select valid article / reference / spotcheck folders.")
            if not env_file.exists():
                raise ValueError("Please select a valid .env file.")

            env           = load_env_file(env_file)
            google_key    = env.get("GOOGLE_API_KEY", "")
            anthropic_key = env.get("ANTHROPIC_API_KEY", "")
            agent_cfg     = load_agent_settings(DEFAULT_MOSAIQ_CONFIG)

            self._log(
                f"Agent settings — step1: {agent_cfg['step1_model']} "
                f"(T={agent_cfg['step1_temp']})  |  "
                f"step2: {agent_cfg['step2_model']} (T={agent_cfg['step2_temp']})"
            )

            articles = list_docx(article_dir)
            refs     = list_docx(ref_dir)
            spots    = list_docx(spot_dir)

            if not articles:
                raise ValueError("No article .docx files found in the selected folder.")

            self._log(
                f"Discovered  {len(articles)} article(s)  ·  "
                f"{len(refs)} reference(s)  ·  {len(spots)} spot-check(s)"
            )

            extraction_dir = out_dir / "extracted"
            annotated_dir  = out_dir / "annotated_articles"
            extraction_dir.mkdir(parents=True, exist_ok=True)
            annotated_dir.mkdir(parents=True, exist_ok=True)

            article_payload: List[Tuple[str, str, Path]] = []
            ref_payload:     List[Tuple[str, str]]       = []
            spot_payload:    List[Tuple[str, str]]        = []
            wc_rows:         List[dict]                   = []

            for p in articles:
                text = docx_to_text(p)
                wc   = word_count(text)
                article_payload.append((p.name, text, p))
                wc_rows.append({"File": p.name, "Category": "Article", "WordCount": wc})
                save_json(extraction_dir / "articles" / f"{p.stem}.json",
                          {"file": p.name, "text": text})

            for p in refs:
                text = docx_to_text(p)
                ref_payload.append((p.name, text))
                wc_rows.append({"File": p.name, "Category": "Reference",
                                 "WordCount": word_count(text)})
                save_json(extraction_dir / "references" / f"{p.stem}.json",
                          {"file": p.name, "text": text})

            for p in spots:
                text = docx_to_text(p)
                spot_payload.append((p.name, text))
                wc_rows.append({"File": p.name, "Category": "SpotCheck",
                                 "WordCount": word_count(text)})
                save_json(extraction_dir / "spotchecks" / f"{p.stem}.json",
                          {"file": p.name, "text": text})

            self._log("Extraction complete — DOCX → JSON")

            # ── Agent 1: Gemini ───────────────────────────────────────────────
            # ── Agent 1: Claude (rules extraction) ───────────────────────────
            if anthropic_key:
                try:
                    self._log("Calling Agent 1 (Claude) — generating QA rules…")
                    rules_bundle = call_claude_extract_rules(
                        anthropic_key,
                        ref_payload,
                        spot_payload,
                        agent_cfg["step1_model"],
                        agent_cfg["step1_temp"],
                    )
                except Exception as e:
                    self._log(f"Claude Agent 1 failed — using fallback rules  ({e})")
                    rules_bundle = fallback_rules(ref_payload, spot_payload)
            else:
                self._log("ANTHROPIC_API_KEY not found — using fallback rules")
                rules_bundle = fallback_rules(ref_payload, spot_payload)

            rules_bundle = normalize_rules_bundle(rules_bundle, ref_payload, spot_payload)
            qa_rules = rules_bundle.get("qa_rules", [])
            save_json(out_dir / "qa_rules.json", rules_bundle)
            self._log(f"QA rules ready  ({len(qa_rules)} rule(s))")

            # ── Agent 2: Claude ───────────────────────────────────────────────
            assessments: List[ArticleAssessment] = []
            for idx, (name, text, source_path) in enumerate(article_payload, 1):
                self._log(
                    f"Agent 2 (Claude) — assessing "
                    f"[{idx}/{len(article_payload)}]  {name}"
                )
                if anthropic_key:
                    try:
                        a = call_claude_assess_article(
                            anthropic_key, name, text, qa_rules,
                            agent_cfg["step2_model"], agent_cfg["step2_temp"],
                        )
                    except Exception as e:
                        self._log(f"Claude failed for {name} — fallback  ({e})")
                        a = fallback_assessment(name, text, qa_rules)
                else:
                    self._log(f"ANTHROPIC_API_KEY missing — fallback for {name}")
                    a = fallback_assessment(name, text, qa_rules)
# test
                assessments.append(a)
                ann_path = annotated_dir / f"{source_path.stem}_annotated.docx"
                create_annotated_docx(source_path, ann_path, a.issues)
                self._log(f"Annotated DOCX saved → {ann_path.name}")

            # ── reports ───────────────────────────────────────────────────────
            save_summary_docx(out_dir / "issues_summary.docx", assessments)
            build_excel_report(out_dir / "qa_readiness_report.xlsx",
                               assessments, wc_rows)

            issues_flat = [
                {
                    "article":        i.article,
                    "severity":       i.severity,
                    "rule_id":        i.rule_id,
                    "rule":           i.rule,
                    "source_refs":    i.source_refs,
                    "source_files":   i.source_files,
                    "evidence":       i.evidence,
                    "recommendation": i.recommendation,
                }
                for a in assessments for i in a.issues
            ]
            save_json(out_dir / "issues.json", {"issues": issues_flat})

            self._log(f"Workflow complete — output saved to: {out_dir}")
            self._progress.stop()
            self._status_bar.set("Completed successfully", "ok")
            messagebox.showinfo(
                "Done",
                f"QA workflow completed.\n\nOutput folder:\n{out_dir}",
            )

        except Exception as e:
            self._progress.stop()
            self._status_bar.set(f"Failed: {e}", "error")
            self._log(f"Fatal error: {e}")
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    app = App()
    app.mainloop()
