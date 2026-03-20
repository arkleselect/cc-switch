from __future__ import annotations

import re
import subprocess
import unicodedata
from dataclasses import dataclass
from pathlib import Path


SOURCE_ROOT = Path("/Users/mortysmith/Desktop/软著/档案大模型轻量化部署应用平台/source")
OUTPUT_DOCX = Path("/Users/mortysmith/Desktop/软著/档案大模型轻量化部署应用平台/3.docx")
TEMP_RTF = OUTPUT_DOCX.with_suffix(".rtf")

PAGES = 60
LINES_PER_PAGE = 60
HEADER_LINES = 4
CODE_LINES_PER_PAGE = LINES_PER_PAGE - HEADER_LINES


@dataclass(frozen=True)
class SourceFile:
    path: Path
    relpath: str
    feature: str


FEATURE_MAP = {
    "app/core/config.py": "Platform Configuration",
    "app/core/rbac.py": "Role Based Access Control",
    "app/core/security.py": "Authentication and Token Security",
    "app/core/validation.py": "Input Validation",
    "app/db/base.py": "Database Access Layer",
    "app/db/migrate.py": "Database Migration",
    "app/db/schema.sql": "Persistent Schema",
    "app/routers/audit.py": "Audit API",
    "app/routers/auth.py": "Authentication API",
    "app/routers/feedback.py": "Feedback API",
    "app/routers/kb.py": "Knowledge Base API",
    "app/routers/model.py": "Model Management API",
    "app/routers/ops.py": "Operations API",
    "app/routers/qa.py": "RAG Question Answering API",
    "app/schemas/audit.py": "Audit Schemas",
    "app/schemas/auth.py": "Authentication Schemas",
    "app/schemas/feedback.py": "Feedback Schemas",
    "app/schemas/kb.py": "Knowledge Base Schemas",
    "app/schemas/model.py": "Model Schemas",
    "app/schemas/qa.py": "Question Answering Schemas",
    "app/services/audit_service.py": "Audit Service",
    "app/services/backup_service.py": "Backup And Restore Service",
    "app/services/feedback_service.py": "Feedback Service",
    "app/services/kb_service.py": "Knowledge Base Processing Service",
    "app/services/model_service.py": "Model Runtime Service",
    "app/services/monitoring_service.py": "Monitoring Service",
    "app/services/qa_service.py": "RAG Answer Generation Service",
    "app/services/search_service.py": "Semantic Search Service",
    "app/services/text_clean_service.py": "Document Cleaning Service",
    "app/services/user_service.py": "User Service",
    "main.py": "Application Entrypoint",
}


def sanitize_ascii(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", text)
    ascii_text = normalized.encode("ascii", "ignore").decode("ascii")
    ascii_text = ascii_text.replace("\t", "    ").rstrip()
    return ascii_text


def escape_rtf(text: str) -> str:
    return text.replace("\\", r"\\").replace("{", r"\{").replace("}", r"\}")


def iter_source_files() -> list[SourceFile]:
    files: list[SourceFile] = []
    for path in sorted(SOURCE_ROOT.rglob("*")):
        if not path.is_file() or path.suffix not in {".py", ".sql"}:
            continue
        relpath = path.relative_to(SOURCE_ROOT).as_posix()
        feature = FEATURE_MAP.get(relpath, "Platform Module")
        files.append(SourceFile(path=path, relpath=relpath, feature=feature))
    return files


def build_code_lines(files: list[SourceFile]) -> list[tuple[str, str]]:
    result: list[tuple[str, str]] = []
    for source_file in files:
        raw_lines = source_file.path.read_text(encoding="utf-8").splitlines()
        for raw_line in raw_lines:
            line = sanitize_ascii(raw_line)
            if not line.strip():
                result.append((source_file.relpath, ""))
                continue

            if any(ord(ch) > 127 for ch in line):
                continue

            result.append((source_file.relpath, line))

    filtered = [(relpath, line) for relpath, line in result if re.search(r"[A-Za-z0-9_#]", line)]
    return filtered


def make_header(page_number: int, relpath: str, feature: str) -> list[str]:
    return [
        "# Source Excerpt Volume 3",
        f"# Page {page_number:03d} of {PAGES:03d}",
        f"# Feature Group: {feature}",
        f"# Source File: {relpath}",
    ]


def build_pages(files: list[SourceFile], code_lines: list[tuple[str, str]]) -> list[list[str]]:
    by_relpath = {f.relpath: f for f in files}
    pages: list[list[str]] = []
    if not code_lines:
        raise RuntimeError("No usable code lines found")

    stride = max(1, len(code_lines) // PAGES)

    for page_number in range(1, PAGES + 1):
        start = ((page_number - 1) * stride) % len(code_lines)
        chunk = [code_lines[(start + offset) % len(code_lines)] for offset in range(CODE_LINES_PER_PAGE)]
        relpath = chunk[0][0]
        feature = by_relpath[relpath].feature
        page_lines = make_header(page_number, relpath, feature)
        page_lines.extend(line for _, line in chunk)
        pages.append(page_lines)
    return pages


def build_rtf(pages: list[list[str]]) -> str:
    parts = [
        r"{\rtf1\ansi\deff0",
        r"{\fonttbl{\f0\fmodern Courier New;}}",
        r"\paperw12240\paperh15840\margl720\margr720\margt720\margb720",
        r"\f0\fs16",
    ]

    for index, page in enumerate(pages):
        for line in page:
            safe_line = escape_rtf(line)
            parts.append(rf"\pard\sl180\slmult1 {safe_line}\par")
        if index != len(pages) - 1:
            parts.append(r"\page")

    parts.append("}")
    return "\n".join(parts)


def main() -> None:
    files = iter_source_files()
    code_lines = build_code_lines(files)
    pages = build_pages(files, code_lines)
    rtf = build_rtf(pages)
    TEMP_RTF.write_text(rtf, encoding="utf-8")

    subprocess.run(
        ["textutil", "-convert", "docx", str(TEMP_RTF), "-output", str(OUTPUT_DOCX)],
        check=True,
    )
    TEMP_RTF.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
