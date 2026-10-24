"""
PDF proof, blank-page check and measured Contents page numbers (plan D29, hard acceptance rules).

A generated .docx is rendered to PDF by a real layout engine, then the PDF is analysed:
  * blank or near-empty pages are listed (hard rule: none may exist),
  * the page on which each heading really lands is measured, so the Contents page numbers can be
    corrected from fact instead of the generator's line-count estimate.

Engines (env PROPOSAL_PDF_ENGINE = libreoffice | word | off; default "auto" = LibreOffice if installed,
otherwise Microsoft Word on a Windows PC that has it):
  libreoffice  `soffice --headless` - used on Railway/Linux (see nixpacks.toml)
  word         Microsoft Word through COM (Windows: exact pagination, the program the proposal is opened in)

Pagination can differ slightly between engines (fonts); Microsoft Word is the final authority.
Every function degrades gracefully: with no engine the proposal is still produced, unverified.
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path

MIN_PAGE_CHARS = 25           # a page with less text than this (after removing headers/footers) is blank
CHROME_SHARE = 0.6            # a line/image on at least this share of pages is header/footer chrome
_SOFFICE_NAMES = ("soffice", "libreoffice")
_SOFFICE_WINDOWS = (
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
)
_WORD_PATH = r"C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE"


class RenderError(RuntimeError):
    pass


# --------------------------------------------------------------------------- engines
def find_soffice() -> str | None:
    explicit = os.getenv("SOFFICE_PATH", "").strip()
    if explicit and Path(explicit).is_file():
        return explicit
    for name in _SOFFICE_NAMES:
        found = shutil.which(name)
        if found:
            return found
    for path in _SOFFICE_WINDOWS:
        if Path(path).is_file():
            return path
    return None


def word_available() -> bool:
    return sys.platform == "win32" and Path(_WORD_PATH).is_file()


def available_engine() -> str | None:
    """The engine that would be used now, or None."""
    wanted = os.getenv("PROPOSAL_PDF_ENGINE", "auto").strip().lower()
    if wanted == "off":
        return None
    if wanted == "word":
        return "word" if word_available() else None
    if wanted == "libreoffice":
        return "libreoffice" if find_soffice() else None
    if find_soffice():
        return "libreoffice"
    # auto: on a Windows PC with Microsoft Word installed, Word is the engine (it is also the program
    # the proposal will be opened in, so the preview is exactly what the user will see)
    return "word" if word_available() else None


def _render_libreoffice(docx: Path, pdf: Path, timeout: int) -> None:
    soffice = find_soffice()
    if not soffice:
        raise RenderError("LibreOffice (soffice) is not installed")
    work = Path(tempfile.mkdtemp(prefix="css_lo_"))
    try:
        profile = (work / "profile").as_uri()
        cmd = [soffice, "--headless", "--norestore", f"-env:UserInstallation={profile}",
               "--convert-to", "pdf", "--outdir", str(work), str(docx)]
        subprocess.run(cmd, check=True, capture_output=True, timeout=timeout)
        produced = work / (docx.stem + ".pdf")
        if not produced.is_file():
            raise RenderError("LibreOffice produced no PDF")
        shutil.move(str(produced), str(pdf))
    except subprocess.TimeoutExpired as exc:
        raise RenderError(f"LibreOffice timed out after {timeout}s") from exc
    except subprocess.CalledProcessError as exc:
        raise RenderError(f"LibreOffice failed: {exc.stderr[-300:]!r}") from exc
    finally:
        shutil.rmtree(work, ignore_errors=True)


_WORD_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
try {
  $word.Visible = $false; $word.DisplayAlerts = 0
  $doc = $word.Documents.Open($env:CSS_DOCX, $false, $true)
  $doc.ExportAsFixedFormat($env:CSS_PDF, 17)
  $doc.Close(0)
} finally { $word.Quit() }
"""


_WORD_LOCK = threading.Lock()   # Word's COM automation is unreliable with simultaneous calls


def _render_word(docx: Path, pdf: Path, timeout: int) -> None:
    if not word_available():
        raise RenderError("Microsoft Word is not available")
    with _WORD_LOCK:
        _render_word_locked(docx, pdf, timeout)


def _render_word_locked(docx: Path, pdf: Path, timeout: int) -> None:
    started = time.time()
    env = dict(os.environ, CSS_DOCX=str(docx), CSS_PDF=str(pdf))
    try:
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", _WORD_SCRIPT],
                       env=env, check=True, capture_output=True, timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        # stop only the Word instance this call started, never the user's own windows
        subprocess.run(["powershell", "-NoProfile", "-Command",
                        f"Get-Process WINWORD -ErrorAction SilentlyContinue | "
                        f"Where-Object {{ $_.StartTime -gt (Get-Date).AddSeconds(-{int(time.time() - started) + 5}) }} | "
                        "Stop-Process -Force"], capture_output=True)
        raise RenderError(f"Word timed out after {timeout}s") from exc
    except subprocess.CalledProcessError as exc:
        raise RenderError(f"Word failed: {exc.stderr[-300:]!r}") from exc
    if not pdf.is_file():
        raise RenderError("Word produced no PDF")


def render_pdf(docx_path: str | Path, pdf_path: str | Path | None = None, *,
               engine: str | None = None, timeout: int = 180) -> Path:
    """Render a .docx to PDF. Raises RenderError when no engine is usable."""
    docx = Path(docx_path).resolve()
    pdf = Path(pdf_path).resolve() if pdf_path else docx.with_suffix(".pdf")
    engine = engine or available_engine()
    if engine == "libreoffice":
        _render_libreoffice(docx, pdf, timeout)
    elif engine == "word":
        _render_word(docx, pdf, timeout)
    else:
        raise RenderError("no PDF engine available (install LibreOffice or set PROPOSAL_PDF_ENGINE)")
    return pdf


# --------------------------------------------------------------------------- analysis
@dataclass
class PdfAnalysis:
    page_count: int = 0
    page_text: list[str] = field(default_factory=list)       # full text of each page
    page_chars: list[int] = field(default_factory=list)      # text length without header/footer lines
    blank_pages: list[int] = field(default_factory=list)     # 1-based page numbers

    def norm_lines(self, page_no: int) -> list[str]:
        return [" ".join(line.split()).lower() for line in self.page_text[page_no - 1].splitlines() if line.strip()]


def _open_pdf(path: str | Path):
    try:
        import pymupdf as fitz
    except ImportError:                                       # older PyMuPDF
        import fitz  # type: ignore
    return fitz.open(str(path))


def analyze_pdf(pdf_path: str | Path, *, min_chars: int = MIN_PAGE_CHARS) -> PdfAnalysis:
    doc = _open_pdf(pdf_path)
    try:
        texts = [page.get_text() for page in doc]
        image_sets = [{img[0] for img in page.get_images(full=True)} for page in doc]
    finally:
        doc.close()
    n = len(texts)
    result = PdfAnalysis(page_count=n, page_text=texts)
    if n == 0:
        return result

    from collections import Counter
    line_pages: Counter = Counter()
    for t in texts:
        for line in {" ".join(l.split()).lower() for l in t.splitlines() if l.strip()}:
            line_pages[line] += 1
    chrome = {line for line, c in line_pages.items() if n >= 3 and c / n >= CHROME_SHARE}
    image_pages: Counter = Counter()
    for imgs in image_sets:
        for x in imgs:
            image_pages[x] += 1
    chrome_images = {x for x, c in image_pages.items() if n >= 3 and c / n >= CHROME_SHARE}

    for i, t in enumerate(texts):
        body = [l for l in t.splitlines() if l.strip() and " ".join(l.split()).lower() not in chrome]
        chars = len("".join(body).replace(" ", ""))
        result.page_chars.append(chars)
        unique_images = image_sets[i] - chrome_images
        if chars < min_chars and not unique_images:
            result.blank_pages.append(i + 1)
    return result


def measure_heading_pages(analysis: PdfAnalysis, annexure_count: int) -> dict[str, int]:
    """Page (1-based) on which each Contents-listed heading really starts.

    Keys: the fixed intro headings (lower-case), 'annexure N', 'terms & conditions',
    'conclusion & pricing'. Headings that cannot be found are omitted.
    """
    static = ["introduction", "our understanding of your requirement", "methodology overview",
              "vulnerability assessment methodology", "proposed timelines & project management", "team overview"]
    pages = range(1, analysis.page_count + 1)
    contents_page = next((p for p in pages if "contents" in analysis.norm_lines(p)
                          and "introduction" in " ".join(analysis.norm_lines(p)).lower()), None)
    start = (contents_page or 1) + 1
    found: dict[str, int] = {}

    def first_line_match(pred, from_page: int) -> int | None:
        for p in range(from_page, analysis.page_count + 1):
            if any(pred(line) for line in analysis.norm_lines(p)):
                return p
        return None

    cursor = start
    for label in static:
        p = first_line_match(lambda line, l=label: line == l, cursor)
        if p:
            found[label] = p
            cursor = p
    cursor = max(cursor, start)
    for n in range(1, annexure_count + 1):
        p = first_line_match(lambda line, k=n: re.match(rf"^annexure {k}\s*:", line) is not None, cursor)
        if p:
            found[f"annexure {n}"] = p
            cursor = p
    p = first_line_match(lambda line: line == "terms & conditions", cursor)
    if p:
        found["terms & conditions"] = p
        cursor = p
    p = first_line_match(lambda line: line == "conclusion & pricing", cursor)
    if p:
        found["conclusion & pricing"] = p
    return found


@dataclass
class PdfCheck:
    engine: str | None
    pdf_path: Path | None = None
    page_count: int = 0
    blank_pages: list[int] = field(default_factory=list)
    headings: dict[str, int] = field(default_factory=dict)
    note: str = ""

    @property
    def verified(self) -> bool:
        return self.engine is not None and self.pdf_path is not None

    def header_value(self) -> str:
        if not self.verified:
            return f"not verified ({self.note or 'no PDF engine'})"
        blanks = f", blank pages {self.blank_pages}" if self.blank_pages else ", no blank pages"
        return f"{self.engine}: {self.page_count} pages{blanks}"


def check_document(docx_path: str | Path, annexure_count: int, *, pdf_path: str | Path | None = None,
                   timeout: int = 180) -> PdfCheck:
    """Render + analyse. Never raises: a missing engine or a render failure comes back in .note."""
    engine = available_engine()
    if engine is None:
        return PdfCheck(engine=None, note="no PDF engine installed")
    try:
        pdf = render_pdf(docx_path, pdf_path, engine=engine, timeout=timeout)
        analysis = analyze_pdf(pdf)
    except Exception as exc:  # RenderError, PyMuPDF errors, ...
        return PdfCheck(engine=None, note=f"render failed: {exc}"[:200])
    return PdfCheck(engine=engine, pdf_path=pdf, page_count=analysis.page_count,
                    blank_pages=analysis.blank_pages,
                    headings=measure_heading_pages(analysis, annexure_count))
