"""Build CSS_Proposal_Source.docx from the organisation's master document.

The master ("CyberSmithSECURE Annexures 2.docx") is the company standard: 8 fixed front pages
(cover, Contents, Introduction, Our Understanding + Team Overview, Methodology, Vulnerability
Assessment Methodology, phases, report + tools), Annexures 1-11, Terms & Conditions and
Conclusion & Pricing, with the standard header and footer on every page.

Two further annexures that the master leaves out (Post Breach Scenario Assessment and Active
Directory Assessment) stay available: they are copied from the previous source template and
placed after Annexure 11 as template annexures 12 and 13. The generator only emits an annexure
when the proposal selects it, so the master's own sequence is what a default proposal shows.

The result also gets:
  * the Contents control unwrapped to a "Contents" heading plus one marker paragraph (the generator
    writes the real, measured Contents there),
  * a ``{{ client_logo }}`` slot between "For" and the client name on page 2.

    python build_source_template.py --master "D:/intern/proposal/CyberSmithSECURE Annexures 2.docx" \
        --legacy CSS_Proposal_Source_legacy.docx --output CSS_Proposal_Source.docx

Then run ``python css_proposal_generator.py --build-template`` to refresh the Base Template.
"""
from __future__ import annotations

import argparse
import copy
import io
import re
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = Path(__file__).resolve().parent
CONTENTS_MARKER = "[[CONTENTS]]"
ANNEXURE_RE = re.compile(r"^Annexure\s+(\d+)\s*:", re.IGNORECASE)
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
LEGACY_KEYS = (1, 2)            # legacy annexure numbers to bring over, in order
FIRST_EXTRA_NUMBER = 12         # master has 1-11; the legacy ones become 12 and 13


def _text(el) -> str:
    return "".join(t.text or "" for t in el.iter(qn("w:t"))).strip()


def _is_page_break_only(el) -> bool:
    if el.tag != qn("w:p") or _text(el):
        return False
    return any(b.get(qn("w:type")) == "page" for b in el.iter(qn("w:br")))


def _heading_positions(body_children) -> list[tuple[int, int, str]]:
    hits = []
    for i, el in enumerate(body_children):
        if el.tag == qn("w:p"):
            t = _text(el)
            m = ANNEXURE_RE.match(t)
            if m:
                hits.append((i, int(m.group(1)), t))
    return hits


# ------------------------------------------------------------------------------------ import helpers
def _used_ids(elements, tag: str) -> list[str]:
    seen: list[str] = []
    for el in elements:
        for node in el.iter(qn(tag)):
            v = node.get(qn("w:val"))
            if v is not None and v not in seen:
                seen.append(v)
    return seen


def _import_numbering(dst, src, elements) -> dict[str, str]:
    """Copy the numbering definitions the elements use; return {old numId: new numId}."""
    src_part = src.part.numbering_part.element
    dst_part = dst.part.numbering_part.element
    nums = {n.get(qn("w:numId")): n for n in src_part.findall(qn("w:num"))}
    abstracts = {a.get(qn("w:abstractNumId")): a for a in src_part.findall(qn("w:abstractNum"))}
    next_abs = max([int(a.get(qn("w:abstractNumId"))) for a in dst_part.findall(qn("w:abstractNum"))] + [-1]) + 1
    next_num = max([int(n.get(qn("w:numId"))) for n in dst_part.findall(qn("w:num"))] + [0]) + 1
    abs_map: dict[str, str] = {}
    num_map: dict[str, str] = {}
    first_num = dst_part.find(qn("w:num"))
    for old in _used_ids(elements, "w:numId"):
        if old == "0" or old not in nums:
            continue
        src_num = nums[old]
        abs_id = src_num.find(qn("w:abstractNumId")).get(qn("w:val"))
        if abs_id not in abs_map:
            new_abs = copy.deepcopy(abstracts[abs_id])
            new_abs.set(qn("w:abstractNumId"), str(next_abs))
            nsid = new_abs.find(qn("w:nsid"))
            if nsid is not None:
                nsid.set(qn("w:val"), f"{(0x5A000000 + next_abs):08X}")          # keep nsid unique
            if first_num is not None:
                first_num.addprevious(new_abs)                                   # abstractNums precede nums
            else:
                dst_part.append(new_abs)
            abs_map[abs_id] = str(next_abs)
            next_abs += 1
        new_num = copy.deepcopy(src_num)
        new_num.set(qn("w:numId"), str(next_num))
        new_num.find(qn("w:abstractNumId")).set(qn("w:val"), abs_map[abs_id])
        dst_part.append(new_num)
        num_map[old] = str(next_num)
        next_num += 1
    return num_map


def _import_styles(dst, src, elements) -> None:
    """Make sure every paragraph/run style the elements name exists in the destination."""
    dst_ids = {s.style_id for s in dst.styles}
    src_styles = {s.style_id: s for s in src.styles}
    for tag in ("w:pStyle", "w:rStyle", "w:tblStyle"):
        for sid in _used_ids(elements, tag):
            if sid in dst_ids or sid not in src_styles:
                continue
            dst.styles.element.append(copy.deepcopy(src_styles[sid].element))
            dst_ids.add(sid)


def _import_relationships(dst, src, elements) -> None:
    """Re-point image relationships (a:blip, v:imagedata) at images copied into the destination."""
    cache: dict[str, str] = {}
    for el in elements:
        for node in el.iter():
            for attr in list(node.attrib):
                if not attr.startswith(f"{{{R_NS}}}"):
                    continue
                old = node.get(attr)
                if old not in src.part.rels:
                    continue
                rel = src.part.rels[old]
                if rel.is_external:
                    continue
                if old not in cache:
                    rid, _img = dst.part.get_or_add_image(io.BytesIO(rel.target_part.blob))
                    cache[old] = rid
                node.set(attr, cache[old])


def import_elements(dst, src, elements) -> list:
    copies = [copy.deepcopy(e) for e in elements]
    num_map = _import_numbering(dst, src, copies)
    for el in copies:
        for node in el.iter(qn("w:numId")):
            old = node.get(qn("w:val"))
            if old in num_map:
                node.set(qn("w:val"), num_map[old])
    _import_styles(dst, src, copies)
    _import_relationships(dst, src, copies)
    return copies


# ------------------------------------------------------------------------------------ assembly
def _legacy_blocks(legacy) -> dict[int, list]:
    children = list(legacy.element.body)
    heads = _heading_positions(children)
    blocks: dict[int, list] = {}
    for k, (idx, number, _t) in enumerate(heads):
        if number not in LEGACY_KEYS:
            continue
        end = heads[k + 1][0] if k + 1 < len(heads) else len(children)
        blocks[number] = children[idx:end]
    missing = [n for n in LEGACY_KEYS if n not in blocks]
    if missing:
        raise SystemExit(f"legacy source is missing annexure(s) {missing}")
    return blocks


def _retitle(elements, number: int) -> None:
    for el in elements:
        if el.tag == qn("w:p") and ANNEXURE_RE.match(_text(el)):
            ts = [t for t in el.iter(qn("w:t"))]
            title = re.sub(r"^Annexure\s+\d+\s*:\s*", "", _text(el))
            ts[0].text = f"Annexure {number}: {title}"
            for t in ts[1:]:
                t.text = ""
            ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            return


def _unwrap_contents(body) -> None:
    """Replace the Contents control by its heading plus one marker paragraph."""
    for sdt in list(body.findall(qn("w:sdt"))):
        content = sdt.find(qn("w:sdtContent"))
        paras = content.findall(qn("w:p")) if content is not None else []
        if not paras or _text(paras[0]) != "Contents":
            continue
        heading = paras[0]
        marker = OxmlElement("w:p")
        run = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = CONTENTS_MARKER
        run.append(t)
        marker.append(run)
        sdt.addprevious(heading)
        sdt.addprevious(marker)
        body.remove(sdt)
        return
    raise SystemExit("Contents control not found in the master")


FRONT_PAGE_STARTS = (
    "OUR UNDERSTANDING OF YOUR REQUIREMENT",                                  # page 4 (with Team Overview)
    "METHODOLOGY OVERVIEW",                                                   # page 5
    "VULNERABILITY ASSESSMENT METHODOLOGY",                                   # page 6
    "PLANNING PHASE",                                                         # page 7 (numbered phases)
    "The vulnerability assessment/penetration testing report consist of:",    # page 8 (report + tools)
)


def _pin_front_pages(body) -> int:
    """Make each fixed front page start on a new page with ``pageBreakBefore``.

    The master pushes these sections onto the next page with rows of empty paragraphs. That fits
    only while the text metrics are exactly the master's: one line more and the spill leaves a
    near-empty page. Each start paragraph gets ``pageBreakBefore`` and the empty paragraphs right
    before it go (anchored pictures and bookmarks are kept). Returns how many pages were pinned.
    """
    wanted = {" ".join(t.split()).lower() for t in FRONT_PAGE_STARTS}
    pinned = 0
    done: set[str] = set()
    for el in list(body):
        if el.tag != qn("w:p"):
            continue
        text = " ".join(_text(el).split())
        if ANNEXURE_RE.match(text):
            break                                     # the annexures have headings of their own
        if text.lower() not in wanted or text.lower() in done:
            continue
        done.add(text.lower())
        prev = el.getprevious()
        while prev is not None:
            before = prev.getprevious()
            if prev.tag == qn("w:p"):
                if _text(prev) or prev.find(f".//{qn('w:drawing')}") is not None \
                        or prev.find(f".//{qn('w:pict')}") is not None or prev.find(f".//{qn('w:sectPr')}") is not None \
                        or _is_page_break_only(prev):
                    break
                body.remove(prev)
            prev = before                     # bookmark markers and the like are skipped, not removed
        ppr = el.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            el.insert(0, ppr)
        if ppr.find(qn("w:pageBreakBefore")) is None:
            node = OxmlElement("w:pageBreakBefore")
            anchor = None
            for tag in ("w:pStyle", "w:keepNext", "w:keepLines"):
                found = ppr.find(qn(tag))
                if found is not None:
                    anchor = found
            if anchor is not None:
                anchor.addnext(node)
            else:
                ppr.insert(0, node)
        pinned += 1
    return pinned


def _pull_up_lead_in(body) -> bool:
    """Keep the "...involves the following phases as identified below:" line on page 6.

    The master leaves a dozen empty paragraphs above it, which push it to the very bottom of
    page 6; without the master's exact metrics it spills onto page 7 and pushes the phases down.
    All but one empty paragraph above it are removed so it follows the text on page 6.
    """
    for el in body:
        if el.tag == qn("w:p") and "Penetration Testing methodology involves the following phases" in _text(el):
            kept = 0
            prev = el.getprevious()
            while prev is not None:
                before = prev.getprevious()
                if prev.tag == qn("w:p"):
                    if _text(prev) or prev.find(f".//{qn('w:drawing')}") is not None \
                            or prev.find(f".//{qn('w:pict')}") is not None or _is_page_break_only(prev):
                        break
                    if kept >= 1:
                        body.remove(prev)
                    kept += 1
                prev = before
            return True
    return False


def _merge_split_text(body, wanted: str) -> int:
    """Join a paragraph whose text is *wanted* but is split over several runs ("15","-0","9",...).

    Word splits a typed date into many runs; the generator replaces literal text inside one run, so
    the master's cover date (in the drawing and in its fallback copy) is merged into the first run.
    Returns how many paragraphs were merged.
    """
    merged = 0
    for p in body.iter(qn("w:p")):
        ts = list(p.iter(qn("w:t")))
        if len(ts) > 1 and "".join(t.text or "" for t in ts).strip() == wanted:
            ts[0].text = wanted
            for t in ts[1:]:
                t.text = ""
            merged += 1
    return merged


def _harden_cover(body) -> int:
    """End the cover with a real page break instead of ~50 empty filler paragraphs.

    In the master the cover paragraph is followed by empty paragraphs that merely push "For" onto
    page 2; that only works while the font metrics match. ``pageBreakBefore`` on "For" does the
    same job exactly. Returns how many filler paragraphs were removed.
    """
    children = list(body)
    for_idx = next((i for i, el in enumerate(children) if el.tag == qn("w:p") and _text(el) == "For"), None)
    if for_idx is None:
        raise SystemExit("master has no 'For' paragraph")
    removed = 0
    for el in children[1:for_idx]:
        if el.tag == qn("w:p") and not _text(el) and el.find(f".//{qn('w:drawing')}") is None \
                and el.find(f".//{qn('w:pict')}") is None and el.find(f".//{qn('w:sectPr')}") is None:
            body.remove(el)
            removed += 1
    p = children[for_idx]
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        p.insert(0, ppr)
    if ppr.find(qn("w:pageBreakBefore")) is None:
        node = OxmlElement("w:pageBreakBefore")
        style = ppr.find(qn("w:pStyle"))
        if style is not None:
            style.addnext(node)
        else:
            ppr.insert(0, node)
    return removed


def _add_logo_slot(body) -> None:
    """Put {{ client_logo }} in the empty paragraph between 'For' and the client-name line (page 2)."""
    children = list(body)
    for i, el in enumerate(children):
        if el.tag == qn("w:p") and _text(el) == "For":
            for nxt in children[i + 1:i + 8]:
                if nxt.tag != qn("w:p"):
                    break
                if "Company name" in _text(nxt):
                    return
                if not _text(nxt) and nxt.find(f".//{qn('w:drawing')}") is None:
                    for child in list(nxt):
                        if child.tag != qn("w:pPr"):
                            nxt.remove(child)
                    ppr = nxt.find(qn("w:pPr"))
                    if ppr is None:
                        ppr = OxmlElement("w:pPr")
                        nxt.insert(0, ppr)
                    jc = ppr.find(qn("w:jc"))
                    if jc is None:
                        jc = OxmlElement("w:jc")
                        ppr.append(jc)
                    jc.set(qn("w:val"), "left")
                    run = OxmlElement("w:r")
                    t = OxmlElement("w:t")
                    t.text = "{{ client_logo }}"
                    run.append(t)
                    nxt.append(run)
                    return
    raise SystemExit("could not find the client-logo slot after 'For' on page 2")


def build(master: Path, legacy: Path, output: Path) -> Path:
    doc = Document(str(master))
    legacy_doc = Document(str(legacy))
    body = doc.element.body
    children = list(body)

    heads = _heading_positions(children)
    if not heads:
        raise SystemExit("master has no annexure headings")
    terms = next((i for i, el in enumerate(children)
                  if el.tag == qn("w:p") and _text(el) == "Terms & Conditions" and i > heads[-1][0]), None)
    if terms is None:
        raise SystemExit("master has no Terms & Conditions heading after the last annexure")
    anchor_idx = terms - 1 if _is_page_break_only(children[terms - 1]) else terms
    anchor = children[anchor_idx]

    for offset, (number, block) in enumerate(sorted(_legacy_blocks(legacy_doc).items())):
        new_number = FIRST_EXTRA_NUMBER + offset
        copies = import_elements(doc, legacy_doc, block)
        _retitle(copies, new_number)
        lead = copies[0]
        if not _is_page_break_only(lead):
            anchor.addprevious(_make_break())
        for el in copies:
            anchor.addprevious(el)

    _unwrap_contents(body)
    if not _merge_split_text(body, "15-09-2025"):
        raise SystemExit("could not find the master's cover date (15-09-2025)")
    if not _pull_up_lead_in(body):
        raise SystemExit("could not find the phases lead-in line in the master")
    if _pin_front_pages(body) != len(FRONT_PAGE_STARTS):
        raise SystemExit("could not find every fixed front-page start in the master")
    _harden_cover(body)
    _add_logo_slot(body)
    output.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output))
    return output


def _make_break():
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r.append(br)
    p.append(r)
    return p


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--master", required=True, type=Path, help="CyberSmithSECURE Annexures 2.docx")
    ap.add_argument("--legacy", required=True, type=Path, help="previous CSS_Proposal_Source.docx (13 annexures)")
    ap.add_argument("--output", type=Path, default=HERE / "CSS_Proposal_Source.docx")
    args = ap.parse_args()
    out = build(args.master, args.legacy, args.output)
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
