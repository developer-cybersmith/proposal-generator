from __future__ import annotations

import io
import math
import os
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

# Bar/donut colours aligned with "Graph of Capita BA 2" PPTX (doughnut chart1.xml order)
PPTX_SEVERITY_HEX: list[tuple[str, str]] = [
    ("Critical", "BA1A16"),
    ("High", "FE0000"),
    ("Medium", "FFC000"),
    ("Low", "05B050"),
    ("Informational", "0070C0"),
]

# Headings / chart titles (Capita BA 2 reference)
CAPITA_TITLE_PURPLE = "#7030A0"
# Default path for the reference deck (override with WEB_REPORT_CAPITA_PPTX)
DEFAULT_CAPITA_REFERENCE_PPTX = os.path.normpath(
    os.path.join("d:", os.sep, "intern", "Graph of Capita BA 2 [Auto-saved].pptx")
)

_NS_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_NS_DRAWING_MAIN = "http://schemas.openxmlformats.org/drawingml/2006/main"

SUMMARY_TOC_BOOKMARK = "_TocWebSummary"


def _capita_reference_pptx_path() -> str:
    return os.environ.get("WEB_REPORT_CAPITA_PPTX", DEFAULT_CAPITA_REFERENCE_PPTX)


def _first_srgb_hex_under(elem: ET.Element) -> str | None:
    for node in elem.iter(f"{{{_NS_DRAWING_MAIN}}}srgbClr"):
        val = (node.get("val") or "").strip()
        if val:
            return val.upper()
    return None


def _collect_series_hexes_from_chart_root(root: ET.Element) -> list[str]:
    """
    Collect solid fill RGB hex values from chart series (dPt first, else series spPr).
    Order matches typical Capita doughnut / clustered bar series order.
    """
    out: list[str] = []
    for ser in root.findall(f".//{{{_NS_CHART}}}ser"):
        dpts = ser.findall(f"{{{_NS_CHART}}}dPt")
        if dpts:
            for dpt in dpts:
                h = _first_srgb_hex_under(dpt)
                if h:
                    out.append(h)
            continue
        h = _first_srgb_hex_under(ser)
        if h:
            out.append(h)
    return out


def extract_capita_severity_palette_from_pptx(pptx_path: str) -> list[tuple[str, str]] | None:
    """
    If pptx_path exists, read ppt/charts/chart*.xml and derive up to five #RRGGBB series colours.
    Maps colours to Critical..Informational in extraction order (matches Capita deck).
    """
    if not pptx_path or not os.path.isfile(pptx_path):
        return None
    try:
        with zipfile.ZipFile(pptx_path, "r") as zf:
            chart_names = [
                n
                for n in zf.namelist()
                if n.startswith("ppt/charts/chart") and n.endswith(".xml")
            ]

            def _chart_idx(name: str) -> int:
                m = re.search(r"chart(\d+)\.xml$", name, re.I)
                return int(m.group(1)) if m else 0

            chart_names.sort(key=_chart_idx)
            for name in chart_names:
                raw = zf.read(name)
                root = ET.fromstring(raw)
                hexes = _collect_series_hexes_from_chart_root(root)
                if len(hexes) < 5:
                    continue
                labels = [pair[0] for pair in PPTX_SEVERITY_HEX]
                return list(zip(labels, hexes[:5]))
    except (OSError, zipfile.BadZipFile, ET.ParseError):
        return None
    return None


def resolve_web_report_chart_palette() -> list[tuple[str, str]]:
    """Prefer colours read from WEB_REPORT_CAPITA_PPTX when present and parseable."""
    custom = extract_capita_severity_palette_from_pptx(_capita_reference_pptx_path())
    return custom if custom else PPTX_SEVERITY_HEX


def _display_hostname_for_chart(vulnerabilities_by_asset: dict[str, list[dict]]) -> str:
    if not vulnerabilities_by_asset:
        return ""
    keys = list(vulnerabilities_by_asset.keys())
    if len(keys) == 1:
        return str(keys[0])
    return "All assessed targets (aggregated)"


def get_severity_colors(severity: str) -> dict[str, str]:
    colors = {
        "Critical": {"row1_bg": "#990000", "row2_bg": "#FF3333", "font": "#990000"},
        "High": {"row1_bg": "#FF0000", "row2_bg": "#FF6666", "font": "#FF0000"},
        "Medium": {"row1_bg": "#FFCC00", "row2_bg": "#FFCC66", "font": "#FFCC00"},
        "Low": {"row1_bg": "#009933", "row2_bg": "#99CC33", "font": "#009933"},
        "Informational": {"row1_bg": "#3399CC", "row2_bg": "#66CCFF", "font": "#3399CC"},
    }
    return colors.get(str(severity or "").title(), colors["Informational"])


def insert_risk_assessment_analysis_table(doc, vulnerabilities_by_asset: dict[str, list[dict]]) -> None:
    """
    Inserts a framework table at _RISK_ASSESSMENT_ANALYSIS_PLACEHOLDER_.
    Web Type-4 uses an \"asset\" key that is typically a URL-ish value.
    """
    placeholder = "_RISK_ASSESSMENT_ANALYSIS_PLACEHOLDER_"
    placeholder_paragraph = next((p for p in doc.paragraphs if placeholder in p.text), None)
    if not placeholder_paragraph:
        print(f"Warning: Placeholder {placeholder} not found in the template.")
        return

    columns = [
        "Sr. No.",
        "Hostname",
        "Instant Purpose",
        "VAPT Status",
        "Critical",
        "High",
        "Medium",
        "Low",
        "Informational",
        "Total",
    ]
    severity_keys = ["Critical", "High", "Medium", "Low", "Informational"]

    data_rows: list[list[str]] = []
    all_severity_totals = {k: 0 for k in severity_keys}
    grand_total = 0

    for idx, (asset, vulns) in enumerate(vulnerabilities_by_asset.items(), 1):
        row = [str(idx), str(asset), "Web App", "Completed"]
        sev_counts = {k: 0 for k in severity_keys}
        for v in vulns:
            sev = (v.get("severity") or "Informational").title()
            if sev in sev_counts:
                sev_counts[sev] += 1
        total = sum(sev_counts.values())
        for k in severity_keys:
            row.append(str(sev_counts[k]))
            all_severity_totals[k] += sev_counts[k]
        row.append(str(total))
        grand_total += total
        data_rows.append(row)

    table = doc.add_table(rows=1 + len(data_rows) + 1, cols=len(columns))
    table.style = "Table Grid"

    widths = [
        Inches(0.8),
        Inches(2.8),
        Inches(1.5),
        Inches(1.3),
        Inches(0.8),
        Inches(0.8),
        Inches(0.8),
        Inches(0.8),
        Inches(1.1),
        Inches(0.8),
    ]
    for i, w in enumerate(widths):
        table.columns[i].width = w

    header_row = table.rows[0]
    tr = header_row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement("w:trHeight")
    trHeight.set(qn("w:val"), str(int(1.1 * 1440)))
    trHeight.set(qn("w:hRule"), "exact")
    trPr.append(trHeight)

    for i, col in enumerate(columns):
        cell = header_row.cells[i]
        p = cell.paragraphs[0]
        p.text = col
        run = p.runs[0] if p.runs else p.add_run(col)
        run.font.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.vertical_alignment = 1

        if col in ["Sr. No.", "Hostname", "Instant Purpose", "VAPT Status"]:
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="8ea9db"/>')
            cell._element.get_or_add_tcPr().append(shading)
        elif col in severity_keys:
            colors = get_severity_colors(col)
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{colors["row1_bg"].lstrip("#")}"/>')
            cell._element.get_or_add_tcPr().append(shading)
            run.font.color.rgb = RGBColor(255, 255, 255)
            tcPr = cell._element.get_or_add_tcPr()
            tcPr.append(parse_xml(r'<w:textDirection w:val="btLr" %s/>' % nsdecls("w")))
        elif col == "Total":
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="6A449A"/>')
            cell._element.get_or_add_tcPr().append(shading)
            run.font.color.rgb = RGBColor(255, 255, 255)
        else:
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="D9D9D9"/>')
            cell._element.get_or_add_tcPr().append(shading)

    for row_idx, row_data in enumerate(data_rows, 1):
        cells = table.rows[row_idx].cells
        for i, val in enumerate(row_data):
            cells[i].text = val
            for para in cells[i].paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cells[i].vertical_alignment = 1
            if columns[i] in ["Sr. No.", "Hostname", "Instant Purpose", "VAPT Status"]:
                shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="bdd7ee"/>')
                cells[i]._element.get_or_add_tcPr().append(shading)
            elif columns[i] in severity_keys:
                colors = get_severity_colors(columns[i])
                shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{colors["row1_bg"].lstrip("#")}"/>')
                cells[i]._element.get_or_add_tcPr().append(shading)
                p = cells[i].paragraphs[0]
                if p.runs:
                    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            elif columns[i] == "Total":
                shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="6A449A"/>')
                cells[i]._element.get_or_add_tcPr().append(shading)
                p = cells[i].paragraphs[0]
                if p.runs:
                    p.runs[0].font.color.rgb = RGBColor(255, 255, 255)

    total_cells = table.rows[len(data_rows) + 1].cells
    merged_cell = total_cells[0].merge(total_cells[1]).merge(total_cells[2]).merge(total_cells[3])
    merged_cell.text = "Overall Findings"
    mp = merged_cell.paragraphs[0]
    if mp.runs:
        mp.runs[0].font.bold = True
    mp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    merged_cell.vertical_alignment = 1
    merged_cell._element.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="bdd7ee"/>'))

    for i, k in enumerate(severity_keys, start=4):
        total_cells[i].text = str(all_severity_totals[k])
        colors = get_severity_colors(k)
        total_cells[i]._element.get_or_add_tcPr().append(
            parse_xml(f'<w:shd {nsdecls("w")} w:fill="{colors["row1_bg"].lstrip("#")}"/>')
        )
        p = total_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        total_cells[i].vertical_alignment = 1
        if p.runs:
            p.runs[0].font.bold = True
            p.runs[0].font.color.rgb = RGBColor(255, 255, 255)

    total_cells[9].text = str(grand_total)
    total_cells[9]._element.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="6A449A"/>'))
    p = total_cells[9].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    total_cells[9].vertical_alignment = 1
    if p.runs:
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = RGBColor(255, 255, 255)

    parent = placeholder_paragraph._p.getparent()
    parent.insert(parent.index(placeholder_paragraph._p), table._element)
    parent.remove(placeholder_paragraph._p)


def create_combined_vulnerability_charts(
    severity_counts: Counter,
    output_path: str,
    *,
    display_hostname: str = "",
) -> str:
    """
    Capita BA 2–style figure: 3D severity bars + donut, legend to the right of the 3D panel,
    data labels on top of bars (black), y-scale from 0 with even ticks (min span 0–14 like reference).
    Colours from WEB_REPORT_CAPITA_PPTX when that file exists and chart XML exposes five sRGB fills.
    """
    import matplotlib.pyplot as plt
    from matplotlib.patches import Patch
    from mpl_toolkits.mplot3d import Axes3D  # noqa: F401

    plt.rcParams.update({
        "font.family": "sans-serif",
        "font.sans-serif": ["Calibri", "Arial", "DejaVu Sans"],
        "font.size": 11,
    })

    palette = resolve_web_report_chart_palette()
    all_severities = [name for name, _ in palette]
    all_colors = [f"#{h}" for _, h in palette]
    counts = [int(severity_counts.get(s, 0)) for s in all_severities]
    total_findings = sum(counts)

    fig = plt.figure(figsize=(13.2, 6.8))
    gs = fig.add_gridspec(1, 2, width_ratios=[1.45, 1.0], wspace=0.28)

    ax3d = fig.add_subplot(gs[0, 0], projection="3d")
    xs = list(range(len(all_severities)))
    dz = [max(0, c) for c in counts]
    ax3d.bar3d(
        xs,
        [0] * len(xs),
        [0] * len(xs),
        [0.62] * len(xs),
        [0.62] * len(xs),
        dz,
        shade=True,
        color=all_colors,
        edgecolor="white",
        linewidth=0.6,
    )
    ax3d.set_xticks([x + 0.31 for x in xs])
    ax3d.set_xticklabels(all_severities, rotation=11, ha="right", fontsize=11)
    ax3d.set_yticks([])
    ax3d.set_zlabel("Number of Vulnerabilities", fontsize=12, labelpad=8)

    raw_max = max(dz) if dz else 0
    zmax = max(14, raw_max)
    if zmax % 2:
        zmax += 1
    ax3d.set_zlim(0, zmax)
    ax3d.set_zticks(list(range(0, zmax + 1, 2)))

    ax3d.set_title(
        "Total Vulnerabilities found",
        fontsize=15,
        fontweight="bold",
        pad=16,
        color=CAPITA_TITLE_PURPLE,
    )
    ax3d.view_init(elev=22, azim=-58)

    z_pad = max(0.15, zmax * 0.02)
    for i, c in enumerate(dz):
        top = c + z_pad
        ax3d.text(
            xs[i] + 0.31,
            0.31,
            top,
            str(c),
            ha="center",
            va="bottom",
            fontsize=13,
            weight="bold",
            color="black",
        )

    legend_handles = [
        Patch(facecolor=all_colors[i], edgecolor="white", label=all_severities[i])
        for i in range(len(all_severities))
    ]
    ax3d.legend(
        handles=legend_handles,
        loc="upper left",
        bbox_to_anchor=(1.02, 0.92),
        fontsize=11,
        frameon=False,
        title=None,
    )

    caption = (display_hostname or "").strip()
    if caption:
        fig.text(0.24, 0.04, caption, ha="center", va="bottom", fontsize=10, color="#333333")

    ax_donut = fig.add_subplot(gs[0, 1])
    if total_findings == 0:
        ax_donut.text(0, 0, "Total Findings:\n0", ha="center", va="center", fontsize=13, weight="bold")
        ax_donut.set_xlim(-1, 1)
        ax_donut.set_ylim(-1, 1)
    else:
        wedges, _ = ax_donut.pie(
            counts,
            colors=all_colors,
            wedgeprops=dict(width=0.42, edgecolor="white", linewidth=1.0),
            startangle=90,
            labels=None,
        )
        for i, w in enumerate(wedges):
            if counts[i] == 0:
                continue
            ang = (w.theta2 + w.theta1) / 2.0
            x = math.cos(math.radians(ang)) * 0.78
            y = math.sin(math.radians(ang)) * 0.78
            ax_donut.text(x, y, str(counts[i]), ha="center", va="center", fontsize=13, weight="bold", color="white")

        centre_circle = plt.Circle((0, 0), 0.58, fc="white")
        ax_donut.add_artist(centre_circle)
        ax_donut.text(
            0,
            0,
            f"Total Findings:\n{total_findings}",
            ha="center",
            va="center",
            fontsize=13,
            weight="bold",
        )
    ax_donut.set_title(
        "Overall Vulnerabilities Identified",
        fontsize=15,
        weight="bold",
        pad=14,
        color=CAPITA_TITLE_PURPLE,
    )
    ax_donut.axis("equal")

    plt.subplots_adjust(left=0.02, right=0.98, top=0.92, bottom=0.12)
    plt.savefig(output_path, bbox_inches="tight", dpi=150)
    plt.close(fig)
    return output_path


def create_3d_vulnerability_bar_chart(
    severity_counts: Counter,
    output_path: str,
    *,
    display_hostname: str = "",
) -> str:
    import matplotlib.pyplot as plt
    from mpl_toolkits.mplot3d import Axes3D  # noqa: F401

    plt.rcParams.update({
        "font.family": "sans-serif",
        "font.sans-serif": ["Calibri", "Arial", "DejaVu Sans"],
    })

    palette = resolve_web_report_chart_palette()
    all_severities = [name for name, _ in palette]
    colors = [f"#{h}" for _, h in palette]
    counts = [int(severity_counts.get(s, 0)) for s in all_severities]

    fig = plt.figure(figsize=(10.5, 5.8))
    ax = fig.add_subplot(111, projection="3d")
    xs = list(range(len(all_severities)))
    dz = [max(0, c) for c in counts]
    ax.bar3d(xs, [0] * len(xs), [0] * len(xs), [0.62] * len(xs), [0.62] * len(xs), dz, shade=True, color=colors, edgecolor="white", linewidth=0.5)
    ax.set_xticks([x + 0.31 for x in xs])
    ax.set_xticklabels(all_severities, rotation=10, ha="right", fontsize=11)
    ax.set_yticks([])
    ax.set_zlabel("Number of Vulnerabilities", fontsize=12, labelpad=8)
    raw_max = max(dz) if dz else 0
    zmax = max(14, raw_max)
    if zmax % 2:
        zmax += 1
    ax.set_zlim(0, zmax)
    ax.set_zticks(list(range(0, zmax + 1, 2)))
    ax.set_title(
        "Total Vulnerabilities found",
        fontsize=15,
        fontweight="bold",
        pad=14,
        color=CAPITA_TITLE_PURPLE,
    )
    ax.view_init(elev=22, azim=-58)
    z_pad = max(0.15, zmax * 0.02)
    for i, c in enumerate(dz):
        ax.text(xs[i] + 0.31, 0.31, c + z_pad, str(c), ha="center", va="bottom", fontsize=12, weight="bold", color="black")
    caption = (display_hostname or "").strip()
    if caption:
        fig.text(0.5, 0.04, caption, ha="center", va="bottom", fontsize=10, color="#333333")
    plt.subplots_adjust(bottom=0.14, top=0.9)
    plt.savefig(output_path, bbox_inches="tight", dpi=150)
    plt.close(fig)
    return output_path


def _severity_counts_from_grouped(vulnerabilities_by_asset: dict[str, list[dict]]) -> Counter:
    all_vulns: list[dict] = []
    for vulns in vulnerabilities_by_asset.values():
        all_vulns.extend(vulns)
    severity_counts: Counter = Counter()
    for vuln in all_vulns:
        severity_counts[(vuln.get("severity") or "Unknown").title()] += 1
    return severity_counts


def insert_overall_vulnerabilities_charts(doc, vulnerabilities_by_asset: dict[str, list[dict]]) -> None:
    placeholder = "_VULNERABILITIES_FOUND_PLACEHOLDER_"
    placeholder_paragraph = next((p for p in doc.paragraphs if placeholder in p.text), None)
    if not placeholder_paragraph:
        print(f"Warning: Placeholder {placeholder} not found in the template.")
        return

    severity_counts = _severity_counts_from_grouped(vulnerabilities_by_asset)
    display_host = _display_hostname_for_chart(vulnerabilities_by_asset)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as chart_file:
        chart_path = create_combined_vulnerability_charts(
            severity_counts, chart_file.name, display_hostname=display_host
        )

    ph_idx = next((i for i, p in enumerate(doc.paragraphs) if placeholder in p.text), None)
    if ph_idx is not None and ph_idx > 0:
        prev = doc.paragraphs[ph_idx - 1]
        txt = (prev.text or "").strip()
        if "Vulnerabilities Found" in txt:
            try:
                prev.style = doc.styles["Heading 1"]
            except KeyError:
                pass
            prev.paragraph_format.page_break_before = True
            for run in prev.runs:
                run.font.name = "Altone Trial"
                run.bold = True
                run.font.color.rgb = RGBColor(106, 68, 154)
        else:
            h = placeholder_paragraph.insert_paragraph_before("Vulnerabilities Found")
            try:
                h.style = doc.styles["Heading 1"]
            except KeyError:
                pass
            h.paragraph_format.page_break_before = True
            rr = h.runs[0] if h.runs else h.add_run("Vulnerabilities Found")
            rr.font.name = "Altone Trial"
            rr.bold = True
            rr.font.color.rgb = RGBColor(106, 68, 154)

    img_paragraph = placeholder_paragraph.insert_paragraph_before()
    img_paragraph.add_run().add_picture(chart_path, width=Inches(6.5))
    img_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    parent = placeholder_paragraph._p.getparent()
    parent.remove(placeholder_paragraph._p)


def insert_overall_vulnerabilities_charts_3d(doc, vulnerabilities_by_asset: dict[str, list[dict]]) -> None:
    placeholder = "_VULNERABILITIES_FOUND_PLACEHOLDER_"
    placeholder_paragraph = next((p for p in doc.paragraphs if placeholder in p.text), None)
    if not placeholder_paragraph:
        print(f"Warning: Placeholder {placeholder} not found in the template.")
        return

    severity_counts = _severity_counts_from_grouped(vulnerabilities_by_asset)
    display_host = _display_hostname_for_chart(vulnerabilities_by_asset)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as chart_file:
        chart_path = create_3d_vulnerability_bar_chart(
            severity_counts, chart_file.name, display_hostname=display_host
        )

    ph_idx = next((i for i, p in enumerate(doc.paragraphs) if placeholder in p.text), None)
    if ph_idx is not None and ph_idx > 0:
        prev = doc.paragraphs[ph_idx - 1]
        txt = (prev.text or "").strip()
        if "Vulnerabilities Found" in txt:
            try:
                prev.style = doc.styles["Heading 1"]
            except KeyError:
                pass
            prev.paragraph_format.page_break_before = True
            for run in prev.runs:
                run.font.name = "Altone Trial"
                run.bold = True
                run.font.color.rgb = RGBColor(106, 68, 154)
        else:
            h = placeholder_paragraph.insert_paragraph_before("Vulnerabilities Found")
            try:
                h.style = doc.styles["Heading 1"]
            except KeyError:
                pass
            h.paragraph_format.page_break_before = True
            rr = h.runs[0] if h.runs else h.add_run("Vulnerabilities Found")
            rr.font.name = "Altone Trial"
            rr.bold = True
            rr.font.color.rgb = RGBColor(106, 68, 154)

    img_paragraph = placeholder_paragraph.insert_paragraph_before()
    img_paragraph.add_run().add_picture(chart_path, width=Inches(6.5))
    img_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    parent = placeholder_paragraph._p.getparent()
    parent.remove(placeholder_paragraph._p)


def _add_bookmark_to_paragraph(paragraph, name: str, bookmark_id: int) -> None:
    """Place bookmark range after w:pPr (required for correct PAGEREF in Word)."""
    p = paragraph._p
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    insert_at = 0
    for idx, child in enumerate(p):
        if child.tag == qn("w:pPr"):
            insert_at = idx + 1
            break
    p.insert(insert_at, start)
    p.append(end)


def _next_bookmark_id(doc) -> int:
    """Return max w:bookmarkStart id in document + 1."""
    m = 0
    body = doc.element.body
    for el in body.iter():
        if el.tag != qn("w:bookmarkStart"):
            continue
        bid = el.get(qn("w:id"))
        if bid and bid.isdigit():
            m = max(m, int(bid))
    return m + 1


def _summary_toc_paragraph_xml(bookmark_name: str, use_wps_toc_style: bool) -> str:
    """
    One TOC row: hyperlink + tab + PAGEREF field. Matches reference Web Application VAPT_report layout.
    """
    style = "WPSOffice1" if use_wps_toc_style else "TOC1"
    if use_wps_toc_style:
        p_pr = (
            f'<w:pPr><w:pStyle w:val="{style}"/><w:tabs>'
            '<w:tab w:val="right" w:leader="dot" w:pos="9360"/></w:tabs>'
            '<w:spacing w:line="360" w:lineRule="auto"/><w:rPr><w:b/></w:rPr></w:pPr>'
        )
        rfonts = (
            '<w:rPr><w:rFonts w:ascii="Altone Trial" w:hAnsi="Altone Trial" w:cs="Altone Trial"/>'
            '<w:b/><w:bCs/><w:szCs w:val="56"/></w:rPr>'
        )
        rb = '<w:rPr><w:b/></w:rPr>'
    else:
        p_pr = (
            f'<w:pPr><w:pStyle w:val="{style}"/><w:tabs>'
            '<w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs>'
            '<w:rPr><w:rFonts w:eastAsiaTheme="minorEastAsia"/><w:noProof/>'
            '<w:sz w:val="24"/><w:szCs w:val="24"/>'
            '<w:lang w:val="en-IN" w:eastAsia="en-IN"/></w:rPr></w:pPr>'
        )
        rfonts = (
            '<w:rPr><w:rStyle w:val="Hyperlink"/><w:rFonts w:ascii="Altone Trial" w:eastAsia="Calibri" '
            'w:hAnsi="Altone Trial"/><w:b/><w:bCs/><w:noProof/><w:lang w:eastAsia="en-IN"/></w:rPr>'
        )
        rb = '<w:rPr><w:noProof/><w:webHidden/></w:rPr>'

    return (
        f"<w:p>{p_pr}"
        f'<w:hyperlink w:anchor="{bookmark_name}" w:history="1">'
        f"<w:r>{rfonts}<w:t>SUMMARY OF FINDINGS &amp; CONCLUSION:</w:t></w:r>"
        f"<w:r>{rb}<w:tab/></w:r>"
        f'<w:r>{rb}<w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r>{rb}<w:instrText xml:space="preserve"> PAGEREF {bookmark_name} \\h </w:instrText></w:r>'
        f"<w:r>{rb}</w:r>"
        f'<w:r>{rb}<w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r>{rb}<w:t>1</w:t></w:r>'
        f'<w:r>{rb}<w:fldChar w:fldCharType="end"/></w:r>'
        "</w:hyperlink></w:p>"
    )


def _find_vulnerabilities_found_toc_insert_at(xml: str, search_limit: int = 900_000) -> int | None:
    """
    Return byte index in xml immediately after the </w:hyperlink></w:p> that closes the
    TOC entry for 'Vulnerabilities Found'. Handles split <w:t> runs and WPS/Word TOC styles.
    """
    head = xml[:search_limit]
    closing = "</w:hyperlink></w:p>"
    pos = 0
    candidates: list[int] = []
    while True:
        j = head.find(closing, pos)
        if j == -1:
            break
        start_win = max(0, j - 12000)
        win = head[start_win : j + len(closing)]
        if (
            "PAGEREF" in win
            and "w:hyperlink" in win
            and "Vulnerabilities" in win
            and "Found" in win
            and ("TOC1" in win or "WPSOffice1" in win or "TOC2" in win)
        ):
            candidates.append(j + len(closing))
        pos = j + 1
    return candidates[-1] if candidates else None


def _inject_summary_toc_into_document_xml(xml: str, bookmark_name: str, use_wps_toc_style: bool) -> tuple[str, bool]:
    """
    After the TOC line for 'Vulnerabilities Found' (first occurrence with PAGEREF in TOC), insert
    SUMMARY OF FINDINGS & CONCLUSION row if not already present.
    """
    if f"PAGEREF {bookmark_name}" in xml:
        return xml, True
    insert_at = None
    pos = 0
    while True:
        i = xml.find("<w:t>Vulnerabilities Found</w:t>", pos)
        if i == -1:
            break
        chunk = xml[i : i + 8000]
        if "PAGEREF" in chunk and "w:hyperlink" in chunk and (
            "TOC1" in chunk or "WPSOffice1" in chunk or "pStyle" in chunk
        ):
            j = chunk.find("</w:hyperlink></w:p>")
            if j != -1:
                insert_at = i + j + len("</w:hyperlink></w:p>")
                break
        pos = i + 20
    if insert_at is None:
        insert_at = _find_vulnerabilities_found_toc_insert_at(xml)
    if insert_at is None:
        return xml, False
    new_row = _summary_toc_paragraph_xml(bookmark_name, use_wps_toc_style)
    return xml[:insert_at] + new_row + xml[insert_at:], True


def _docx_uses_wps_toc_style(zip_bytes: bytes) -> bool:
    with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zf:
        s = zf.read("word/document.xml").decode("utf-8", errors="replace")
    return "WPSOffice1" in s[:800000]


def web_report_finalize(doc, output_path: str, *, summary_heading_paragraph) -> None:
    """
    Add bookmark to SUMMARY heading and inject TOC row so Contents lists
    'SUMMARY OF FINDINGS & CONCLUSION' with page ref (matches manual report).
    Call instead of doc.save when the summary section was appended at the end.
    """
    bm_id = _next_bookmark_id(doc)
    _add_bookmark_to_paragraph(summary_heading_paragraph, SUMMARY_TOC_BOOKMARK, bm_id)

    buf = io.BytesIO()
    doc.save(buf)
    raw = buf.getvalue()
    use_wps = _docx_uses_wps_toc_style(raw)
    with zipfile.ZipFile(io.BytesIO(raw), "r") as zin:
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    s = data.decode("utf-8", errors="replace")
                    s2, ok = _inject_summary_toc_into_document_xml(s, SUMMARY_TOC_BOOKMARK, use_wps)
                    if not ok:
                        print(
                            "Warning: Could not inject SUMMARY OF FINDINGS TOC row; "
                            "Contents may be missing that entry until you Update Field in Word."
                        )
                    else:
                        data = s2.encode("utf-8")
                zout.writestr(item, data)
