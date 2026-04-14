from __future__ import annotations

import math
import tempfile
from collections import Counter

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor


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


def create_combined_vulnerability_charts(severity_counts: Counter, output_path: str) -> str:
    import matplotlib.pyplot as plt
    from matplotlib.patches import Patch

    all_severities = ["Critical", "High", "Medium", "Low", "Informational"]
    all_colors = [get_severity_colors(s)["row1_bg"] for s in all_severities]
    counts = [int(severity_counts.get(s, 0)) for s in all_severities]
    total_findings = sum(counts)

    fig = plt.figure(figsize=(12, 6))
    gs = fig.add_gridspec(2, 2, height_ratios=[5, 1])

    ax_bar = fig.add_subplot(gs[0, 0])
    x_pos = range(len(all_severities))
    bars = ax_bar.bar(x_pos, counts, color=all_colors)
    ax_bar.set_xticks(list(x_pos))
    ax_bar.set_xticklabels(all_severities, rotation=10, ha="right", fontsize=14)
    ax_bar.set_ylabel("Number of Vulnerabilities", fontsize=14)
    ax_bar.set_title("Total Vulnerabilities Found", fontsize=18)
    for i, bar in enumerate(bars):
        height = bar.get_height()
        ax_bar.text(
            bar.get_x() + bar.get_width() / 2,
            height / 2 if height else 0.1,
            str(counts[i]),
            ha="center",
            va="center",
            fontsize=16,
            weight="bold",
            color="white",
        )

    ax_donut = fig.add_subplot(gs[0, 1])
    wedges, _ = ax_donut.pie(
        counts,
        colors=all_colors,
        wedgeprops=dict(width=0.4, edgecolor="white"),
        startangle=90,
        labels=None,
    )
    for i, w in enumerate(wedges):
        ang = (w.theta2 + w.theta1) / 2.0
        x = math.cos(math.radians(ang)) * 0.8
        y = math.sin(math.radians(ang)) * 0.8
        ax_donut.text(x, y, str(counts[i]), ha="center", va="center", fontsize=16, weight="bold", color="white")

    centre_circle = plt.Circle((0, 0), 0.60, fc="white")
    ax_donut.add_artist(centre_circle)
    ax_donut.text(0, 0, f"Total\nFindings\n{total_findings}", ha="center", va="center", fontsize=14, weight="bold")
    ax_donut.set_title("Overall Vulnerabilities Identified", fontsize=18, weight="bold", pad=20, color="#660099")
    ax_donut.axis("equal")

    for spine in ("top", "right", "left", "bottom"):
        ax_bar.spines[spine].set_visible(False)
    ax_bar.tick_params(left=False, bottom=False)
    ax_bar.yaxis.set_ticks_position("none")
    ax_bar.xaxis.set_ticks_position("none")

    legend_ax = fig.add_subplot(gs[1, :])
    legend_ax.axis("off")
    legend_handles = [Patch(facecolor=all_colors[i], label=all_severities[i]) for i in range(len(all_severities))]
    legend_ax.legend(handles=legend_handles, loc="center", ncol=len(all_severities), fontsize=16, frameon=False)
    legend_ax.set_xticks([])
    legend_ax.set_yticks([])

    plt.tight_layout(rect=[0, 0.08, 1, 1])
    plt.savefig(output_path, bbox_inches="tight")
    plt.close(fig)
    return output_path


def create_3d_vulnerability_bar_chart(severity_counts: Counter, output_path: str) -> str:
    import matplotlib.pyplot as plt
    from mpl_toolkits.mplot3d import Axes3D  # noqa: F401

    all_severities = ["Critical", "High", "Medium", "Low", "Informational"]
    colors = [get_severity_colors(s)["row1_bg"] for s in all_severities]
    counts = [int(severity_counts.get(s, 0)) for s in all_severities]

    fig = plt.figure(figsize=(10, 5.5))
    ax = fig.add_subplot(111, projection="3d")
    xs = list(range(len(all_severities)))
    ax.bar3d(xs, [0] * len(xs), [0] * len(xs), [0.6] * len(xs), [0.6] * len(xs), counts, shade=True, color=colors)
    ax.set_xticks([x + 0.3 for x in xs])
    ax.set_xticklabels(all_severities, rotation=10, ha="right", fontsize=12)
    ax.set_yticks([])
    ax.set_zlabel("Count", fontsize=12)
    ax.set_title("Total Vulnerabilities Found", fontsize=16, pad=14)
    for i, c in enumerate(counts):
        ax.text(xs[i] + 0.3, 0.3, c + 0.1, str(c), ha="center", va="bottom", fontsize=12, weight="bold")
    plt.tight_layout()
    plt.savefig(output_path, bbox_inches="tight")
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
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as chart_file:
        chart_path = create_combined_vulnerability_charts(severity_counts, chart_file.name)

    heading = placeholder_paragraph.insert_paragraph_before("Overall Vulnerability Identified")
    heading.style = placeholder_paragraph.style
    for run in heading.runs:
        run.font.size = Pt(32)
        run.font.bold = True
        run.font.color.rgb = RGBColor(106, 68, 154)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

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
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as chart_file:
        chart_path = create_3d_vulnerability_bar_chart(severity_counts, chart_file.name)

    heading = placeholder_paragraph.insert_paragraph_before("Overall Vulnerability Identified")
    heading.style = placeholder_paragraph.style
    for run in heading.runs:
        run.font.size = Pt(32)
        run.font.bold = True
        run.font.color.rgb = RGBColor(106, 68, 154)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    img_paragraph = placeholder_paragraph.insert_paragraph_before()
    img_paragraph.add_run().add_picture(chart_path, width=Inches(6.5))
    img_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    parent = placeholder_paragraph._p.getparent()
    parent.remove(placeholder_paragraph._p)

