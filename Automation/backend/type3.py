import os
import sys

# Project root (user_use-main) for db, excel_parser, Automation.*
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

# Windows consoles can default to cp1252; emoji/debug prints from this module and
# imported modules should never crash request handling because of encoding errors.
for _stream_name in ("stdout", "stderr"):
    _stream = getattr(sys, _stream_name, None)
    if _stream and hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

from fastapi import FastAPI, UploadFile, File, HTTPException, Form, Request, Depends
from fastapi.responses import JSONResponse, FileResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from datetime import datetime
import traceback
import shutil
import tempfile
import json
import re
import zipfile
import io
from sqlalchemy.orm import Session
from db import get_db, AuditLog, CertINReport

from excel_parser import parse_excel_data, format_for_template

# Shared Excel converter (any-format -> feature-format)
from Automation.backend.excel_converter import convert_any_excel_to_format

VERSION = "2.2.0"

app = FastAPI()
# Session cookie is owned only by Report-Generator-IP-main/app.py (no nested SessionMiddleware here).

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

def sanitize_filename(filename):
    if not filename:
        return "upload.xlsx"
    safe_name = re.sub(r"[^A-Za-z0-9._-]", "_", filename)
    return safe_name

def read_vulnerability_excel(file_path):
    """
    Read vulnerability data from Excel using standardized parser with fallback to legacy parsing.
    Converts the new format to legacy format for compatibility with existing code.
    """
    # Try standardized parser first
    try:
        parsed_data = parse_excel_data(file_path)
        
        # Check if we got valid data
        if parsed_data and parsed_data.get('vulnerabilities'):
            print(f"✅ Parsed Excel with standardized parser")
            print(f"   Metadata: {parsed_data['metadata']}")
            print(f"   Assets: {len(parsed_data['assets'])} assets")
            print(f"   Vulnerabilities: {len(parsed_data['vulnerabilities'])} vulnerabilities")
            
            # Convert to legacy format expected by rest of type3.py code
            vulnerabilities = []
            for idx, vuln in enumerate(parsed_data['vulnerabilities'], 1):
                vuln_data = {
                    # Map new format to old format
                    'observation_number': str(vuln.get('sr_no', idx)),
                    'observation': vuln.get('observation', vuln.get('sr_no', idx)),
                    'new_or_repeat': vuln.get('new_or_re', 'New'),
                    'asset': vuln.get('affected_asset', ''),
                    'purpose': '',  # Not directly in vuln, would come from assets
                    'vapt_status': vuln.get('status', 'Open'),
                    'tester_name': vuln.get('tester', parsed_data['metadata']['tester']),
                    'project': vuln.get('project', parsed_data['metadata']['project']),
                    'client': vuln.get('client', parsed_data['metadata']['client']),
                    'cve_cwe': vuln.get('cve_cwe', ''),
                    'cvss': str(vuln.get('cvss', '')),
                    'cvss_version': '',  # Can be extracted from cvss_vector if needed
                    'cvss_vector': vuln.get('cvss_vector', ''),
                    'affected_asset': vuln.get('affected_asset', '') or vuln.get('ip_url_app', ''),
                    'title': vuln.get('observation', ''),
                    'description': vuln.get('detailed_observation', '') or vuln.get('observation_summary', ''),
                    'recommendation': vuln.get('recommendation', ''),
                    'reference': vuln.get('reference', ''),
                    'evidence': vuln.get('evidence', ''),
                    'severity': vuln.get('severity', 'Medium'),
                    'status': vuln.get('status', 'Open'),
                    # Count fields (not used in type3 but kept for compatibility)
                    'critical': '',
                    'high': '',
                    'medium': '',
                    'low': '',
                    'informational': '',
                    'total': ''
                }
                
                # If evidence is empty but we have steps, create evidence text from steps
                if not vuln_data['evidence'] and vuln.get('steps'):
                    evidence_lines = []
                    for step in vuln['steps']:
                        evidence_lines.append(f"Step {step['number']}: {step['content']}")
                    vuln_data['evidence'] = '\n'.join(evidence_lines)
                
                vulnerabilities.append(vuln_data)
                print(f"📄 Row {idx}: Tester={vuln_data['tester_name']}, Project={vuln_data['project']}, Client={vuln_data['client']}, Title={vuln_data['title'][:50]}")
            
            print(f"✅ Converted {len(vulnerabilities)} vulnerabilities to legacy format")
            return vulnerabilities
        else:
            print("⚠️ Standardized parser returned no vulnerabilities, falling back to legacy parsing")
            raise Exception("No vulnerabilities found in standardized parser")
            
    except Exception as e:
        print(f"⚠️ Standardized parser failed: {e}")
        print("📋 Falling back to legacy Excel parsing...")
        traceback.print_exc()
        
        # FALLBACK: Use original/legacy parsing logic
        # If you had a legacy parser before, restore it here
        # For now, return empty to avoid breaking
        return []

async def process_poc_zip_files(poc_files, vulnerabilities):
    """
    Process PoC zip with structure: POC_certIN.zip/POC_certIN/#1, #2, etc.
    Each folder contains images for that specific observation only
    """
    poc_mapping = {}
    
    for poc_file in poc_files:
        if not poc_file.filename:
            continue
            
        temp_zip_path = os.path.join(UPLOAD_DIR, f"temp_{poc_file.filename}")
        content = await poc_file.read()
        with open(temp_zip_path, "wb") as f:
            f.write(content)
        
        extract_dir = os.path.join(UPLOAD_DIR, "poc_images")
        # Clean extract directory to prevent contamination from previous uploads
        if os.path.exists(extract_dir):
            import shutil
            shutil.rmtree(extract_dir)
        os.makedirs(extract_dir, exist_ok=True)
        
        try:
            with zipfile.ZipFile(temp_zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
            
            print(f"✅ Extracted zip to: {extract_dir}")
            
            # Find all observation folders (#1, #2, #3, etc.)
            for root, dirs, files in os.walk(extract_dir):
                for dir_name in dirs:
                    # Check if directory name matches #1, #2, etc.
                    match = re.match(r'#(\d+)', dir_name)
                    if match:
                        obs_num = int(match.group(1))
                        obs_key = f"OBS-{obs_num:03d}"
                        obs_folder = os.path.join(root, dir_name)
                        
                        print(f"📁 Found observation folder: {obs_folder} for {obs_key}")
                        
                        # Get all images in this observation folder ONLY
                        obs_images = []
                        for img_file in sorted(os.listdir(obs_folder)):
                            if img_file.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                                img_path = os.path.join(obs_folder, img_file)
                                
                                # Extract step number from filename
                                step_match = re.search(r'step[_\-\s]?(\d+)', img_file.lower())
                                if step_match:
                                    step_num = int(step_match.group(1))
                                else:
                                    # Try just extracting a number
                                    num_match = re.search(r'(\d+)', img_file)
                                    step_num = int(num_match.group(1)) if num_match else len(obs_images) + 1
                                
                                obs_images.append({
                                    'filename': img_file,
                                    'path': img_path,
                                    'step_number': step_num
                                })
                                print(f"  🖼️  Found image: {img_file} (Step {step_num})")
                        
                        # Sort images by step number
                        obs_images.sort(key=lambda x: x['step_number'])
                        
                        if obs_images:
                            poc_mapping[obs_key] = obs_images
                            print(f"✅ Mapped {len(obs_images)} images to {obs_key}")
            
        except Exception as e:
            print(f"Error extracting PoC zip: {e}")
            traceback.print_exc()
        finally:
            if os.path.exists(temp_zip_path):
                os.remove(temp_zip_path)
    
    print(f"🎯 Final PoC mapping keys: {list(poc_mapping.keys())}")
    for key, images in poc_mapping.items():
        print(f"   {key}: {len(images)} images")
    return poc_mapping

def parse_evidence_steps(evidence_text):
    """Parse evidence/PoC text to extract step descriptions"""
    if not evidence_text or evidence_text.lower() in ['nan', 'none', '']:
        return []
    
    steps = []
    
    # Try to split by "Step X:" pattern
    step_pattern = re.compile(r'Step\s*(\d+)\s*[:\-]?\s*(.+?)(?=Step\s*\d+|$)', re.IGNORECASE | re.DOTALL)
    matches = step_pattern.findall(evidence_text)
    
    if matches:
        for step_num, step_desc in matches:
            steps.append({
                'number': int(step_num),
                'description': step_desc.strip()
            })
    else:
        # If no step pattern, split by newlines and number each
        lines = [line.strip() for line in evidence_text.split('\n') if line.strip()]
        for i, line in enumerate(lines, 1):
            steps.append({
                'number': i,
                'description': line
            })
    
    return steps

def safe_str(val, default=''):
    """Convert value to string, handling NaN/None/float safely"""
    if val is None:
        return default
    if isinstance(val, float) and pd.isna(val):
        return default
    return str(val) if val else default

def generate_vulnerability_sections(vulnerabilities, poc_mapping):
    """Generate vulnerability sections with proper data extraction"""
    vulnerability_sections = []
    
    for i, vuln in enumerate(vulnerabilities, 1):
        print(f"\nProcessing vulnerability {i}:")
        print(f"Available keys: {list(vuln.keys())}")
        
        # Extract data using standard keys from normalized mapping - use safe_str to handle NaN
        obs_num = safe_str(vuln.get('observation_number'), str(i))
        new_or_repeat = safe_str(vuln.get('new_or_repeat'), 'New')
        cve_cwe = safe_str(vuln.get('cve_cwe'))
        cvss_version = safe_str(vuln.get('cvss_version'))
        affected_asset = safe_str(vuln.get('affected_asset'))
        title = safe_str(vuln.get('title'))
        description = safe_str(vuln.get('description'))
        recommendation = safe_str(vuln.get('recommendation'))
        reference = safe_str(vuln.get('reference'))
        evidence_text = safe_str(vuln.get('evidence'))
        
        # Extract severity and other metadata
        severity = safe_str(vuln.get('severity'), 'Medium')
        cvss = safe_str(vuln.get('cvss'))
        cvss_vector = safe_str(vuln.get('cvss_vector'))
        status = safe_str(vuln.get('status'), 'Open')
        
        print(f"Extracted - Title: {title}, CVE: {cve_cwe}, Asset: {affected_asset}")
        
        # Parse evidence steps from Excel
        evidence_steps = parse_evidence_steps(evidence_text)
        
        # Get PoC images for this observation
        obs_key = f"OBS-{i:03d}"
        poc_images = poc_mapping.get(obs_key, [])
        
        print(f"Found {len(poc_images)} PoC images for {obs_key}")
        print(f"Found {len(evidence_steps)} evidence steps from Excel")
        
        # Match images with step descriptions
        matched_poc = []
        for idx, img in enumerate(poc_images):
            step_desc = ''
            if idx < len(evidence_steps):
                step_desc = evidence_steps[idx]['description']
            
            matched_poc.append({
                'filename': img['filename'],
                'path': img['path'],
                'step_number': idx + 1,
                'description': step_desc
            })
        
        # Process recommendations
        recommendations_list = []
        if recommendation and isinstance(recommendation, str):
            rec_lines = [line.strip() for line in recommendation.split('\n') if line.strip()]
            for line in rec_lines:
                clean_line = re.sub(r'^\d+\.\s*', '', line)
                if clean_line:
                    recommendations_list.append(clean_line)
        
        vulnerability_section = {
            'observation_number': i,
            'new_or_repeat': new_or_repeat,
            'severity': severity,
            'status': status,
            'cve_cwe': cve_cwe,
            'cvss': cvss,
            'cvss_version': cvss_version,
            'cvss_vector': cvss_vector,
            'affected_asset': affected_asset,
            'title': title,
            'description': description,
            'recommendations': recommendations_list,
            'reference': reference,
            'poc_images': matched_poc,
            'has_poc': len(matched_poc) > 0
        }
        
        vulnerability_sections.append(vulnerability_section)
        print(f"Created section with {len(matched_poc)} PoC images")
    
    return vulnerability_sections



def create_landscape_vulnerability_box(doc, vulnerability_section):
    """Create properly formatted vulnerability box matching expected layout"""
    import os
    from docx.shared import Inches, Pt, RGBColor, Twips, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls, qn

    def _clear_cell(cell):
        """Remove extra paragraphs and clear the first paragraph safely."""
        for p in cell.paragraphs[1:]:
            p._element.getparent().remove(p._element)
        cell.paragraphs[0].clear()
        return cell.paragraphs[0]

    # Get severity colors
    severity = safe_str(vulnerability_section.get('severity', 'Medium')).strip()
    severity_colors = {
        'Critical': {'bg': '800000', 'text': 'FFFFFF'},
        'High': {'bg': 'FF0000', 'text': 'FFFFFF'},
        'Medium': {'bg': 'FFC000', 'text': '000000'},
        'Low': {'bg': '92D050', 'text': 'FFFFFF'},
        'Informational': {'bg': '00B0F0', 'text': 'FFFFFF'}
    }
    colors = severity_colors.get(severity, severity_colors['Medium'])

    # Create main table - Row 1: Obs | Severity | Status, Row 2: New/Repeat | CVE/CWE
    table = doc.add_table(rows=2, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Set total table width to page width
    tbl_elem = table._tbl
    tblPr_elem = tbl_elem.tblPr if tbl_elem.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="9350" w:type="dxa"/>')
    existing_tblW = tblPr_elem.find(qn('w:tblW'))
    if existing_tblW is not None:
        tblPr_elem.remove(existing_tblW)
    tblPr_elem.append(tblW)

    # Set column widths
    col_widths_twips = [3740, 2805, 2805]
    for i, col in enumerate(table.columns):
        for cell in col.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{col_widths_twips[i]}" w:type="dxa"/>')
            existing = tcPr.find(qn('w:tcW'))
            if existing is not None:
                tcPr.remove(existing)
            tcPr.append(tcW)

    # Set light gray border color for the table
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:color="CCCCCC"/>'
        '<w:left w:val="single" w:sz="4" w:color="CCCCCC"/>'
        '<w:bottom w:val="single" w:sz="4" w:color="CCCCCC"/>'
        '<w:right w:val="single" w:sz="4" w:color="CCCCCC"/>'
        '<w:insideH w:val="single" w:sz="4" w:color="CCCCCC"/>'
        '<w:insideV w:val="single" w:sz="4" w:color="CCCCCC"/>'
        '</w:tblBorders>'
    )
    tblPr.append(tblBorders)

    # ---------- Row 1: Observation | Severity | Status ----------
    row1 = table.rows[0]

    # Observation cell
    cell_obs = row1.cells[0]
    p = _clear_cell(cell_obs)
    run = p.add_run(f"Observation: #{vulnerability_section.get('observation_number', '1')}")
    run.font.name = 'Altone Trial'
    run.font.size = Pt(11)
    run.font.bold = True

    # Severity cell with colored background
    cell_sev = row1.cells[1]
    p = _clear_cell(cell_sev)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"Severity: {severity}")
    run.font.name = 'Altone Trial'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255) if colors['text'] == 'FFFFFF' else RGBColor(0, 0, 0)
    cell_sev._element.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{colors["bg"]}"/>'))

    # Status cell
    cell_status = row1.cells[2]
    p = _clear_cell(cell_status)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"Status: {safe_str(vulnerability_section.get('status', 'Open'))}")
    run.font.name = 'Altone Trial'
    run.font.size = Pt(11)
    run.font.bold = True
    cell_status._element.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="FFD966"/>'))

    # ---------- Row 2: New/Repeat | CVE/CWE ----------
    row2 = table.rows[1]

    # New or Repeat Observation cell
    cell_nor = row2.cells[0]
    p = _clear_cell(cell_nor)
    run = p.add_run("New or Repeat Observation: ")
    run.font.name = 'Altone Trial'
    run.font.size = Pt(10)
    run.font.bold = True
    run = p.add_run(safe_str(vulnerability_section.get('new_or_repeat', 'New')))
    run.font.name = 'Altone Trial'
    run.font.size = Pt(10)

    # CVE/CWE cell (merge cells 1 and 2)
    cell_cve = row2.cells[1]
    cell_cve.merge(row2.cells[2])
    p = _clear_cell(cell_cve)
    run = p.add_run("CVE/CWE: ")
    run.font.name = 'Altone Trial'
    run.font.size = Pt(10)
    run.font.bold = True
    run = p.add_run(safe_str(vulnerability_section.get('cve_cwe', '')))
    run.font.name = 'Altone Trial'
    run.font.size = Pt(10)

    # ---------- Content rows ----------
    def add_content_row(label, value, is_url=False):
        if not value:
            return
        row = table.add_row()
        cell = row.cells[0]
        cell.merge(row.cells[1])
        cell.merge(row.cells[2])

        p = _clear_cell(cell)
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)

        run = p.add_run(label)
        run.font.name = 'Altone Trial'
        run.font.size = Pt(10)
        run.font.bold = True

        run = p.add_run(safe_str(value))
        run.font.name = 'Altone Trial'
        run.font.size = Pt(10)
        if is_url:
            run.font.color.rgb = RGBColor(0, 0, 255)
            run.font.underline = True

    # CVSS row
    add_content_row("CVSS: ", vulnerability_section.get('cvss'))

    # CVSS Version Ref row
    add_content_row("CVSS Version Ref: ", vulnerability_section.get('cvss_vector'))

    # Affected Asset row
    add_content_row("Affected Asset i.e. IP/URL/Application etc.: ",
                    vulnerability_section.get('affected_asset'), is_url=True)

    # Observation/Vulnerability Title row
    add_content_row("Observation/ Vulnerability Title: ", vulnerability_section.get('title'))

    # Detailed Observation row
    add_content_row("Detailed Observation/ Vulnerable Point: ", vulnerability_section.get('description'))

    # Recommendation row
    recommendations = vulnerability_section.get('recommendations', [])
    if recommendations:
        row = table.add_row()
        cell = row.cells[0]
        cell.merge(row.cells[1])
        cell.merge(row.cells[2])

        p = _clear_cell(cell)
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)

        run = p.add_run("Recommendation:")
        run.font.name = 'Altone Trial'
        run.font.size = Pt(10)
        run.font.bold = True

        for i, rec in enumerate(recommendations, 1):
            rec_p = cell.add_paragraph()
            rec_p.paragraph_format.space_before = Pt(0)
            rec_p.paragraph_format.space_after = Pt(0)
            run = rec_p.add_run(f"{i}. {safe_str(rec)}")
            run.font.name = 'Altone Trial'
            run.font.size = Pt(10)

    # Reference row
    reference = vulnerability_section.get('reference')
    if reference:
        row = table.add_row()
        cell = row.cells[0]
        cell.merge(row.cells[1])
        cell.merge(row.cells[2])

        p = _clear_cell(cell)
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)

        run = p.add_run("Reference:")
        run.font.name = 'Altone Trial'
        run.font.size = Pt(10)
        run.font.bold = True

        ref_p = cell.add_paragraph()
        run = ref_p.add_run(safe_str(reference))
        run.font.name = 'Altone Trial'
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0, 0, 255)
        run.font.underline = True

    # Evidence/Proof of Concept row
    poc_images = vulnerability_section.get('poc_images', [])
    if poc_images:
        row = table.add_row()
        cell = row.cells[0]
        cell.merge(row.cells[1])
        cell.merge(row.cells[2])

        p = _clear_cell(cell)
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)

        run = p.add_run("Evidence / Proof of Concept:")
        run.font.name = 'Altone Trial'
        run.font.size = Pt(10)
        run.font.bold = True

        for poc_img in poc_images:
            step_p = cell.add_paragraph()
            step_p.paragraph_format.space_before = Pt(6)
            step_p.paragraph_format.space_after = Pt(3)

            step_text = f"Step {poc_img.get('step_number', '')}: "
            if poc_img.get('description'):
                step_text += safe_str(poc_img['description'])

            run = step_p.add_run(step_text)
            run.font.name = 'Altone Trial'
            run.font.size = Pt(10)
            run.font.bold = False

            img_path = poc_img.get('path') or poc_img.get('filename')
            if img_path and os.path.exists(img_path):
                img_p = cell.add_paragraph()
                img_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                img_run = img_p.add_run()
                try:
                    img_run.add_picture(img_path, width=Inches(4.5))
                except Exception as e:
                    err_p = cell.add_paragraph(f"[Image error: {os.path.basename(img_path)}: {e}]")
                    err_p.runs[0].font.italic = True
            else:
                missing_p = cell.add_paragraph(f"[Image missing: {poc_img.get('filename', 'unknown')}]")
                missing_p.runs[0].font.italic = True

    return table


def create_asset_findings_table(doc, vulnerability_data):
    """Create the Asset/Hostname findings table with vulnerability counts"""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls, qn
    from docx.enum.table import WD_TABLE_ALIGNMENT

    if not vulnerability_data:
        return None

    def _clear_cell(cell):
        """Remove extra paragraphs and clear the first paragraph safely."""
        for p in cell.paragraphs[1:]:
            p._element.getparent().remove(p._element)
        cell.paragraphs[0].clear()
        return cell.paragraphs[0]

    def _set_cell_text(cell, text, bold=False, center=False, font_size=10, color=None):
        p = _clear_cell(cell)
        if center:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(str(text))
        run.font.name = 'Altone Trial'
        run.font.size = Pt(font_size)
        run.font.bold = bold
        if color:
            run.font.color.rgb = color

    # Aggregate data by asset
    asset_data = {}
    for vuln in vulnerability_data:
        asset = safe_str(vuln.get('asset', '')) or safe_str(vuln.get('affected_asset', '')) or 'Unknown'
        purpose = safe_str(vuln.get('purpose', '')) or safe_str(vuln.get('vapt_status', ''))
        status = safe_str(vuln.get('vapt_status', '')) or safe_str(vuln.get('status', 'Open'))
        severity = safe_str(vuln.get('severity', 'Medium')).lower()

        if asset not in asset_data:
            asset_data[asset] = {
                'purpose': purpose,
                'status': status,
                'critical': 0,
                'high': 0,
                'medium': 0,
                'low': 0,
                'informational': 0
            }

        if severity == 'critical':
            asset_data[asset]['critical'] += 1
        elif severity == 'high':
            asset_data[asset]['high'] += 1
        elif severity == 'medium':
            asset_data[asset]['medium'] += 1
        elif severity == 'low':
            asset_data[asset]['low'] += 1
        elif severity in ['informational', 'info']:
            asset_data[asset]['informational'] += 1

    # Calculate totals
    overall = {'critical': 0, 'high': 0, 'medium': 0, 'low': 0, 'informational': 0}
    for asset, counts in asset_data.items():
        counts['total'] = counts['critical'] + counts['high'] + counts['medium'] + counts['low'] + counts['informational']
        overall['critical'] += counts['critical']
        overall['high'] += counts['high']
        overall['medium'] += counts['medium']
        overall['low'] += counts['low']
        overall['informational'] += counts['informational']
    overall['total'] = overall['critical'] + overall['high'] + overall['medium'] + overall['low'] + overall['informational']

    # Correct direct table creation
    num_rows = 1 + len(asset_data) + 1  # header + assets + overall
    table = doc.add_table(rows=num_rows, cols=10)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Set total table width and column widths
    tbl_elem = table._tbl
    tblPr_elem = tbl_elem.tblPr if tbl_elem.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
    # Keep findings table visually full-width like other report tables.
    tblW_elem = parse_xml(f'<w:tblW {nsdecls("w")} w:w="5000" w:type="pct"/>')
    existing_tblW = tblPr_elem.find(qn('w:tblW'))
    if existing_tblW is not None:
        tblPr_elem.remove(existing_tblW)
    tblPr_elem.append(tblW_elem)

    col_widths = [650, 2400, 1450, 1200, 750, 750, 850, 750, 1000, 800]
    for col_i, col in enumerate(table.columns):
        for cell in col.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{col_widths[col_i]}" w:type="dxa"/>')
            existing = tcPr.find(qn('w:tcW'))
            if existing is not None:
                tcPr.remove(existing)
            tcPr.append(tcW)

    # Header row
    header_row = table.rows[0]
    headers = ['Sr.\nNo.', 'Asset/Hostname', 'Instant\npurpose', 'VAPT\nStatus', 'Critical', 'High', 'Medium', 'Low', 'Informational', 'Total']

    header_colors = {
        0: '6923D0',
        1: '6923D0',
        2: '6923D0',
        3: '6923D0',
        4: '800000',
        5: 'FF0000',
        6: 'FFC000',
        7: '92D050',
        8: '00B0F0',
        9: '7030A0'
    }

    for idx, header_text in enumerate(headers):
        cell = header_row.cells[idx]
        _set_cell_text(cell, header_text, bold=True, center=True, font_size=9, color=RGBColor(255, 255, 255))
        color = header_colors.get(idx, '6923D0')
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>')
        cell._element.get_or_add_tcPr().append(shading)

    # Data rows
    row_idx = 1
    for sr_no, (asset, counts) in enumerate(asset_data.items(), 1):
        row = table.rows[row_idx]

        _set_cell_text(row.cells[0], f"{sr_no}.", center=True)
        _set_cell_text(row.cells[1], asset, center=True)
        _set_cell_text(row.cells[2], counts['purpose'], center=True)
        _set_cell_text(row.cells[3], counts['status'], center=True)

        severity_vals = [
            (4, counts['critical'], '800000'),
            (5, counts['high'], 'FF0000'),
            (6, counts['medium'], 'FFC000'),
            (7, counts['low'], '92D050'),
            (8, counts['informational'], '00B0F0'),
            (9, counts['total'], '7030A0')
        ]

        for col_idx, val, color in severity_vals:
            cell = row.cells[col_idx]
            _set_cell_text(cell, val, bold=True, center=True, color=RGBColor(255, 255, 255))
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>')
            cell._element.get_or_add_tcPr().append(shading)

        row_idx += 1

    # Overall Findings row
    overall_row = table.rows[row_idx]

    # Merge first 4 cells FIRST
    overall_row.cells[0].merge(overall_row.cells[1])
    overall_row.cells[0].merge(overall_row.cells[2])
    overall_row.cells[0].merge(overall_row.cells[3])

    # Then set text
    _set_cell_text(overall_row.cells[0], 'Overall Findings', bold=True, center=True)

    overall_vals = [
        (4, overall['critical'], '800000'),
        (5, overall['high'], 'FF0000'),
        (6, overall['medium'], 'FFC000'),
        (7, overall['low'], '92D050'),
        (8, overall['informational'], '00B0F0'),
        (9, overall['total'], '7030A0')
    ]

    for col_idx, val, color in overall_vals:
        cell = overall_row.cells[col_idx]
        _set_cell_text(cell, val, bold=True, center=True, color=RGBColor(255, 255, 255))
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>')
        cell._element.get_or_add_tcPr().append(shading)

    return table


def _certin_tbl_clear_left_indent(table) -> None:
    tbl_pr = table._tbl.tblPr
    if tbl_pr is None:
        return
    for old in tbl_pr.findall(qn("w:tblInd")):
        tbl_pr.remove(old)


def _certin_template_table_header_text(table) -> str:
    if not table.rows:
        return ""
    parts = []
    for cell in table.rows[0].cells:
        parts.append(cell.text.replace("\n", " ").replace("\r", " "))
    return " ".join(parts).lower()


def _certin_center_template_front_sections(doc: Document) -> None:
    """
    Cert-IN template tables are often left-indented in Word XML. Center Engagement Scope,
    Auditing team, Audit Activities, and Tools tables like programmatic tables; center section titles.
    """
    from docx.enum.table import WD_TABLE_ALIGNMENT

    for p in doc.paragraphs:
        raw = p.text.strip()
        if not raw:
            continue
        norm = re.sub(r"\s+", " ", raw).strip().lower()
        if norm == "engagement scope":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            continue
        if norm == "details of the auditing team":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            continue
        if norm == "audit activities and timelines":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            continue
        if re.match(r"^tools\s*/\s*software\s*used$", norm):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            continue

    for table in doc.tables:
        h = _certin_template_table_header_text(table)
        if not h or len(table.rows[0].cells) < 2:
            continue
        if "asset description" in h and "criticality" in h:
            _certin_tbl_clear_left_indent(table)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            continue
        if (
            "designation" in h
            and "email" in h
            and "organization" not in h
            and ("qualification" in h or "cert" in h or "listed" in h)
        ):
            _certin_tbl_clear_left_indent(table)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            continue
        if "activity" in h and "date" in h:
            _certin_tbl_clear_left_indent(table)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            continue
        if "tool" in h and "software" in h and ("version" in h or "open source" in h or "licensed" in h):
            _certin_tbl_clear_left_indent(table)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            continue


def _certin_insert_page_break_before_risk_level_section(doc: Document) -> None:
    """Force 'Risk Level & Description' (and following graphics) onto a new page to avoid clipping."""
    from docx.text.paragraph import Paragraph as DocxParagraph
    from docx.oxml import OxmlElement

    target = None
    for p in doc.paragraphs:
        t = " ".join(p.text.split()).strip().lower()
        if t == "risk level & description" or (t.startswith("risk level") and "description" in t):
            target = p
            break
    if target is None:
        return
    new_p = OxmlElement("w:p")
    target._element.addprevious(new_p)
    pb_para = DocxParagraph(new_p, doc)
    pb_para.paragraph_format.space_before = Pt(0)
    pb_para.paragraph_format.space_after = Pt(0)
    run = pb_para.add_run()
    run.add_break(WD_BREAK.PAGE)


def _drop_empty_items(items):
    """Remove fully empty dict/list records to avoid blank table rows in template loops."""
    cleaned = []
    for item in items or []:
        if isinstance(item, dict):
            has_value = any(str(v or "").strip() for v in item.values())
            if has_value:
                cleaned.append(item)
            continue
        if isinstance(item, (list, tuple)):
            has_value = any(str(v or "").strip() for v in item)
            if has_value:
                cleaned.append(item)
            continue
        if str(item or "").strip():
            cleaned.append(item)
    return cleaned


def _certin_remove_empty_rows_and_normalize_borders(doc: Document) -> None:
    """
    Fix front-matter tables that often render with blank spacer rows and broken lines.
    Keeps content intact while removing empty rows and forcing connected grid borders.
    """
    from docx.oxml import OxmlElement

    def _set_border(node, edge, val="single", sz="8", color="000000"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), val)
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        node.append(el)

    def _normalize_table_grid(table):
        tbl_pr = table._tbl.tblPr
        if tbl_pr is None:
            return

        # Replace table-level borders to keep all lines connected.
        for old in tbl_pr.findall(qn("w:tblBorders")):
            tbl_pr.remove(old)
        tbl_borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            _set_border(tbl_borders, edge)
        tbl_pr.append(tbl_borders)

    # Remove fully-empty rows from all non-vulnerability tables.
    for table in doc.tables:
        header_text = _certin_template_table_header_text(table)
        if "asset/hostname" in header_text:
            # Programmatic findings table is handled separately.
            continue

        # Remove blank rows (except first/header row).
        for ridx in range(len(table.rows) - 1, 0, -1):
            row = table.rows[ridx]
            row_text = "".join((c.text or "").strip() for c in row.cells)
            if not row_text:
                table._tbl.remove(row._tr)

        _normalize_table_grid(table)


def generate_certin_report_from_form(data, template_path, output_path, vulnerability_data=None, poc_mapping=None):
    """Generate Cert-IN report from form data and template"""
    try:
        safe_template_path = _sanitize_docx_jinja_placeholders(template_path)
        doc = DocxTemplate(safe_template_path)
        
        document_change_history = _drop_empty_items(json.loads(data.get('document_change_history', '[]')))
        distribution_list = _drop_empty_items(json.loads(data.get('distribution_list', '[]')))
        engagement_scope = _drop_empty_items(json.loads(data.get('engagement_scope', '[]')))
        auditing_team = _drop_empty_items(json.loads(data.get('auditing_team', '[]')))
        audit_activities = _drop_empty_items(json.loads(data.get('audit_activities', '[]')))
        tools_software = _drop_empty_items(json.loads(data.get('tools_software', '[]')))
        
        vulnerability_sections = []
        client_name_from_excel = ''
        project_name_from_excel = ''
        if vulnerability_data and poc_mapping is not None:
            vulnerability_sections = generate_vulnerability_sections(vulnerability_data, poc_mapping)
            # Extract client/project from first vulnerability row if not provided in form
            if vulnerability_data:
                client_name_from_excel = vulnerability_data[0].get('client', '')
                project_name_from_excel = vulnerability_data[0].get('project', '')
                print(f"📊 Extracted from Excel - Client: {client_name_from_excel}, Project: {project_name_from_excel}")
        
        # Use form data as primary, fall back to Excel data
        final_client_name = data.get('client_name', '') or client_name_from_excel
        final_company_name = data.get('company_name', '') or final_client_name
        final_report_name = data.get('report_name', '') or project_name_from_excel
        
        print(f"📝 Form client_name: '{data.get('client_name', '')}', Final: '{final_client_name}'")
        print(f"📝 Form report_name: '{data.get('report_name', '')}', Final: '{final_report_name}'")
        
        context = {
            'CLIENT_NAME': final_client_name,
            'COMPANY_NAME': final_company_name,
            'REPORT_NAME': final_report_name,
            # Also add variations in case template uses different format
            'TITLE': data.get('document_title', ''),
            'REPORT_TITLE': data.get('document_title', ''),
            'CLIENT__NAME': final_client_name,
            'COMPANY__NAME': final_company_name,
            'REPORT__NAME': final_report_name,
            'CLIENT_NAME_': final_client_name,
            'COMPANY_NAME_': final_company_name,
            'REPORT_NAME_': final_report_name,
            '_CLIENT_NAME': final_client_name,
            '_COMPANY_NAME': final_company_name,
            '_REPORT_NAME': final_report_name,
            'REPORT_RELEASE_DATE': data.get('report_release_date', ''),
            'TYPE_OF_AUDIT': data.get('type_of_audit', ''),
            'TYPE_OF_AUDIT_REPORT': data.get('type_of_audit_report', ''),
            'PERIOD': data.get('period', ''),
            'DOCUMENT_TITLE': data.get('document_title', ''),
            'DOCUMENT_ID': data.get('document_id', ''),
            'DOCUMENT_VERSION': data.get('document_version', ''),
            'PREPARED_BY': data.get('prepared_by', ''),
            'REVIEWED_BY': data.get('reviewed_by', ''),
            'APPROVED_BY': data.get('approved_by', ''),
            'RELEASED_BY': data.get('released_by', ''),
            'RELEASE_DATE': data.get('release_date', ''),
            'DOCUMENT_CHANGE_HISTORY': document_change_history,
            'DISTRIBUTION_LIST': distribution_list,
            'ENGAGEMENT_SCOPE': engagement_scope,
            'AUDITING_TEAM': auditing_team,
            'AUDIT_ACTIVITIES': audit_activities,
            'TOOLS_SOFTWARE': tools_software,
            'VULNERABILITIES': vulnerability_sections,
            'HAS_VULNERABILITIES': len(vulnerability_sections) > 0,
            'VULNERABILITY_COUNT': len(vulnerability_sections),
            'GENERATION_DATE': datetime.now().strftime('%d %B %Y'),
            'GENERATION_TIME': datetime.now().strftime('%H:%M:%S'),
        }
        
        doc.render(context)
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            doc.save(tmp.name)
            tmp_path = tmp.name
        
        doc = Document(tmp_path)
        _certin_center_template_front_sections(doc)
        _certin_insert_page_break_before_risk_level_section(doc)
        _certin_remove_empty_rows_and_normalize_borders(doc)
        
        # In generate_certin_report_from_form(), replace lines 985-1005 with:

        if vulnerability_sections:
            table = create_asset_findings_table(doc, vulnerability_data)

            if table is not None:
                # Keep this table immediately below the "Total Vulnerabilities" graph block
                # and before "Detailed Observations".
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                table._element.getparent().remove(table._element)
                inserted = False
                for para in doc.paragraphs:
                    txt = " ".join((para.text or "").split()).strip().lower()
                    if txt == "total vulnerabilities" or "total vulnerabilities" in txt:
                        spacer_before = doc.add_paragraph()
                        spacer_before.paragraph_format.space_before = Pt(6)
                        spacer_before.paragraph_format.space_after = Pt(6)
                        para._element.addnext(spacer_before._element)
                        spacer_before._element.addnext(table._element)
                        spacer_after = doc.add_paragraph()
                        spacer_after.paragraph_format.space_before = Pt(6)
                        spacer_after.paragraph_format.space_after = Pt(6)
                        table._element.addnext(spacer_after._element)
                        inserted = True
                        break
                if not inserted:
                    # Fallback: append right before Detailed Observations section creation.
                    doc.element.body.append(table._element)
                    spacer_after = doc.add_paragraph()
                    spacer_after.paragraph_format.space_before = Pt(6)
                    spacer_after.paragraph_format.space_after = Pt(6)

            # Detailed Observations heading on new page
            doc.add_page_break()
            vuln_heading = doc.add_paragraph()
            vuln_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
            vuln_run = vuln_heading.add_run("Detailed Observations")
            vuln_run.font.name = 'Altone Trial'
            vuln_run.font.size = Pt(18)
            vuln_run.font.bold = True
            vuln_run.font.color.rgb = RGBColor(106, 68, 154)

            for i, vuln_section in enumerate(vulnerability_sections):
                create_landscape_vulnerability_box(doc, vuln_section)

                if i < len(vulnerability_sections) - 1:
                    doc.add_page_break()

        doc.save(output_path)
        os.remove(tmp_path)
        
        return True
        
    except Exception as e:
        print(f"Error generating report: {str(e)}")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Error generating report: {str(e)}")

def _sanitize_docx_jinja_placeholders(template_path: str) -> str:
    try:
        with open(template_path, 'rb') as f:
            original_bytes = f.read()
        zin = zipfile.ZipFile(io.BytesIO(original_bytes))
        out_buf = io.BytesIO()
        zout = zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED)
        
        pattern_expr = re.compile(r"(\{\{\s*)([^}]+?)(\s*\}\})")
        
        def _normalize_print_statement(inner: str) -> str:
            inner = inner.strip()
            if inner == "Audit.sr_no":
                inner = "team.sr_no"
            
            chars = []
            for ch in inner:
                if re.match(r"[A-Za-z0-9_.]", ch):
                    chars.append(ch)
                elif ch.isspace():
                    chars.append('_')
                else:
                    chars.append('_')
            return ''.join(chars)
        
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith('word/') and item.filename.endswith('.xml'):
                xml = data.decode('utf-8', errors='ignore')
                xml = pattern_expr.sub(lambda m: m.group(1) + _normalize_print_statement(m.group(2)) + m.group(3), xml)
                data = xml.encode('utf-8')
            zout.writestr(item, data)
        zin.close()
        zout.close()
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmpf:
            tmpf.write(out_buf.getvalue())
            return tmpf.name
    except Exception:
        return template_path

@app.get("/", response_class=HTMLResponse)
async def root():
    """Serve the Cert-IN form interface"""
    html_content = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Cert-IN Report Generator</title>
        <style>
            body { font-family: Calibri, sans-serif; margin: 20px; background-color: #f5f5f5; }
            .container { max-width: 1200px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
            h1 { color: #6923d0; text-align: center; margin-bottom: 30px; }
            .form-section { margin-bottom: 30px; padding: 20px; border: 1px solid #ddd; border-radius: 5px; }
            .form-section h3 { color: #6923d0; margin-top: 0; border-bottom: 2px solid #6923d0; padding-bottom: 10px; }
            .form-row { display: flex; gap: 20px; margin-bottom: 15px; }
            .form-group { flex: 1; }
            .form-group label { display: block; margin-bottom: 5px; font-weight: bold; color: #333; }
            .form-group input, .form-group select, .form-group textarea { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; }
            .form-group textarea { height: 60px; resize: vertical; }
            .dynamic-list { border: 1px solid #ddd; padding: 15px; margin: 10px 0; border-radius: 4px; }
            .dynamic-item { display: flex; gap: 10px; margin-bottom: 10px; align-items: end; flex-wrap: wrap; }
            .dynamic-item input, .dynamic-item select { flex: 1; min-width: 120px; box-sizing: border-box; }
            /* Engagement scope: many fields per row + many rows — match MitKat layout, stay on screen */
            #engagementScopeList {
                max-height: min(480px, 60vh);
                overflow-y: auto;
                overflow-x: hidden;
                padding: 10px;
                border: 1px solid #ddd;
                border-radius: 6px;
                background: #fafafa;
            }
            #engagementScopeList .dynamic-item {
                align-items: center;
                margin-bottom: 12px;
                padding: 10px;
                border: 1px solid #e8e8e8;
                border-radius: 4px;
                background: #fff;
            }
            #engagementScopeList .dynamic-item input,
            #engagementScopeList .dynamic-item select {
                flex: 1 1 140px;
                min-width: 130px;
                max-width: 100%;
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
            }
            #engagementScopeList .dynamic-item .btn { flex: 0 0 auto; }
            .btn { background-color: #6923d0; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; margin: 5px; }
            .btn:hover { background-color: #5a1fb8; }
            .btn-secondary { background-color: #6c757d; }
            .btn-secondary:hover { background-color: #5a6268; }
            .btn-danger { background-color: #dc3545; }
            .btn-danger:hover { background-color: #c82333; }
            .btn-success { background-color: #28a745; }
            .btn-success:hover { background-color: #218838; }
            .progress { display: none; margin: 20px 0; }
            .progress-bar { background-color: #6923d0; height: 20px; border-radius: 4px; width: 0%; transition: width 0.3s; }
            .result { display: none; margin: 20px 0; padding: 15px; border-radius: 4px; }
            .result.success { background-color: #d4edda; border: 1px solid #c3e6cb; color: #155724; }
            .result.error { background-color: #f8d7da; border: 1px solid #f5c6cb; color: #721c24; }
            .info-box { background-color: #e7f3ff; border-left: 4px solid #2196F3; padding: 15px; margin: 15px 0; }
            .info-box h4 { margin-top: 0; color: #2196F3; }
            .info-box code { background-color: #fff; padding: 2px 6px; border-radius: 3px; font-family: monospace; }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>🔒 Cert-IN Report Generator v2.2</h1>
            <div style="text-align:center;margin-bottom:16px;">
                <a href="/report_formats.html" class="btn btn-secondary" style="text-decoration:none;display:inline-block;">← Back to Report Formats</a>
            </div>
            <p style="text-align: center; color: #666; margin-bottom: 30px;">
                Generate Cert-IN compliant reports with document control pages
            </p>

            <!-- STEP 1: Convert any Excel to Cert-IN Format (OPTIONAL) -->
            <div class="form-section" style="border: 2px solid #28a745; background: #f0fff4;">
                <h3 style="color:#28a745; border-bottom-color:#28a745;">
                    <span style="background:#28a745;color:white;padding:2px 10px;border-radius:10px;font-size:0.8em;margin-right:8px;text-transform:uppercase;letter-spacing:1px;">Step 1 — Optional</span>
                    Convert Any Excel to Cert-IN Format Excel
                    <span style="font-size:0.7em;color:#555;font-weight:normal;"> (Skip if your Excel is already in Cert-IN format)</span>
                </h3>
                <p style="color:#555;margin-top:0;">Only needed if your Excel is not already in the required Cert-IN format. Upload any VAPT tracker Excel and optional POC ZIP to get a properly formatted Cert-IN Excel.</p>
                <div class="form-row">
                    <div class="form-group">
                        <label for="conv_vapt_file">Upload Any VAPT Excel (.xlsx) *</label>
                        <input type="file" id="conv_vapt_file" accept=".xlsx,.xls">
                    </div>
                    <div class="form-group">
                        <label for="conv_poc_zip">Upload POC ZIP (optional) — for screenshot mapping</label>
                        <input type="file" id="conv_poc_zip" accept=".zip">
                    </div>
                </div>
                <button type="button" class="btn btn-success" id="convertBtn">Convert to Cert-IN Format Excel</button>
                <div class="result" id="convertResult" style="display:none;"></div>
            </div>

            <hr style="border: 2px dashed #6923d0; margin: 10px 0 30px 0;">

            <!-- STEP 2: Generate Report -->
            <div class="form-section" style="border: 2px solid #6923d0;">
                <h3 style="color:#6923d0; border-bottom-color:#6923d0;">
                    <span style="background:#6923d0;color:white;padding:2px 10px;border-radius:10px;font-size:0.8em;margin-right:8px;text-transform:uppercase;letter-spacing:1px;">Step 2</span>
                    Generate Cert-IN Word Report
                </h3>
                <p style="color:#555;margin-top:0;">Fill in the report details and upload the Cert-IN Format Excel (output from Step 1, or any correctly formatted Cert-IN Excel) to generate the final Word report.</p>
            </div>

            <form id="certinForm">
                <!-- Basic Report Information -->
                <div class="form-section">
                    <h3>📋 Basic Report Information</h3>
                    <div class="form-row">
                        <div class="form-group">
                            <label for="client_name">Client Name *</label>
                            <input type="text" id="client_name" name="client_name" required>
                        </div>
                        <div class="form-group">
                            <label for="company_name">Company Name *</label>
                            <input type="text" id="company_name" name="company_name" required>
                        </div>
                        <div class="form-group">
                            <label for="report_name">Web/Network Initial Report Name *</label>
                            <select id="report_name" name="report_name" required>
                                <option value="">Select Initial Report Name</option>
                                <option value="Web Application">Web Application</option>
                                <option value="Network">Network</option>
                                <option value="Web and Network">Web and Network</option>
                            </select>
                        </div>
                    </div>
                    <div class="form-row">
                        <div class="form-group">
                            <label for="report_release_date">Report Release Date *</label>
                            <input type="date" id="report_release_date" name="report_release_date" required>
                        </div>
                        <div class="form-group">
                            <label for="type_of_audit">Type of Audit *</label>
                            <select id="type_of_audit" name="type_of_audit" required>
                                <option value="">Select Type</option>
                                <option value="Security assessment (VAPT) - WEB & Network">Security assessment (VAPT) - WEB & Network</option>
                                <option value="Security assessment (VAPT) - WEB">Security assessment (VAPT) - WEB</option>
                                <option value="Security assessment (VAPT) - Network">Security assessment (VAPT) - Network</option>
                                <option value="Penetration Testing">Penetration Testing</option>
                                <option value="Vulnerability Assessment">Vulnerability Assessment</option>
                            </select>
                        </div>
                    </div>
                    <div class="form-row">
                        <div class="form-group">
                            <label for="type_of_audit_report">Type of Audit Report *</label>
                            <select id="type_of_audit_report" name="type_of_audit_report" required>
                                <option value="">Select Type</option>
                                <option value="First Audit Report">First Audit Report</option>
                                <option value="Follow-up Audit Report">Follow-up Audit Report</option>
                                <option value="Final Audit Report">Final Audit Report</option>
                            </select>
                        </div>
                        <div class="form-group">
                            <label for="period">Period *</label>
                            <input type="text" id="period" name="period" placeholder="e.g., 23-07-2025 to 02-08-2025" required>
                        </div>
                    </div>
                </div>

                <!-- Document Control -->
                <div class="form-section">
                    <h3>📄 Document Control</h3>
                    <div class="form-row">
                        <div class="form-group">
                            <label for="document_title">Document Title *</label>
                            <input type="text" id="document_title" name="document_title" required>
                        </div>
                        <div class="form-group">
                            <label for="document_id">Document ID *</label>
                            <input type="text" id="document_id" name="document_id" required>
                        </div>
                    </div>
                    <div class="form-row">
                        <div class="form-group">
                            <label for="document_version">Document Version *</label>
                            <input type="text" id="document_version" name="document_version" value="1.0" required>
                        </div>
                        <div class="form-group">
                            <label for="prepared_by">Prepared by *</label>
                            <input type="text" id="prepared_by" name="prepared_by" required>
                        </div>
                    </div>
                    <div class="form-row">
                        <div class="form-group">
                            <label for="reviewed_by">Reviewed by *</label>
                            <input type="text" id="reviewed_by" name="reviewed_by" required>
                        </div>
                        <div class="form-group">
                            <label for="approved_by">Approved by *</label>
                            <input type="text" id="approved_by" name="approved_by" required>
                        </div>
                    </div>
                    <div class="form-row">
                        <div class="form-group">
                            <label for="released_by">Released by *</label>
                            <input type="text" id="released_by" name="released_by" required>
                        </div>
                        <div class="form-group">
                            <label for="release_date">Release Date *</label>
                            <input type="date" id="release_date" name="release_date" required>
                        </div>
                    </div>
                </div>

                <!-- Document Change History -->
                <div class="form-section">
                    <h3>📝 Document Change History</h3>
                    <div id="changeHistoryList">
                        <div class="dynamic-item">
                            <input type="text" name="change_version[]" placeholder="Version" required>
                            <input type="date" name="change_date[]" required>
                            <input type="text" name="change_remarks[]" placeholder="Remarks/Reason of change" required>
                            <button type="button" class="btn btn-danger" onclick="removeChangeHistory(this)">Remove</button>
                        </div>
                    </div>
                    <button type="button" class="btn btn-secondary" onclick="addChangeHistory()">+ Add Change History</button>
                </div>

                <!-- Document Distribution List -->
                <div class="form-section">
                    <h3>📧 Document Distribution List</h3>
                    <div id="distributionList">
                        <div class="dynamic-item">
                            <input type="text" name="dist_name[]" placeholder="Name" required>
                            <input type="text" name="dist_organization[]" placeholder="Organization" required>
                            <input type="text" name="dist_designation[]" placeholder="Designation" required>
                            <input type="email" name="dist_email[]" placeholder="Email ID" required>
                            <button type="button" class="btn btn-danger" onclick="removeDistribution(this)">Remove</button>
                        </div>
                    </div>
                    <button type="button" class="btn btn-secondary" onclick="addDistribution()">+ Add Distribution</button>
                </div>

                <!-- Engagement Scope -->
                <div class="form-section">
                    <h3>🎯 Engagement Scope</h3>
                    <p style="color:#666;font-size:0.9em;margin-top:0;">Fields wrap on smaller screens; scroll the list when you have many assets.</p>
                    <div id="engagementScopeList">
                        <div class="dynamic-item">
                            <input type="number" name="scope_sr_no[]" placeholder="S. No">
                            <input type="text" name="scope_asset[]" placeholder="Asset Description">
                            <input type="text" name="scope_criticality[]" placeholder="Criticality">
                            <input type="text" name="scope_internal_ip[]" placeholder="Internal IP">
                            <input type="text" name="scope_url[]" placeholder="URL or NA">
                            <input type="text" name="scope_public_ip[]" placeholder="Public IP">
                            <input type="text" name="scope_location[]" placeholder="Location">
                            <input type="text" name="scope_hash[]" placeholder="Hash Value">
                            <input type="text" name="scope_version[]" placeholder="Version">
                            <input type="text" name="scope_other[]" placeholder="Other Details">
                            <button type="button" class="btn btn-danger" onclick="removeEngagementScope(this)">Remove</button>
                        </div>
                    </div>
                    <button type="button" class="btn btn-secondary" onclick="addEngagementScope()">+ Add Asset</button>
                </div>

                <!-- Details of Auditing Team -->
                <div class="form-section">
                    <h3>👥 Details of Auditing Team</h3>
                    <div id="auditingTeamList">
                        <div class="dynamic-item">
                            <input type="text" name="team_name[]" placeholder="Name" required>
                            <input type="text" name="team_designation[]" placeholder="Designation" required>
                            <input type="email" name="team_email[]" placeholder="Email ID" required>
                            <input type="text" name="team_qualifications[]" placeholder="Qualifications/Certifications" required>
                            <select name="team_certin_listed[]" required>
                                <option value="">CERT-In Listed?</option>
                                <option value="Yes">Yes</option>
                                <option value="No">No</option>
                            </select>
                            <button type="button" class="btn btn-danger" onclick="removeAuditingTeam(this)">Remove</button>
                        </div>
                    </div>
                    <button type="button" class="btn btn-secondary" onclick="addAuditingTeam()">+ Add Team Member</button>
                </div>

                <!-- Audit Activities and Timelines -->
                <div class="form-section">
                    <h3>📅 Audit Activities and Timelines</h3>
                    <div id="auditActivitiesList">
                        <div class="dynamic-item">
                            <input type="text" name="activity_task[]" placeholder="Activity/Task" required>
                            <input type="date" name="activity_date[]" required>
                            <button type="button" class="btn btn-danger" onclick="removeAuditActivity(this)">Remove</button>
                        </div>
                    </div>
                    <button type="button" class="btn btn-secondary" onclick="addAuditActivity()">+ Add Activity</button>
                </div>

                <!-- Tools/Software Used -->
                <div class="form-section">
                    <h3>🛠️ Tools/Software Used</h3>
                    <div id="toolsSoftwareList">
                        <div class="dynamic-item">
                            <input type="number" name="tool_sr_no[]" placeholder="S. No" required>
                            <input type="text" name="tool_name[]" placeholder="Name of Tool/Software" required>
                            <input type="text" name="tool_version[]" placeholder="Version" required>
                            <select name="tool_type[]" required>
                                <option value="">Type</option>
                                <option value="Opensource">Opensource</option>
                                <option value="Licensed">Licensed</option>
                            </select>
                            <button type="button" class="btn btn-danger" onclick="removeToolSoftware(this)">Remove</button>
                        </div>
                    </div>
                    <button type="button" class="btn btn-secondary" onclick="addToolSoftware()">+ Add Tool</button>
                </div>

                <!-- Vulnerability Data Upload -->
                <div class="form-section">
                    <h3>🔍 Vulnerability Data (Required)</h3>
                    
                    <div class="info-box">
                        <h4>📋 Excel Format Requirements</h4>
                        <p><strong>Required columns (any case/spacing):</strong></p>
                        <ul>
                            <li><code>New or Repeat Observation</code></li>
                            <li><code>CVE/CWE</code></li>
                            <li><code>CVSS Version Ref</code></li>
                            <li><code>Affected Asset i.e. IP/URL/Application etc.</code></li>
                            <li><code>Observation/ Vulnerability Title</code></li>
                            <li><code>Detailed Observation/ Vulnerable Point</code></li>
                            <li><code>Recommendation</code></li>
                            <li><code>Reference</code></li>
                            <li><code>Evidence / Proof of Concept</code> - Write step descriptions here</li>
                        </ul>
                        <p><strong>Note:</strong> Column names are case-insensitive and flexible with spaces</p>
                    </div>
                    
                    <div class="form-group">
                        <label for="vulnerability_file">Upload Vulnerability Excel File *</label>
                        <input type="file" id="vulnerability_file" name="vulnerability_file" accept=".xlsx,.xls" required>
                    </div>
                    
                    <div class="info-box">
                        <h4>📁 PoC Zip Structure Requirements</h4>
                        <p><strong>Folder Structure:</strong></p>
                        <code>POC_certIN.zip/POC_certIN/#1, #2, #3</code>
                        <p><strong>Example:</strong></p>
                        <ul>
                            <li><code>POC_certIN/#1/step1.png</code> → Observation #1, Step 1</li>
                            <li><code>POC_certIN/#1/step2.png</code> → Observation #1, Step 2</li>
                            <li><code>POC_certIN/#2/step1.png</code> → Observation #2, Step 1</li>
                        </ul>
                        <p><strong>Important:</strong> Each <code>#N</code> folder contains ONLY images for that observation</p>
                    </div>
                    
                    <div class="form-group">
                        <label for="poc_files">Upload PoC Zip File *</label>
                        <input type="file" id="poc_files" name="poc_files" accept=".zip" required>
                    </div>
                </div>

                <!-- Submit Button -->
                <div style="text-align: center; margin-top: 30px;">
                    <button type="submit" class="btn btn-success" style="font-size: 18px; padding: 15px 30px;">
                        🚀 Generate Cert-IN Report
                    </button>
                </div>
            </form>

            <!-- Progress Bar -->
            <div class="progress" id="progress">
                <div style="background-color: #f0f0f0; border-radius: 4px; padding: 10px;">
                    <div id="progressText">Processing...</div>
                    <div class="progress-bar" id="progressBar"></div>
                </div>
            </div>

            <!-- Result -->
            <div class="result" id="result"></div>
        </div>

        <script>
            let CSRF_TOKEN = null;
            (async () => {
                try {
                    const res = await fetch('/csrf-token', { credentials: 'same-origin' });
                    if (res.ok) {
                        const data = await res.json().catch(() => ({}));
                        CSRF_TOKEN = data.csrf_token || data.token || data.csrf || null;
                    }
                } catch (_) {}
            })();

            async function refreshCsrf() {
                try {
                    const res = await fetch('/csrf-token', { credentials: 'same-origin' });
                    if (res.ok) {
                        const data = await res.json().catch(() => ({}));
                        CSRF_TOKEN = data.csrf_token || data.token || data.csrf || CSRF_TOKEN;
                    }
                } catch (_) {}
            }

            // ── Step 1: Convert Excel ──────────────────────────────────────
            document.addEventListener('DOMContentLoaded', function() {
                const convertBtn = document.getElementById('convertBtn');
                if (convertBtn) {
                    convertBtn.addEventListener('click', async function () {
                        const resultEl = document.getElementById('convertResult');
                        resultEl.style.display = 'none';
                        const vaptFile = document.getElementById('conv_vapt_file').files[0];
                        if (!vaptFile) {
                            resultEl.innerHTML = '<p>Please choose a VAPT Excel file first.</p>';
                            resultEl.className = 'result error';
                            resultEl.style.display = 'block';
                            return;
                        }
                        this.disabled = true;
                        this.textContent = 'Converting...';
                        await refreshCsrf();
                        const fd = new FormData();
                        fd.append('vapt_file', vaptFile);
                        const pocFile = document.getElementById('conv_poc_zip').files[0];
                        if (pocFile) fd.append('poc_zip', pocFile);
                        if (CSRF_TOKEN) fd.append('csrf_token', CSRF_TOKEN);
                        try {
                            const headers = {};
                            if (CSRF_TOKEN) headers['X-CSRF-Token'] = CSRF_TOKEN;
                            const response = await fetch('/type3/convert-excel/', {
                                method: 'POST',
                                headers: headers,
                                body: fd,
                                credentials: 'same-origin'
                            });
                            if (response.status === 401) {
                                throw new Error('Session expired. <a href="/auth/test-login" style="color:#28a745;">Click here to log in again</a>.');
                            }
                            if (!response.ok) {
                                const err = await response.json().catch(() => ({}));
                                throw new Error(err.detail || 'Conversion failed (' + response.status + ')');
                            }
                            const blob = await response.blob();
                            const url = window.URL.createObjectURL(blob);
                            const a = document.createElement('a');
                            a.href = url;
                            const cd = response.headers.get('content-disposition') || '';
                            a.download = cd.match(/filename="?([^";\\n]+)"?/)?.[1] || 'CertIN_format_output.xlsx';
                            document.body.appendChild(a);
                            a.click();
                            window.URL.revokeObjectURL(url);
                            document.body.removeChild(a);
                            resultEl.innerHTML = '<h4>Conversion successful!</h4><p>Cert-IN Format Excel downloaded. Use it in Step 2 below.</p>';
                            resultEl.className = 'result success';
                            resultEl.style.display = 'block';
                        } catch (error) {
                            resultEl.innerHTML = '<h4>Conversion failed</h4><p>' + error.message + '</p>';
                            resultEl.className = 'result error';
                            resultEl.style.display = 'block';
                        } finally {
                            this.disabled = false;
                            this.textContent = 'Convert to Cert-IN Format Excel';
                        }
                    });
                }
            });

            function addChangeHistory() {
                const container = document.getElementById('changeHistoryList');
                const div = document.createElement('div');
                div.className = 'dynamic-item';
                div.innerHTML = `
                    <input type="text" name="change_version[]" placeholder="Version" required>
                    <input type="date" name="change_date[]" required>
                    <input type="text" name="change_remarks[]" placeholder="Remarks/Reason of change" required>
                    <button type="button" class="btn btn-danger" onclick="removeChangeHistory(this)">Remove</button>
                `;
                container.appendChild(div);
            }

            function removeChangeHistory(button) {
                button.parentElement.remove();
            }

            function addDistribution() {
                const container = document.getElementById('distributionList');
                const div = document.createElement('div');
                div.className = 'dynamic-item';
                div.innerHTML = `
                    <input type="text" name="dist_name[]" placeholder="Name" required>
                    <input type="text" name="dist_organization[]" placeholder="Organization" required>
                    <input type="text" name="dist_designation[]" placeholder="Designation" required>
                    <input type="email" name="dist_email[]" placeholder="Email ID" required>
                    <button type="button" class="btn btn-danger" onclick="removeDistribution(this)">Remove</button>
                `;
                container.appendChild(div);
            }

            function removeDistribution(button) {
                button.parentElement.remove();
            }

            function addEngagementScope() {
                const container = document.getElementById('engagementScopeList');
                const div = document.createElement('div');
                div.className = 'dynamic-item';
                div.innerHTML = `
                    <input type="number" name="scope_sr_no[]" placeholder="S. No" required>
                    <input type="text" name="scope_asset[]" placeholder="Asset Description" required>
                    <input type="text" name="scope_criticality[]" placeholder="Criticality" required>
                    <input type="text" name="scope_internal_ip[]" placeholder="Internal IP">
                    <input type="url" name="scope_url[]" placeholder="URL">
                    <input type="text" name="scope_public_ip[]" placeholder="Public IP">
                    <input type="text" name="scope_location[]" placeholder="Location">
                    <input type="text" name="scope_hash[]" placeholder="Hash Value">
                    <input type="text" name="scope_version[]" placeholder="Version">
                    <input type="text" name="scope_other[]" placeholder="Other Details">
                    <button type="button" class="btn btn-danger" onclick="removeEngagementScope(this)">Remove</button>
                `;
                container.appendChild(div);
            }

            function removeEngagementScope(button) {
                button.parentElement.remove();
            }

            function addAuditingTeam() {
                const container = document.getElementById('auditingTeamList');
                const div = document.createElement('div');
                div.className = 'dynamic-item';
                div.innerHTML = `
                    <input type="text" name="team_name[]" placeholder="Name" required>
                    <input type="text" name="team_designation[]" placeholder="Designation" required>
                    <input type="email" name="team_email[]" placeholder="Email ID" required>
                    <input type="text" name="team_qualifications[]" placeholder="Qualifications/Certifications" required>
                    <select name="team_certin_listed[]" required>
                        <option value="">CERT-In Listed?</option>
                        <option value="Yes">Yes</option>
                        <option value="No">No</option>
                    </select>
                    <button type="button" class="btn btn-danger" onclick="removeAuditingTeam(this)">Remove</button>
                `;
                container.appendChild(div);
            }

            function removeAuditingTeam(button) {
                button.parentElement.remove();
            }

            function addAuditActivity() {
                const container = document.getElementById('auditActivitiesList');
                const div = document.createElement('div');
                div.className = 'dynamic-item';
                div.innerHTML = `
                    <input type="text" name="activity_task[]" placeholder="Activity/Task" required>
                    <input type="date" name="activity_date[]" required>
                    <button type="button" class="btn btn-danger" onclick="removeAuditActivity(this)">Remove</button>
                `;
                container.appendChild(div);
            }

            function removeAuditActivity(button) {
                button.parentElement.remove();
            }

            function addToolSoftware() {
                const container = document.getElementById('toolsSoftwareList');
                const div = document.createElement('div');
                div.className = 'dynamic-item';
                div.innerHTML = `
                    <input type="number" name="tool_sr_no[]" placeholder="S. No" required>
                    <input type="text" name="tool_name[]" placeholder="Name of Tool/Software" required>
                    <input type="text" name="tool_version[]" placeholder="Version" required>
                    <select name="tool_type[]" required>
                        <option value="">Type</option>
                        <option value="Opensource">Opensource</option>
                        <option value="Licensed">Licensed</option>
                    </select>
                    <button type="button" class="btn btn-danger" onclick="removeToolSoftware(this)">Remove</button>
                `;
                container.appendChild(div);
            }

            function removeToolSoftware(button) {
                button.parentElement.remove();
            }

            document.getElementById('certinForm').addEventListener('submit', async function(e) {
                e.preventDefault();
                
                const formData = new FormData();
                const form = document.getElementById('certinForm');
                
                for (let element of form.elements) {
                    if (element.name && element.type !== 'file') {
                        if (element.type === 'checkbox') {
                            if (element.checked) {
                                formData.append(element.name, element.value);
                            }
                        } else {
                            formData.append(element.name, element.value);
                        }
                    }
                }
                
                const vulnerabilityFile = document.getElementById('vulnerability_file').files[0];
                if (vulnerabilityFile) {
                    formData.append('vulnerability_file', vulnerabilityFile);
                }
                
                const pocFile = document.getElementById('poc_files').files[0];
                if (pocFile) {
                    formData.append('poc_files', pocFile);
                }
                
                const changeVersions = document.querySelectorAll('input[name="change_version[]"]');
                const changeDates = document.querySelectorAll('input[name="change_date[]"]');
                const changeRemarks = document.querySelectorAll('input[name="change_remarks[]"]');
                
                const changeHistory = [];
                for (let i = 0; i < changeVersions.length; i++) {
                    changeHistory.push({
                        version: changeVersions[i].value,
                        date: changeDates[i].value,
                        remarks: changeRemarks[i].value
                    });
                }
                formData.append('document_change_history', JSON.stringify(changeHistory));
                
                const distNames = document.querySelectorAll('input[name="dist_name[]"]');
                const distOrgs = document.querySelectorAll('input[name="dist_organization[]"]');
                const distDesignations = document.querySelectorAll('input[name="dist_designation[]"]');
                const distEmails = document.querySelectorAll('input[name="dist_email[]"]');
                
                const distributionList = [];
                for (let i = 0; i < distNames.length; i++) {
                    distributionList.push({
                        name: distNames[i].value,
                        organization: distOrgs[i].value,
                        designation: distDesignations[i].value,
                        email: distEmails[i].value
                    });
                }
                formData.append('distribution_list', JSON.stringify(distributionList));
                
                const scopeSrNos = document.querySelectorAll('input[name="scope_sr_no[]"]');
                const scopeAssets = document.querySelectorAll('input[name="scope_asset[]"]');
                const scopeCriticalities = document.querySelectorAll('input[name="scope_criticality[]"]');
                const scopeInternalIPs = document.querySelectorAll('input[name="scope_internal_ip[]"]');
                const scopeURLs = document.querySelectorAll('input[name="scope_url[]"]');
                const scopePublicIPs = document.querySelectorAll('input[name="scope_public_ip[]"]');
                const scopeLocations = document.querySelectorAll('input[name="scope_location[]"]');
                const scopeHashes = document.querySelectorAll('input[name="scope_hash[]"]');
                const scopeVersions = document.querySelectorAll('input[name="scope_version[]"]');
                const scopeOthers = document.querySelectorAll('input[name="scope_other[]"]');
                
                const engagementScope = [];
                for (let i = 0; i < scopeSrNos.length; i++) {
                    engagementScope.push({
                        sr_no: scopeSrNos[i].value,
                        asset_description: scopeAssets[i].value,
                        criticality: scopeCriticalities[i].value,
                        internal_ip: scopeInternalIPs[i].value,
                        url: (scopeURLs[i].value && scopeURLs[i].value.toUpperCase() !== 'NA') ? scopeURLs[i].value : '',
                        public_ip: scopePublicIPs[i].value,
                        location: scopeLocations[i].value,
                        hash_value: scopeHashes[i].value,
                        version: scopeVersions[i].value,
                        other_details: scopeOthers[i].value
                    });
                }
                formData.append('engagement_scope', JSON.stringify(engagementScope));
                
                const teamNames = document.querySelectorAll('input[name="team_name[]"]');
                const teamDesignations = document.querySelectorAll('input[name="team_designation[]"]');
                const teamEmails = document.querySelectorAll('input[name="team_email[]"]');
                const teamQualifications = document.querySelectorAll('input[name="team_qualifications[]"]');
                const teamCertinListed = document.querySelectorAll('select[name="team_certin_listed[]"]');
                
                const auditingTeam = [];
                for (let i = 0; i < teamNames.length; i++) {
                    auditingTeam.push({
                        sr_no: i + 1,
                        name: teamNames[i].value,
                        designation: teamDesignations[i].value,
                        email: teamEmails[i].value,
                        qualifications: teamQualifications[i].value,
                        certin_listed: teamCertinListed[i].value
                    });
                }
                formData.append('auditing_team', JSON.stringify(auditingTeam));
                
                const activityTasks = document.querySelectorAll('input[name="activity_task[]"]');
                const activityDates = document.querySelectorAll('input[name="activity_date[]"]');
                
                const auditActivities = [];
                for (let i = 0; i < activityTasks.length; i++) {
                    auditActivities.push({
                        task: activityTasks[i].value,
                        date: activityDates[i].value
                    });
                }
                formData.append('audit_activities', JSON.stringify(auditActivities));
                
                const toolSrNos = document.querySelectorAll('input[name="tool_sr_no[]"]');
                const toolNames = document.querySelectorAll('input[name="tool_name[]"]');
                const toolVersions = document.querySelectorAll('input[name="tool_version[]"]');
                const toolTypes = document.querySelectorAll('select[name="tool_type[]"]');
                
                const toolsSoftware = [];
                for (let i = 0; i < toolSrNos.length; i++) {
                    toolsSoftware.push({
                        sr_no: toolSrNos[i].value,
                        name: toolNames[i].value,
                        version: toolVersions[i].value,
                        type: toolTypes[i].value
                    });
                }
                formData.append('tools_software', JSON.stringify(toolsSoftware));
                
                try {
                    document.getElementById('progress').style.display = 'block';
                    document.getElementById('progressText').textContent = 'Generating report...';
                    document.getElementById('progressBar').style.width = '50%';

                    await refreshCsrf();
                    if (CSRF_TOKEN) formData.append('csrf_token', CSRF_TOKEN);

                    const response = await fetch('/type3/generate-report/', {
                        method: 'POST',
                        headers: CSRF_TOKEN ? { 'X-CSRF-Token': CSRF_TOKEN } : {},
                        body: formData,
                        credentials: 'same-origin'
                    });
                    
                    if (response.status === 401) {
                        throw new Error('Session expired. <a href="/auth/test-login" style="color:#28a745;">Click here to log in again</a>.');
                    }
                    if (response.status === 403) {
                        throw new Error('CSRF validation failed. Refresh this page and try again.');
                    }
                    if (!response.ok) {
                        let msg = 'Report generation failed';
                        try {
                            const errorData = await response.json();
                            msg = errorData.detail || JSON.stringify(errorData);
                        } catch (e) {
                            const txt = await response.text();
                            if (txt) msg = txt;
                        }
                        throw new Error(msg);
                    }

                    const cd = response.headers.get('content-disposition') || '';
                    let saveName = 'Cert-IN_Report.docx';
                    const utf8m = cd.match(/filename\\*=UTF-8''([^;\\n]+)/i);
                    if (utf8m) {
                        try { saveName = decodeURIComponent(utf8m[1]); } catch (_) {}
                    } else {
                        const qm = cd.match(/filename="([^"]+)"/i);
                        const nm = cd.match(/filename=([^;\\n]+)/i);
                        if (qm) saveName = qm[1];
                        else if (nm) saveName = nm[1].trim().replace(/^["']|["']$/g, '');
                    }

                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = saveName;
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    document.body.removeChild(a);

                    const vulnN = response.headers.get('X-Report-Vuln-Count') || '—';
                    const pocN = response.headers.get('X-Report-Poc-Count') || '—';

                    document.getElementById('progressText').textContent = 'Report generated successfully!';
                    document.getElementById('progressBar').style.width = '100%';

                    document.getElementById('result').innerHTML = `
                        <h3>✅ Report Generated Successfully!</h3>
                        <p><strong>Filename:</strong> ${saveName}</p>
                        <p><strong>Vulnerabilities:</strong> ${vulnN}</p>
                        <p><strong>PoC observations with images:</strong> ${pocN}</p>
                        <p style="margin-top:10px;color:#555;">The Word file should download automatically. If it did not, run the flow again or check your browser download settings.</p>
                    `;
                    document.getElementById('result').className = 'result success';
                    document.getElementById('result').style.display = 'block';
                    
                } catch (error) {
                    document.getElementById('progress').style.display = 'none';
                    document.getElementById('result').innerHTML = `
                        <h3>❌ Error</h3>
                        <p>${error.message}</p>
                    `;
                    document.getElementById('result').className = 'result error';
                    document.getElementById('result').style.display = 'block';
                }
            });
        </script>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)

@app.get("/csrf-token")
async def csrf_token():
    """
    Compatibility endpoint for UI flows expecting a CSRF token route.
    This standalone service does not enforce CSRF validation.
    """
    return {"csrf_token": None}

@app.get("/type3/csrf-token")
async def csrf_token_type3():
    """Alias route for prefixed deployments."""
    return {"csrf_token": None}

@app.post("/type3/convert-excel/")
@app.post("/convert-excel/")
async def certin_convert_excel(
    vapt_file: UploadFile = File(...),
    poc_zip: UploadFile = File(None),
):
    """
    Step 1 for Cert-IN reports.
    Convert any VAPT Excel to the 29-column Cert-IN format Excel.
    """
    temp_vapt_path = None
    temp_zip_path = None
    extract_dir = None
    output_path = None
    try:
        if not vapt_file or not vapt_file.filename:
            raise HTTPException(status_code=400, detail="VAPT Excel file is required")

        suffix = os.path.splitext(vapt_file.filename)[1] or ".xlsx"
        temp_vapt_path = os.path.join(
            tempfile.gettempdir(),
            f"certin_conv_vapt_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}{suffix}",
        )
        with open(temp_vapt_path, "wb") as f:
            f.write(await vapt_file.read())

        if poc_zip and poc_zip.filename:
            temp_zip_path = os.path.join(
                tempfile.gettempdir(),
                f"certin_conv_poc_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.zip",
            )
            with open(temp_zip_path, "wb") as f:
                f.write(await poc_zip.read())

        output_path, extract_dir = convert_any_excel_to_format(
            temp_vapt_path,
            poc_zip_path=temp_zip_path,
            poc_layout="certin",
        )

        return FileResponse(
            path=output_path,
            filename="CertIN_format_output.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Cert-IN Excel conversion failed: {str(e)}")
    finally:
        for p in [temp_vapt_path, temp_zip_path]:
            try:
                if p and os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass
        try:
            if extract_dir and os.path.exists(extract_dir):
                shutil.rmtree(extract_dir, ignore_errors=True)
        except Exception:
            pass


@app.post("/type3/generate-report/")
@app.post("/generate-report/")
async def generate_report(
    request: Request,
    db: Session = Depends(get_db)
):
    """Generate Cert-IN report from form data"""
    try:
        print(f"Starting Cert-IN report generation - Version {VERSION}")

        form = await request.form()
        data = {k: form.get(k) for k in form.keys()}

        vulnerability_data = []
        poc_mapping = {}

        if 'vulnerability_file' in data and data['vulnerability_file']:
            vulnerability_file = data['vulnerability_file']
            temp_path = os.path.join(UPLOAD_DIR, f"temp_vuln_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            content = await vulnerability_file.read()
            with open(temp_path, "wb") as f:
                f.write(content)

            vulnerability_data = read_vulnerability_excel(temp_path)
            os.remove(temp_path)

        if 'poc_files' in data and data['poc_files']:
            poc_file = data['poc_files']
            poc_mapping = await process_poc_zip_files([poc_file], vulnerability_data)

        template_path = os.path.join(os.path.dirname(__file__), "CSS Certin temp.docx")
        if not os.path.exists(template_path):
            raise HTTPException(status_code=404, detail="Cert-IN template not found")

        client_name = data.get('client_name', 'Client')
        raw_filename = f"{client_name}_Cert-IN_Report_{datetime.now().strftime('%Y-%m-%d')}.docx"
        output_path = os.path.join(UPLOAD_DIR, raw_filename)

        generate_certin_report_from_form(data, template_path, output_path, vulnerability_data, poc_mapping)

        safe_filename = os.path.basename(output_path).replace(" ", "_")
        safe_output_path = os.path.join(os.path.dirname(output_path), safe_filename)

        if output_path != safe_output_path:
            os.rename(output_path, safe_output_path)

        try:
            user = request.session.get('user') or {}
            certin_report = CertINReport(
                client_name=data.get('client_name', ''),
                report_name=data.get('report_name', ''),
                report_release_date=data.get('report_release_date', ''),
                type_of_audit=data.get('type_of_audit', ''),
                type_of_audit_report=data.get('type_of_audit_report', ''),
                period=data.get('period', ''),
                document_title=data.get('document_title', ''),
                document_id=data.get('document_id', ''),
                document_version=data.get('document_version', ''),
                prepared_by=data.get('prepared_by', ''),
                reviewed_by=data.get('reviewed_by', ''),
                approved_by=data.get('approved_by', ''),
                released_by=data.get('released_by', ''),
                release_date=data.get('release_date', ''),
                document_change_history=data.get('document_change_history', '[]'),
                distribution_list=data.get('distribution_list', '[]'),
                engagement_scope=data.get('engagement_scope', '[]'),
                auditing_team=data.get('auditing_team', '[]'),
                audit_activities=data.get('audit_activities', '[]'),
                tools_software=data.get('tools_software', '[]'),
                created_by_email=user.get('email', 'unknown'),
                created_by_name=user.get('name'),
                file_path=safe_output_path
            )
            db.add(certin_report)
            db.commit()
        except Exception as e:
            print(f"Error saving to database: {e}")

        print(f"Report generated successfully - Version {VERSION}")

        try:
            user = request.session.get('user') or {}
            db.add(AuditLog(
                user_email=user.get('email'),
                user_name=user.get('name'),
                action='generate-certin-report-form',
                metadata_json=json.dumps({
                    'client_name': data.get('client_name'),
                    'template_used': 'CSS Certin temp.docx',
                    'version': VERSION,
                    'has_vulnerability_data': len(vulnerability_data) > 0,
                    'has_poc_data': len(poc_mapping) > 0,
                    'vulnerability_count': len(vulnerability_data),
                    'poc_count': len(poc_mapping)
                }),
                ip_address=request.client.host if request.client else None,
                user_agent=request.headers.get('user-agent')
            ))
            db.commit()
        except Exception as e:
            print(f"Error logging audit: {e}")

        # Return the file directly so the browser does not need a second GET to /download/
        # (avoids extra session/CSRF surface and matches Type 1/2 blob-download flow).
        return FileResponse(
            path=safe_output_path,
            filename=safe_filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "X-Report-Vuln-Count": str(len(vulnerability_data)),
                "X-Report-Poc-Count": str(len(poc_mapping)),
            },
        )

    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/download/{filename}")
async def download_report(filename: str):
    """Download generated report"""
    try:
        # Don't over-sanitize - just use the filename as-is but ensure it's safe
        # The filename was already sanitized when created
        file_path = os.path.join(UPLOAD_DIR, filename)
        
        if not os.path.exists(file_path):
            # Try with sanitized filename as fallback
            safe_filename = sanitize_filename(filename)
            file_path = os.path.join(UPLOAD_DIR, safe_filename)
            if not os.path.exists(file_path):
                raise HTTPException(status_code=404, detail=f"Report not found: {filename}")
        
        return FileResponse(
            path=file_path,
            filename=os.path.basename(file_path),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/reports/")
async def list_reports(request: Request, db: Session = Depends(get_db)):
    """List all generated Cert-IN reports"""
    try:
        user = request.session.get('user')
        if not user and os.getenv("TEST_MODE", "false").lower() != "true":
            raise HTTPException(status_code=401, detail="Not authenticated")

        reports = db.query(CertINReport).order_by(CertINReport.created_at.desc()).limit(50).all()
        
        return [
            {
                "id": report.id,
                "client_name": report.client_name,
                "report_name": report.report_name,
                "document_title": report.document_title,
                "created_at": report.created_at.isoformat(),
                "created_by": report.created_by_name,
                "download_url": f"/type3/download/{os.path.basename(report.file_path)}" if report.file_path else None
            }
            for report in reports
        ]
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="localhost", port=8003)