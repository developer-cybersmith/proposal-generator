# Executive Summary template assets

PNG files extracted from the CSS Executive Summary reference DOCX:

- `cover_bg.png` — full-page cover background
- `header_logo.png` — CyberSmithSECURE header banner
- `footer_banner.png` — footer with contact details + branding
- `icon_small.png` — optional small icon (reserved)

The objectives “Completed” column uses a small green checkmark PNG generated at runtime (matplotlib); no `checkmark.png` asset is required.

Regenerate `exec_summary_template.docx` after replacing assets:

```bash
python Automation/backend/build_exec_template.py
```

To re-extract from a fresh CSS reference template, copy `word/media/image1.png` … `image4.png` from that `.docx` into this folder with the names above, then run the build script.
