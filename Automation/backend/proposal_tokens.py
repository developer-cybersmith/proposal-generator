"""
Formatting tokens for generated proposal content.

Single source of every font, colour, size, spacing and table metric. The values come from
docs/CyberSmithSECURE_Proposal_Formatting_Specification.md (the binding standard, plan section 0).
No other module may contain raw hex colours or font names for generated content; import from here.

Open brand decisions (plan section 12, Track B) are marked TODO and never guessed.
"""
from __future__ import annotations

# ---------------------------------------------------------------- colours (hex, no '#')
BRAND_PURPLE = "7030A0"        # headings, annexure titles, table header fill, accent lines
BODY_GRAY = "595959"           # body copy (value verified in the source document)
BODY_DARK = "404040"           # table body text; spec allows 404040 or 595959 for body
BODY_BLACK = "000000"
WHITE = "FFFFFF"
TABLE_HEADER_TEXT = WHITE
TABLE_BODY_FILL = "F2F2F2"     # optional alternating rows and the Total row
TABLE_BORDER = "7F7F7F"
RULE_GRAY = "B7B7B7"           # header/footer separators
COVER_NAVY = None              # TODO(brand owner): exact value not yet confirmed (spec section 4)

BODY_TEXT = BODY_GRAY          # TODO(Track B): body colour 404040 vs 595959 still to be decided

# Colours that may appear in generated runs. The lint warns on anything else.
ALLOWED_COLORS = frozenset({BRAND_PURPLE, BODY_GRAY, BODY_DARK, BODY_BLACK, WHITE})

# ---------------------------------------------------------------- fonts
FONT = "Corbel"

# ---------------------------------------------------------------- sizes (pt)
SIZE_HEADING_1 = 18
SIZE_HEADING_2 = 16
SIZE_HEADING_3 = 14            # spec allows 13-14; the source uses 14
SIZE_ANNEXURE_TITLE = 16
SIZE_BODY = 11
SIZE_LIST = 11
SIZE_TABLE_HEADER = 10
SIZE_TABLE_BODY = 10
SIZE_TABLE_TOTAL = 10.5

# ---------------------------------------------------------------- paragraph spacing
BODY_SPACE_BEFORE_PT = 0
BODY_SPACE_AFTER_PT = 6
LIST_SPACE_AFTER_PT = 3
LIST_LEVEL_INDENT_IN = (0.25, 0.50, 0.75)
LIST_HANGING_IN = 0.25
LIST_GLYPHS = ("•", "o")  # round bullet, open circle (the source's glyphs)

# ---------------------------------------------------------------- pricing table (plan section 7)
PRICING_HEADERS = ("Sr. No.", "Assessment Type", "Quantity", "Frequency", "Cost Per Annexure")
PRICING_COLUMN_WIDTH_PCT = (8, 36, 20, 16, 20)   # starting point, tuned after a rendered check
PRICING_TOTAL_LABEL = "Total"
PRICING_TOTAL_EMPTY_TEXT = "₹ /- * + 18% GST"
PRICING_ALIGN = ("center", "left", "center", "center", "right")
TABLE_BORDER_EIGHTHS_PT = 6    # Word border size unit: 6 = 0.75 pt
TABLE_CELL_PAD_V_IN = 0.08
TABLE_CELL_PAD_H_IN = 0.10

# ---------------------------------------------------------------- page (spec section 3)
PAGE_WIDTH_IN = 8.27
PAGE_HEIGHT_IN = 11.69
PAGE_MARGIN_IN = 1.0           # Track B target; the current template differs and is fixed there


def rgb(hex_value: str) -> tuple[int, int, int]:
    """'7030A0' -> (112, 48, 160), for python-docx RGBColor(*rgb(...))."""
    return int(hex_value[0:2], 16), int(hex_value[2:4], 16), int(hex_value[4:6], 16)
