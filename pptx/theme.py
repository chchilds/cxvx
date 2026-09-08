"""Dark theme color palette for Cisco EA PowerPoint decks."""

from pptx.dml.color import RGBColor

# Core palette
DARK_NAVY = RGBColor(0x1A, 0x2B, 0x4A)
DARK_ROW = RGBColor(0x2D, 0x3A, 0x52)
DARK_BANNER = RGBColor(0x24, 0x3B, 0x5C)
CISCO_BLUE = RGBColor(0x00, 0x7B, 0xC7)
TEAL_ACCENT = RGBColor(0x00, 0xA9, 0x8E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BODY_TEXT = RGBColor(0xE0, 0xE0, 0xE0)
FOOTER_TEXT = RGBColor(0x99, 0x99, 0x99)
METADATA_TEXT = RGBColor(0xCC, 0xCC, 0xCC)

# Hex strings for XML-level replacements (no # prefix, uppercase)
COLOR_MAP = {
    "333333": "E0E0E0",  # body text
    "F5F5F5": "2D3A52",  # table alternating rows
    "CCE8F7": "243B5C",  # subtitle banners
    "666666": "999999",  # footer / disclaimer text
}
