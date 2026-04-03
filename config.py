from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# ──────────────────────────────────────────────
# Files
# ──────────────────────────────────────────────
INPUT_FILE    = "Input Sample file.xlsx"
TEMPLATE_FILE = "ppt_template.pptx"
OUTPUT_DIR    = "output"

BLUEPRINT_IDX = 1  # blueprint is slide 2 (index 1) in the template

# ──────────────────────────────────────────────
# Card layout
# ──────────────────────────────────────────────
CARD_W       = Inches(2.2)
CARD_H       = Inches(0.8)
CARD_GAP_X   = Inches(0.1)
CARD_GAP_Y   = Inches(0.8)
START_TOP    = Inches(0.5)
START_LEFT   = Inches(0.5)
RUBRIC_LABEL_W = Inches(1.5)
RUBRIC_GAP     = Inches(0.1)

# ──────────────────────────────────────────────
# Colors
# ──────────────────────────────────────────────
BORDER_COLOR  = RGBColor(255, 98,  87)
TEXT_COLOR    = RGBColor(255, 255, 255)
HEADER_COLOR  = RGBColor(249, 214, 5)
CARD_BG_COLOR = RGBColor(30,  30,  30)

# ──────────────────────────────────────────────
# Metrics
# ──────────────────────────────────────────────
TV_METRICS_BY_CHANNEL = [
    'standard spots',
    'standard spots seconds',
    'live ad spots',
    'live ad spots seconds',
]

TV_METRICS_OTHER = [
    'event promo count',
    'break bumper count',
    'graphic overlay count',
    'commentator announcement',
    'average daily reach',
    'commentator announcement seconds',
    'event promo seconds',
    'break bumper seconds',
    'graphic overlay seconds',
]

TV_SUMMARY_METRICS = [
    'break bumper count',
    'event promo count',
    'graphic overlay count',
    'commentator announcement',
]

OTT_METRICS = [
    'pre-roll views',
    'live ad views',
    'customers',
    'break bumpers',
    'hero banner impressions',
    'events count',
    'spots count',
    'spots seconds',
    'landing page match views',
]

SMM_METRICS = [
    'post count',
    'post impressions',
    'video post views',
    'video count',
    'video views',
    'video impressions',
]