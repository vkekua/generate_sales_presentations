from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor

from config import BLUEPRINT_IDX, TEMPLATE_FILE, SMM_METRICS, TV_METRICS_BY_CHANNEL
from utils import _copy_slide
from ppt_helpers import (
    add_centered_textbox,
    build_channel_kpi_cards,
    build_kpi_cards_grid,
    build_youtube_rubric_cards,
    build_totals_slide,
    is_valid_value,
    style_card,
    style_text_frame,
    START_TOP, START_LEFT, CARD_W, CARD_H, CARD_GAP_X, HEADER_COLOR,
)
from pptx.util import Pt


YELLOW = RGBColor(249, 214, 5)
WHITE  = RGBColor(255, 255, 255)


def build_presentation(partner: str, country: str, month: str, data: dict, output_path: str):
    """
    Build the full presentation for one partner and save it.

    Args:
        partner:     Partner name string.
        country:     Country string.
        month:       Month string (e.g. "January 2025").
        data:        Dict of prepared dataframes from data.prepare_partner_data().
        output_path: File path to save the .pptx file.
    """
    prs     = Presentation(TEMPLATE_FILE)
    SLIDE_W = prs.slide_width
    SLIDE_H = prs.slide_height

    def copy(offset: int):
        """Shortcut: copy blueprint and insert at blueprint + offset."""
        return _copy_slide(prs, src_idx=BLUEPRINT_IDX, dst_prs=prs, insert_idx=BLUEPRINT_IDX + offset)

    def title_slide(slide, text: str):
        """Add a centered title to a slide."""
        add_centered_textbox(slide, text=text, font_size=54, font_color=YELLOW,
                             top=SLIDE_H * 0.38, height=Inches(1.0), slide_w=SLIDE_W)

    # ──────────────────────────────────────────────
    # SLIDE 1 — Partner / Country / Month
    # ──────────────────────────────────────────────
    slide1 = copy(1)
    add_centered_textbox(slide1, text=partner,                font_size=40, font_color=YELLOW,
                         top=SLIDE_H * 0.30, height=Inches(0.8), slide_w=SLIDE_W)
    add_centered_textbox(slide1, text=f"{country}  |  {month}", font_size=24, font_color=WHITE,
                         top=SLIDE_H * 0.42, height=Inches(0.6), slide_w=SLIDE_W, bold=False)

    # ──────────────────────────────────────────────
    # SLIDE 2 — TV REPORT title
    # ──────────────────────────────────────────────
    title_slide(copy(2), "TV REPORT")

    # ──────────────────────────────────────────────
    # SLIDE 3 — TV Summary (channel cards + other metrics)
    # ──────────────────────────────────────────────
    tv_slide = copy(3)
    build_channel_kpi_cards(tv_slide, data["tv_channel"], TV_METRICS_BY_CHANNEL)

    right_start = START_LEFT + len(TV_METRICS_BY_CHANNEL) * (CARD_W + CARD_GAP_X) + Inches(0.3)
    visible_k   = 0
    for _, row_other in data["tv_other"].iterrows():
        value = row_other["Value"]
        if not is_valid_value(value):
            continue
        top  = START_TOP + Inches(0.5) + visible_k * (CARD_H + Inches(0.1))
        from pptx.enum.shapes import MSO_SHAPE
        card = tv_slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, right_start, top, CARD_W, CARD_H)
        style_card(card)
        card.text_frame.text = f"{row_other['Metric']}\n{value}"
        style_text_frame(card.text_frame, font_size=12)
        visible_k += 1

    # ──────────────────────────────────────────────
    # SLIDE 4 — OTT REPORT title
    # ──────────────────────────────────────────────
    title_slide(copy(4), "OTT REPORT")

    # ──────────────────────────────────────────────
    # SLIDE 5 — OTT Summary (2-row centered grid)
    # ──────────────────────────────────────────────
    build_kpi_cards_grid(copy(5), data["ott"], SLIDE_W, SLIDE_H)

    # ──────────────────────────────────────────────
    # SLIDE 6 — SOCIAL MEDIA REPORT title
    # ──────────────────────────────────────────────
    title_slide(copy(6), "SOCIAL MEDIA REPORT")

    # ──────────────────────────────────────────────
    # SLIDE 7 — SMM Summary (channel cards)
    # ──────────────────────────────────────────────
    build_channel_kpi_cards(copy(7), data["smm_channel"], SMM_METRICS)

    # ──────────────────────────────────────────────
    # SLIDE 8 — YouTube Summary (rubric rows)
    # ──────────────────────────────────────────────
    build_youtube_rubric_cards(copy(8), data["yt_rubric"], SMM_METRICS)

    # ──────────────────────────────────────────────
    # SLIDE 9 — Totals (one fitted row per platform)
    # ──────────────────────────────────────────────
    build_totals_slide(copy(9), data["all_totals"], SLIDE_W)

    # ──────────────────────────────────────────────
    # Delete blueprint slide and save
    # ──────────────────────────────────────────────
    del prs.slides._sldIdLst[BLUEPRINT_IDX]
    prs.save(output_path)