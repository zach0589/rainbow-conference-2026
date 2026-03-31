#!/usr/bin/env python3
"""Build Rainbow Conference 2026 PowerPoint - 15 slides."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import os

# ── Constants ──
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
NAVY = RGBColor(0x1B, 0x30, 0x57)
RED = RGBColor(0xC8, 0x10, 0x2E)
DARKRED = RGBColor(0x9B, 0x1B, 0x30)
GRAY = RGBColor(0x58, 0x59, 0x5B)
LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MUTED_WHITE = RGBColor(0xAA, 0xAA, 0xAA)
FONT = "Segoe UI"

ASSETS = "/private/tmp/rainbow-conference-2026/assets"
RAINBOW_LOGO = os.path.join(ASSETS, "rainbow-logo.png")
SOS_LOGO = os.path.join(ASSETS, "sos-logo.png")
SOS_LOGO_WHITE = os.path.join(ASSETS, "sos-logo-white.png")

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H

BLANK_LAYOUT = prs.slide_layouts[6]  # blank


# ── Helpers ──
def add_slide():
    return prs.slides.add_slide(BLANK_LAYOUT)

def set_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_text(slide, left, top, width, height, text, font_size=18,
             color=WHITE, bold=False, alignment=PP_ALIGN.LEFT,
             font_name=FONT, line_spacing=None):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    if line_spacing:
        p.line_spacing = Pt(line_spacing)
    return txBox

def add_rich_text(slide, left, top, width, height, runs, alignment=PP_ALIGN.LEFT, line_spacing=None):
    """runs = list of (text, font_size, color, bold)"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    if alignment:
        p.alignment = alignment
    if line_spacing:
        p.line_spacing = Pt(line_spacing)
    for i, (txt, fs, clr, bld) in enumerate(runs):
        if i == 0:
            r = p.runs[0] if p.runs else p.add_run()
            r.text = txt
        else:
            r = p.add_run()
            r.text = txt
        r.font.size = Pt(fs)
        r.font.color.rgb = clr
        r.font.bold = bld
        r.font.name = FONT
    return txBox

def add_multiline(slide, left, top, width, height, lines, font_size=16,
                  color=GRAY, bold=False, alignment=PP_ALIGN.LEFT,
                  bullet_color=None, line_spacing=None):
    """lines = list of strings, each becomes a paragraph."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.font.bold = bold
        p.font.name = FONT
        p.alignment = alignment
        if line_spacing:
            p.line_spacing = Pt(line_spacing)
        p.space_after = Pt(6)
    return txBox

def add_bullet_list(slide, left, top, width, height, items, font_size=16,
                    color=GRAY, line_spacing=None):
    """items = list of str. Each gets a bullet character prefix."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(8)
        if line_spacing:
            p.line_spacing = Pt(line_spacing)
        # Red bullet dot
        r_dot = p.add_run()
        r_dot.text = "\u2022  "
        r_dot.font.size = Pt(font_size)
        r_dot.font.color.rgb = RED
        r_dot.font.name = FONT
        r = p.add_run()
        r.text = item
        r.font.size = Pt(font_size)
        r.font.color.rgb = color
        r.font.name = FONT
    return txBox

def add_rich_bullet_list(slide, left, top, width, height, items, font_size=16, line_spacing=None):
    """items = list of lists of (text, font_size, color, bold) tuples."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, runs in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(10)
        if line_spacing:
            p.line_spacing = Pt(line_spacing)
        # Bullet
        r_dot = p.add_run()
        r_dot.text = "\u2022  "
        r_dot.font.size = Pt(font_size)
        r_dot.font.color.rgb = RED
        r_dot.font.name = FONT
        for (txt, fs, clr, bld) in runs:
            r = p.add_run()
            r.text = txt
            r.font.size = Pt(fs)
            r.font.color.rgb = clr
            r.font.bold = bld
            r.font.name = FONT
    return txBox

def add_logos_light(slide):
    """Logos for dark bg slides — rainbow left, sos-white right."""
    slide.shapes.add_picture(RAINBOW_LOGO, Inches(0.4), Inches(6.7), height=Inches(0.55))
    slide.shapes.add_picture(SOS_LOGO_WHITE, Inches(11.5), Inches(6.85), height=Inches(0.35))

def add_logos_dark(slide):
    """Logos for white bg slides — rainbow left, sos right."""
    slide.shapes.add_picture(RAINBOW_LOGO, Inches(0.4), Inches(6.7), height=Inches(0.55))
    slide.shapes.add_picture(SOS_LOGO, Inches(11.5), Inches(6.85), height=Inches(0.35))

def add_rect(slide, left, top, width, height, fill_color=None, border_color=None, border_width=Pt(0)):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.line.fill.background()
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = border_width
        shape.line.fill.solid()
    return shape

def add_rounded_rect(slide, left, top, width, height, fill_color=None, border_color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(2)
    else:
        shape.line.fill.background()
    return shape

def add_stat_card(slide, left, top, width, height, num_text, label_text, border_color=RED):
    """Add a stat card with colored top border."""
    # Background
    card = add_rect(slide, left, top, width, height, fill_color=LIGHT_GRAY)
    # Top border line
    add_rect(slide, left, top, width, Pt(5), fill_color=border_color)
    # Number
    add_text(slide, left, top + Inches(0.25), width, Inches(0.6),
             num_text, font_size=32, color=NAVY, bold=True, alignment=PP_ALIGN.CENTER)
    # Label
    add_text(slide, left + Inches(0.1), top + Inches(0.85), width - Inches(0.2), Inches(0.8),
             label_text, font_size=11, color=GRAY, alignment=PP_ALIGN.CENTER)

def add_callout(slide, left, top, width, height, text, border_color=NAVY, font_size=13, text_color=GRAY):
    """Callout box with left border."""
    bg_color = RGBColor(0xEE, 0xEF, 0xF2) if border_color == NAVY else RGBColor(0xFB, 0xF0, 0xF1)
    card = add_rect(slide, left, top, width, height, fill_color=bg_color)
    add_rect(slide, left, top, Pt(5), height, fill_color=border_color)
    add_text(slide, left + Inches(0.25), top + Inches(0.08), width - Inches(0.35), height - Inches(0.16),
             text, font_size=font_size, color=text_color, alignment=PP_ALIGN.LEFT)

def add_source(slide, text, top=Inches(7.0)):
    add_text(slide, Inches(0), top, SLIDE_W, Inches(0.3),
             text, font_size=9, color=MUTED_WHITE if False else RGBColor(0xAA, 0xAA, 0xAA),
             alignment=PP_ALIGN.CENTER)

def add_funnel_trapezoid(slide, left, top, top_width, bottom_width, height, fill_color):
    """Draw a trapezoid using a freeform shape."""
    center_x = left + top_width // 2
    # Calculate offsets
    top_left_x = left
    top_right_x = left + top_width
    bottom_left_x = left + (top_width - bottom_width) // 2
    bottom_right_x = bottom_left_x + bottom_width

    # Use freeform builder
    freeform = slide.shapes.build_freeform(top_left_x, top)
    freeform.add_line_segments([
        (top_right_x, top),
        (bottom_right_x, top + height),
        (bottom_left_x, top + height),
        (top_left_x, top),
    ])
    shape = freeform.convert_to_shape()
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape

def add_funnel(slide, center_x, top_y, total_width, total_height, labels=True):
    """Draw the 4-stage funnel centered at center_x."""
    gap = Inches(0.06)
    stage_h = (total_height - 3 * gap) / 4

    # Widths narrow from top to bottom
    widths = [total_width, total_width * 0.72, total_width * 0.48, total_width * 0.28]
    colors = [RED, DARKRED, NAVY, GRAY]
    texts = ["AWARENESS", "CONSIDERATION", "DECISION", "CONVERSION"]
    font_sizes = [16, 14, 16, 12]

    y = top_y
    for i in range(4):
        if i == 0:
            tw = widths[0]
        else:
            tw = widths[i-1]  # top width = previous bottom width
        bw = widths[i] if i < 3 else widths[i]

        # For proper funnel: top of each stage = bottom of previous
        if i == 0:
            tw = total_width
            bw = widths[1]
        elif i == 1:
            tw = widths[1]
            bw = widths[2]
        elif i == 2:
            tw = widths[2]
            bw = widths[3]
        else:
            tw = widths[3]
            bw = total_width * 0.18

        left = center_x - tw // 2
        shape = add_funnel_trapezoid(slide, left, int(y), int(tw), int(bw), int(stage_h), colors[i])

        if labels:
            # Add text centered on the trapezoid
            text_w = max(int(tw), int(bw))
            text_left = center_x - text_w // 2
            add_text(slide, text_left, int(y + stage_h * 0.25), text_w, int(stage_h * 0.5),
                     texts[i], font_size=font_sizes[i], color=WHITE, bold=True,
                     alignment=PP_ALIGN.CENTER)

        y += stage_h + gap


# ══════════════════════════════════════════════════════════════
# SLIDE 1: Title
# ══════════════════════════════════════════════════════════════
s = add_slide()
set_bg(s, NAVY)
add_text(s, Inches(1), Inches(1.8), Inches(11.3), Inches(2),
         "The 2026 Restoration\nMarketing Reality",
         font_size=48, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER, line_spacing=56)
add_text(s, Inches(1), Inches(3.9), Inches(11.3), Inches(0.8),
         "Winning with the Right Media Mix",
         font_size=24, color=RED, alignment=PP_ALIGN.CENTER)
add_text(s, Inches(1), Inches(5.2), Inches(11.3), Inches(0.5),
         "Rainbow Restoration National Conference 2026",
         font_size=14, color=RGBColor(0x80, 0x80, 0x90), alignment=PP_ALIGN.CENTER)
add_logos_light(s)

# ══════════════════════════════════════════════════════════════
# SLIDE 2: The Problem (split bg)
# ══════════════════════════════════════════════════════════════
s = add_slide()
# Top half white
add_rect(s, Inches(0), Inches(0), SLIDE_W, Inches(3.75), fill_color=WHITE)
# Bottom half navy
add_rect(s, Inches(0), Inches(3.75), SLIDE_W, Inches(3.75), fill_color=NAVY)
# Red divider line (80% width centered)
line_w = Inches(10.667)
line_left = (SLIDE_W - line_w) // 2
add_rect(s, line_left, Inches(3.72), line_w, Pt(4), fill_color=RED)
# Top text
add_text(s, Inches(1), Inches(1.8), Inches(11.3), Inches(1.5),
         "How your customers find you\nhas changed",
         font_size=40, color=NAVY, bold=True, alignment=PP_ALIGN.CENTER, line_spacing=50)
# Bottom text
add_text(s, Inches(1), Inches(4.5), Inches(11.3), Inches(1),
         "Has your marketing kept up?",
         font_size=26, color=MUTED_WHITE, alignment=PP_ALIGN.CENTER)
# Logos — bottom half is navy so use white SOS
s.shapes.add_picture(RAINBOW_LOGO, Inches(0.4), Inches(6.7), height=Inches(0.55))
s.shapes.add_picture(SOS_LOGO_WHITE, Inches(11.5), Inches(6.85), height=Inches(0.35))

# ══════════════════════════════════════════════════════════════
# SLIDE 3: Media Mix
# ══════════════════════════════════════════════════════════════
s = add_slide()
# Title with red underline
add_text(s, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6),
         "Your Media Mix Is Your System", font_size=28, color=NAVY, bold=True)
add_rect(s, Inches(0.8), Inches(0.95), Inches(4.5), Pt(3), fill_color=RED)

# Bullets
add_bullet_list(s, Inches(0.8), Inches(1.2), Inches(11), Inches(1.2),
    [
        "It's every marketing channel you invest in \u2014 and how they connect",
        "Most operators have a list of channels. The best operators have a system.",
    ], font_size=16, line_spacing=24)

# Three boxes
box_w = Inches(3.5)
box_h = Inches(1.6)
box_y = Inches(3.0)
box_gap = Inches(0.35)
box_start = Inches(1.2)

boxes = [
    ("CAPTURE", "Demand that\nalready exists", NAVY, RGBColor(0xE8, 0xEC, 0xF2)),
    ("CREATE", "Demand before the\nemergency happens", RED, RGBColor(0xFB, 0xF0, 0xF1)),
    ("CONVERT", "Attention into\nbooked jobs", GRAY, RGBColor(0xF0, 0xF0, 0xF1)),
]
for i, (title, desc, color, bg) in enumerate(boxes):
    x = box_start + i * (box_w + box_gap)
    add_rect(s, x, box_y, box_w, box_h, fill_color=bg)
    add_rect(s, x, box_y, box_w, Pt(5), fill_color=color)
    add_text(s, x, box_y + Inches(0.2), box_w, Inches(0.5),
             title, font_size=26, color=color, bold=True, alignment=PP_ALIGN.CENTER)
    add_text(s, x, box_y + Inches(0.7), box_w, Inches(0.7),
             desc, font_size=13, color=GRAY, alignment=PP_ALIGN.CENTER)

# Callout
add_callout(s, Inches(0.8), Inches(5.1), Inches(11.5), Inches(0.7),
            "Most restoration companies are good at one of these. The ones growing are covering all three.",
            border_color=NAVY, font_size=14)
add_logos_dark(s)

# ══════════════════════════════════════════════════════════════
# SLIDE 4: Funnel
# ══════════════════════════════════════════════════════════════
s = add_slide()
add_text(s, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6),
         "The Marketing Funnel", font_size=28, color=NAVY, bold=True)
add_rect(s, Inches(0.8), Inches(0.95), Inches(3.5), Pt(3), fill_color=RED)

# Funnel on the left
funnel_cx = int(Inches(3.5))
funnel_top = int(Inches(1.4))
funnel_w = int(Inches(4.0))
funnel_h = int(Inches(4.2))
add_funnel(s, funnel_cx, funnel_top, funnel_w, funnel_h, labels=True)

# Labels on the right
right_x = Inches(6.5)
label_y_start = Inches(1.5)
label_gap = Inches(1.1)

stages = [
    (RED, "Awareness", "Social media, sponsorships, YouTube, vehicle wraps"),
    (DARKRED, "Consideration", "Retargeting, blog content, SEO, email nurture"),
    (NAVY, "Decision", "LSA, Google Ads, SEO (high-intent), GBP"),
    (GRAY, "Conversion", "Website, reviews, speed to lead, follow-up"),
]
for i, (color, name, desc) in enumerate(stages):
    y = label_y_start + i * label_gap
    # Dot
    dot = add_rect(s, right_x, y + Inches(0.05), Inches(0.15), Inches(0.15), fill_color=color)
    # Name
    add_text(s, right_x + Inches(0.3), y - Inches(0.05), Inches(5.5), Inches(0.35),
             name, font_size=18, color=color, bold=True)
    # Desc
    add_text(s, right_x + Inches(0.3), y + Inches(0.3), Inches(5.5), Inches(0.5),
             desc, font_size=12, color=GRAY)

# Red callout
add_callout(s, Inches(0.8), Inches(5.8), Inches(11.5), Inches(0.85),
            "Restoration is emergency-driven. The consideration phase is far shorter than most industries. Trust has to be built BEFORE the emergency.",
            border_color=RED, font_size=13)
add_logos_dark(s)

# ══════════════════════════════════════════════════════════════
# SLIDE 5: Landscape
# ══════════════════════════════════════════════════════════════
s = add_slide()
add_text(s, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6),
         "The Landscape Is Shifting", font_size=28, color=NAVY, bold=True)
add_rect(s, Inches(0.8), Inches(0.95), Inches(4.0), Pt(3), fill_color=RED)

# 3 stat cards
card_w = Inches(3.5)
card_h = Inches(1.6)
card_y = Inches(1.3)
card_gap = Inches(0.4)
card_start = Inches(1.2)

stats5 = [
    ("29%", "of homeowners already have a\ncompany name in mind\nwhen they search", RED),
    ("+27.8%", "Google Ads spend growth\nacross our restoration clients", NAVY),
    ("+63.9%", "LSA spend growth\nacross our restoration clients", RED),
]
for i, (num, label, border_c) in enumerate(stats5):
    x = card_start + i * (card_w + card_gap)
    add_stat_card(s, x, card_y, card_w, card_h, num, label, border_color=border_c)

# Two smaller bullets
add_bullet_list(s, Inches(0.8), Inches(3.3), Inches(11.5), Inches(1.0),
    [
        "Businesses in LSA results get 25-30% more calls than those relying on organic alone",
        "31% of consumers now require 4.5+ stars to consider a business \u2014 up from 17% last year",
    ], font_size=14, line_spacing=22)

# Callout
add_callout(s, Inches(0.8), Inches(4.7), Inches(11.5), Inches(0.85),
            "Your competitors are spending more every quarter. But the operators winning aren't just spending \u2014 they're building recognition before the emergency.",
            border_color=NAVY, font_size=13)

add_source(s, "Sources: SOS x C&R Magazine Report 2026; AdMall/SalesFuel 2025; BrightLocal LCRS 2026; Boomcycle 2026", top=Inches(6.6))
add_logos_dark(s)

# ══════════════════════════════════════════════════════════════
# SLIDE 6: Customer Behavior
# ══════════════════════════════════════════════════════════════
s = add_slide()
add_text(s, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6),
         "What Your Customers Respond To", font_size=28, color=NAVY, bold=True)
add_rect(s, Inches(0.8), Inches(0.95), Inches(5.0), Pt(3), fill_color=RED)

stats6 = [
    ("97%", "of restoration operators say referrals\nare their #1 source of new business", RED),
    ("58%", "of homeowners were influenced\nby a blog post in the last 30 days", NAVY),
    ("55,352", "restoration establishments\ncompeting for attention nationwide", RED),
]
for i, (num, label, border_c) in enumerate(stats6):
    x = card_start + i * (card_w + card_gap)
    add_stat_card(s, x, card_y, card_w, card_h, num, label, border_color=border_c)

add_bullet_list(s, Inches(0.8), Inches(3.3), Inches(11.5), Inches(1.4),
    [
        "Homeowners are 12% more likely to check a company\u2019s blog to evaluate expertise",
        "41% called or visited the advertiser after seeing an ad \u2014 136% higher than the general population",
        "Only 3.5% of restoration customers say they don\u2019t read online reviews \u2014 nearly everyone checks",
    ], font_size=14, line_spacing=22)

add_source(s, "Sources: AdMall/SalesFuel 2025; Cleanfax Restoration Benchmarking Survey 2024", top=Inches(6.6))
add_logos_dark(s)

# ══════════════════════════════════════════════════════════════
# SLIDE 7: AI
# ══════════════════════════════════════════════════════════════
s = add_slide()
add_text(s, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6),
         "AI Is Changing How Customers Find You", font_size=28, color=NAVY, bold=True)
add_rect(s, Inches(0.8), Inches(0.95), Inches(6.0), Pt(3), fill_color=RED)

# 3 stat cards - smaller
card_h7 = Inches(1.2)
stats7 = [
    ("14.2%", "AI Conversion Rate", RED),
    ("2.8%", "Google Conversion Rate", NAVY),
    ("5x", "AI Converts Higher", GRAY),
]
for i, (num, label, border_c) in enumerate(stats7):
    x = card_start + i * (card_w + card_gap)
    add_stat_card(s, x, Inches(1.3), card_w, card_h7, num, label, border_color=border_c)

# Two columns
col_left = Inches(0.8)
col_right = Inches(6.8)
col_w = Inches(5.5)
col_top = Inches(2.9)

add_text(s, col_left, col_top, col_w, Inches(0.4),
         "THE REALITY", font_size=13, color=NAVY, bold=True)
add_bullet_list(s, col_left, col_top + Inches(0.35), col_w, Inches(1.3),
    [
        "Google still handles 90% of search",
        "AI converts at 5x the rate",
        "41% of restoration customers use AI to research",
    ], font_size=13, line_spacing=20)

add_text(s, col_right, col_top, col_w, Inches(0.4),
         "THE GOOD NEWS", font_size=13, color=NAVY, bold=True)
# For "good news" column, last item needs special color
txBox = slide_shapes = add_bullet_list(s, col_right, col_top + Inches(0.35), col_w, Inches(1.5),
    [
        "Google and AI reward the same thing",
        "Authoritative content, strong reviews, managed GBP",
        'A brand worth citing \u2014 AI recommends what it "knows"',
    ], font_size=13, line_spacing=20)

# Callout
add_callout(s, Inches(0.8), Inches(5.1), Inches(11.5), Inches(0.75),
            "You don't need two strategies. Google and AI are two systems rewarding the same thing.",
            border_color=RED, font_size=14)

add_source(s, "Source: Superprompt, 12M Visit Study, 2025", top=Inches(6.6))
add_logos_dark(s)

# ══════════════════════════════════════════════════════════════
# SLIDE 8: What This Means
# ══════════════════════════════════════════════════════════════
s = add_slide()
add_text(s, Inches(0.8), Inches(0.4), Inches(11), Inches(0.6),
         "What This Means for Your Funnel", font_size=28, color=NAVY, bold=True)
add_rect(s, Inches(0.8), Inches(0.95), Inches(5.0), Pt(3), fill_color=RED)

# Three rich bullets with red bold lead + gray continuation
bullets8 = [
    [
        ("The bottom of the funnel still matters most. ", 16, RED, True),
        ("LSA and PPC are where leads come from. But costs are up 20.9% \u2014 the bottom alone isn't enough anymore.", 16, GRAY, False),
    ],
    [
        ("The top of the funnel makes the bottom work harder. ", 16, RED, True),
        ("Branded searches convert at 6x the rate of generic terms. Awareness doesn't replace demand capture \u2014 it supercharges it.", 16, GRAY, False),
    ],
    [
        ("AI is the new front door. ", 16, RED, True),
        ("41% of your potential customers are using AI to research you. Your content, reviews, and GBP feed what AI recommends.", 16, GRAY, False),
    ],
]
add_rich_bullet_list(s, Inches(0.8), Inches(1.3), Inches(11.5), Inches(3.5),
                     bullets8, font_size=16, line_spacing=24)

# Navy callout
add_callout(s, Inches(0.8), Inches(5.0), Inches(11.5), Inches(0.85),
            "Keep investing in the bottom. But the operators pulling ahead are the ones adding layers above it.",
            border_color=NAVY, font_size=14)

add_source(s, "Sources: SOS x C&R Magazine Report 2026; AdMall/SalesFuel 2025", top=Inches(6.6))
add_logos_dark(s)

# ══════════════════════════════════════════════════════════════
# SLIDE 9: Workshop Divider
# ══════════════════════════════════════════════════════════════
s = add_slide()
set_bg(s, RED)
add_text(s, Inches(1), Inches(1.5), Inches(11.3), Inches(1.5),
         "WORKSHOP", font_size=60, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)
add_text(s, Inches(1), Inches(3.0), Inches(11.3), Inches(0.8),
         "Build Your Marketing Funnel", font_size=26, color=WHITE, alignment=PP_ALIGN.CENTER)
# QR placeholder
qr_left = Inches(4.5)
qr_top = Inches(4.2)
qr_size = Inches(1.2)
qr = add_rounded_rect(s, qr_left, qr_top, qr_size, qr_size,
                       fill_color=RGBColor(0xD0, 0x20, 0x3E),
                       border_color=RGBColor(0xE0, 0x50, 0x60))
add_text(s, qr_left, qr_top + Inches(0.4), qr_size, Inches(0.4),
         "QR CODE", font_size=10, color=RGBColor(0xE0, 0x70, 0x80), alignment=PP_ALIGN.CENTER)
# Text next to QR
add_text(s, Inches(5.9), Inches(4.2), Inches(4), Inches(1.2),
         "Scan to open the\nWorkshop Worksheet\nor use the printed copy at your seat",
         font_size=14, color=RGBColor(0xFF, 0xCC, 0xCC), line_spacing=22)
add_logos_light(s)

# ══════════════════════════════════════════════════════════════
# SLIDE 10: Funnel Reference
# ══════════════════════════════════════════════════════════════
s = add_slide()
add_text(s, Inches(0.5), Inches(0.4), Inches(12.3), Inches(0.6),
         "Where Does Your Marketing Fit?", font_size=30, color=NAVY, bold=True, alignment=PP_ALIGN.CENTER)
add_text(s, Inches(0.5), Inches(0.95), Inches(12.3), Inches(0.4),
         "Map each activity to one of these four stages.", font_size=16, color=GRAY, alignment=PP_ALIGN.CENTER)

# Large centered funnel
funnel_cx = int(Inches(6.666))  # center of slide
funnel_top = int(Inches(1.6))
funnel_w = int(Inches(6.0))
funnel_h = int(Inches(5.0))
add_funnel(s, funnel_cx, funnel_top, funnel_w, funnel_h, labels=True)
add_logos_dark(s)

# ══════════════════════════════════════════════════════════════
# SLIDE 11: Step 1
# ══════════════════════════════════════════════════════════════
s = add_slide()
# Red border
border = add_rect(s, Inches(0.3), Inches(0.3), Inches(12.7), Inches(6.9),
                  border_color=RED, border_width=Pt(4))
add_text(s, Inches(1.2), Inches(1.0), Inches(11), Inches(0.7),
         "Step 1: Map What You\u2019re Doing", font_size=32, color=RED, bold=True)

add_bullet_list(s, Inches(1.2), Inches(2.2), Inches(10.5), Inches(3.5),
    [
        "Write down every marketing activity you're currently doing",
        "For each one, answer: WHY are you doing this?",
        "Then answer: WHERE does it fit in the funnel?",
        "Don't overthink it. Just get it all on paper.",
    ], font_size=20, line_spacing=34)
add_logos_dark(s)

# ══════════════════════════════════════════════════════════════
# SLIDE 12: Step 2
# ══════════════════════════════════════════════════════════════
s = add_slide()
border = add_rect(s, Inches(0.3), Inches(0.3), Inches(12.7), Inches(6.9),
                  border_color=RED, border_width=Pt(4))
add_text(s, Inches(1.2), Inches(1.0), Inches(11), Inches(0.7),
         "Step 2: Find the Gaps", font_size=32, color=RED, bold=True)

# Last bullet is red bold
bullets12 = [
    [("Which funnel stages have the most activity? Which have none?", 20, GRAY, False)],
    [("Heavy on Decision (LSA, PPC) but light on everything else?", 20, GRAY, False)],
    [("No Awareness activity at all?", 20, GRAY, False)],
    [('Nothing in Conversion beyond "we answer the phone"?', 20, GRAY, False)],
    [("Circle the biggest gap \u2014 that\u2019s your #1 priority", 20, RED, True)],
]
add_rich_bullet_list(s, Inches(1.2), Inches(2.0), Inches(10.5), Inches(4),
                     bullets12, font_size=20, line_spacing=34)
add_logos_dark(s)

# ══════════════════════════════════════════════════════════════
# SLIDE 13: Step 3
# ══════════════════════════════════════════════════════════════
s = add_slide()
border = add_rect(s, Inches(0.3), Inches(0.3), Inches(12.7), Inches(6.9),
                  border_color=RED, border_width=Pt(4))
add_text(s, Inches(1.2), Inches(1.0), Inches(11), Inches(0.7),
         "Step 3: Pick Your Next Move", font_size=32, color=RED, bold=True)

bullets13 = [
    [("Based on your gaps, choose ", 20, GRAY, False), ("1-3 things", 20, GRAY, True), (" you want to add to your marketing", 20, GRAY, False)],
    [("For each one, write down: ", 20, GRAY, False), ("What is it?", 20, GRAY, True)],
    [("Where does it fit in the funnel?", 20, GRAY, False)],
    [("Why this one?", 20, GRAY, False)],
    [("Don't worry about the \"how\" yet \u2014 ", 20, GRAY, False), ("that's what we're here for", 20, RED, True)],
]
add_rich_bullet_list(s, Inches(1.2), Inches(2.0), Inches(10.5), Inches(4),
                     bullets13, font_size=20, line_spacing=34)
add_logos_dark(s)

# ══════════════════════════════════════════════════════════════
# SLIDE 14: Close
# ══════════════════════════════════════════════════════════════
s = add_slide()
set_bg(s, NAVY)
# Logos centered at top
logo_y = Inches(0.8)
s.shapes.add_picture(RAINBOW_LOGO, Inches(4.2), logo_y, height=Inches(1.0))
s.shapes.add_picture(SOS_LOGO_WHITE, Inches(7.5), logo_y + Inches(0.2), height=Inches(0.65))

add_text(s, Inches(1), Inches(2.5), Inches(11.3), Inches(2.5),
         "Build the System.\nCover the Funnel.\nWin the Market.",
         font_size=40, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER, line_spacing=52)
add_text(s, Inches(1), Inches(5.2), Inches(11.3), Inches(0.6),
         "Authoritative content. Strong reviews. A brand worth citing.",
         font_size=18, color=RGBColor(0x99, 0x99, 0xAA), alignment=PP_ALIGN.CENTER)
# No footer logos on this slide

# ══════════════════════════════════════════════════════════════
# SLIDE 15: QR Codes
# ══════════════════════════════════════════════════════════════
s = add_slide()
set_bg(s, NAVY)
# Logos centered at top
s.shapes.add_picture(RAINBOW_LOGO, Inches(4.2), Inches(0.5), height=Inches(1.0))
s.shapes.add_picture(SOS_LOGO_WHITE, Inches(7.5), Inches(0.7), height=Inches(0.65))

add_text(s, Inches(1), Inches(1.8), Inches(11.3), Inches(0.7),
         "Scan Before You Leave", font_size=32, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)

# Three QR placeholders
qr_labels = ["Download the Report", "Workshop Worksheet", "Talk to an SOS\nMarketing Consultant"]
qr_size = Inches(2.0)
qr_y = Inches(3.0)
qr_gap = Inches(0.8)
total_qr_w = 3 * qr_size + 2 * qr_gap
qr_start_x = (SLIDE_W - total_qr_w) // 2

for i, label in enumerate(qr_labels):
    x = int(qr_start_x) + i * int(qr_size + qr_gap)
    # Rounded rect placeholder
    qr = add_rounded_rect(s, x, int(qr_y), int(qr_size), int(qr_size),
                          fill_color=RGBColor(0x2A, 0x40, 0x68),
                          border_color=RGBColor(0x4A, 0x60, 0x88))
    add_text(s, x, int(qr_y) + int(qr_size * 0.35), int(qr_size), Inches(0.4),
             "QR CODE", font_size=12, color=RGBColor(0x60, 0x70, 0x90), alignment=PP_ALIGN.CENTER)
    # Label below
    add_text(s, x - Inches(0.2), int(qr_y) + int(qr_size) + Inches(0.15), int(qr_size) + Inches(0.4), Inches(0.7),
             label, font_size=14, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)
# No footer logos on this slide

# ── Save ──
output = "/private/tmp/rainbow-conference-2026/Rainbow_Conference_2026_Media_Mix.pptx"
prs.save(output)
print(f"Saved {output}")
print(f"Total slides: {len(prs.slides)}")
