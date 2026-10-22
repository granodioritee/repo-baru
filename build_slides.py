"""Builds digital-marketing-tools.pptx (Types of Mobile Marketing and SEO slides).

Run: python3 build_slides.py
Icons come from assets/icons (regenerate with scripts/make_icons.js).
"""
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = Path(__file__).parent
ICONS = ROOT / "assets" / "icons"
OUTPUT = ROOT / "digital-marketing-tools.pptx"

FONT = "Calibri"
ACCENT = RGBColor(0x6D, 0x28, 0xD9)  # the one accent colour (violet)
TINT = RGBColor(0xF3, 0xEF, 0xFD)  # light tint of the accent for cards
TITLE = RGBColor(0x1F, 0x1B, 0x2E)
BODY = RGBColor(0x37, 0x41, 0x51)
MUTED = RGBColor(0x6B, 0x72, 0x80)
LINE = RGBColor(0xE5, 0xE7, 0xEB)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SLIDE_W, SLIDE_H = 13.333, 7.5
MARGIN = 0.6
GAP = 0.35
GRID_TOP = 1.45
GRID_BOTTOM = 7.0

TOOLS = [
    ("display", "Display Ads", "Visual ads on websites or apps.",
     "A banner showing a discount."),
    ("search", "Search Ads", "Ads appearing in search results.",
     "A restaurant ad on Google."),
    ("video", "Video Advertising", "Promotions through videos.",
     "A YouTube advertisement."),
    ("sms", "Text Messages", "Promotions sent through SMS.",
     "A discount code sent by text."),
    ("inapp", "In-app Advertising", "Ads inside mobile apps.",
     "An ad appearing in a mobile game."),
    ("qr", "QR Codes, Digital Watermarks, and 2D Barcodes",
     "Technologies connecting users to digital information.",
     "Scanning a QR code to register for an event."),
    ("geo", "Geo-targeting", "Targeting customers based on location.",
     "A café promotes offers to people nearby."),
]

NOTES_TOOLS_1 = [
    "Display ads: images, banners or short animations placed through ad networks like the "
    "Google Display Network. Mainly for brand awareness, and can retarget people who already "
    "visited the brand's website.",
    "Search ads: triggered by the keywords people type and labeled \"Sponsored\". Use "
    "pay-per-click (PPC), so the business only pays when someone clicks. They reach people who "
    "are already looking for the product.",
    "Video advertising: runs on YouTube, TikTok, Instagram Reels and streaming. Formats include "
    "skippable, non-skippable and 6-second bumper ads. Sound and storytelling make them more "
    "memorable.",
    "Text messages: very high open rate and good for short, time-sensitive offers. Customers "
    "must opt in and be able to opt out.",
]

NOTES_TOOLS_2 = [
    "In-app advertising: banners, full-screen (interstitial) ads and rewarded ads, where users "
    "get something in return like coins or extra lives.",
    "QR codes, digital watermarks and 2D barcodes: QR codes open a website, form or payment "
    "page; digital watermarks are invisible codes hidden in images or packaging; 2D barcodes "
    "store more information than normal barcodes. Scans can be tracked.",
    "Geo-targeting: uses GPS or IP address to show location-based ads. Geofencing triggers an "
    "ad when someone enters a set area. Great for local businesses.",
]

# (indent level, bold label, text)
SEO_BULLETS = [
    (0, "", "SEO is the process of increasing the likelihood that a website appears in organic "
     "(unpaid) search engine results, by improving the website's content and structure"),
    (0, "Key elements:", ""),
    (1, "Keywords: ", "words people type into search engines"),
    (1, "Quality content: ", "useful and relevant information for readers"),
    (1, "Fast, mobile-friendly website: ", "loads quickly and is easy to use on phones"),
    (1, "Backlinks: ", "links from other trusted websites"),
    (0, "Example: ", "A shoe store optimizes its website so customers can find it when "
     "searching for running shoes on Google"),
]

NOTES_SEO = (
    "SEO improves a website so it ranks higher in organic search results. Key elements are "
    "using relevant keywords in titles, headings and content, publishing useful content, making "
    "the site fast and mobile-friendly, and getting backlinks from trusted websites. SEO "
    "does not cost money for ads and its results last a long time, but it takes time to work. "
    "For example, a shoe store could write a blog post like \"Best "
    "Running Shoes for Beginners\" and use \"running shoes\" in its page titles so it shows up "
    "when people search for running shoes."
)


def style_run(run, size, color=BODY, bold=False, italic=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic


def add_text(slide, x, y, w, h, runs, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT, name=None):
    """runs: list of (text, size, color, bold, italic) in one paragraph."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        box.name = name
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    for text, size, color, bold, italic in runs:
        style_run(p.add_run(), size, color, bold, italic)
        p.runs[-1].text = text
    return box


def add_box(slide, shape, x, y, w, h, fill, line=None, name=None):
    shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        shp.name = name
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(1)
    style = shp._element.find(qn("p:style"))  # drop theme style so no default shadow is applied
    if style is not None:
        shp._element.remove(style)
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        shp.adjustments[0] = min(0.5, 0.15 / min(w, h))  # ~0.15" corner radius
    return shp


def add_icon(slide, icon, x, y, d):
    add_box(slide, MSO_SHAPE.OVAL, x, y, d, d, ACCENT, name=f"{icon} icon circle")
    pad = d * 0.25
    pic = slide.shapes.add_picture(str(ICONS / f"{icon}.png"), Inches(x + pad), Inches(y + pad),
                                   Inches(d - 2 * pad), Inches(d - 2 * pad))
    pic.name = f"{icon} icon"


def add_title(slide, text):
    add_text(slide, MARGIN, 0.45, SLIDE_W - 2 * MARGIN, 0.8,
             [(text, 36, TITLE, True, False)], anchor=MSO_ANCHOR.MIDDLE, name="Title")


def add_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def add_example(slide, x, y, w, h, example):
    add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, WHITE, name="Example box")
    add_text(slide, x + 0.2, y, w - 0.4, h,
             [("Example: ", 14, ACCENT, True, False), (example, 14, BODY, False, True)],
             anchor=MSO_ANCHOR.MIDDLE, name="Example")


def blank_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = WHITE
    return slide


def tools_grid_slide(prs, tools, notes):
    """Tools 1 to 4 as a 2 x 2 grid of cards."""
    slide = blank_slide(prs)
    add_title(slide, "Types of Mobile Marketing (1/2)")
    card_w = (SLIDE_W - 2 * MARGIN - GAP) / 2
    card_h = (GRID_BOTTOM - GRID_TOP - GAP) / 2
    for i, (icon, name, definition, example) in enumerate(tools):
        x = MARGIN + (i % 2) * (card_w + GAP)
        y = GRID_TOP + (i // 2) * (card_h + GAP)
        add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w, card_h, TINT, name=f"{name} card")
        add_icon(slide, icon, x + 0.3, y + 0.3, 0.75)
        add_text(slide, x + 1.25, y + 0.3, card_w - 1.55, 0.75,
                 [(name, 22, TITLE, True, False)], anchor=MSO_ANCHOR.MIDDLE, name=f"{name} name")
        add_text(slide, x + 0.3, y + 1.2, card_w - 0.6, 0.45,
                 [(definition, 17, BODY, False, False)], name=f"{name} definition")
        add_example(slide, x + 0.3, y + 1.75, card_w - 0.6, 0.6, example)
    add_notes(slide, "\n\n".join(notes))


def tools_column_slide(prs, tools, notes):
    """Tools 5 to 7 as three tall column cards."""
    slide = blank_slide(prs)
    add_title(slide, "Types of Mobile Marketing (2/2)")
    card_w = (SLIDE_W - 2 * MARGIN - 2 * GAP) / 3
    card_h = GRID_BOTTOM - GRID_TOP
    for i, (icon, name, definition, example) in enumerate(tools):
        x = MARGIN + i * (card_w + GAP)
        y = GRID_TOP
        inner_w = card_w - 0.7
        add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w, card_h, TINT, name=f"{name} card")
        add_icon(slide, icon, x + 0.35, y + 0.4, 0.95)
        add_text(slide, x + 0.35, y + 1.55, inner_w, 1.2,
                 [(name, 22, TITLE, True, False)], anchor=MSO_ANCHOR.BOTTOM, name=f"{name} name")
        add_text(slide, x + 0.35, y + 2.9, inner_w, 1.0,
                 [(definition, 17, BODY, False, False)], name=f"{name} definition")
        add_example(slide, x + 0.35, y + card_h - 1.45, inner_w, 1.1, example)
    add_notes(slide, "\n\n".join(notes))


def seo_slide(prs):
    slide = blank_slide(prs)
    add_title(slide, "Search Engine Optimization (SEO)")

    # Left: bullets
    box = slide.shapes.add_textbox(Inches(MARGIN), Inches(GRID_TOP + 0.1), Inches(6.1), Inches(5.3))
    box.name = "SEO bullets"
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, (level, label, text) in enumerate(SEO_BULLETS):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        size = 20 if level == 0 else 17
        nxt = SEO_BULLETS[i + 1][0] if i + 1 < len(SEO_BULLETS) else 0
        p.space_after = Pt(6 if nxt == 1 else 18)
        indent = 0.3 + 0.35 * level
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(Inches(indent)))
        pPr.set("indent", str(-Inches(0.3)))
        bu_clr = pPr.makeelement("{http://schemas.openxmlformats.org/drawingml/2006/main}buClr", {})
        srgb = bu_clr.makeelement("{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr",
                                  {"val": "6D28D9"})
        bu_clr.append(srgb)
        pPr.append(bu_clr)
        bu = pPr.makeelement("{http://schemas.openxmlformats.org/drawingml/2006/main}buChar",
                             {"char": "\u2022" if level == 0 else "\u25E6"})
        pPr.append(bu)
        is_example = label.startswith("Example")
        if label:
            style_run(p.add_run(), size, ACCENT if is_example else TITLE, bold=True)
            p.runs[-1].text = label
        if text:
            style_run(p.add_run(), size, BODY, italic=is_example)
            p.runs[-1].text = text

    # Right: a simple search results mock-up (paid vs organic)
    px, py, pw, ph = 7.2, GRID_TOP, SLIDE_W - MARGIN - 7.2, GRID_BOTTOM - GRID_TOP
    add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, px, py, pw, ph, TINT, name="Search results panel")

    ix, iw = px + 0.35, pw - 0.7
    add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, ix, py + 0.35, iw, 0.6, WHITE, line=LINE,
            name="Search bar").adjustments[0] = 0.5
    add_text(slide, ix + 0.3, py + 0.35, iw - 1.2, 0.6, [("running shoes", 16, TITLE, False, False)],
             anchor=MSO_ANCHOR.MIDDLE, name="Search query")
    add_icon(slide, "seo", ix + iw - 0.55, py + 0.42, 0.46)

    # The shoe store's page ranks first thanks to SEO; the other results are plain
    results = [
        ("stridestore.com", "Best Running Shoes for Beginners", True),
        ("shoereviews.com", "Top 10 Running Shoes This Year", False),
        ("footwearhub.com", "Running Shoes Buying Guide", False),
    ]
    ry = py + 1.25
    rh = 1.2
    for label, headline, highlight in results:
        add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, ix, ry, iw, rh, WHITE,
                line=ACCENT if highlight else LINE, name=f"{label} result")
        add_text(slide, ix + 0.25, ry + 0.15, iw - 2.6, 0.3,
                 [(label, 12, MUTED, False, False)], name="Result source")
        if highlight:
            chip_w = 2.1
            chip = add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, ix + iw - chip_w - 0.2, ry + 0.15,
                           chip_w, 0.32, ACCENT, name="Top result chip")
            chip.adjustments[0] = 0.5
            add_text(slide, ix + iw - chip_w - 0.2, ry + 0.15, chip_w, 0.32,
                     [("#1 thanks to SEO", 12, WHITE, True, False)],
                     anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, name="Top result label")
        add_text(slide, ix + 0.25, ry + 0.52, iw - 0.5, 0.35,
                 [(headline, 16, ACCENT if highlight else TITLE, True, False)], name="Result headline")
        add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, ix + 0.25, ry + 0.93, iw * 0.65, 0.1, LINE,
                name="Result snippet").adjustments[0] = 0.5
        ry += rh + 0.2

    add_notes(slide, NOTES_SEO)


def main():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    tools_grid_slide(prs, TOOLS[:4], NOTES_TOOLS_1)
    tools_column_slide(prs, TOOLS[4:], NOTES_TOOLS_2)
    seo_slide(prs)
    prs.save(OUTPUT)
    print(f"Saved {OUTPUT}")


if __name__ == "__main__":
    main()
