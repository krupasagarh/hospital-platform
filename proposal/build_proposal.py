#!/usr/bin/env python3
"""Build the Tiptur Health Centre proposal as a PowerPoint deck and a Markdown leave-behind.

Prices live in this file. Run it again after a change:

    python3 proposal/build_proposal.py
"""

from __future__ import annotations

from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parent
PPTX_PATH = ROOT / "Tiptur-Health-Centre-proposal.pptx"
MD_PATH = ROOT / "Tiptur-Health-Centre-proposal.md"

HOSPITAL = "Tiptur Health Centre"
PLACE = "Krishna Raja Extension, Tiptur"
MONTH = "November 2026"
FONT = "Calibri"

# One-time tasks
SOFTWARE = 45_000
WEBSITE = 28_000
PRESENCE = 12_000
ADS_SETUP = 8_000
LIST_TOTAL = SOFTWARE + WEBSITE + PRESENCE + ADS_SETUP
LAUNCH = 78_000
SAVE = LIST_TOTAL - LAUNCH
HALF = LAUNCH // 2

# Monthly care
CARE_SOFT = 4_500
CARE_MKT = 9_000
CARE_LIST = CARE_SOFT + CARE_MKT
CARE = 12_000

# Paid by the hospital to the ad platforms
AD_START = 10_000
AD_LOW = 8_000
AD_HIGH = 15_000

MONTHS = 3
CARE_3 = CARE * MONTHS
AD_3 = AD_START * MONTHS
TO_US = LAUNCH + CARE_3
ALL_IN = TO_US + AD_3

# Later, only if they ask
WHATSAPP = 15_000
BEDS = 35_000
PHARMACY = 40_000
OT = 25_000
AYUSHMAN = 30_000
KANNADA_UI = 18_000

assert LIST_TOTAL == 93_000
assert SAVE == 15_000
assert LAUNCH % 2 == 0
assert CARE_LIST == 13_500
assert CARE_3 == 36_000
assert AD_3 == 30_000
assert TO_US == 114_000
assert ALL_IN == 144_000

# Palette
BG = RGBColor(0xF3, 0xEF, 0xE6)
INK = RGBColor(0x1B, 0x2A, 0x27)
TEAL = RGBColor(0x0F, 0x6E, 0x64)
DEEP = RGBColor(0x0B, 0x3D, 0x38)
WHITE = RGBColor(0xFF, 0xFC, 0xF8)
SAGE = RGBColor(0xE5, 0xF0, 0xEA)
MUTED = RGBColor(0x5E, 0x6F, 0x69)
GOLD = RGBColor(0xC6, 0xA1, 0x5A)
SAND = RGBColor(0xF8, 0xEB, 0xD9)
LINE = RGBColor(0xD5, 0xDD, 0xD6)
CREAM = RGBColor(0xF7, 0xF3, 0xEA)

SLIDE_W = 13.333
SLIDE_H = 7.5
TOTAL = 13


def inr(amount: int) -> str:
    sign = "-" if amount < 0 else ""
    raw = str(abs(amount))
    if len(raw) <= 3:
        body = raw
    else:
        head, tail = raw[:-3], raw[-3:]
        groups = []
        while head:
            groups.append(head[-2:])
            head = head[:-2]
        body = ",".join(reversed(groups)) + "," + tail
    return f"{sign}₹{body}"


def style_run(run, size, bold, color, italic=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    r_pr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:cs", "a:ea"):
        el = r_pr.find(qn(tag))
        if el is None:
            el = etree.SubElement(r_pr, qn(tag))
        el.set("typeface", FONT)


def write(
    slide,
    x,
    y,
    w,
    h,
    content,
    size=18,
    bold=False,
    color=INK,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    spacing=1.0,
    after=2,
):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = Inches(0.02)
    tf.margin_right = Inches(0.02)
    tf.margin_top = Inches(0.0)
    tf.margin_bottom = Inches(0.0)
    anchor_map = {
        MSO_ANCHOR.TOP: "t",
        MSO_ANCHOR.MIDDLE: "ctr",
        MSO_ANCHOR.BOTTOM: "b",
    }
    tf._txBody.bodyPr.set("anchor", anchor_map.get(anchor, "t"))

    if isinstance(content, str):
        paragraphs = [content]
    else:
        paragraphs = content

    for i, para in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_before = Pt(0)
        p.space_after = Pt(after)
        p.line_spacing = spacing
        if isinstance(para, str):
            runs = [(para, bold, color, size, False)]
        else:
            runs = []
            for item in para:
                text, b, c, s = item[:4]
                italic = item[4] if len(item) > 4 else False
                runs.append((text, b, c, s, italic))
        for text, b, c, s, italic in runs:
            run = p.add_run()
            run.text = text
            style_run(run, s, b, c, italic)
    return box


def add_rect(slide, x, y, w, h, fill, line=None, radius=0.1, rounded=True):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(1.0)
    if rounded:
        try:
            shape.adjustments[0] = radius
        except Exception:
            pass
    return shape


def paint(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def new_slide(prs, color):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    paint(slide, color)
    return slide


def footer(slide, number):
    add_rect(slide, 0.58, 7.08, 12.17, 0.012, LINE, rounded=False)
    write(slide, 0.58, 7.14, 8.2, 0.28, HOSPITAL, size=12, color=MUTED, anchor=MSO_ANCHOR.MIDDLE)
    write(
        slide,
        9.4,
        7.14,
        3.35,
        0.28,
        f"{number}   /   {TOTAL}",
        size=12,
        color=MUTED,
        align=PP_ALIGN.RIGHT,
        anchor=MSO_ANCHOR.MIDDLE,
    )


def heading(slide, text, y=0.32):
    write(slide, 0.58, y, 12.1, 0.58, text, size=30, bold=True, color=DEEP)
    add_rect(slide, 0.60, y + 0.62, 1.05, 0.045, TEAL, rounded=False)


def bullet_block(slide, x, y, w, h, items, size=16, color=INK):
    paragraphs = []
    for item in items:
        paragraphs.append([(f"•  {item}", False, color, size, False)])
    return write(slide, x, y, w, h, paragraphs, size=size, color=color, after=8, spacing=1.05)


def build_pptx():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    prs.core_properties.title = f"Proposal for {HOSPITAL}"
    prs.core_properties.subject = "Custom hospital software, website, and online marketing"
    prs.core_properties.category = "Proposal"
    blank = prs.slide_layouts[6]
    if blank.name != "Blank":
        raise RuntimeError(f"Expected a blank layout, found {blank.name!r}")

    cover(prs)
    three_things(prs)
    desk(prs)
    personalized(prs)
    website(prs)
    marketing(prs)
    path(prs)
    prices_once(prs)
    prices_month(prs)
    three_months(prs)
    start(prs)
    later(prs)
    close(prs)

    if len(prs.slides) != TOTAL:
        raise RuntimeError(f"Expected {TOTAL} slides, built {len(prs.slides)}")
    prs.save(PPTX_PATH)


def cover(prs):
    slide = new_slide(prs, DEEP)
    add_rect(slide, 0, 0, 0.16, SLIDE_H, GOLD, rounded=False)
    write(slide, 0.72, 0.48, 11.5, 0.36, "FOR THE DOCTORS AND STAFF", size=14, bold=True, color=GOLD)
    write(slide, 0.70, 1.15, 12.0, 1.35, HOSPITAL, size=54, bold=True, color=CREAM)
    add_rect(slide, 0.74, 2.62, 1.35, 0.045, GOLD, rounded=False)
    write(
        slide,
        0.70,
        2.95,
        11.2,
        2.15,
        [
            [("Your own hospital software.", True, CREAM, 26, False)],
            [("Your own website, built and handed to you.", True, CREAM, 26, False)],
            [("Online marketing and social media ads. We will take care.", True, CREAM, 26, False)],
        ],
        after=10,
    )
    write(slide, 0.72, 6.45, 8.5, 0.55, f"{PLACE}\n{MONTH}", size=16, color=RGBColor(0xD5, 0xE4, 0xDF), spacing=1.15)
    notes(
        slide,
        "Open with their name. Say this sitting is for Tiptur Health Centre: their doctors, their staff, their patients. "
        "Then read the three lines slowly. Do not open the laptop on the software yet.",
    )


def three_things(prs):
    slide = new_slide(prs, BG)
    heading(slide, f"What we will do for {HOSPITAL}")
    write(
        slide,
        0.58,
        1.18,
        12.0,
        0.45,
        "Three things, in this hospital’s name. We set them up, and we look after them.",
        size=18,
        color=MUTED,
    )
    cards = [
        (
            "01",
            "Your software",
            "Customized for this hospital. Your doctors, your fees, your staff, your bills. Patient records stay here.",
        ),
        (
            "02",
            "Your website",
            "A site in the hospital’s name, in English and Kannada. We build it, and we give it to you.",
        ),
        (
            "03",
            "Your reach",
            "Google, Facebook, Instagram, and the ads. We set them up, we link them to you, and we run them.",
        ),
    ]
    x = 0.58
    for number, title, body in cards:
        add_rect(slide, x, 1.90, 3.90, 4.35, WHITE, line=LINE, radius=0.08)
        add_rect(slide, x, 1.90, 3.90, 0.10, TEAL, rounded=False)
        write(slide, x + 0.28, 2.25, 3.3, 0.45, number, size=18, bold=True, color=GOLD)
        write(slide, x + 0.28, 2.80, 3.3, 0.9, title, size=26, bold=True, color=DEEP)
        write(slide, x + 0.28, 3.90, 3.35, 1.9, body, size=16, color=INK, spacing=1.15)
        x += 4.12
    footer(slide, 2)
    notes(
        slide,
        "Stay on this slide until someone nods. If a doctor asks “what do I have to do?”, say: you see patients. "
        "We set up the desk, the site, and the ads. Reception uses the desk. You get a monthly report.",
    )


def desk(prs):
    slide = new_slide(prs, BG)
    heading(slide, "The hospital, on one screen")
    columns = [
        (
            "Reception",
            [
                "Appointment, walk-in, or emergency",
                "A token queue for each doctor",
                "Emergencies called first",
                "Report saved on that patient",
            ],
        ),
        (
            "Doctor",
            [
                "Today’s patients at this hospital",
                "Only the patients under your care",
                "Order a test or a scan",
                "Your fees, and only your fees",
            ],
        ),
        (
            "Owner",
            [
                "The day’s collection, cash and UPI",
                "What is still due",
                "Doctor’s share and hospital’s share",
                "A discount needs your approval",
            ],
        ),
    ]
    x = 0.58
    for title, items in columns:
        add_rect(slide, x, 1.22, 3.90, 4.55, WHITE, line=LINE, radius=0.08)
        add_rect(slide, x + 0.28, 1.48, 0.42, 0.08, GOLD, rounded=False)
        write(slide, x + 0.28, 1.68, 3.35, 0.5, title, size=22, bold=True, color=DEEP)
        bullet_block(slide, x + 0.28, 2.40, 3.35, 3.05, items, size=16)
        x += 4.12
    add_rect(slide, 0.58, 5.95, 12.17, 0.90, DEEP, radius=0.08)
    write(
        slide,
        0.85,
        6.10,
        11.6,
        0.62,
        "It runs on a computer already in the hospital. Reception and the doctor open it on a phone connected to the hospital Wi-Fi.",
        size=16,
        color=CREAM,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    footer(slide, 3)
    notes(
        slide,
        "Walk one patient across the three columns: Geetha-style visit is fine as a story, but use a made-up patient, not a real one. "
        "Register at reception, token, doctor opens the record, a report is attached, the owner sees the bill. "
        "If a doctor asks to open this from home, say the records stay on the hospital network. Home access is a later conversation, and it is not in this price.",
    )


def personalized(prs):
    slide = new_slide(prs, BG)
    heading(slide, "Customized for this hospital")
    add_rect(slide, 0.58, 1.22, 7.35, 5.55, WHITE, line=LINE, radius=0.08)
    write(slide, 0.88, 1.48, 6.7, 0.45, "Set up in your name", size=20, bold=True, color=TEAL)
    bullet_block(
        slide,
        0.88,
        2.15,
        6.7,
        4.2,
        [
            "The hospital name on every screen",
            "Each doctor’s role, consultation fee, and share",
            "Logins for the owner, reception, billing, and lab",
            "Your test and scan list, with your prices",
            "One running bill for the visit, cash or UPI",
            "Two training sittings with your staff, at the hospital",
        ],
        size=17,
    )
    add_rect(slide, 8.15, 1.22, 4.60, 5.55, DEEP, radius=0.08)
    write(slide, 8.45, 1.55, 4.05, 0.8, "What stays here", size=20, bold=True, color=GOLD)
    write(
        slide,
        8.45,
        2.55,
        4.05,
        3.7,
        [
            [("Patient records stay on the hospital computer.", False, CREAM, 16, False)],
            [("A doctor sees only the patients under their care at Tiptur Health Centre.", False, CREAM, 16, False)],
            [("Reception can take a payment. The fee split is for the owner.", False, CREAM, 16, False)],
        ],
        after=16,
        spacing=1.15,
    )
    footer(slide, 4)
    notes(
        slide,
        "Ask for the doctor list and today’s consultation fees before you leave, even if they have not said yes. "
        "A doctor who also works at another hospital keeps one login. At this hospital they see only their patients here, and only their fee here. Say that if a visiting doctor is in the room.",
    )


def website(prs):
    slide = new_slide(prs, BG)
    heading(slide, "A website we build and give you")
    pages = [
        ("Home", "Phone, map, hours, and a WhatsApp button."),
        ("Doctors", "Names, specialties, and the days they are here."),
        ("Services", "What the hospital actually does, in plain words."),
        ("Visit", "An appointment request that reaches reception."),
        ("Language", "The same pages in English and in Kannada."),
    ]
    y = 1.22
    for title, body in pages:
        add_rect(slide, 0.58, y, 7.55, 0.98, WHITE, line=LINE, radius=0.12)
        write(slide, 0.82, y + 0.14, 2.1, 0.7, title, size=18, bold=True, color=TEAL, anchor=MSO_ANCHOR.MIDDLE)
        write(slide, 2.95, y + 0.14, 4.9, 0.7, body, size=16, color=INK, anchor=MSO_ANCHOR.MIDDLE)
        y += 1.08
    add_rect(slide, 8.35, 1.22, 4.40, 5.40, SAND, radius=0.08)
    write(slide, 8.65, 1.55, 3.85, 0.9, "The site is yours", size=22, bold=True, color=DEEP)
    write(
        slide,
        8.65,
        2.65,
        3.85,
        3.5,
        [
            [("We hand it over to the hospital.", False, INK, 16, False)],
            [("The domain is registered in the hospital’s name.", False, INK, 16, False)],
            [("A patient can call, ask for directions, or message you from their phone.", False, INK, 16, False)],
        ],
        after=16,
        spacing=1.12,
    )
    footer(slide, 5)
    notes(
        slide,
        "Say the website is the public face. The software is the desk. They are linked: an appointment request from the site lands with reception. "
        "Domain and yearly hosting are billed at cost, in their name, and are outside the ₹28,000. Usually a few thousand rupees a year.",
    )


def marketing(prs):
    slide = new_slide(prs, BG)
    heading(slide, "Marketing and ads. We will take care.")
    cards = [
        ("Google", "The hospital on Google Maps and Search. Hours, phone, photos, and the way to Krishna Raja Extension."),
        ("Facebook and Instagram", "Pages in the hospital’s name. Eight posts a month, in Kannada and English. We design them and we publish them."),
        ("Ads", "A search ad for people looking for a hospital in Tiptur. A Facebook and Instagram ad for families nearby. We write them and we adjust them."),
        ("A monthly report", "Calls, direction requests, messages, the amount spent, and what we will change next month."),
    ]
    positions = [(0.58, 1.22), (6.68, 1.22), (0.58, 3.95), (6.68, 3.95)]
    for (title, body), (x, y) in zip(cards, positions):
        add_rect(slide, x, y, 5.90, 2.55, WHITE, line=LINE, radius=0.08)
        write(slide, x + 0.28, y + 0.22, 5.35, 0.5, title, size=20, bold=True, color=DEEP)
        write(slide, x + 0.28, y + 0.85, 5.35, 1.45, body, size=15, color=INK, spacing=1.08)
    footer(slide, 6)
    notes(
        slide,
        "This is the slide to slow down on. Say: you do not have to learn the ad screens. We run them. "
        "The ad money itself is paid to Google and Meta from the hospital account. Our fee is for the work. "
        "Eight posts a month are designed from photographs they give us. A camera visit or a reel shoot is a separate conversation. YouTube is not in this fee.",
    )


def path(prs):
    slide = new_slide(prs, BG)
    heading(slide, "From a phone search to your desk")
    write(
        slide,
        0.58,
        1.18,
        12.1,
        0.5,
        "The website, the ads, and the software are linked. A patient moves along one path.",
        size=18,
        color=MUTED,
    )
    steps = [
        ("1", "They search", "Hospital in Tiptur, or a doctor they already know."),
        ("2", "They find you", "Your Google listing, your ad, or your website."),
        ("3", "They reach you", "A call, a direction request, or WhatsApp."),
        ("4", "Reception books", "The visit is entered in your software."),
        ("5", "The doctor sees it", "The patient is on today’s list."),
    ]
    x = 0.48
    for number, title, body in steps:
        add_rect(slide, x, 2.05, 2.38, 3.55, WHITE, line=LINE, radius=0.1)
        oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.90), Inches(2.32), Inches(0.52), Inches(0.52))
        oval.fill.solid()
        oval.fill.fore_color.rgb = TEAL
        oval.line.fill.background()
        write(
            slide,
            x + 0.90,
            2.38,
            0.52,
            0.52,
            number,
            size=16,
            bold=True,
            color=CREAM,
            align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        write(slide, x + 0.14, 3.05, 2.10, 0.85, title, size=16, bold=True, color=DEEP, align=PP_ALIGN.CENTER)
        write(slide, x + 0.16, 3.90, 2.06, 1.4, body, size=13, color=INK, align=PP_ALIGN.CENTER, spacing=1.05)
        x += 2.55
    footer(slide, 7)
    notes(
        slide,
        "Trace the path with a finger. Stop on step 4 and say the enquiry does not sit in a notebook. It becomes a visit on the token list. "
        "The WhatsApp here is the hospital’s business chat, the button on the website and the ads. Automatic messages to the patient, such as “doctor is late” or “report is ready”, are a separate setup on a later slide.",
    )


def prices_once(prs):
    slide = new_slide(prs, BG)
    heading(slide, "Each task, and the launch price")
    rows = [
        ("Hospital software, customized and set up", "Your name, doctors, fees, tests, staff logins, and two trainings.", SOFTWARE),
        ("Website, built and handed over", "English and Kannada. Doctors, services, map, call, and WhatsApp.", WEBSITE),
        ("Google, Facebook, and Instagram", "Pages created in the hospital’s name and linked to the website.", PRESENCE),
        ("First ad campaigns set up", "One search campaign and one social campaign, ready to run.", ADS_SETUP),
    ]
    y = 1.18
    for i, (title, scope, amount) in enumerate(rows):
        fill = WHITE if i % 2 == 0 else SAGE
        add_rect(slide, 0.58, y, 12.17, 0.92, fill, line=LINE, radius=0.08)
        write(slide, 0.82, y + 0.10, 8.6, 0.38, title, size=18, bold=True, color=DEEP)
        write(slide, 0.82, y + 0.46, 8.6, 0.36, scope, size=13, color=MUTED)
        write(
            slide,
            9.40,
            y + 0.18,
            3.05,
            0.56,
            inr(amount),
            size=22,
            bold=True,
            color=TEAL,
            align=PP_ALIGN.RIGHT,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        y += 1.00
    add_rect(slide, 0.58, 5.28, 12.17, 1.55, DEEP, radius=0.08)
    write(slide, 0.86, 5.46, 7.4, 0.4, f"Launch pack for {HOSPITAL}", size=18, bold=True, color=GOLD)
    write(
        slide,
        0.86,
        5.95,
        7.5,
        0.65,
        f"All four tasks together. Taken one by one they are {inr(LIST_TOTAL)}. You save {inr(SAVE)}.",
        size=15,
        color=CREAM,
    )
    write(
        slide,
        8.35,
        5.55,
        4.1,
        1.05,
        inr(LAUNCH),
        size=36,
        bold=True,
        color=CREAM,
        align=PP_ALIGN.RIGHT,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    footer(slide, 8)
    notes(
        slide,
        f"Say {inr(LAUNCH)} first, then show the four lines so they can see the breakup. "
        "Domain and yearly hosting are extra, at actual cost, registered in the hospital’s name. Add GST if it applies. "
        "Beds, pharmacy, operation theatre, Ayushman, and patient WhatsApp messages are not inside this number. They are on the later slide.",
    )


def prices_month(prs):
    slide = new_slide(prs, BG)
    heading(slide, "Every month, we take care")
    add_rect(slide, 0.58, 1.22, 7.55, 2.35, WHITE, line=LINE, radius=0.08)
    write(slide, 0.86, 1.40, 4.6, 0.4, "Software and website", size=18, bold=True, color=DEEP)
    write(
        slide,
        0.86,
        1.95,
        4.25,
        1.35,
        "A new doctor, a fee, a test price, a visiting day. Help when the desk is stuck. Small changes on the website.",
        size=15,
        color=INK,
    )
    write(slide, 5.30, 1.48, 2.55, 0.62, inr(CARE_SOFT), size=26, bold=True, color=TEAL, align=PP_ALIGN.RIGHT)
    write(slide, 5.30, 2.12, 2.55, 0.32, "a month", size=13, color=MUTED, align=PP_ALIGN.RIGHT)

    add_rect(slide, 0.58, 3.75, 7.55, 2.35, WHITE, line=LINE, radius=0.08)
    write(slide, 0.86, 3.93, 4.8, 0.4, "Posts and the ads", size=18, bold=True, color=DEEP)
    write(
        slide,
        0.86,
        4.48,
        4.25,
        1.35,
        "Eight posts, published. Replies on the pages. The campaigns kept running. One report a month.",
        size=15,
        color=INK,
    )
    write(slide, 5.30, 4.00, 2.55, 0.62, inr(CARE_MKT), size=26, bold=True, color=TEAL, align=PP_ALIGN.RIGHT)
    write(slide, 5.30, 4.64, 2.55, 0.32, "a month", size=13, color=MUTED, align=PP_ALIGN.RIGHT)

    add_rect(slide, 8.35, 1.22, 4.40, 4.88, DEEP, radius=0.08)
    write(slide, 8.62, 1.48, 3.9, 0.7, "Both together", size=16, bold=True, color=GOLD)
    write(slide, 8.62, 2.15, 3.9, 0.85, inr(CARE), size=36, bold=True, color=CREAM)
    write(slide, 8.62, 3.05, 3.9, 0.4, "a month", size=16, color=CREAM)
    write(
        slide,
        8.62,
        3.60,
        3.9,
        1.15,
        f"Taken separately, these two are {inr(CARE_LIST)} a month.",
        size=15,
        color=CREAM,
        spacing=1.08,
    )
    write(
        slide,
        8.62,
        4.90,
        3.9,
        0.95,
        "One month’s notice to stop. The software and the website stay with the hospital.",
        size=14,
        color=CREAM,
        spacing=1.08,
    )
    footer(slide, 9)
    notes(
        slide,
        f"Offer the combined {inr(CARE)} a month, not the two lines, unless they only want one. "
        "The monthly care can stop with one month’s notice. The software and the website stay with the hospital. "
        "This slide is our fee. The next numbers are the ad budget, which is their money.",
    )


def three_months(prs):
    slide = new_slide(prs, BG)
    heading(slide, "The first three months")
    write(
        slide,
        0.58,
        1.08,
        12.1,
        0.40,
        f"{inr(ALL_IN)} in all. {inr(TO_US)} is the work. {inr(AD_3)} is paid to Google and Meta.",
        size=16,
        color=MUTED,
    )
    blocks = [
        (SAGE, TEAL, inr(LAUNCH), "Launch pack", "Paid once, to us"),
        (SAGE, TEAL, inr(CARE_3), "Three months of care", f"{inr(CARE)} × {MONTHS}, to us"),
        (SAND, DEEP, inr(AD_3), "Ad budget", "Paid to Google and Meta"),
    ]
    x = 0.58
    for fill, num_color, amount, title, subtitle in blocks:
        add_rect(slide, x, 1.62, 3.90, 2.28, fill, line=LINE, radius=0.08)
        write(slide, x + 0.22, 1.76, 3.46, 0.68, amount, size=28, bold=True, color=num_color, align=PP_ALIGN.CENTER)
        write(slide, x + 0.22, 2.48, 3.46, 0.40, title, size=16, bold=True, color=DEEP, align=PP_ALIGN.CENTER)
        write(slide, x + 0.22, 2.92, 3.46, 0.70, subtitle, size=14, color=MUTED, align=PP_ALIGN.CENTER)
        x += 4.12

    add_rect(slide, 0.58, 4.12, 6.00, 2.62, WHITE, line=LINE, radius=0.08)
    write(slide, 0.86, 4.30, 5.45, 0.36, "To us", size=14, bold=True, color=MUTED)
    write(slide, 0.86, 4.68, 5.45, 0.68, inr(TO_US), size=32, bold=True, color=TEAL)
    write(slide, 0.86, 5.45, 5.45, 0.70, "Launch pack plus three months of care.", size=15, color=INK)
    write(slide, 0.86, 6.15, 5.45, 0.36, "GST is extra, if it applies.", size=13, color=MUTED)

    add_rect(slide, 6.78, 4.12, 5.97, 2.62, SAND, radius=0.08)
    write(slide, 7.05, 4.30, 5.45, 0.36, "Suggested ad budget", size=14, bold=True, color=MUTED)
    write(slide, 7.05, 4.72, 5.45, 0.62, f"{inr(AD_START)} a month", size=26, bold=True, color=DEEP)
    write(
        slide,
        7.05,
        5.45,
        5.4,
        1.05,
        f"A sensible range is {inr(AD_LOW)} to {inr(AD_HIGH)}. This money goes to Google and Meta. We add nothing on top.",
        size=14,
        color=INK,
    )
    footer(slide, 10)
    notes(
        slide,
        f"Write the two piles in the air if you need to: {inr(TO_US)} for the work, {inr(AD_3)} for the platforms, {inr(ALL_IN)} moving in three months. "
        "Say the ad budget can start lower, at ₹8,000, if they want to try for a month. After three months, care and ads continue only if they want them. "
        "Recommend they pay Google and Meta from their own card, so the account is theirs.",
    )


def start(prs):
    slide = new_slide(prs, BG)
    heading(slide, "How we begin, and how you pay")
    steps = [
        ("1", "We sit at the hospital", "Doctors, fees, services, photographs, and the phone patients should call."),
        ("2", "The desk goes live", "Software on your computer. Two trainings for reception, billing, and the doctors."),
        ("3", "The public side goes live", "Website, Google, Facebook, Instagram, then the ads."),
    ]
    y = 1.20
    for number, title, body in steps:
        oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.62), Inches(y), Inches(0.46), Inches(0.46))
        oval.fill.solid()
        oval.fill.fore_color.rgb = TEAL
        oval.line.fill.background()
        write(slide, 0.62, y, 0.46, 0.46, number, size=16, bold=True, color=CREAM, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        write(slide, 1.28, y - 0.02, 6.3, 0.38, title, size=18, bold=True, color=DEEP)
        write(slide, 1.28, y + 0.38, 6.4, 0.7, body, size=14, color=INK)
        y += 1.25

    add_rect(slide, 8.15, 1.20, 4.60, 5.15, WHITE, line=LINE, radius=0.08)
    write(slide, 8.45, 1.42, 4.05, 0.4, "Launch pack", size=14, bold=True, color=GOLD)
    write(slide, 8.45, 1.85, 4.05, 0.45, f"{inr(HALF)} to begin", size=20, bold=True, color=DEEP)
    write(
        slide,
        8.45,
        2.40,
        4.05,
        0.85,
        f"{inr(HALF)} when the desk is in daily use and the website is live.",
        size=16,
        color=INK,
    )
    add_rect(slide, 8.45, 3.45, 4.00, 0.015, LINE, rounded=False)
    write(slide, 8.45, 3.65, 4.05, 0.4, "After that", size=14, bold=True, color=GOLD)
    write(
        slide,
        8.45,
        4.15,
        4.05,
        1.85,
        [
            [(f"{inr(CARE)} at the start of each month.", False, INK, 16, False)],
            [("One month’s notice to stop the monthly care.", False, INK, 16, False)],
            [("The software and the website stay with you.", False, INK, 16, False)],
        ],
        after=8,
    )
    footer(slide, 11)
    notes(
        slide,
        "Half and half is the line to be clear about. Do not start the build on a verbal yes. "
        "GST is extra if it applies. Say that once, here, and do not negotiate it away. "
        "If they want only the software, the price is ₹45,000. If they want only the website, ₹28,000. The saving is for taking the launch pack.",
    )


def later(prs):
    slide = new_slide(prs, BG)
    heading(slide, "Add these only when you ask")
    rows = [
        ("WhatsApp messages to the patient", "Booking confirmed, doctor running late, report ready, and the bill.", WHATSAPP, "setup"),
        ("Beds and admission", "Daily bed charges, rounds, and a discharge summary.", BEDS, ""),
        ("Pharmacy", "Stock, batch, expiry, and GST.", PHARMACY, ""),
        ("Operation theatre", "Surgeon and anaesthetist fees on the same bill.", OT, ""),
        ("Ayushman Bharat", "Insurance cases recorded with the visit.", AYUSHMAN, ""),
        ("Kannada on the desk screens", "The software itself, in Kannada. The website is already in both languages.", KANNADA_UI, ""),
    ]
    y = 1.18
    for i, (title, scope, amount, suffix) in enumerate(rows):
        fill = WHITE if i % 2 == 0 else SAGE
        add_rect(slide, 0.58, y, 12.17, 0.86, fill, line=LINE, radius=0.1)
        write(slide, 0.82, y + 0.08, 8.3, 0.34, title, size=16, bold=True, color=DEEP)
        write(slide, 0.82, y + 0.42, 8.3, 0.34, scope, size=13, color=MUTED)
        price = inr(amount) + ("  " + suffix if suffix else "")
        write(
            slide,
            9.15,
            y + 0.16,
            3.3,
            0.54,
            price,
            size=18,
            bold=True,
            color=TEAL,
            align=PP_ALIGN.RIGHT,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        y += 0.92
    footer(slide, 12)
    notes(
        slide,
        "Do not present this slide unless they ask what else you can do, or they assume beds and pharmacy are already in the launch pack. "
        "Patient WhatsApp is a setup fee plus the message charges of the provider, at actual cost. "
        "The website is already bilingual. Kannada inside the desk software is the separate ₹18,000.",
    )


def close(prs):
    slide = new_slide(prs, DEEP)
    add_rect(slide, 0, 0, 0.16, SLIDE_H, GOLD, rounded=False)
    write(slide, 0.72, 0.48, 11.5, 0.4, HOSPITAL.upper(), size=14, bold=True, color=GOLD)
    write(slide, 0.70, 1.15, 11.6, 1.6, "We will sit with you\nand walk through one patient.", size=36, bold=True, color=CREAM, spacing=1.02)
    write(
        slide,
        0.72,
        3.15,
        11.2,
        1.5,
        "Register. Token. Report. Bill.\nThen the website, and where the ads will appear.",
        size=22,
        color=CREAM,
        spacing=1.15,
    )
    add_rect(slide, 0.72, 4.95, 7.6, 1.70, RGBColor(0x0E, 0x4A, 0x44), radius=0.1)
    write(slide, 0.98, 5.12, 7.1, 0.36, "Bring to the sitting", size=14, bold=True, color=GOLD)
    write(
        slide,
        0.98,
        5.52,
        7.1,
        0.95,
        "The doctor list and consultation fees.\nA phone number for patients, and a few photographs.",
        size=16,
        color=CREAM,
        spacing=1.12,
    )
    write(
        slide,
        8.55,
        5.05,
        4.15,
        1.70,
        [
            [(PLACE, False, RGBColor(0xD5, 0xE4, 0xDF), 15, False)],
            [("Your name", True, CREAM, 20, False)],
            [("Your phone", True, CREAM, 20, False)],
        ],
        after=8,
    )
    notes(
        slide,
        "Stop talking after the four words: register, token, report, bill. Ask for a time to sit at reception. "
        "Replace the words Your name and Your phone on this slide with your own before you walk in. "
        "Leave the deck with them, or send the PDF the same day.",
    )


def build_markdown():
    rows_once = "\n".join(
        [
            f"| Hospital software, customized and set up | Your name, doctors, fees, tests, staff logins, and two trainings at the hospital | {inr(SOFTWARE)} |",
            f"| Website, built and handed over | English and Kannada. Doctors, services, map, call, WhatsApp, appointment request | {inr(WEBSITE)} |",
            f"| Google, Facebook, and Instagram | Pages in the hospital’s name, linked to the website | {inr(PRESENCE)} |",
            f"| First ad campaigns | One Google search campaign and one Facebook and Instagram campaign | {inr(ADS_SETUP)} |",
            f"| **Launch pack for {HOSPITAL}** | **All four tasks** | **{inr(LAUNCH)}** |",
        ]
    )
    text = f"""# {HOSPITAL}

**Proposal for the doctors and staff**
{PLACE} · {MONTH}

Your own hospital software. Your own website, built and handed to you. Online marketing and social media ads — we will take care.

---

## Message you can send before the visit

Hello Doctor,

This is for {HOSPITAL}.

We will set up hospital software in your name. Your doctors, your fees, your reception, and your bills. Patient records stay at the hospital.

We will also build your website, in English and Kannada, and give it to you.

Google, Facebook, Instagram, and the ads: we will link them to the hospital and we will take care of them. You get a short report every month.

All of this together, once: {inr(LAUNCH)}.
After that, {inr(CARE)} a month for the software, the website, the posts, and the ads.
A separate {inr(AD_START)} a month is a sound start for the ads themselves. That amount is paid to Google and Meta.

I can come to {PLACE}, sit with your staff, and walk through one patient.

[Your name]
[Your phone]

---

## What we will do

Three things, in this hospital’s name.

1. **Your software.** Customized for {HOSPITAL}. Your doctors, your fees, your staff, your bills. Patient records stay on the hospital computer.
2. **Your website.** Built for you, in English and Kannada, and handed over. The domain is registered in the hospital’s name.
3. **Your reach.** Google, Facebook, Instagram, and the ads. We set them up, link them to the hospital, and run them.

## The desk

| Who | What they get |
| --- | --- |
| Reception | Appointment, walk-in, or emergency. A token queue for each doctor. Emergencies first. The report saved on that patient. |
| Doctor | Today’s patients at this hospital. Only the patients under your care. Order a test or a scan. Your fees, and only your fees. |
| Owner | The day’s collection, cash and UPI. What is still due. The doctor’s share and the hospital’s share. A discount needs your approval. |

It runs on a computer already in the hospital. Reception and the doctor open it on a phone connected to the hospital Wi-Fi.

A doctor who also visits another hospital keeps one login. At {HOSPITAL} they see only their patients here.

## The website

- Home, with phone, map, hours, and a WhatsApp button
- Doctors, specialties, and the days they are here
- Services, in plain words
- An appointment request that reaches reception
- The same pages in English and in Kannada

The site belongs to the hospital. Domain and yearly hosting are registered in the hospital’s name and charged at actual cost, outside the website fee. That is usually a few thousand rupees a year.

## Marketing we take care of

- **Google.** The hospital on Google Maps and Search, with hours, phone, photos, and the way to Krishna Raja Extension.
- **Facebook and Instagram.** Pages in the hospital’s name. Eight posts a month, in Kannada and English. We design them and publish them.
- **Ads.** One search ad for people looking for a hospital in Tiptur. One Facebook and Instagram ad for families nearby. We write them and adjust them.
- **Report.** Once a month: calls, direction requests, messages, the amount spent, and what we will change.

The path is one line. A patient searches, finds the listing or the website, calls or sends a WhatsApp, reception books the visit, and the doctor sees them on today’s list.

Photographs come from the hospital. A camera visit, a reel shoot, and YouTube are outside this fee.

## Price of each task

### Once

| Task | What it covers | Fee |
| --- | --- | --- |
{rows_once}

Taken one by one, the four tasks are {inr(LIST_TOTAL)}. The launch pack is {inr(LAUNCH)}. The hospital saves {inr(SAVE)}.

### Every month

| Care | What it covers | Fee |
| --- | --- | --- |
| Software and website | A new doctor, a fee, a test price, a visiting day. Help when the desk is stuck. Small changes on the website. | {inr(CARE_SOFT)} |
| Posts and the ads | Eight posts, published. Replies on the pages. Campaigns kept running. One report. | {inr(CARE_MKT)} |
| **Both together** | **We take care of the software, the website, the posts, and the ads** | **{inr(CARE)}** |

Taken separately, the two monthly tasks are {inr(CARE_LIST)}. Together they are {inr(CARE)}.

### Ad budget

This is paid to Google and Meta, from the hospital’s own account. We add nothing on top.

| | Amount |
| --- | --- |
| A sound place to start | {inr(AD_START)} a month |
| A sensible range | {inr(AD_LOW)} to {inr(AD_HIGH)} a month |

### The first three months

| | Amount | Paid to |
| --- | --- | --- |
| Launch pack | {inr(LAUNCH)} | Us, once |
| Care, three months | {inr(CARE_3)} | Us |
| Ad budget, three months | {inr(AD_3)} | Google and Meta |
| **Work we do** | **{inr(TO_US)}** | |
| **All money that moves** | **{inr(ALL_IN)}** | |

After three months, the care and the ads continue only if the hospital wants them. The software and the website stay with the hospital.

Add GST if it applies.

## How we begin

1. We sit at the hospital and note the doctors, fees, services, photographs, and the phone patients should call.
2. We put the software on the hospital computer and hold two trainings for reception, billing, and the doctors.
3. We publish the website, Google, Facebook, and Instagram, and then start the ads.

**Payment of the launch pack**

- {inr(HALF)} to begin
- {inr(HALF)} when the desk is in daily use and the website is live
- {inr(CARE)} at the start of each month after that
- One month’s notice to stop the monthly care

Any single task can be taken on its own, at the price in the table. The saving is for the launch pack.

## Add these only when the hospital asks

These are not part of the launch pack.

| Task | What it covers | Fee |
| --- | --- | --- |
| WhatsApp messages to the patient | Booking confirmed, doctor running late, report ready, and the bill. Message charges of the provider are extra, at actual cost. | {inr(WHATSAPP)} setup |
| Beds and admission | Daily bed charges, rounds, and a discharge summary | {inr(BEDS)} |
| Pharmacy | Stock, batch, expiry, and GST | {inr(PHARMACY)} |
| Operation theatre | Surgeon and anaesthetist fees on the same bill | {inr(OT)} |
| Ayushman Bharat | Insurance cases recorded with the visit | {inr(AYUSHMAN)} |
| Kannada on the desk screens | The software itself in Kannada. The website is already in both languages. | {inr(KANNADA_UI)} |

## Next step

We sit at reception and walk through one patient. Register. Token. Report. Bill. Then the website, and where the ads will appear.

[Your name]
[Your phone]

---

## For you, before the meeting

This copy is written for {HOSPITAL}, {PLACE}. To present it to another hospital, change the name and the place on the cover and keep the same prices, unless you have a reason to change them.

Put your name and phone on the last slide before you walk in.

Say {inr(LAUNCH)} first. Then show the four lines. In the same breath, say that {inr(AD_START)} a month is paid by the hospital to Google and Meta.

The launch pack is the desk as it works today: registration, the token queue, reports on the patient record, and the bill, plus the website and the marketing setup. Beds, pharmacy, the operation theatre, Ayushman Bharat, and automatic WhatsApp messages to patients have their own prices. The WhatsApp button on the website is already included.

The desk opens on the hospital computer and on phones on the hospital Wi-Fi. Home access for a doctor is a later conversation.

Show the software only if it is already running on your laptop. This file is the presentation.

Slide notes in the PowerPoint carry the same cues.
"""
    MD_PATH.write_text(text, encoding="utf-8")


def main():
    build_pptx()
    build_markdown()
    print(f"Wrote {PPTX_PATH}")
    print(f"Wrote {MD_PATH}")


if __name__ == "__main__":
    main()
