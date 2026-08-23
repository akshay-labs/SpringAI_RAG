"""
Generates the Day 2 training deck (Vector Embeddings, Semantic Similarity, Document
Chunking, Vector Database & Vector Search) as a .pptx file.

    pip install python-pptx
    python tools/generate_day2_slides.py

Output: Day2-Vector-Embeddings-and-Search.pptx (repo root)
"""
import os

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------------------------------------------------------------------------
# Palette
# ---------------------------------------------------------------------------
NAVY = RGBColor(0x0B, 0x12, 0x20)
NAVY_2 = RGBColor(0x11, 0x1B, 0x30)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
INK = RGBColor(0x0F, 0x17, 0x2A)
MUTED = RGBColor(0x5B, 0x67, 0x7A)
CYAN = RGBColor(0x22, 0xD3, 0xEE)
TEAL = RGBColor(0x14, 0xB8, 0xA6)
AMBER = RGBColor(0xF5, 0x9E, 0x0B)
ROSE = RGBColor(0xFB, 0x71, 0x85)
CARD_BG = RGBColor(0xF1, 0xF5, 0xF9)
CARD_LINE = RGBColor(0xD9, 0xE2, 0xEC)
GREEN = RGBColor(0x22, 0xC5, 0x5E)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H
BLANK = prs.slide_layouts[6]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def add_slide():
    return prs.slides.add_slide(BLANK)


def set_bg(slide, color):
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = color


def add_rect(slide, x, y, w, h, fill=None, line=None, line_w=None, shadow=False, round_=False):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if round_ else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, x, y, w, h)
    if round_:
        try:
            shp.adjustments[0] = 0.06
        except Exception:
            pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = line_w or Pt(1)
    shp.shadow.inherit = False
    return shp


def add_text(slide, x, y, w, h, text, size=18, color=INK, bold=False, italic=False,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font="Calibri", line_spacing=1.0,
             wrap=True):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
        r.font.name = font
    return tb


def add_bullets(slide, x, y, w, h, items, size=18, color=INK, marker_color=None,
                 gap=0.30, bold_lead=False, font="Calibri"):
    """items: list of str, or (lead, rest) tuples for a bold lead-in phrase."""
    marker_color = marker_color or CYAN
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap * 72)
        p.line_spacing = 1.08
        bullet_r = p.add_run()
        bullet_r.text = "›  "
        bullet_r.font.size = Pt(size)
        bullet_r.font.bold = True
        bullet_r.font.color.rgb = marker_color
        bullet_r.font.name = font
        if isinstance(item, tuple):
            lead, rest = item
            r1 = p.add_run()
            r1.text = lead
            r1.font.size = Pt(size)
            r1.font.bold = True
            r1.font.color.rgb = color
            r1.font.name = font
            r2 = p.add_run()
            r2.text = rest
            r2.font.size = Pt(size)
            r2.font.color.rgb = color
            r2.font.name = font
        else:
            r = p.add_run()
            r.text = item
            r.font.size = Pt(size)
            r.font.color.rgb = color
            r.font.name = font
    return tb


def add_pill(slide, x, y, w, h, text, fill, text_color=WHITE, size=13):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    try:
        shp.adjustments[0] = 0.5
    except Exception:
        pass
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    tf = shp.text_frame
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.color.rgb = text_color
    return shp


def kicker_and_title(slide, kicker, title, dark=False, title_size=32):
    kicker_color = CYAN if dark else TEAL
    title_color = WHITE if dark else INK
    add_text(slide, Inches(0.7), Inches(0.45), Inches(8), Inches(0.4), kicker.upper(),
              size=14, color=kicker_color, bold=True)
    add_text(slide, Inches(0.68), Inches(0.78), Inches(11.6), Inches(0.9), title,
              size=title_size, color=title_color, bold=True)
    add_rect(slide, Inches(0.7), Inches(1.55), Inches(1.1), Pt(4), fill=(CYAN if dark else TEAL))


def footer(slide, page_no, section, dark=False):
    color = RGBColor(0x8A, 0x9B, 0xB5) if dark else MUTED
    add_text(slide, Inches(0.7), Inches(7.12), Inches(6), Inches(0.3),
              f"Day 2  ·  {section}", size=11, color=color)
    add_text(slide, Inches(12.2), Inches(7.12), Inches(0.5), Inches(0.3),
              str(page_no), size=11, color=color, align=PP_ALIGN.RIGHT)


def right_arrow(slide, x, y, w, h, color=TEAL):
    shp = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def flow_box(slide, x, y, w, h, title, subtitle=None, fill=WHITE, line=CARD_LINE,
             title_color=INK, subtitle_color=MUTED, title_size=15, subtitle_size=11.5):
    add_rect(slide, x, y, w, h, fill=fill, line=line, line_w=Pt(1.25), round_=True)
    if subtitle:
        add_text(slide, x + Inches(0.12), y + Inches(0.1), w - Inches(0.24), Inches(0.4),
                  title, size=title_size, color=title_color, bold=True, align=PP_ALIGN.CENTER,
                  anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, x + Inches(0.12), y + h - Inches(0.42), w - Inches(0.24), Inches(0.36),
                  subtitle, size=subtitle_size, color=subtitle_color, align=PP_ALIGN.CENTER,
                  anchor=MSO_ANCHOR.MIDDLE)
    else:
        add_text(slide, x + Inches(0.12), y, w - Inches(0.24), h, title, size=title_size,
                  color=title_color, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def set_notes(slide, text):
    notes = slide.notes_slide
    notes.notes_text_frame.text = text


def vector_arrow(slide, origin_x, origin_y, length, angle_deg, color, label, label_color=None):
    """Draws a RIGHT_ARROW rotated about its own center, then repositions so its tail sits at origin."""
    label_color = label_color or color
    thickness = Inches(0.16)
    shp = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, origin_x, origin_y - thickness / 2,
                                  length, thickness)
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    try:
        shp.adjustments[0] = 0.55
        shp.adjustments[1] = 0.55
    except Exception:
        pass
    shp.rotation = angle_deg
    import math
    rad = math.radians(angle_deg)
    lx = origin_x + Emu(int(length * 1.12 * math.cos(rad)))
    ly = origin_y + Emu(int(length * 1.12 * math.sin(rad))) - Inches(0.15)
    add_text(slide, lx, ly, Inches(2.2), Inches(0.35), label, size=13, color=label_color, bold=True)


# ---------------------------------------------------------------------------
# Slide 1 — Title
# ---------------------------------------------------------------------------
s = add_slide()
set_bg(s, NAVY)
add_rect(s, Inches(0), Inches(0), SLIDE_W, Inches(0.09), fill=CYAN)
add_pill(s, Inches(0.7), Inches(1.05), Inches(1.7), Inches(0.4), "DAY 2", CYAN, text_color=NAVY)
add_text(s, Inches(0.68), Inches(1.7), Inches(11.8), Inches(1.9),
          "Vector Embeddings,\nDocument Chunking & Vector Search",
          size=42, color=WHITE, bold=True, line_spacing=1.05)
add_text(s, Inches(0.7), Inches(3.75), Inches(10.5), Inches(0.6),
          "The building blocks behind Retrieval-Augmented Generation (RAG)",
          size=19, color=RGBColor(0xB9, 0xC6, 0xDD))
add_rect(s, Inches(0.7), Inches(4.35), Inches(0.9), Pt(3), fill=CYAN)
add_text(s, Inches(0.7), Inches(6.5), Inches(9), Inches(0.4),
          "How Spring AI + MariaDB turn an HR policy PDF into answerable knowledge",
          size=14, color=RGBColor(0x8A, 0x9B, 0xB5), italic=True)
set_notes(s, "Welcome slide. Frame the session: today is the theory behind the RAG demo they'll "
             "see end-to-end on Day 3. Everything here maps directly onto the HR-policy demo app.")

# ---------------------------------------------------------------------------
# Slide 2 — Agenda
# ---------------------------------------------------------------------------
s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "Agenda", "Four ideas that make RAG possible")
items = [
    ("01", "Introduction to Vector Embeddings", "Turning text into numbers a computer can compare", CYAN),
    ("02", "Understanding Semantic Similarity", "How we measure “closeness in meaning”", TEAL),
    ("03", "Document Chunking for RAG", "Splitting documents so retrieval stays precise", AMBER),
    ("04", "Vector Database & Vector Search", "Storing embeddings and finding the best matches fast", ROSE),
]
y = Inches(2.15)
for num, title, sub, color in items:
    add_rect(s, Inches(0.7), y, Inches(11.9), Inches(1.05), fill=CARD_BG, line=CARD_LINE, line_w=Pt(1), round_=True)
    add_rect(s, Inches(0.7), y, Inches(0.12), Inches(1.05), fill=color)
    add_text(s, Inches(1.05), y + Inches(0.12), Inches(0.9), Inches(0.8), num, size=26, color=color, bold=True)
    add_text(s, Inches(1.95), y + Inches(0.14), Inches(9.5), Inches(0.45), title, size=19, color=INK, bold=True)
    add_text(s, Inches(1.95), y + Inches(0.58), Inches(9.5), Inches(0.4), sub, size=13.5, color=MUTED)
    y += Inches(1.2)
footer(s, 2, "Agenda")
set_notes(s, "Set expectations: each topic builds on the last. By the end, the audience should be "
             "able to trace exactly what happens between 'PDF' and 'answer' in the demo app.")

# ---------------------------------------------------------------------------
# SECTION 1 — Vector Embeddings
# ---------------------------------------------------------------------------
def section_header(no, title, subtitle, color):
    s = add_slide()
    set_bg(s, NAVY)
    add_rect(s, Inches(0), Inches(0), Inches(0.14), SLIDE_H, fill=color)
    add_text(s, Inches(0.9), Inches(2.55), Inches(3), Inches(1.4), no, size=90, color=RGBColor(0x24, 0x30, 0x48), bold=True)
    add_text(s, Inches(0.95), Inches(3.35), Inches(11), Inches(1.0), title, size=36, color=WHITE, bold=True)
    add_text(s, Inches(0.97), Inches(4.25), Inches(10.5), Inches(0.7), subtitle, size=17, color=RGBColor(0xB9, 0xC6, 0xDD))
    add_rect(s, Inches(0.97), Inches(4.85), Inches(0.9), Pt(3), fill=color)
    return s

s = section_header("01", "Introduction to Vector Embeddings", "How do you make “meaning” something a database can compare?", CYAN)
set_notes(s, "Transition slide. Ask the audience: how would YOU teach a computer that 'vacation' and "
             "'PTO' mean almost the same thing, but 'vacation' and 'invoice' don't? That's the problem "
             "embeddings solve.")

# --- What is an embedding
s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "01 · Vector Embeddings", "What is an embedding?")
add_text(s, Inches(0.7), Inches(1.95), Inches(5.6), Inches(1.3),
          "An embedding model converts a piece of text into a list of numbers "
          "(a vector) that captures its meaning.",
          size=18, color=INK, line_spacing=1.15)
add_bullets(s, Inches(0.7), Inches(3.3), Inches(5.7), Inches(3),
             [("Similar meaning ", "→ vectors point in a similar direction"),
              ("Unrelated meaning ", "→ vectors point in very different directions"),
              ("Every model has a fixed size ", "→ e.g. 1536 numbers per vector for OpenAI's text-embedding-3-small"),
              ("It's just an API call ", "→ text in, array of floats out")],
             size=16.5)
# right-side visual: text -> vector
bx, by, bw, bh = Inches(6.9), Inches(2.15), Inches(2.5), Inches(0.85)
flow_box(s, bx, by, bw, bh, "“How many vacation\ndays do I get?”", fill=CARD_BG, line=CARD_LINE, title_size=14)
right_arrow(s, bx + bw + Inches(0.15), by + Inches(0.28), Inches(0.55), Inches(0.3), color=TEAL)
bx2 = bx + bw + Inches(0.85)
add_rect(s, bx2, by, Inches(2.9), bh, fill=INK, round_=True)
add_text(s, bx2 + Inches(0.15), by + Inches(0.12), Inches(2.6), Inches(0.6),
          "[0.021, -0.144, 0.078, …]", size=12.5, color=CYAN, font="Consolas", anchor=MSO_ANCHOR.MIDDLE)
add_text(s, bx2, by + bh + Inches(0.08), Inches(2.9), Inches(0.3), "1,536 numbers", size=12, color=MUTED, align=PP_ALIGN.CENTER)
add_text(s, bx + Inches(0.15), by + Inches(1.35), Inches(6), Inches(0.4), "embedding model", size=12, color=MUTED, italic=True)
footer(s, 4, "Introduction to Vector Embeddings")
set_notes(s, "Live in the demo: this is exactly what DocumentIngestionService/RagQueryService send to "
             "OpenAI's embedding endpoint. Show the config: spring.ai.openai.embedding.options.model="
             "text-embedding-3-small if you want to point at the code.")

# --- Embedding space visual
s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "01 · Vector Embeddings", "Meaning becomes geometry")
add_text(s, Inches(0.7), Inches(1.85), Inches(11.6), Inches(0.5),
          "Plot embeddings in space and related concepts land near each other — even without shared words.",
          size=15.5, color=MUTED)
plot_x, plot_y, plot_w, plot_h = Inches(0.9), Inches(2.55), Inches(11.4), Inches(4.2)
add_rect(s, plot_x, plot_y, plot_w, plot_h, fill=CARD_BG, line=CARD_LINE, line_w=Pt(1))

def dot(cx, cy, label, color, dx=0.18, dy=-0.42):
    d = Inches(0.16)
    slide_shapes = s.shapes.add_shape(MSO_SHAPE.OVAL, cx - d/2, cy - d/2, d, d)
    slide_shapes.fill.solid()
    slide_shapes.fill.fore_color.rgb = color
    slide_shapes.line.fill.background()
    slide_shapes.shadow.inherit = False
    add_text(s, cx + Inches(dx) - Inches(0.9), cy + Inches(dy), Inches(1.8), Inches(0.3),
              label, size=12.5, color=INK, bold=True, align=PP_ALIGN.CENTER)

# cluster A: leave / time off
leave_cluster = [
    (Inches(2.6), Inches(3.3), "“annual leave”"),
    (Inches(3.15), Inches(3.75), "“vacation days”"),
    (Inches(2.35), Inches(4.05), "“time off”"),
    (Inches(3.0), Inches(4.35), "“PTO balance”"),
]
for cx, cy, label in leave_cluster:
    dot(cx, cy, label, TEAL)
add_rect(s, Inches(1.85), Inches(2.95), Inches(2.0), Inches(1.75), fill=None, line=TEAL, line_w=Pt(1.5), round_=True)
add_text(s, Inches(1.85), Inches(2.7), Inches(2.2), Inches(0.3), "leave & time off", size=12, color=TEAL, bold=True)

# cluster B: expenses
expense_cluster = [
    (Inches(8.6), Inches(4.5), "“expense report”"),
    (Inches(9.3), Inches(4.9), "“reimbursement”"),
    (Inches(8.9), Inches(5.25), "“receipt”"),
]
for cx, cy, label in expense_cluster:
    dot(cx, cy, label, AMBER)
add_rect(s, Inches(7.9), Inches(4.2), Inches(2.1), Inches(1.4), fill=None, line=AMBER, line_w=Pt(1.5), round_=True)
add_text(s, Inches(7.9), Inches(3.95), Inches(2.3), Inches(0.3), "expenses", size=12, color=AMBER, bold=True)

# outlier
dot(Inches(10.8), Inches(2.95), "“code of conduct”", ROSE, dx=0.05, dy=0.22)

add_text(s, plot_x + Inches(0.2), plot_y + plot_h - Inches(0.4), Inches(6), Inches(0.3),
          "(a simplified 2-D sketch — real embeddings use 1,000+ dimensions)", size=11.5, color=MUTED, italic=True)
footer(s, 5, "Introduction to Vector Embeddings")
set_notes(s, "This is a simplification for intuition — real vectors aren't 2D and you can't literally "
             "plot them like this, but the clustering behavior is real. This is why a query about "
             "'vacation days' can retrieve a chunk that never uses the word 'vacation'.")

# ---------------------------------------------------------------------------
# SECTION 2 — Semantic Similarity
# ---------------------------------------------------------------------------
s = section_header("02", "Understanding Semantic Similarity", "Once meaning is a vector, “similar” becomes something we can compute", TEAL)
set_notes(s, "Bridge from section 1: now that we have vectors, how do we actually score how similar two "
             "of them are? That's cosine similarity.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "02 · Semantic Similarity", "Cosine similarity: comparing direction, not distance")
add_text(s, Inches(0.7), Inches(1.95), Inches(5.6), Inches(1.0),
          "We measure the angle between two vectors. A smaller angle means the two "
          "pieces of text mean something closer.",
          size=17, color=INK, line_spacing=1.15)
add_rect(s, Inches(0.7), Inches(3.05), Inches(5.7), Inches(1.0), fill=CARD_BG, line=CARD_LINE, round_=True)
add_text(s, Inches(0.95), Inches(3.2), Inches(5.2), Inches(0.35), "cos(θ) = (A · B) / (|A| × |B|)",
          size=18, color=INK, bold=True, font="Consolas")
add_text(s, Inches(0.95), Inches(3.62), Inches(5.2), Inches(0.35), "result ranges from -1 (opposite) to 1 (identical direction)",
          size=12.5, color=MUTED)
add_bullets(s, Inches(0.7), Inches(4.35), Inches(5.7), Inches(2.2),
             ["1.0  →  same meaning", "≈ 0.8  →  strongly related", "≈ 0.3  →  loosely related",
              "≈ 0.0  →  unrelated"], size=16)

# vector diagram
diagram_x, diagram_y = Inches(7.6), Inches(2.6)
add_rect(s, diagram_x - Inches(0.3), diagram_y - Inches(0.3), Inches(5.4), Inches(3.9), fill=CARD_BG, line=CARD_LINE, round_=True)
origin_x, origin_y = diagram_x + Inches(0.4), diagram_y + Inches(3.1)
vector_arrow(s, origin_x, origin_y, Inches(3.6), -32, TEAL, "“annual leave days” (query)")
vector_arrow(s, origin_x, origin_y, Inches(3.3), -55, RGBColor(0x0d, 0x94, 0x88), "“vacation policy” (chunk)")
add_text(s, origin_x + Inches(0.65), origin_y - Inches(0.55), Inches(2), Inches(0.3), "small angle → high similarity", size=12, color=MUTED, italic=True)
footer(s, 7, "Understanding Semantic Similarity")
set_notes(s, "This is what vectorStore.similaritySearch() computes under the hood, for every stored "
             "chunk, against the embedded question. MariaDB is configured with distance-type: COSINE "
             "(see application.yml) — point that out if showing code.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "02 · Semantic Similarity", "Seen through the demo app")
add_text(s, Inches(0.7), Inches(1.95), Inches(11.6), Inches(0.5),
          "The same question scores very differently against two real chunks from the HR policy.",
          size=15.5, color=MUTED)
q_box = flow_box
add_rect(s, Inches(0.7), Inches(2.55), Inches(11.9), Inches(0.85), fill=INK, round_=True)
add_text(s, Inches(1.0), Inches(2.55), Inches(11), Inches(0.85),
          "“How many days of annual leave do employees get?”", size=18, color=WHITE, bold=True,
          anchor=MSO_ANCHOR.MIDDLE)

row_y = Inches(3.75)
add_rect(s, Inches(0.7), row_y, Inches(5.7), Inches(2.7), fill=CARD_BG, line=TEAL, line_w=Pt(1.5), round_=True)
add_pill(s, Inches(0.95), row_y + Inches(0.22), Inches(1.6), Inches(0.4), "cos ≈ 0.86", TEAL)
add_text(s, Inches(0.95), row_y + Inches(0.8), Inches(5.2), Inches(0.4), "Section 2 — Annual Leave and Time Off", size=15, color=INK, bold=True)
add_text(s, Inches(0.95), row_y + Inches(1.25), Inches(5.2), Inches(1.3),
          "“Full-time employees accrue 18 days of paid annual leave per calendar year…”",
          size=13, color=MUTED, italic=True)

add_rect(s, Inches(6.7), row_y, Inches(5.9), Inches(2.7), fill=CARD_BG, line=ROSE, line_w=Pt(1.5), round_=True)
add_pill(s, Inches(6.95), row_y + Inches(0.22), Inches(1.6), Inches(0.4), "cos ≈ 0.19", ROSE)
add_text(s, Inches(6.95), row_y + Inches(0.8), Inches(5.4), Inches(0.4), "Section 5 — Expense Reimbursement", size=15, color=INK, bold=True)
add_text(s, Inches(6.95), row_y + Inches(1.25), Inches(5.4), Inches(1.3),
          "“Meal expenses while traveling on company business are capped at $60 per day…”",
          size=13, color=MUTED, italic=True)
footer(s, 8, "Understanding Semantic Similarity")
set_notes(s, "Illustrative scores for the talk — exact numbers will vary slightly by embedding model, "
             "but the ordering (leave chunk wins, expense chunk loses) is the real, reproducible behavior "
             "you can show live by asking this exact question in the UI and inspecting sources[].")

# ---------------------------------------------------------------------------
# SECTION 3 — Document Chunking
# ---------------------------------------------------------------------------
s = section_header("03", "Document Chunking for RAG", "You don't embed a whole document — you embed pieces of it. Which pieces, and why?", AMBER)
set_notes(s, "Bridge: before anything gets embedded, the source document has to be broken into chunks. "
             "This is the step people underestimate most in RAG.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "03 · Document Chunking", "Why not embed the whole document at once?")
add_bullets(s, Inches(0.7), Inches(2.1), Inches(11.7), Inches(3.2),
             [("Dilution ", "— one vector for a whole document blurs many different topics into a single average meaning"),
              ("Precision ", "— you want the retrieved text to be just the relevant paragraph, not an entire policy handbook"),
              ("Context limits ", "— embedding models and LLM prompts both have practical size limits"),
              ("Cost & speed ", "— smaller, targeted chunks mean cheaper, faster retrieval and generation")],
             size=18, gap=0.42)
footer(s, 10, "Document Chunking for RAG")
set_notes(s, "Keep this brief — the real payoff is the next slide where we look at the actual chunks "
             "the demo produced.")

# --- chunking pipeline
s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "03 · Document Chunking", "From PDF to stored chunks")
fy = Inches(2.3)
fw, fh = Inches(2.35), Inches(1.15)
gap = Inches(0.55)
labels = [
    ("HR-Company-\nPolicy.pdf", "1 file"),
    ("PagePdfDocument\nReader", "2 pages"),
    ("TokenTextSplitter", "~800 tokens/chunk"),
    ("3 chunks", "embedded & stored"),
]
x = Inches(0.7)
for i, (title, sub) in enumerate(labels):
    color = CARD_BG if i not in (2,) else RGBColor(0xFF, 0xF3, 0xDA)
    line_c = CARD_LINE if i not in (2,) else AMBER
    flow_box(s, x, fy, fw, fh, title, sub, fill=color, line=line_c)
    if i < len(labels) - 1:
        right_arrow(s, x + fw + Inches(0.06), fy + fh/2 - Inches(0.15), gap - Inches(0.12), Inches(0.3), color=MUTED)
    x += fw + gap

add_text(s, Inches(0.7), Inches(3.85), Inches(11.6), Inches(0.4),
          "TokenTextSplitter packs text greedily up to a token-count target, breaking near a paragraph "
          "boundary — it does not know about your section headings.",
          size=14.5, color=MUTED)

# real chunk sizes
ty = Inches(4.55)
add_text(s, Inches(0.7), ty, Inches(4), Inches(0.35), "Actual chunks from the demo run:", size=14, color=INK, bold=True)
chunks = [("Chunk 1", "5,676 chars", "Northwind Retail Group … (opening + policies)", TEAL),
          ("Chunk 2", "4,615 chars", "…arrangements must be agreed upon with the employee's direct manager…", TEAL),
          ("Chunk 3", "967 chars", "…figures, must never be shared outside the company without written au[thorization]…", ROSE)]
cy = ty + Inches(0.5)
cw = Inches(3.83)
cx = Inches(0.7)
for name, size_lbl, preview, accent in chunks:
    add_rect(s, cx, cy, cw, Inches(1.75), fill=CARD_BG, line=CARD_LINE, round_=True)
    add_rect(s, cx, cy, Inches(0.09), Inches(1.75), fill=accent)
    add_text(s, cx + Inches(0.22), cy + Inches(0.12), cw - Inches(0.4), Inches(0.3), name, size=13.5, color=INK, bold=True)
    add_pill(s, cx + cw - Inches(1.55), cy + Inches(0.12), Inches(1.35), Inches(0.32), size_lbl, accent, size=11)
    add_text(s, cx + Inches(0.22), cy + Inches(0.55), cw - Inches(0.42), Inches(1.1), "“" + preview + "”",
              size=11.5, color=MUTED, italic=True, line_spacing=1.1)
    cx += cw + Inches(0.18)
add_text(s, Inches(0.7), Inches(6.85), Inches(11), Inches(0.3),
          "Chunk 3 is a mid-sentence leftover — proof the split is token-driven, not heading-aware.",
          size=12.5, color=AMBER, bold=True, italic=True)
footer(s, 11, "Document Chunking for RAG")
set_notes(s, "This is a real query result from the demo DB: "
             "SELECT id, CHAR_LENGTH(content), LEFT(content,80) FROM vector_store; "
             "Point out chunk 3 starts mid-sentence in the Code of Conduct section — great live moment "
             "to show the DB rows on screen if you have time.")

# --- trade-offs
s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "03 · Document Chunking", "The size trade-off")
colw = Inches(5.65)
gap2 = Inches(0.6)
add_rect(s, Inches(0.7), Inches(2.15), colw, Inches(3.9), fill=CARD_BG, line=ROSE, line_w=Pt(1.5), round_=True)
add_text(s, Inches(1.0), Inches(2.4), colw - Inches(0.6), Inches(0.4), "Chunks too LARGE", size=18, color=ROSE, bold=True)
add_bullets(s, Inches(1.0), Inches(2.95), colw - Inches(0.6), Inches(3),
             ["Mixes multiple topics into one vector", "Retrieval gets less precise — relevant text is diluted by irrelevant text",
              "Wastes LLM context on unrelated content"], size=15, marker_color=ROSE)

x2 = Inches(0.7) + colw + gap2
add_rect(s, x2, Inches(2.15), colw, Inches(3.9), fill=CARD_BG, line=AMBER, line_w=Pt(1.5), round_=True)
add_text(s, x2 + Inches(0.3), Inches(2.4), colw - Inches(0.6), Inches(0.4), "Chunks too SMALL", size=18, color=AMBER, bold=True)
add_bullets(s, x2 + Inches(0.3), Inches(2.95), colw - Inches(0.6), Inches(3),
             ["Loses surrounding context a sentence needs", "More chunks to search — more noise, higher cost",
              "Easy to split a fact away from its qualifier (e.g. the number away from what it's counting)"],
             size=15, marker_color=AMBER)
add_text(s, Inches(0.7), Inches(6.35), Inches(11.7), Inches(0.5),
          "There's no universal right size — it's a tunable parameter you choose per document type.",
          size=14.5, color=INK, bold=True, align=PP_ALIGN.CENTER)
footer(s, 12, "Document Chunking for RAG")
set_notes(s, "If time allows: mention chunk overlap as the common mitigation for the 'too small' problem "
             "— repeating a bit of text at chunk boundaries so a split fact still has its neighbor "
             "sentence nearby.")

# ---------------------------------------------------------------------------
# SECTION 4 — Vector DB & Search
# ---------------------------------------------------------------------------
s = section_header("04", "Vector Database & Vector Search", "Where do all these vectors live, and how do we search millions of them fast?", ROSE)
set_notes(s, "Bridge: now we have chunks and we know how to score similarity between two vectors. "
             "At scale, comparing a query against every single row one by one wouldn't be practical "
             "— that's what a vector database is for.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "04 · Vector Database & Search", "Why not just a normal SQL table?")
add_bullets(s, Inches(0.7), Inches(2.1), Inches(11.7), Inches(2.6),
             [("A vector database ", "stores each chunk's embedding alongside its text and metadata, and adds an index built specifically for high-dimensional similarity search"),
              ("Curse of dimensionality ", "— with hundreds/thousands of dimensions, a plain row-by-row scan to compare a query against every vector doesn't scale"),
              ("Approximate Nearest Neighbor (ANN) indexes ", "trade a tiny bit of accuracy for massive speed — sub-second search over millions of vectors")],
             size=17, gap=0.4)
footer(s, 14, "Vector Database & Vector Search")
set_notes(s, "Keep conceptual here — the next slide grounds it directly in the demo's actual table.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "04 · Vector Database & Search", "MariaDB as the vector store")
add_text(s, Inches(0.7), Inches(1.95), Inches(11.6), Inches(0.5),
          "MariaDB 11.7+ has a native VECTOR column type — embeddings live in an ordinary table, right next to the text.",
          size=15.5, color=MUTED)
# table sketch
tx, tyy, tw = Inches(0.9), Inches(2.65), Inches(11.2)
cols = [("id", 1.6, WHITE), ("content", 3.6, WHITE), ("metadata", 2.8, WHITE), ("embedding", 3.2, RGBColor(0xE9, 0xFB, 0xF3))]
header_h = Inches(0.55)
row_h = Inches(0.62)
cxp = tx
add_rect(s, tx, tyy, tw, header_h, fill=INK, round_=True)
cxp = tx
for name, w, _ in cols:
    add_text(s, cxp + Inches(0.15), tyy, Inches(w) - Inches(0.2), header_h, name, size=14, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    cxp += Inches(w)
sample_rows = [
    ("f8d4…", "“Full-time employees accrue 18 days…”", "{page: 1}", "[0.021,-0.144,…]"),
    ("a3e1…", "“…arrangements must be agreed upon…”", "{page: 2}", "[0.093, 0.017,…]"),
    ("9b02…", "“…figures, must never be shared…”", "{page: 2}", "[-0.061,0.052,…]"),
]
ry = tyy + header_h
for r_i, row in enumerate(sample_rows):
    bgc = WHITE if r_i % 2 == 0 else CARD_BG
    add_rect(s, tx, ry, tw, row_h, fill=bgc, line=CARD_LINE, line_w=Pt(0.75))
    cxp = tx
    for (name, w, hl), val in zip(cols, row):
        cellc = RGBColor(0xEA, 0xFB, 0xF3) if hl != WHITE and r_i % 2 == 0 else (RGBColor(0xDF, 0xF7, 0xEC) if hl != WHITE else None)
        if cellc:
            add_rect(s, cxp, ry, Inches(w), row_h, fill=cellc)
        add_text(s, cxp + Inches(0.15), ry, Inches(w) - Inches(0.25), row_h, val, size=11.5,
                  color=(RGBColor(0x0d,0x7a,0x5a) if name=="embedding" else INK),
                  anchor=MSO_ANCHOR.MIDDLE, font=("Consolas" if name in ("id","embedding") else "Calibri"))
        cxp += Inches(w)
    ry += row_h
add_text(s, tx, ry + Inches(0.15), tw, Inches(0.35),
          "vector_store   ·   indexed for COSINE distance   ·   dimensions = 1536 (text-embedding-3-small)",
          size=12.5, color=MUTED, italic=True)
footer(s, 15, "Vector Database & Vector Search")
set_notes(s, "This mirrors the real schema: id (uuid), content (text), metadata (json), embedding "
             "(vector(1536)) with a VECTOR KEY index. spring.ai.vectorstore.mariadb.distance-type=COSINE "
             "in application.yml controls the metric shown here.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "04 · Vector Database & Search", "similaritySearch(question, topK)")
fy = Inches(2.4)
fw, fh = Inches(2.6), Inches(1.2)
gap3 = Inches(0.35)
steps = [
    ("Question", "“How many leave days?”"),
    ("Embed the question", "same model as ingestion"),
    ("Compare vs every\nstored vector", "COSINE distance"),
    ("Return top K", "topK = 4"),
]
x = Inches(0.7)
for i, (title, sub) in enumerate(steps):
    fill_c = RGBColor(0xE9, 0xFB, 0xF3) if i == 3 else CARD_BG
    line_c = TEAL if i == 3 else CARD_LINE
    flow_box(s, x, fy, fw, fh, title, sub, fill=fill_c, line=line_c)
    if i < len(steps) - 1:
        right_arrow(s, x + fw + Inches(0.04), fy + fh/2 - Inches(0.15), gap3 - Inches(0.08), Inches(0.3), color=MUTED)
    x += fw + gap3
right_arrow(s, x + Inches(0.05), fy + fh/2 - Inches(0.15), Inches(0.55), Inches(0.3), color=TEAL)
add_rect(s, x + Inches(0.7), fy - Inches(0.1), Inches(2.4), fh + Inches(0.2), fill=NAVY, round_=True)
add_text(s, x + Inches(0.85), fy + Inches(0.15), Inches(2.1), Inches(0.9),
          "Top-4 chunks\n→ context for the LLM", size=13.5, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)

add_text(s, Inches(0.7), Inches(4.2), Inches(11.6), Inches(0.4),
          "This whole step happens before the LLM is ever called — retrieval decides what the model is allowed to see.",
          size=15, color=INK, bold=True)
add_rect(s, Inches(0.7), Inches(4.85), Inches(11.9), Inches(1.7), fill=CARD_BG, line=CARD_LINE, round_=True)
add_text(s, Inches(1.0), Inches(5.05), Inches(11.2), Inches(0.35), "In the demo code (RagQueryService.java):", size=13, color=MUTED, bold=True)
add_text(s, Inches(1.0), Inches(5.4), Inches(11.2), Inches(1.0),
          "vectorStore.similaritySearch(\n    SearchRequest.builder().query(question).topK(topK).build());",
          size=14, color=INK, font="Consolas")
footer(s, 16, "Vector Database & Vector Search")
set_notes(s, "This is the literal line from RagQueryService.answerWithSources(). Good moment to flip to "
             "the IDE/terminal if doing a hybrid slide+code walkthrough.")

# ---------------------------------------------------------------------------
# Closing — Putting it together
# ---------------------------------------------------------------------------
s = add_slide()
set_bg(s, NAVY)
kicker_and_title(s, "Putting it together", "This is the full pipeline behind the demo", dark=True)
fy = Inches(2.5)
fw, fh = Inches(2.0), Inches(1.1)
gap4 = Inches(0.28)
pipeline = [
    ("PDF", CYAN), ("Pages", CYAN), ("Chunks", AMBER), ("Embeddings", TEAL),
    ("MariaDB", TEAL), ("Question → top-K", ROSE), ("LLM answer", ROSE),
]
x = Inches(0.55)
for i, (title, color) in enumerate(pipeline):
    shp = slide_shapes = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, fy, fw, fh)
    shp.fill.solid(); shp.fill.fore_color.rgb = NAVY_2
    shp.line.color.rgb = color; shp.line.width = Pt(1.5)
    shp.shadow.inherit = False
    try:
        shp.adjustments[0] = 0.12
    except Exception:
        pass
    tf = shp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = title
    r.font.size = Pt(13.5); r.font.bold = True; r.font.color.rgb = WHITE
    if i < len(pipeline) - 1:
        right_arrow(s, x + fw + Inches(0.02), fy + fh/2 - Inches(0.12), gap4 - Inches(0.04), Inches(0.24), color=RGBColor(0x4B, 0x5A, 0x78))
    x += fw + gap4

add_text(s, Inches(0.7), Inches(4.15), Inches(11.6), Inches(0.5),
          "Day 3 picks up right here — wiring this into a Spring Boot REST API and a React front end.",
          size=16, color=RGBColor(0xB9, 0xC6, 0xDD))
add_rect(s, Inches(0.7), Inches(4.85), Inches(11.9), Inches(1.7), fill=NAVY_2, round_=True)
add_bullets(s, Inches(1.0), Inches(5.05), Inches(11.3), Inches(1.4),
             ["Embeddings turn text into comparable vectors",
              "Cosine similarity scores how close two vectors are in meaning",
              "Chunking decides what unit of text gets its own vector",
              "A vector database indexes those vectors for fast top-K search"],
             size=14.5, color=WHITE, marker_color=CYAN, gap=0.18)
footer(s, 17, "Recap", dark=True)
set_notes(s, "Recap slide before switching to the live demo. Then: open the app, click Ingest, ask a "
             "question, and show the sources[] array in the response — that's this entire diagram "
             "happening in under a second.")

# ---------------------------------------------------------------------------
out_path = os.path.join(os.path.dirname(__file__), "..", "Day2-Vector-Embeddings-and-Search.pptx")
prs.save(out_path)
print(f"Wrote {os.path.abspath(out_path)}  ({len(prs.slides.__iter__.__self__._sldIdLst)} slides)")
