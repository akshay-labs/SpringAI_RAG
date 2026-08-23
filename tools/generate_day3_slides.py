"""
Generates the Day 3 training deck (RAG Application Architecture, Document Ingestion,
Embeddings & Vector Storage, Retrieval & Context Generation, LLM Integration,
Spring Boot backend, React frontend, End-to-End Demo) as a .pptx file.

    pip install python-pptx
    python tools/generate_day3_slides.py

Output: Day3-Building-RAG-Application.pptx (repo root)
"""
import os
import math

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---------------------------------------------------------------------------
# Palette (matches the Day 2 deck)
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
VIOLET = RGBColor(0x8B, 0x5C, 0xF6)
CARD_BG = RGBColor(0xF1, 0xF5, 0xF9)
CARD_LINE = RGBColor(0xD9, 0xE2, 0xEC)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H
BLANK = prs.slide_layouts[6]


# ---------------------------------------------------------------------------
# Helpers (same conventions as the Day 2 deck)
# ---------------------------------------------------------------------------
def add_slide():
    return prs.slides.add_slide(BLANK)


def set_bg(slide, color):
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = color


def add_rect(slide, x, y, w, h, fill=None, line=None, line_w=None, round_=False):
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
                 gap=0.30, font="Calibri"):
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


def kicker_and_title(slide, kicker, title, dark=False, title_size=30):
    kicker_color = CYAN if dark else TEAL
    title_color = WHITE if dark else INK
    add_text(slide, Inches(0.7), Inches(0.42), Inches(10), Inches(0.4), kicker.upper(),
              size=14, color=kicker_color, bold=True)
    add_text(slide, Inches(0.68), Inches(0.74), Inches(11.8), Inches(0.85), title,
              size=title_size, color=title_color, bold=True)
    add_rect(slide, Inches(0.7), Inches(1.48), Inches(1.1), Pt(4), fill=(CYAN if dark else TEAL))


def footer(slide, page_no, section, dark=False):
    color = RGBColor(0x8A, 0x9B, 0xB5) if dark else MUTED
    add_text(slide, Inches(0.7), Inches(7.12), Inches(7), Inches(0.3),
              f"Day 3  ·  {section}", size=11, color=color)
    add_text(slide, Inches(12.2), Inches(7.12), Inches(0.5), Inches(0.3),
              str(page_no), size=11, color=color, align=PP_ALIGN.RIGHT)


def right_arrow(slide, x, y, w, h, color=TEAL):
    shp = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def down_arrow(slide, x, y, w, h, color=MUTED):
    shp = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def flow_box(slide, x, y, w, h, title, subtitle=None, fill=WHITE, line=CARD_LINE,
             title_color=INK, subtitle_color=MUTED, title_size=14, subtitle_size=11):
    add_rect(slide, x, y, w, h, fill=fill, line=line, line_w=Pt(1.25), round_=True)
    if subtitle:
        add_text(slide, x + Inches(0.1), y + Inches(0.08), w - Inches(0.2), Inches(0.42),
                  title, size=title_size, color=title_color, bold=True, align=PP_ALIGN.CENTER,
                  anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, x + Inches(0.1), y + h - Inches(0.4), w - Inches(0.2), Inches(0.34),
                  subtitle, size=subtitle_size, color=subtitle_color, align=PP_ALIGN.CENTER,
                  anchor=MSO_ANCHOR.MIDDLE)
    else:
        add_text(slide, x + Inches(0.1), y, w - Inches(0.2), h, title, size=title_size,
                  color=title_color, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def code_box(slide, x, y, w, h, code, size=12.5, color=RGBColor(0xE2, 0xE8, 0xF0),
             bg=INK, title=None):
    add_rect(slide, x, y, w, h, fill=bg, round_=True)
    top = y + Inches(0.12)
    if title:
        add_text(slide, x + Inches(0.2), y + Inches(0.08), w - Inches(0.4), Inches(0.3),
                  title, size=11, color=CYAN, bold=True)
        top = y + Inches(0.42)
    add_text(slide, x + Inches(0.2), top, w - Inches(0.4), h - (top - y) - Inches(0.1),
              code, size=size, color=color, font="Consolas", line_spacing=1.15)


def set_notes(slide, text):
    notes = slide.notes_slide
    notes.notes_text_frame.text = text


def section_header(no, title, subtitle, color):
    s = add_slide()
    set_bg(s, NAVY)
    add_rect(s, Inches(0), Inches(0), Inches(0.14), SLIDE_H, fill=color)
    add_text(s, Inches(0.9), Inches(2.5), Inches(3), Inches(1.4), no, size=88,
              color=RGBColor(0x24, 0x30, 0x48), bold=True)
    add_text(s, Inches(0.95), Inches(3.3), Inches(11), Inches(1.0), title, size=34, color=WHITE, bold=True)
    add_text(s, Inches(0.97), Inches(4.2), Inches(10.7), Inches(0.9), subtitle, size=16,
              color=RGBColor(0xB9, 0xC6, 0xDD), line_spacing=1.2)
    add_rect(s, Inches(0.97), Inches(4.95), Inches(0.9), Pt(3), fill=color)
    return s


def pipeline_row(slide, x, y, fw, fh, gap, boxes, arrow_color=MUTED):
    cx = x
    for i, (title, sub, fill, line) in enumerate(boxes):
        flow_box(slide, cx, y, fw, fh, title, sub, fill=fill, line=line)
        if i < len(boxes) - 1:
            right_arrow(slide, cx + fw + Inches(0.04), y + fh / 2 - Inches(0.13), gap - Inches(0.08),
                        Inches(0.26), color=arrow_color)
        cx += fw + gap
    return cx


# ===========================================================================
# Slide 1 — Title
# ===========================================================================
s = add_slide()
set_bg(s, NAVY)
add_rect(s, Inches(0), Inches(0), SLIDE_W, Inches(0.09), fill=TEAL)
add_pill(s, Inches(0.7), Inches(1.05), Inches(1.7), Inches(0.4), "DAY 3", TEAL, text_color=NAVY)
add_text(s, Inches(0.68), Inches(1.7), Inches(11.8), Inches(1.9),
          "Building an End-to-End\nRAG Application",
          size=42, color=WHITE, bold=True, line_spacing=1.05)
add_text(s, Inches(0.7), Inches(3.75), Inches(10.5), Inches(0.6),
          "From a PDF on disk to a working REST API and React front end",
          size=19, color=RGBColor(0xB9, 0xC6, 0xDD))
add_rect(s, Inches(0.7), Inches(4.35), Inches(0.9), Pt(3), fill=TEAL)
add_text(s, Inches(0.7), Inches(6.5), Inches(10), Inches(0.4),
          "Spring AI  ·  MariaDB vector store  ·  OpenAI  ·  React",
          size=14, color=RGBColor(0x8A, 0x9B, 0xB5), italic=True)
set_notes(s, "Recap in one line: Day 2 was the theory (embeddings, similarity, chunking, vector "
             "search). Today we wire all of it into a real running application, end to end, live.")

# ===========================================================================
# Slide 2 — Agenda
# ===========================================================================
s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "Agenda", "Eight steps from PDF to product")
items = [
    ("01", "RAG Application Architecture", "The shape of the whole system", CYAN),
    ("02", "Document Ingestion & Processing", "PDF → pages → chunks", TEAL),
    ("03", "Embeddings & Vector Storage", "Chunks → vectors → MariaDB", AMBER),
    ("04", "Retrieval & Context Generation", "Question → top-K → augmented prompt", ROSE),
    ("05", "LLM Integration", "Swappable chat models via Spring AI", VIOLET),
    ("06", "Building the Backend (Spring Boot)", "REST API, services, config", CYAN),
    ("07", "Building the Frontend (React)", "A minimal UI for ingest + ask", TEAL),
    ("08", "End-to-End Demonstration", "Watch it all run, live", AMBER),
]
y = Inches(1.95)
col_w = Inches(5.85)
row_h = Inches(1.18)
for i, (num, title, sub, color) in enumerate(items):
    col = i // 4
    row = i % 4
    x = Inches(0.7) + col * (col_w + Inches(0.3))
    yy = y + row * row_h
    add_rect(s, x, yy, col_w, Inches(1.02), fill=CARD_BG, line=CARD_LINE, line_w=Pt(1), round_=True)
    add_rect(s, x, yy, Inches(0.1), Inches(1.02), fill=color)
    add_text(s, x + Inches(0.28), yy + Inches(0.1), Inches(0.7), Inches(0.75), num, size=22, color=color, bold=True)
    add_text(s, x + Inches(1.05), yy + Inches(0.12), col_w - Inches(1.2), Inches(0.42), title, size=15, color=INK, bold=True)
    add_text(s, x + Inches(1.05), yy + Inches(0.54), col_w - Inches(1.2), Inches(0.4), sub, size=11.5, color=MUTED)
footer(s, 2, "Agenda")
set_notes(s, "8 topics, roughly in the order the request actually flows through the system - "
             "architecture first, then ingestion, storage, retrieval, generation, then the two "
             "codebases (backend/frontend), then the live demo.")

# ===========================================================================
# 01 — RAG Application Architecture
# ===========================================================================
s = section_header("01", "RAG Application Architecture", "Two pipelines share the same storage: one fills it, one reads from it.", CYAN)
set_notes(s, "Frame the whole day around this one idea: ingestion and query are two separate flows "
             "that meet at the vector store.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "01 · Architecture", "The system in layers")
layers = [
    ("React UI  (rag-demo-ui)", "Ingest button  ·  question box  ·  answer + sources", VIOLET),
    ("REST API  (RagController)", "POST /documents/ingest   GET /documents/status   POST /rag/query", CYAN),
    ("Services  (DocumentIngestionService, RagQueryService)", "chunking · retrieval · prompt augmentation", TEAL),
    ("Spring AI abstractions", "EmbeddingModel  ·  ChatModel  ·  VectorStore  (provider-agnostic interfaces)", AMBER),
    ("External systems", "OpenAI API  (chat + embeddings)      MariaDB  (vector_store table)", ROSE),
]
ly = Inches(1.95)
lh = Inches(0.92)
for i, (title, sub, color) in enumerate(layers):
    add_rect(s, Inches(0.7), ly, Inches(11.9), lh, fill=CARD_BG, line=color, line_w=Pt(1.5), round_=True)
    add_text(s, Inches(1.0), ly + Inches(0.1), Inches(10.5), Inches(0.4), title, size=16, color=INK, bold=True)
    add_text(s, Inches(1.0), ly + Inches(0.52), Inches(10.9), Inches(0.35), sub, size=12, color=MUTED)
    if i < len(layers) - 1:
        down_arrow(s, Inches(6.35), ly + lh, Inches(0.3), Inches(0.18), color=RGBColor(0xB7,0xC1,0xD1))
    ly += lh + Inches(0.18)
footer(s, 4, "RAG Application Architecture")
set_notes(s, "Each layer only talks to the one directly below it. The Services layer is the only "
             "place that knows this is a 'RAG app' - everything below it is generic Spring AI "
             "plumbing, everything above it is generic REST/HTTP.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "01 · Architecture", "Two pipelines, one shared store")
add_text(s, Inches(0.7), Inches(1.72), Inches(3), Inches(0.35), "INGESTION", size=13, color=TEAL, bold=True)
pipeline_row(s, Inches(0.7), Inches(2.1), Inches(1.85), Inches(1.0), Inches(0.22),
             [("PDF", None, CARD_BG, CARD_LINE), ("Pages", None, CARD_BG, CARD_LINE),
              ("Chunks", None, CARD_BG, CARD_LINE), ("Embed", None, CARD_BG, CARD_LINE),
              ("MariaDB", None, RGBColor(0xE9,0xFB,0xF3), TEAL)], arrow_color=TEAL)

add_text(s, Inches(0.7), Inches(3.55), Inches(3), Inches(0.35), "QUERY", size=13, color=ROSE, bold=True)
pipeline_row(s, Inches(0.7), Inches(3.93), Inches(1.85), Inches(1.0), Inches(0.22),
             [("Question", None, CARD_BG, CARD_LINE), ("Embed", None, CARD_BG, CARD_LINE),
              ("Search\nMariaDB", None, CARD_BG, CARD_LINE), ("Augment\nprompt", None, CARD_BG, CARD_LINE),
              ("LLM answer", None, RGBColor(0xFF,0xEC,0xEF), ROSE)], arrow_color=ROSE)

add_rect(s, Inches(4.35), Inches(2.1), Inches(0.02), Inches(2.83), fill=None)  # spacer no-op
add_text(s, Inches(0.7), Inches(5.15), Inches(11.7), Inches(0.4),
          "Both pipelines call the SAME EmbeddingModel — a question and a policy chunk must land in the same vector space to be comparable.",
          size=14.5, color=INK, bold=True)
add_rect(s, Inches(0.7), Inches(5.75), Inches(11.9), Inches(1.1), fill=CARD_BG, line=CARD_LINE, round_=True)
add_text(s, Inches(1.0), Inches(5.9), Inches(11.2), Inches(0.8),
          "Ingestion runs once (or whenever the source document changes). Query runs on every single "
          "question. That asymmetry is why ingestion can afford to be a bit slower/heavier.",
          size=13.5, color=MUTED, line_spacing=1.2)
footer(s, 5, "RAG Application Architecture")
set_notes(s, "This slide is the one to keep on screen while explaining the rest of the day - every "
             "later topic slots into exactly one of these two rows.")

# ===========================================================================
# 02 — Document Ingestion & Processing
# ===========================================================================
s = section_header("02", "Document Ingestion & Processing", "Turning a PDF file into a list of Document chunks Spring AI can embed.", TEAL)
set_notes(s, "This is the left-hand pipeline from the architecture slide, now with real code.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "02 · Document Ingestion", "DocumentIngestionService.ingest()")
code_box(s, Inches(0.7), Inches(1.95), Inches(7.35), Inches(3.55), (
    "PagePdfDocumentReader pdfReader =\n"
    "    new PagePdfDocumentReader(documentResource);\n"
    "List<Document> pages = pdfReader.get();\n\n"
    "List<Document> chunks =\n"
    "    new TokenTextSplitter().apply(pages);\n\n"
    "vectorStore.add(chunks);\n\n"
    "return chunks.size();"
), title="DocumentIngestionService.java")
add_bullets(s, Inches(8.3), Inches(2.05), Inches(4.3), Inches(3.4),
             [("PagePdfDocumentReader ", "— PDFBox under the hood, one Document per page"),
              ("TokenTextSplitter ", "— re-chunks by token count (Day 2, topic 3)"),
              ("vectorStore.add(chunks) ", "— embeds every chunk and inserts it as a row"),
              ("Result ", "— chunk count returned straight to the caller")],
             size=14.5)
footer(s, 7, "Document Ingestion & Processing")
set_notes(s, "vectorStore.add() is doing two things per chunk under the hood: calling the "
             "EmbeddingModel to get a vector, then inserting {content, metadata, embedding} as a "
             "row. Both steps are hidden behind that one call.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "02 · Document Ingestion", "Exposed as a REST endpoint")
add_rect(s, Inches(0.7), Inches(2.1), Inches(11.9), Inches(1.3), fill=CARD_BG, line=TEAL, line_w=Pt(1.5), round_=True)
add_pill(s, Inches(1.0), Inches(2.35), Inches(1.1), Inches(0.4), "POST", TEAL)
add_text(s, Inches(2.25), Inches(2.35), Inches(6), Inches(0.4), "/api/v1/documents/ingest", size=17, color=INK, bold=True, font="Consolas")
add_text(s, Inches(1.0), Inches(2.85), Inches(10.5), Inches(0.45), "Reads the bundled PDF, chunks it, embeds every chunk, stores it → returns { chunksIndexed }", size=13, color=MUTED)

add_rect(s, Inches(0.7), Inches(3.6), Inches(11.9), Inches(1.3), fill=CARD_BG, line=CARD_LINE, line_w=Pt(1.5), round_=True)
add_pill(s, Inches(1.0), Inches(3.85), Inches(1.1), Inches(0.4), "GET", RGBColor(0x64,0x74,0x8B))
add_text(s, Inches(2.25), Inches(3.85), Inches(6), Inches(0.4), "/api/v1/documents/status", size=17, color=INK, bold=True, font="Consolas")
add_text(s, Inches(1.0), Inches(4.35), Inches(10.5), Inches(0.45), "SELECT COUNT(*) FROM vector_store → { chunksStored } — a quick sanity check the UI polls on load", size=13, color=MUTED)

code_box(s, Inches(0.7), Inches(5.15), Inches(11.9), Inches(1.5),
          "curl -X POST http://localhost:8080/api/v1/documents/ingest\n"
          "curl http://localhost:8080/api/v1/documents/status", size=14)
footer(s, 8, "Document Ingestion & Processing")
set_notes(s, "Deliberately idempotent-ish for a demo: re-running ingest just adds another copy of "
             "the chunks. Good moment to mention that a production version would want a "
             "delete-then-reload or upsert-by-source-hash strategy.")

# ===========================================================================
# 03 — Embeddings & Vector Storage
# ===========================================================================
s = section_header("03", "Embeddings & Vector Storage", "Which model turns text into vectors, and where those vectors actually live.", AMBER)
set_notes(s, "Two config blocks, one for the embedding model, one for the vector store. Both are "
             "pure configuration - no custom code needed for either.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "03 · Embeddings & Vector Storage", "It's almost entirely configuration")
code_box(s, Inches(0.7), Inches(1.95), Inches(5.85), Inches(2.6), (
    "spring:\n"
    "  ai:\n"
    "    openai:\n"
    "      api-key: ${OPENAI_API_KEY}\n"
    "      embedding:\n"
    "        options:\n"
    "          model: text-embedding-3-small"
), title="application.yml — embedding model")
code_box(s, Inches(6.75), Inches(1.95), Inches(5.85), Inches(2.6), (
    "spring:\n"
    "  ai:\n"
    "    vectorstore:\n"
    "      mariadb:\n"
    "        initialize-schema: true\n"
    "        dimensions: 1536\n"
    "        distance-type: COSINE"
), title="application.yml — vector store")
add_bullets(s, Inches(0.7), Inches(4.85), Inches(11.9), Inches(1.9),
             [("No Java code talks to OpenAI or MariaDB directly ", "— DocumentIngestionService and RagQueryService only see the EmbeddingModel / VectorStore interfaces"),
              ("dimensions: 1536 ", "must match the embedding model's output size — text-embedding-3-small produces 1,536-number vectors"),
              ("initialize-schema: true ", "auto-creates the vector_store table (and its VECTOR index) on first startup")],
             size=14.5)
footer(s, 10, "Embeddings & Vector Storage")
set_notes(s, "Worth calling out explicitly: this is the exact spot where the earlier version of the "
             "project plugged in different providers (Ollama, a local ONNX model) without touching "
             "a single line of Java - just swapping this config block. That's the payoff of coding "
             "against Spring AI's interfaces instead of a specific vendor SDK.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "03 · Embeddings & Vector Storage", "What actually lands in the table")
tx, tyy, tw = Inches(0.9), Inches(2.15), Inches(11.2)
cols = [("id", 1.6), ("content", 3.6), ("metadata", 2.8), ("embedding", 3.2)]
header_h = Inches(0.55)
add_rect(s, tx, tyy, tw, header_h, fill=INK, round_=True)
cxp = tx
for name, w in cols:
    add_text(s, cxp + Inches(0.15), tyy, Inches(w) - Inches(0.2), header_h, name, size=14, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    cxp += Inches(w)
sample_rows = [
    ("f8d4…", "“Full-time employees accrue 18 days…”", '{"page":1}', "vector(1536)  [0.021,…]"),
    ("a3e1…", "“…arrangements must be agreed upon…”", '{"page":2}', "vector(1536)  [0.093,…]"),
    ("9b02…", "“…figures, must never be shared…”", '{"page":2}', "vector(1536)  [-0.061,…]"),
]
row_h = Inches(0.62)
ry = tyy + header_h
for r_i, row in enumerate(sample_rows):
    bgc = WHITE if r_i % 2 == 0 else CARD_BG
    add_rect(s, tx, ry, tw, row_h, fill=bgc, line=CARD_LINE, line_w=Pt(0.75))
    cxp = tx
    for (name, w), val in zip(cols, row):
        add_text(s, cxp + Inches(0.15), ry, Inches(w) - Inches(0.25), row_h, val, size=11.5,
                  color=(RGBColor(0xB4,0x5A,0x00) if name=="embedding" else INK),
                  anchor=MSO_ANCHOR.MIDDLE, font=("Consolas" if name in ("id","embedding") else "Calibri"))
        cxp += Inches(w)
    ry += row_h
add_text(s, tx, ry + Inches(0.2), tw, Inches(0.4), "VECTOR KEY vector_store_embedding_idx (embedding)  — the ANN index from Day 2, topic 4",
          size=13, color=MUTED, italic=True)
add_rect(s, tx, ry + Inches(0.7), tw, Inches(1.1), fill=CARD_BG, line=CARD_LINE, round_=True)
add_text(s, tx + Inches(0.25), ry + Inches(0.85), tw - Inches(0.5), Inches(0.8),
          "One row per chunk. The text, its origin metadata, and its vector all live together — "
          "there's no separate embeddings service to keep in sync.",
          size=13.5, color=INK, line_spacing=1.2)
footer(s, 11, "Embeddings & Vector Storage")
set_notes(s, "If you have a terminal handy, this is a great moment to actually run the SELECT query "
             "against ragdb live instead of showing the sketch.")

# ===========================================================================
# 04 — Retrieval & Context Generation
# ===========================================================================
s = section_header("04", "Retrieval & Context Generation", "Finding the right chunks, then handing them to the model as context.", ROSE)
set_notes(s, "This is RagQueryService - the piece that turns a bare question into a grounded prompt.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "04 · Retrieval & Context Generation", "RagQueryService.answerWithSources()")
code_box(s, Inches(0.7), Inches(1.95), Inches(7.5), Inches(4.6), (
    "List<Document> matches = vectorStore.similaritySearch(\n"
    "    SearchRequest.builder()\n"
    "        .query(question)\n"
    "        .topK(topK)\n"
    "        .build());\n\n"
    "String context = matches.stream()\n"
    "    .map(Document::getText)\n"
    "    .reduce(\"\", (a, b) -> a + \"\\n---\\n\" + b);\n\n"
    "String answer = chatClient.prompt()\n"
    "    .system(SYSTEM_PROMPT.formatted(context))\n"
    "    .user(question)\n"
    "    .call()\n"
    "    .content();"
), title="RagQueryService.java")
add_bullets(s, Inches(8.4), Inches(2.05), Inches(4.2), Inches(4.3),
             [("similaritySearch ", "does the embed-and-compare step from Day 2, topic 4 in one call"),
              ("topK ", "= 4 — how many chunks get pulled in as context (config: rag.top-k)"),
              ("System prompt ", "instructs the model to answer ONLY from the supplied context"),
              ("sources[] ", "returned to the caller so the UI can show what backed the answer")],
             size=13.5)
footer(s, 13, "Retrieval & Context Generation")
set_notes(s, "Emphasize the guardrail: the system prompt explicitly tells the model to say 'I don't "
             "know' rather than invent an answer when the context doesn't cover the question. Good "
             "to demo live with an out-of-scope question later.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "04 · Retrieval & Context Generation", "Why write this by hand instead of an Advisor?")
add_bullets(s, Inches(0.7), Inches(2.1), Inches(11.7), Inches(2.4),
             [("Spring AI ships a QuestionAnswerAdvisor ", "that does retrieval + augmentation in one line — great for production code"),
              ("This project writes it out explicitly ", "so every step of \"Retrieval & Context Generation\" is visible in one method, for teaching"),
              ("It also makes sources[] easy ", "to return to the API caller — the advisor version hides the retrieved chunks inside the chat response's internal context")],
             size=17, gap=0.4)
add_rect(s, Inches(0.7), Inches(5.1), Inches(11.9), Inches(1.3), fill=CARD_BG, line=CARD_LINE, round_=True)
add_text(s, Inches(1.0), Inches(5.3), Inches(11.2), Inches(0.9),
          "One-liner version (for reference): chatClient.prompt().advisors(QuestionAnswerAdvisor."
          "builder(vectorStore).build()).user(question).call().content()",
          size=13.5, color=MUTED, font="Consolas", line_spacing=1.2)
footer(s, 14, "Retrieval & Context Generation")
set_notes(s, "Good spot for a judgment-call discussion: advisors are the right default for real "
             "projects; hand-rolling it is a deliberate teaching choice here, not a best practice "
             "recommendation.")

# ===========================================================================
# 05 — LLM Integration
# ===========================================================================
s = section_header("05", "LLM Integration", "One interface, several possible providers behind it.", VIOLET)
set_notes(s, "Tell the real story here if it fits the time: this project tried Claude, then local "
             "Ollama, then a local ONNX embedding model, before settling on OpenAI for both chat and "
             "embeddings - each swap was a config/dependency change, never a code change.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "05 · LLM Integration", "ChatModel is the abstraction — the provider is a dependency")
code_box(s, Inches(0.7), Inches(2.0), Inches(5.7), Inches(1.55), (
    "<dependency>\n"
    "  <groupId>org.springframework.ai</groupId>\n"
    "  <artifactId>\n"
    "    spring-ai-starter-model-openai\n"
    "  </artifactId>\n"
    "</dependency>"
), title="pom.xml", size=12)
code_box(s, Inches(6.55), Inches(2.0), Inches(6.05), Inches(1.55), (
    "chatClient.prompt()\n"
    "    .system(systemPrompt)\n"
    "    .user(question)\n"
    "    .call()\n"
    "    .content();"
), title="RagQueryService.java — never changes", size=13)
add_text(s, Inches(0.7), Inches(3.85), Inches(11.9), Inches(0.4),
          "This app's real history of provider swaps — same code, different dependency + config each time:",
          size=14.5, color=INK, bold=True)
providers = [
    ("Anthropic Claude", "Needs a funded console.anthropic.com API key", AMBER, "tried"),
    ("Local Ollama", "Blocked by corporate endpoint-security policy on this machine", ROSE, "blocked"),
    ("Local ONNX (BAAI/bge)", "Ran into native-library extraction issues on Windows", ROSE, "blocked"),
    ("OpenAI (chat + embeddings)", "One key, one provider, both chat and embeddings — what's running now", TEAL, "in use"),
]
py = Inches(4.35)
for name, desc, color, tag in providers:
    add_rect(s, Inches(0.7), py, Inches(11.9), Inches(0.62), fill=CARD_BG, line=color, line_w=Pt(1.25), round_=True)
    add_pill(s, Inches(0.9), py + Inches(0.1), Inches(1.15), Inches(0.4), tag, color)
    add_text(s, Inches(2.25), py + Inches(0.08), Inches(3.2), Inches(0.46), name, size=13.5, color=INK, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(5.6), py + Inches(0.08), Inches(6.8), Inches(0.46), desc, size=12, color=MUTED, anchor=MSO_ANCHOR.MIDDLE)
    py += Inches(0.72)
footer(s, 16, "LLM Integration")
set_notes(s, "This is a genuinely useful real-world story: LLM integration decisions are often driven "
             "as much by what your network/security policy allows as by model quality. Worth 60 "
             "seconds of honest war-story if the audience is technical.")

# ===========================================================================
# 06 — Building the Backend
# ===========================================================================
s = section_header("06", "Building the Backend using Spring Boot", "A REST controller, two services, and a handful of config properties.", CYAN)
set_notes(s, "Now zoom out from individual features to the whole backend project structure.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "06 · Backend (Spring Boot)", "Project structure")
code_box(s, Inches(0.7), Inches(1.95), Inches(5.9), Inches(4.7), (
    "com.training.ragdemo\n"
    "├── RagDemoApplication.java\n"
    "├── config/\n"
    "│   └── CorsConfig.java\n"
    "├── service/\n"
    "│   ├── DocumentIngestionService.java\n"
    "│   └── RagQueryService.java\n"
    "├── web/\n"
    "│   ├── RagController.java\n"
    "│   └── dto/\n"
    "│       ├── QueryRequest.java\n"
    "│       ├── QueryResponse.java\n"
    "│       ├── SourceChunk.java\n"
    "│       ├── IngestResponse.java\n"
    "│       └── StatusResponse.java\n"
    "resources/\n"
    "├── application.yml\n"
    "└── documents/HR-Company-Policy.pdf"
), title="src/main/…", size=13)
add_bullets(s, Inches(6.9), Inches(2.05), Inches(5.7), Inches(4.4),
             [("web ", "— HTTP in/out only: controller + request/response records"),
              ("service ", "— the actual RAG logic, framework-agnostic-ish"),
              ("config ", "— cross-cutting setup (here: CORS for the React dev server)"),
              ("Records for DTOs ", "keep request/response shapes explicit and immutable"),
              ("No repository/DAO layer ", "needed — VectorStore and JdbcTemplate cover persistence")],
             size=14)
footer(s, 18, "Building the Backend using Spring Boot")
set_notes(s, "Small enough to read top-to-bottom in a couple of minutes — worth actually opening the "
             "IDE here if time allows instead of just showing the tree.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "06 · Backend (Spring Boot)", "The full REST API")
rows = [
    ("POST", "/api/v1/documents/ingest", "—", "{ chunksIndexed }", TEAL),
    ("GET", "/api/v1/documents/status", "—", "{ chunksStored }", RGBColor(0x64,0x74,0x8B)),
    ("POST", "/api/v1/rag/query", '{ "question": "..." }', "{ answer, sources[] }", ROSE),
]
headers = ["Method", "Path", "Request body", "Response"]
tx, tyy = Inches(0.7), Inches(2.1)
colw = [Inches(1.3), Inches(3.4), Inches(3.5), Inches(3.7)]
add_rect(s, tx, tyy, sum(colw, Emu(0)), Inches(0.55), fill=INK, round_=True)
cxp = tx
for h, w in zip(headers, colw):
    add_text(s, cxp + Inches(0.15), tyy, w - Inches(0.2), Inches(0.55), h, size=13, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    cxp += w
ry = tyy + Inches(0.55)
for method, path, body, resp, color in rows:
    rh = Inches(0.85)
    add_rect(s, tx, ry, sum(colw, Emu(0)), rh, fill=CARD_BG, line=CARD_LINE, line_w=Pt(0.75))
    cxp = tx
    add_pill(s, cxp + Inches(0.15), ry + Inches(0.22), Inches(0.9), Inches(0.4), method, color)
    cxp += colw[0]
    add_text(s, cxp + Inches(0.15), ry, colw[1] - Inches(0.25), rh, path, size=13.5, color=INK, bold=True, font="Consolas", anchor=MSO_ANCHOR.MIDDLE)
    cxp += colw[1]
    add_text(s, cxp + Inches(0.15), ry, colw[2] - Inches(0.25), rh, body, size=12, color=MUTED, font="Consolas", anchor=MSO_ANCHOR.MIDDLE)
    cxp += colw[2]
    add_text(s, cxp + Inches(0.15), ry, colw[3] - Inches(0.25), rh, resp, size=12, color=MUTED, font="Consolas", anchor=MSO_ANCHOR.MIDDLE)
    ry += rh
add_text(s, Inches(0.7), ry + Inches(0.3), Inches(11.9), Inches(0.4),
          "CorsConfig allows http://localhost:5173 (the Vite dev server) to call all three — otherwise the browser blocks it.",
          size=13, color=MUTED, italic=True)
footer(s, 19, "Building the Backend using Spring Boot")
set_notes(s, "This table is the contract the React frontend is built against — good bridge into "
             "the next section.")

# ===========================================================================
# 07 — Building the Frontend
# ===========================================================================
s = section_header("07", "Building the Frontend using React", "A single page, two actions: ingest, then ask.", TEAL)
set_notes(s, "Deliberately minimal — this is a Vite + plain React app, no router, no state library, "
             "no UI kit.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "07 · Frontend (React)", "One page, two calls to the backend")
pipeline_row(s, Inches(0.7), Inches(2.0), Inches(2.55), Inches(1.15), Inches(0.35),
             [("Ingest button", "onClick", CARD_BG, CARD_LINE),
              ("POST /documents\n/ingest", None, RGBColor(0xE9,0xFB,0xF3), TEAL),
              ("setIngestMessage", "shows chunk count", CARD_BG, CARD_LINE)], arrow_color=TEAL)
pipeline_row(s, Inches(0.7), Inches(3.55), Inches(2.55), Inches(1.15), Inches(0.35),
             [("Question form", "onSubmit", CARD_BG, CARD_LINE),
              ("POST /rag/query", None, RGBColor(0xFF,0xEC,0xEF), ROSE),
              ("setAnswer + setSources", "renders answer & sources", CARD_BG, CARD_LINE)], arrow_color=ROSE)
add_bullets(s, Inches(0.7), Inches(5.05), Inches(11.9), Inches(1.9),
             [("Plain fetch() ", "— no axios, no query library; two async handlers, two try/catch blocks"),
              ("useState only ", "— chunksStored, ingesting/asking flags, answer, sources, error"),
              ("API_BASE = \"http://localhost:8080\" ", "— a single constant, easy to point at a deployed backend later")],
             size=14)
footer(s, 21, "Building the Frontend using React")
set_notes(s, "The whole App.jsx is under 120 lines — worth mentioning explicitly that a demo/training "
             "UI doesn't need Redux, React Query, or a design system to be useful.")

# ===========================================================================
# 08 — End-to-End Demonstration
# ===========================================================================
s = section_header("08", "End-to-End RAG Application Demonstration", "Everything from today, running live.", AMBER)
set_notes(s, "This is the payoff slide - transition straight into the actual terminal/browser now.")

s = add_slide()
set_bg(s, WHITE)
kicker_and_title(s, "08 · Live Demo", "Run of show")
steps = [
    ("1", "Start the backend", "mvn spring-boot:run   (Spring Boot on :8080, connects to MariaDB)"),
    ("2", "Start the frontend", "cd rag-demo-ui && npm run dev   (Vite dev server on :5173)"),
    ("3", "Ingest the policy", "Click “Ingest HR Policy PDF” → watch chunksIndexed come back"),
    ("4", "Ask an in-scope question", "“How many days of annual leave do employees get?” → grounded answer + sources"),
    ("5", "Ask an out-of-scope question", "“What's the company's stock price?” → model says it doesn't know — proves it's grounded, not guessing"),
]
y = Inches(1.95)
for num, title, desc in steps:
    add_rect(s, Inches(0.7), y, Inches(0.65), Inches(0.85), fill=INK, round_=True)
    add_text(s, Inches(0.7), y, Inches(0.65), Inches(0.85), num, size=22, color=WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(1.55), y + Inches(0.02), Inches(4), Inches(0.4), title, size=15.5, color=INK, bold=True)
    add_text(s, Inches(1.55), y + Inches(0.42), Inches(10.7), Inches(0.4), desc, size=12.5, color=MUTED)
    y += Inches(0.98)
footer(s, 23, "End-to-End RAG Application Demonstration")
set_notes(s, "Step 5 is the most convincing moment for a skeptical audience - it demonstrates the "
             "system prompt guardrail from topic 04 actually working, not just retrieval succeeding.")

# ---------------------------------------------------------------------------
# Closing
# ---------------------------------------------------------------------------
s = add_slide()
set_bg(s, NAVY)
kicker_and_title(s, "Recap", "One PDF. One pipeline. Two days.", dark=True)
fy = Inches(2.5)
fw, fh = Inches(1.85), Inches(1.05)
gap4 = Inches(0.22)
pipeline = [
    ("PDF", CYAN), ("Chunks", AMBER), ("Embeddings", TEAL), ("MariaDB", TEAL),
    ("Question", ROSE), ("Top-K", ROSE), ("LLM", VIOLET), ("Answer", VIOLET),
]
x = Inches(0.5)
for i, (title, color) in enumerate(pipeline):
    shp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, fy, fw, fh)
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
    r.font.size = Pt(12.5); r.font.bold = True; r.font.color.rgb = WHITE
    if i < len(pipeline) - 1:
        right_arrow(s, x + fw + Inches(0.02), fy + fh/2 - Inches(0.11), gap4 - Inches(0.04), Inches(0.22), color=RGBColor(0x4B,0x5A,0x78))
    x += fw + gap4

add_rect(s, Inches(0.7), Inches(4.05), Inches(11.9), Inches(1.9), fill=NAVY_2, round_=True)
add_bullets(s, Inches(1.0), Inches(4.3), Inches(11.3), Inches(1.5),
             ["Architecture: two pipelines (ingest, query) sharing one vector store",
              "Ingestion & embeddings: PDF → chunks → vectors, almost entirely via config",
              "Retrieval & LLM integration: similaritySearch() + ChatClient, provider-agnostic",
              "Backend + frontend: a small REST API and a small React page — nothing more than needed"],
             size=14.5, color=WHITE, marker_color=CYAN, gap=0.16)
add_text(s, Inches(0.7), Inches(6.2), Inches(10), Inches(0.5), "Questions?", size=22, color=CYAN, bold=True)
footer(s, 24, "Recap", dark=True)
set_notes(s, "Close by pointing back at the Day 2 deck's closing slide - it's literally the same "
             "pipeline, and today filled in every box with real, running code.")

# ---------------------------------------------------------------------------
out_path = os.path.join(os.path.dirname(__file__), "..", "Day3-Building-RAG-Application.pptx")
prs.save(out_path)
print(f"Wrote {os.path.abspath(out_path)}  ({len(prs.slides._sldIdLst)} slides)")
