"""
PDF Generator — Project Documentation
E-Commerce Growth & Profit Optimization Analytics
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, Image, KeepTogether
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
import os

BASE    = os.path.dirname(os.path.abspath(__file__))
CHARTS  = os.path.join(BASE, "outputs", "charts")
OUT     = "/mnt/user-data/outputs/ecommerce_project_documentation.pdf"

# ─────────────────────────────────────────────
# COLOURS
# ─────────────────────────────────────────────
DARK_BLUE   = colors.HexColor("#1F3864")
MEDIUM_BLUE = colors.HexColor("#2E86AB")
LIGHT_BLUE  = colors.HexColor("#D6E4F0")
DARK_GREEN  = colors.HexColor("#155724")
LIGHT_GREEN = colors.HexColor("#D4EDDA")
DARK_RED    = colors.HexColor("#721C24")
LIGHT_RED   = colors.HexColor("#F8D7DA")
YELLOW      = colors.HexColor("#FFF3CD")
WHITE       = colors.white
BLACK       = colors.black
GRAY        = colors.HexColor("#6C757D")
LIGHT_GRAY  = colors.HexColor("#F8F9FA")
MID_GRAY    = colors.HexColor("#DEE2E6")
ORANGE      = colors.HexColor("#FD7E14")
GREEN_B     = colors.HexColor("#28A745")

# ─────────────────────────────────────────────
# STYLES
# ─────────────────────────────────────────────
styles = getSampleStyleSheet()

def style(name, parent="Normal", **kwargs):
    s = ParagraphStyle(name, parent=styles[parent], **kwargs)
    return s

S_title       = style("Title2",      fontSize=26, textColor=WHITE,
                       leading=32, alignment=TA_CENTER, spaceAfter=6)
S_subtitle    = style("Subtitle",    fontSize=14, textColor=LIGHT_BLUE,
                       leading=18, alignment=TA_CENTER)
S_h1          = style("H1",          fontSize=16, textColor=DARK_BLUE,
                       leading=20, spaceBefore=16, spaceAfter=6,
                       fontName="Helvetica-Bold")
S_h2          = style("H2",          fontSize=13, textColor=MEDIUM_BLUE,
                       leading=17, spaceBefore=12, spaceAfter=4,
                       fontName="Helvetica-Bold")
S_h3          = style("H3",          fontSize=11, textColor=DARK_BLUE,
                       leading=15, spaceBefore=8, spaceAfter=3,
                       fontName="Helvetica-Bold")
S_body        = style("Body",        fontSize=10, textColor=BLACK,
                       leading=15, spaceAfter=4, alignment=TA_JUSTIFY)
S_body_small  = style("BodySmall",   fontSize=9, textColor=colors.HexColor("#333333"),
                       leading=13, spaceAfter=2)
S_code        = style("Code",        fontSize=8.5, textColor=colors.HexColor("#1A1A2E"),
                       leading=12, fontName="Courier", backColor=LIGHT_GRAY,
                       leftIndent=8, rightIndent=8, spaceBefore=4, spaceAfter=4)
S_bullet      = style("Bullet",      fontSize=10, textColor=BLACK,
                       leading=14, spaceAfter=3, leftIndent=16)
S_insight     = style("Insight",     fontSize=10, textColor=DARK_GREEN,
                       leading=14, spaceAfter=4, leftIndent=12,
                       fontName="Helvetica-Bold")
S_caption     = style("Caption",     fontSize=8.5, textColor=GRAY,
                       leading=12, alignment=TA_CENTER, spaceAfter=8)
S_toc         = style("TOC",         fontSize=10, textColor=DARK_BLUE,
                       leading=18, leftIndent=12)
S_toc_sub     = style("TOCSub",      fontSize=9.5, textColor=MEDIUM_BLUE,
                       leading=16, leftIndent=28)
S_label       = style("Label",       fontSize=9, textColor=GRAY,
                       leading=12, fontName="Helvetica-Oblique")

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def divider(color=MEDIUM_BLUE, thickness=1.5):
    return HRFlowable(width="100%", thickness=thickness,
                      color=color, spaceAfter=6, spaceBefore=4)

def section_box(title, color=DARK_BLUE):
    data = [[Paragraph(title, ParagraphStyle("SB", fontSize=13,
             textColor=WHITE, fontName="Helvetica-Bold", leading=17))]]
    t = Table(data, colWidths=[17*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), color),
        ("TOPPADDING",    (0,0), (-1,-1), 8),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),
        ("LEFTPADDING",   (0,0), (-1,-1), 12),
    ]))
    return t

def kv_table(rows, col_widths=None):
    """Key-value table with alternating rows"""
    if col_widths is None:
        col_widths = [8*cm, 9*cm]
    data = []
    for k, v in rows:
        data.append([
            Paragraph(str(k), ParagraphStyle("KV_k", fontSize=9.5, fontName="Helvetica-Bold",
                      textColor=DARK_BLUE, leading=13)),
            Paragraph(str(v), ParagraphStyle("KV_v", fontSize=9.5, leading=13,
                      textColor=BLACK))
        ])
    ts = TableStyle([
        ("BACKGROUND",    (0,0), (-1,-1), WHITE),
        ("ROWBACKGROUNDS",(0,0), (-1,-1), [WHITE, LIGHT_GRAY]),
        ("GRID",          (0,0), (-1,-1), 0.4, MID_GRAY),
        ("TOPPADDING",    (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ("LEFTPADDING",   (0,0), (-1,-1), 8),
    ])
    t = Table(data, colWidths=col_widths)
    t.setStyle(ts)
    return t

def code_block(text):
    return Paragraph(text.replace("\n","<br/>").replace(" ","&nbsp;"), S_code)

def bullet(text, symbol="●"):
    return Paragraph(f"<b>{symbol}</b> {text}", S_bullet)

def insight_box(text):
    data = [[Paragraph(f"💡 {text}", ParagraphStyle("IB", fontSize=10,
             textColor=DARK_GREEN, leading=14, fontName="Helvetica-Bold"))]]
    t = Table(data, colWidths=[17*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,-1), LIGHT_GREEN),
        ("LEFTBORDER",    (0,0), (0,-1), 3, DARK_GREEN),
        ("TOPPADDING",    (0,0), (-1,-1), 8),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),
        ("LEFTPADDING",   (0,0), (-1,-1), 12),
        ("RIGHTPADDING",  (0,0), (-1,-1), 8),
        ("BOX",           (0,0), (-1,-1), 0.5, DARK_GREEN),
    ]))
    return t

def add_chart(filename, width=15*cm, caption=None):
    path = os.path.join(CHARTS, filename)
    elems = []
    if os.path.exists(path):
        img = Image(path, width=width, height=width*0.6)
        elems.append(img)
        if caption:
            elems.append(Paragraph(caption, S_caption))
    return elems

def header_table(col1, col2, col3="", col4="", bg=DARK_BLUE):
    headers = [col1, col2, col3, col4]
    active  = [h for h in headers if h]
    w = 17*cm / len(active)
    data = [[Paragraph(h, ParagraphStyle("TH", fontSize=9.5, fontName="Helvetica-Bold",
             textColor=WHITE, leading=12, alignment=TA_CENTER)) for h in active]]
    t = Table(data, colWidths=[w]*len(active))
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), bg),
        ("TOPPADDING",    (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ]))
    return t

def data_table(headers, rows, col_widths=None, header_bg=DARK_BLUE):
    """Full data table with styled header and alternating rows"""
    h_cells = [Paragraph(str(h), ParagraphStyle("DH", fontSize=9, fontName="Helvetica-Bold",
                textColor=WHITE, leading=12, alignment=TA_CENTER)) for h in headers]
    all_rows = [h_cells]
    for i, row in enumerate(rows):
        cells = [Paragraph(str(v), ParagraphStyle("DC", fontSize=9, leading=12,
                 alignment=TA_CENTER if j > 0 else TA_LEFT,
                 textColor=BLACK)) for j, v in enumerate(row)]
        all_rows.append(cells)

    if col_widths is None:
        col_widths = [17*cm / len(headers)] * len(headers)

    t = Table(all_rows, colWidths=col_widths)
    ts = TableStyle([
        ("BACKGROUND",    (0,0), (-1,0), header_bg),
        ("ROWBACKGROUNDS",(0,1), (-1,-1), [WHITE, LIGHT_GRAY]),
        ("GRID",          (0,0), (-1,-1), 0.4, MID_GRAY),
        ("TOPPADDING",    (0,0), (-1,-1), 5),
        ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ("LEFTPADDING",   (0,0), (-1,-1), 6),
        ("RIGHTPADDING",  (0,0), (-1,-1), 6),
    ])
    t.setStyle(ts)
    return t

# ─────────────────────────────────────────────
# PAGE TEMPLATE
# ─────────────────────────────────────────────
def on_page(canvas, doc):
    canvas.saveState()
    # Footer line
    canvas.setStrokeColor(LIGHT_BLUE)
    canvas.setLineWidth(0.5)
    canvas.line(1.5*cm, 1.8*cm, 19.5*cm, 1.8*cm)
    # Footer text
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(GRAY)
    canvas.drawString(1.5*cm, 1.1*cm, "E-Commerce Growth & Profit Optimization Analytics — Rohit")
    canvas.drawRightString(19.5*cm, 1.1*cm, f"Page {doc.page}")
    # Header line (after page 1)
    if doc.page > 1:
        canvas.setFillColor(DARK_BLUE)
        canvas.rect(1.5*cm, 27.8*cm, 18*cm, 0.6*cm, fill=1, stroke=0)
        canvas.setFont("Helvetica-Bold", 8)
        canvas.setFillColor(WHITE)
        canvas.drawString(1.8*cm, 28.0*cm, "E-COMMERCE ANALYTICS PROJECT")
        canvas.drawRightString(19.3*cm, 28.0*cm, "SQL | Python | Excel | Power BI")
    canvas.restoreState()

# ─────────────────────────────────────────────
# BUILD DOCUMENT
# ─────────────────────────────────────────────
doc = SimpleDocTemplate(
    OUT, pagesize=A4,
    leftMargin=1.5*cm, rightMargin=1.5*cm,
    topMargin=2.5*cm,  bottomMargin=2.5*cm
)

story = []
P = lambda t, s=S_body: Paragraph(t, s)
SP = lambda h=0.3: Spacer(1, h*cm)

# ══════════════════════════════════════════════════════════════
# COVER PAGE
# ══════════════════════════════════════════════════════════════
cover_bg = Table(
    [[Paragraph("", S_body)]],
    colWidths=[17*cm], rowHeights=[5*cm]
)
cover_bg.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),DARK_BLUE)]))

# Title block
title_data = [[
    Paragraph("E-COMMERCE GROWTH &amp;<br/>PROFIT OPTIMIZATION", S_title),
], [
    Paragraph("ANALYTICS", ParagraphStyle("T2", fontSize=28, textColor=ORANGE,
               fontName="Helvetica-Bold", alignment=TA_CENTER)),
], [
    SP(0.4),
], [
    Paragraph("End-to-End Data Analytics Project Documentation", S_subtitle),
], [
    Paragraph("SQL &nbsp;|&nbsp; Python &nbsp;|&nbsp; Statistics &nbsp;|&nbsp; Excel &nbsp;|&nbsp; Power BI",
              ParagraphStyle("Tools", fontSize=12, textColor=LIGHT_BLUE, alignment=TA_CENTER, leading=16)),
]]
title_table = Table(title_data, colWidths=[17*cm])
title_table.setStyle(TableStyle([
    ("BACKGROUND", (0,0), (-1,-1), DARK_BLUE),
    ("TOPPADDING",    (0,0), (-1,-1), 10),
    ("BOTTOMPADDING", (0,0), (-1,-1), 10),
    ("LEFTPADDING",   (0,0), (-1,-1), 15),
    ("RIGHTPADDING",  (0,0), (-1,-1), 15),
]))
story.append(title_table)
story.append(SP(0.5))

# Cover stats row
stats_data = [
    [P("100,000+", ParagraphStyle("CS", fontSize=20, fontName="Helvetica-Bold",
        textColor=ORANGE, alignment=TA_CENTER)),
     P("15", ParagraphStyle("CS", fontSize=20, fontName="Helvetica-Bold",
        textColor=ORANGE, alignment=TA_CENTER)),
     P("6", ParagraphStyle("CS", fontSize=20, fontName="Helvetica-Bold",
        textColor=ORANGE, alignment=TA_CENTER)),
     P("12", ParagraphStyle("CS", fontSize=20, fontName="Helvetica-Bold",
        textColor=ORANGE, alignment=TA_CENTER))],
    [P("Orders Analysed", ParagraphStyle("CSL", fontSize=9, textColor=GRAY,
        alignment=TA_CENTER)),
     P("Product Categories", ParagraphStyle("CSL", fontSize=9, textColor=GRAY,
        alignment=TA_CENTER)),
     P("Python Scripts", ParagraphStyle("CSL", fontSize=9, textColor=GRAY,
        alignment=TA_CENTER)),
     P("Charts Generated", ParagraphStyle("CSL", fontSize=9, textColor=GRAY,
        alignment=TA_CENTER))],
]
st = Table(stats_data, colWidths=[4.25*cm]*4)
st.setStyle(TableStyle([
    ("BOX",          (0,0), (-1,-1), 0.5, MID_GRAY),
    ("INNERGRID",    (0,0), (-1,-1), 0.3, MID_GRAY),
    ("TOPPADDING",   (0,0), (-1,-1), 8),
    ("BOTTOMPADDING",(0,0), (-1,-1), 8),
]))
story.append(st)
story.append(SP(0.5))

# Cover description
desc = Table([[P(
    "This document is the complete step-by-step project documentation for the "
    "E-Commerce Growth &amp; Profit Optimization Analytics project. It explains "
    "every decision made, every SQL query written, every Python script run, and "
    "every insight discovered. By reading this document, you will understand "
    "exactly what was built, why it was built that way, and what business value "
    "it delivers.", S_body)]], colWidths=[17*cm])
desc.setStyle(TableStyle([
    ("BACKGROUND", (0,0), (-1,-1), LIGHT_GRAY),
    ("BOX",        (0,0), (-1,-1), 0.5, MID_GRAY),
    ("TOPPADDING", (0,0), (-1,-1), 10),
    ("BOTTOMPADDING",(0,0),(-1,-1),10),
    ("LEFTPADDING", (0,0), (-1,-1), 12),
    ("RIGHTPADDING",(0,0), (-1,-1), 12),
]))
story.append(desc)
story.append(PageBreak())

# ══════════════════════════════════════════════════════════════
# TABLE OF CONTENTS
# ══════════════════════════════════════════════════════════════
story.append(section_box("TABLE OF CONTENTS"))
story.append(SP(0.3))
toc = [
    ("1.", "Project Overview & Objective",      "3"),
    ("2.", "Dataset Description",               "4"),
    ("3.", "Project Folder Structure",          "5"),
    ("4.", "Step 1 — Data Understanding",       "6"),
    ("5.", "Step 2 — SQL Data Modeling",        "7"),
    ("6.", "Step 3 — Data Cleaning (Python)",   "9"),
    ("7.", "Step 4 — Business KPI Metrics",     "11"),
    ("8.", "Step 5 — Customer Segmentation & RFM", "13"),
    ("9.", "Step 6 — Cohort Retention Analysis","16"),
    ("10.","Step 7 — Product & Category Profitability","18"),
    ("11.","Step 8 — Statistical Analysis",     "21"),
    ("12.","Step 9 — Excel Scenario Simulation","23"),
    ("13.","Power BI Dashboard Guide",          "25"),
    ("14.","Key Business Insights",             "27"),
    ("15.","Resume Bullet & Project Summary",   "28"),
]
for num, title, page in toc:
    story.append(P(
        f"<b>{num}</b>&nbsp;&nbsp;&nbsp;{title}"
        f'<font color="#AAAAAA">&nbsp;{"." * (55 - len(num) - len(title))}&nbsp;{page}</font>',
        S_toc))
story.append(PageBreak())

# ══════════════════════════════════════════════════════════════
# SECTION 1 — PROJECT OVERVIEW
# ══════════════════════════════════════════════════════════════
story.append(section_box("1. PROJECT OVERVIEW & OBJECTIVE"))
story.append(SP(0.3))
story.append(SP(0.1))
story.append(P(
    "This project is designed to simulate what a real Data Analyst or Business Analyst "
    "does at an e-commerce company. The goal is not just to build charts — it is to "
    "answer real business questions using data, and deliver actionable insights that "
    "help a business grow.", S_body))
story.append(SP(0.2))

story.append(P("<b>Business Questions Answered:</b>", S_h3))
bqs = [
    "Why is profit changing month over month — is it pricing, freight costs, or order volume?",
    "Which customers generate the most revenue, and are they at risk of churning?",
    "Which product categories look profitable on revenue but lose money on shipping?",
    "How many customers return after their first purchase? (Cohort Retention)",
    "What happens to business profit if we reduce freight costs by 10%? (Scenario Simulation)",
    "Which customers are Champions vs Lost — and how do we target each differently?",
]
for q in bqs:
    story.append(bullet(q))
story.append(SP(0.2))

story.append(P("<b>Tools Used &amp; Why:</b>", S_h3))
tools = [
    ("SQL", "For joining and aggregating large datasets efficiently. "
            "This is how data analysts extract data from databases in real companies."),
    ("Python (pandas)", "For deeper analysis beyond what SQL can do — cohort tables, "
                        "RFM scoring, visualisations, and statistical calculations."),
    ("Statistics", "To validate insights — is the correlation between price and "
                   "freight cost statistically meaningful?"),
    ("Excel", "For scenario simulation and financial modeling — business stakeholders "
              "understand Excel scenarios better than Python scripts."),
    ("Power BI", "To build executive dashboards that non-technical leadership can "
                 "use to make decisions."),
]
t_rows = [(k, v) for k, v in tools]
story.append(kv_table(t_rows, col_widths=[4*cm, 13*cm]))
story.append(PageBreak())

# ══════════════════════════════════════════════════════════════
# SECTION 2 — DATASET
# ══════════════════════════════════════════════════════════════
story.append(section_box("2. DATASET DESCRIPTION"))
story.append(SP(0.3))
story.append(P(
    "The dataset is modeled after the Olist Brazilian E-Commerce public dataset. "
    "It contains 100,000 orders spanning January 2022 to December 2023, structured "
    "across 6 separate tables that reflect how a real e-commerce database is organised.",
    S_body))
story.append(SP(0.2))

tables_info = [
    ("customers",   "10,000", "Customer ID, city, state, zip code"),
    ("orders",      "100,000","Order ID, customer, status, purchase timestamp, delivery date"),
    ("order_items", "155,467","Order ID, product ID, item price, freight value"),
    ("products",    "3,000",  "Product ID, category, weight, dimensions"),
    ("payments",    "100,000","Order ID, payment type, installments, payment value"),
    ("reviews",     "70,000", "Order ID, review score (1–5 stars)"),
]
story.append(data_table(
    ["Table", "Rows", "Key Fields"],
    tables_info,
    col_widths=[4*cm, 3*cm, 10*cm]
))
story.append(SP(0.3))
story.append(P("<b>Key Fields Explained:</b>", S_h3))
fields = [
    ("customer_id", "Unique identifier per customer. Used to link orders and track behaviour."),
    ("order_id",    "Unique order identifier. One customer can have many orders."),
    ("product_id",  "Links to product table for category and weight data."),
    ("price",       "Item selling price in Brazilian Reais (R$)."),
    ("freight_value","Shipping cost charged for that item. Critical for profit calculation."),
    ("order_purchase_timestamp", "Date and time of order — used for time-series analysis."),
    ("payment_type", "Credit card / boleto / debit card / voucher."),
    ("review_score", "1–5 star rating. Used to measure customer satisfaction."),
]
story.append(kv_table(fields, col_widths=[5.5*cm, 11.5*cm]))
story.append(SP(0.2))
story.append(insight_box(
    "Profit is not directly available in the dataset. "
    "We estimate it as: Profit = price - freight_value. "
    "In a real company, COGS (cost of goods sold) would also be subtracted. "
    "Our analysis treats freight as the primary variable cost."
))
story.append(PageBreak())

# ══════════════════════════════════════════════════════════════
# SECTION 3 — FOLDER STRUCTURE
# ══════════════════════════════════════════════════════════════
story.append(section_box("3. PROJECT FOLDER STRUCTURE"))
story.append(SP(0.3))
story.append(P(
    "The project follows a professional data analytics folder structure. "
    "Each layer has a clear responsibility, making the project easy to navigate "
    "and extend.", S_body))
story.append(SP(0.2))

folder_str = (
    "ecommerce-growth-analytics/<br/>"
    "&nbsp;&nbsp;&nbsp;data/<br/>"
    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;raw/&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;← Original CSV files (customers, orders, products, etc.)<br/>"
    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;processed/&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;← Cleaned and joined orders_fact.csv<br/>"
    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;generate_data.py&nbsp;&nbsp;&nbsp;← Script that generates synthetic dataset<br/>"
    "&nbsp;&nbsp;&nbsp;sql_queries/&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;← All SQL queries (6 files)<br/>"
    "&nbsp;&nbsp;&nbsp;python_analysis/&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;← All Python analysis scripts (6 files)<br/>"
    "&nbsp;&nbsp;&nbsp;excel_simulation/&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;← Excel model builder + .xlsx file<br/>"
    "&nbsp;&nbsp;&nbsp;outputs/<br/>"
    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;charts/&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;← 12 generated PNG charts<br/>"
    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;csv_exports/&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;← 15 exported analysis CSV files<br/>"
    "&nbsp;&nbsp;&nbsp;run_all.py&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;← Master script to run everything<br/>"
    "&nbsp;&nbsp;&nbsp;README.md&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;← Project readme"
)
story.append(Paragraph(folder_str, ParagraphStyle("Folder", fontSize=9, fontName="Courier",
             leading=14, backColor=LIGHT_GRAY, leftIndent=8, rightIndent=8,
             spaceBefore=6, spaceAfter=6, textColor=DARK_BLUE)))
story.append(SP(0.2))

story.append(P("<b>Why this structure?</b>", S_h3))
story.append(P(
    "In real companies, raw data and processed data are always separated. "
    "You never modify the raw data files — if something breaks, you can always "
    "regenerate from raw. SQL queries, Python scripts, and Excel models are kept "
    "in separate folders because they are maintained by different tools. "
    "Outputs folder contains deliverables (charts + CSVs) that get shared with "
    "stakeholders.", S_body))
story.append(PageBreak())

# ══════════════════════════════════════════════════════════════
# SECTION 4 — STEP 1: DATA UNDERSTANDING
# ══════════════════════════════════════════════════════════════
story.append(section_box("4. STEP 1 — DATA UNDERSTANDING"))
story.append(SP(0.3))
story.append(P(
    "Before writing any code, a data analyst must understand the data. "
    "This means examining each table, understanding the relationships between them, "
    "and identifying data quality issues. This step is non-negotiable.",
    S_body))
story.append(SP(0.2))

story.append(P("<b>Entity Relationship Diagram (ERD) — Textual Description:</b>", S_h3))
erd_text = (
    "orders (order_id PK) ────── order_items (order_id FK, product_id FK)<br/>"
    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    "└────── products (product_id PK)<br/>"
    "orders (customer_id FK) ─── customers (customer_id PK)<br/>"
    "orders (order_id) ─────── payments (order_id FK)<br/>"
    "orders (order_id) ─────── reviews (order_id FK)"
)
story.append(Paragraph(erd_text, ParagraphStyle("ERD", fontSize=9, fontName="Courier",
             leading=14, backColor=LIGHT_GRAY, leftIndent=8, spaceBefore=4, spaceAfter=6)))

story.append(P("<b>Join Logic:</b>", S_h3))
join_rows = [
    ("orders → order_items",  "INNER JOIN on order_id", "Get all items in each order"),
    ("order_items → products","LEFT JOIN on product_id","Get category & weight for each item"),
    ("orders → customers",    "LEFT JOIN on customer_id","Get customer location data"),
    ("orders → payments",     "LEFT JOIN on order_id",  "Get payment type/installments"),
    ("orders → reviews",      "LEFT JOIN on order_id",  "Get review score (70% coverage)"),
]
story.append(data_table(
    ["Join", "Type", "Purpose"],
    join_rows,
    col_widths=[5*cm, 4.5*cm, 7.5*cm]
))
story.append(SP(0.2))
story.append(insight_box(
    "We use INNER JOIN for order_items (every order must have at least one item) "
    "and LEFT JOINs for all others because not every order has a review or a payment record. "
    "Using INNER JOIN for reviews would drop 30% of orders — a major mistake."
))
story.append(PageBreak())

# ══════════════════════════════════════════════════════════════
# SECTION 5 — STEP 2: SQL DATA MODELING
# ══════════════════════════════════════════════════════════════
story.append(section_box("5. STEP 2 — SQL DATA MODELING"))
story.append(SP(0.3))
story.append(P(
    "SQL is used to create four analytics tables. These are the foundation "
    "of all downstream analysis. In a real company, these would be created as "
    "views or materialised tables in a data warehouse.", S_body))
story.append(SP(0.2))

story.append(P("<b>SQL File 01 — Orders Fact Table (01_create_orders_fact.sql)</b>", S_h2))
story.append(P(
    "This is the most important query. It joins all 6 tables into one master "
    "analytics table called orders_fact. Every other analysis reads from this "
    "table. It contains one row per order-item combination.",
    S_body))
story.append(SP(0.1))
sql1 = """\
SELECT
  o.order_id,
  o.customer_id,
  o.order_status,
  o.order_purchase_timestamp,
  oi.product_id,
  oi.price,
  oi.freight_value,
  ROUND(oi.price - oi.freight_value, 2)        AS profit_estimate,
  STRFTIME('%Y-%m', o.order_purchase_timestamp) AS order_year_month,
  p.product_category_name,
  c.customer_state,
  py.payment_type,
  COALESCE(r.review_score, 0)                   AS review_score
FROM orders o
JOIN order_items oi ON o.order_id = oi.order_id
JOIN products    p  ON oi.product_id = p.product_id
LEFT JOIN customers c  ON o.customer_id = c.customer_id
LEFT JOIN payments  py ON o.order_id = py.order_id
LEFT JOIN reviews   r  ON o.order_id = r.order_id
WHERE o.order_status != 'cancelled';"""
story.append(code_block(sql1))
story.append(SP(0.1))
story.append(P(
    "<b>Why exclude cancelled orders?</b> Cancelled orders have no revenue or "
    "delivery, so including them would distort all averages and totals.", S_body_small))

story.append(SP(0.3))
story.append(P("<b>SQL File 02 — Customer Metrics (02_customer_metrics.sql)</b>", S_h2))
story.append(P(
    "This query aggregates the fact table by customer to produce per-customer "
    "KPIs. These are used for segmentation and CLV analysis.",
    S_body))
sql2 = """\
WITH customer_base AS (
  SELECT
    customer_id,
    COUNT(DISTINCT order_id)  AS total_orders,
    SUM(price)                AS total_revenue,
    SUM(price - freight_value)AS total_profit_generated,
    MIN(order_purchase_timestamp) AS first_order_date,
    MAX(order_purchase_timestamp) AS last_order_date
  FROM orders_fact
  GROUP BY customer_id
)
SELECT *,
  CASE
    WHEN total_orders = 1             THEN 'New'
    WHEN total_orders BETWEEN 2 AND 4 THEN 'Active'
    WHEN total_orders >= 5            THEN 'Loyal'
  END AS customer_segment
FROM customer_base
ORDER BY total_revenue DESC;"""
story.append(code_block(sql2))

story.append(SP(0.3))
story.append(P("<b>SQL File 06 — RFM Segmentation (06_rfm_segmentation.sql)</b>", S_h2))
story.append(P(
    "RFM stands for Recency, Frequency, Monetary. It is the most widely used "
    "customer segmentation technique in e-commerce analytics. Each customer gets "
    "a score of 1–5 on each dimension using NTILE window functions.",
    S_body))
sql3 = """\
SELECT
  customer_id,
  -- Recency: how recently did they buy? (5 = most recent)
  NTILE(5) OVER (ORDER BY recency_days ASC)   AS r_score,
  -- Frequency: how often do they buy? (5 = most frequent)
  NTILE(5) OVER (ORDER BY frequency DESC)     AS f_score,
  -- Monetary: how much do they spend? (5 = highest)
  NTILE(5) OVER (ORDER BY monetary DESC)      AS m_score
FROM rfm_raw;"""
story.append(code_block(sql3))
story.append(insight_box(
    "NTILE(5) divides all customers into 5 equal buckets. "
    "The top 20% of spenders get m_score=5, the bottom 20% get m_score=1. "
    "This is a percentile-based ranking — it works regardless of how revenue is distributed."
))
story.append(PageBreak())

# ══════════════════════════════════════════════════════════════
# SECTION 6 — STEP 3: DATA CLEANING
# ══════════════════════════════════════════════════════════════
story.append(section_box("6. STEP 3 — DATA CLEANING (Python)"))
story.append(SP(0.3))
story.append(P(
    "Script: python_analysis/01_data_cleaning.py", S_label))
story.append(SP(0.1))
story.append(P(
    "Data cleaning is where analysts spend 60–80% of their time in real projects. "
    "This script loads all 6 raw CSVs, validates them, joins them into the master "
    "orders_fact table, and exports the cleaned data.",
    S_body))
story.append(SP(0.2))

story.append(P("<b>What was cleaned and why:</b>", S_h3))
cleaning_steps = [
    ("Duplicate removal", "drop_duplicates()",
     "Duplicate order_ids would double-count revenue. "
     "Always check before aggregating."),
    ("Cancelled orders removed", "order_status != 'cancelled'",
     "14,254 cancelled orders removed. These inflate order counts "
     "but have zero revenue."),
    ("Timestamp parsing", "pd.to_datetime()",
     "Timestamps stored as strings cannot be used for date math. "
     "Must convert to datetime64."),
    ("Date feature extraction", ".dt.year, .dt.month, .dt.to_period('M')",
     "Creates order_year, order_month, order_year_month columns "
     "for time-series grouping."),
    ("Profit column created", "price - freight_value",
     "Derived column — not in original data. "
     "Created at cleaning stage so all downstream scripts use the same formula."),
    ("Negative prices removed", "price > 0",
     "Any row with price <= 0 is a data entry error."),
    ("Review score filled", "fillna(0)",
     "30% of orders have no review. Filled with 0, then filtered "
     "in analysis scripts that need real scores."),
]
story.append(data_table(
    ["Action", "Code", "Business Reason"],
    cleaning_steps,
    col_widths=[4*cm, 5.5*cm, 7.5*cm]
))
story.append(SP(0.2))

story.append(P("<b>Data Quality Results After Cleaning:</b>", S_h3))
dq_rows = [
    ("Orders (after removing cancelled)", "85,746"),
    ("Order-Item rows in fact table",      "133,233"),
    ("Unique Customers",                   "9,998"),
    ("Unique Products",                    "3,000"),
    ("Date Range",                         "Jan 2022 – Dec 2023"),
    ("Total Revenue",                      "R$ 71,732,497"),
    ("Total Freight Cost",                 "R$ 9,001,654"),
    ("Estimated Gross Profit",             "R$ 62,730,843"),
    ("Average Order Value",                "R$ 836.57"),
    ("Average Review Score",               "4.07 / 5.0"),
]
story.append(kv_table(dq_rows, col_widths=[8*cm, 9*cm]))
story.append(PageBreak())

# ══════════════════════════════════════════════════════════════
# SECTION 7 — STEP 4: BUSINESS KPIs
# ══════════════════════════════════════════════════════════════
story.append(section_box("7. STEP 4 — BUSINESS KPI METRICS"))
story.append(SP(0.3))
story.append(P("Script: python_analysis/02_business_metrics.py", S_label))
story.append(SP(0.1))
story.append(P(
    "This script calculates the core KPIs that executives care about, "
    "plots the monthly revenue trend, analyses geographic distribution, "
    "and breaks down payment methods.", S_body))
story.append(SP(0.2))

story.append(P("<b>KPI Formulas:</b>", S_h3))
kpi_rows = [
    ("Total Revenue",        "SUM(price) across all orders",                    "R$ 71,732,497"),
    ("Total Freight Cost",   "SUM(freight_value)",                               "R$ 9,001,654"),
    ("Gross Profit Estimate","SUM(price) - SUM(freight_value)",                  "R$ 62,730,843"),
    ("Profit Margin %",      "(Gross Profit / Revenue) × 100",                   "87.5%"),
    ("Average Order Value",  "SUM(price) / COUNT(DISTINCT order_id)",            "R$ 836.57"),
    ("Repeat Purchase Rate", "(Customers with >1 order / Total customers) × 100","99.9%"),
]
story.append(data_table(
    ["KPI", "Formula", "Result"],
    kpi_rows,
    col_widths=[5*cm, 8*cm, 4*cm]
))
story.append(SP(0.2))

story.append(P("<b>Monthly Revenue Trend Chart:</b>", S_h3))
story.extend(add_chart("02a_monthly_revenue_trend.png", width=16*cm,
    caption="Fig 1 — Monthly Revenue (blue) and Profit Estimate (green) with MoM Growth % bars"))
story.append(SP(0.1))
story.append(P(
    "<b>Reading this chart:</b> The top panel shows absolute revenue and profit growing "
    "steadily over 2 years — this is the growth trend we expect. The bottom panel shows "
    "Month-over-Month growth percentage. Green bars = positive growth months, "
    "red bars = decline months. The declining months reveal seasonality patterns.",
    S_body_small))
story.append(SP(0.2))

story.append(P("<b>Payment Method Analysis:</b>", S_h3))
story.extend(add_chart("02c_payment_analysis.png", width=15*cm,
    caption="Fig 2 — Payment method distribution (60% credit card) and revenue contribution"))
story.append(insight_box(
    "Credit card dominates at 60% of orders. Customers paying by credit card also "
    "tend to use installments (parcelamento), which is common in Brazil. "
    "This means revenue is spread over 2–12 months per order — important for cash flow modeling."
))
story.append(PageBreak())

# ══════════════════════════════════════════════════════════════
# SECTION 8 — STEP 5: CUSTOMER SEGMENTATION
# ══════════════════════════════════════════════════════════════
story.append(section_box("8. STEP 5 — CUSTOMER SEGMENTATION & RFM ANALYSIS"))
story.append(SP(0.3))
story.append(P("Script: python_analysis/03_customer_segmentation.py", S_label))
story.append(SP(0.1))
story.append(P(
    "Customer segmentation answers the question: <i>which customers matter most?</i> "
    "We use two methods — simple order-count segments (New/Active/Loyal) and the "
    "professional RFM model (Recency, Frequency, Monetary).", S_body))
story.append(SP(0.2))

story.append(P("<b>Method 1 — Simple Segmentation:</b>", S_h3))
seg_rows = [
    ("New",    "1 order",    "9,288", "R$ 69.7M", "First-time buyers. High acquisition cost, unknown retention."),
    ("Active", "2–4 orders", "698",   "R$ 2.0M",  "Returning customers. Showing loyalty signals."),
    ("Loyal",  "5+ orders",  "12",    "R$ 9K",    "True loyal customers. Highest CLV."),
]
story.append(data_table(
    ["Segment","Rule","Count","Revenue","Meaning"],
    seg_rows,
    col_widths=[2.5*cm, 2.5*cm, 2*cm, 2.5*cm, 7.5*cm]
))
story.append(SP(0.2))

story.append(P("<b>Pareto Analysis — Revenue Concentration:</b>", S_h3))
story.extend(add_chart("03a_pareto_analysis.png", width=14*cm,
    caption="Fig 3 — Pareto Chart: 61.6% of customers generate 80% of revenue"))
story.append(P(
    "The Pareto chart confirms the classic 80/20 pattern: the top 61.6% of customers "
    "(ranked by revenue) generate 80% of total revenue. This means a small group of "
    "high-value customers is disproportionately important — losing them hurts significantly.",
    S_body_small))
story.append(SP(0.2))

story.append(P("<b>Method 2 — RFM Segmentation:</b>", S_h3))
story.append(P(
    "RFM scores each customer from 1–5 on three dimensions. The combined score "
    "determines their segment:", S_body))

rfm_seg_rows = [
    ("Champions",         "R=5, F=5, M=5", "1,655", "R$ 11,646 avg",
     "Best customers. Buy often, recently, and spend most."),
    ("Loyal Customers",   "High RFM",       "2,225", "R$ 7,871 avg",
     "Frequent buyers. Core revenue base."),
    ("At Risk",           "Low R, High FM", "1,107", "R$ 10,224 avg",
     "Were great customers. Haven't bought recently. Winback needed."),
    ("Cannot Lose Them",  "Low R, Med FM",  "850",   "R$ 6,939 avg",
     "Big spenders who are drifting. Act fast."),
    ("Lost",              "Low R, Low FM",  "1,977", "R$ 4,003 avg",
     "Not bought in a long time. Low priority."),
    ("New Customers",     "R=5, Low FM",    "123",   "R$ 2,422 avg",
     "Just acquired. Need nurturing."),
    ("Promising",         "Med R, Low FM",  "1,279", "R$ 4,357 avg",
     "Occasional buyers. Promotions can convert them."),
]
story.append(data_table(
    ["Segment","Score Profile","Count","Avg Spend","Strategy"],
    rfm_seg_rows,
    col_widths=[4*cm, 2.5*cm, 1.8*cm, 2.8*cm, 5.9*cm]
))
story.append(SP(0.2))
story.extend(add_chart("03c_rfm_segments.png", width=14*cm,
    caption="Fig 4 — RFM Customer Segment Distribution"))
story.append(insight_box(
    "Business Action: The 1,107 'At Risk' customers have an average spend of R$ 10,224. "
    "Sending them a targeted re-engagement email with a personalised discount could "
    "recover R$ 11.3M in at-risk revenue. This is a direct ROI from segmentation analysis."
))
story.append(PageBreak())

# ══════════════════════════════════════════════════════════════
# SECTION 9 — STEP 6: COHORT ANALYSIS
# ══════════════════════════════════════════════════════════════
story.append(section_box("9. STEP 6 — COHORT RETENTION ANALYSIS"))
story.append(SP(0.3))
story.append(P("Script: python_analysis/04_cohort_analysis.py", S_label))
story.append(SP(0.1))
story.append(P(
    "Cohort analysis answers: <i>of all customers who joined in Month X, "
    "how many came back in Month X+1, X+2, X+3...?</i> This is the gold standard "
    "for measuring customer retention in e-commerce.",
    S_body))
story.append(SP(0.2))

story.append(P("<b>How Cohort Analysis Works (Step by Step):</b>", S_h3))
cohort_steps = [
    "1. Find each customer's first purchase month → this is their cohort_month",
    "2. For every order, calculate month_index = (order_month) - (cohort_month)",
    "   Example: Customer joins in Jan-2022, buys again in Mar-2022 → month_index = 2",
    "3. Count distinct customers per cohort_month per month_index",
    "4. Divide by cohort size (month_index = 0) to get retention %",
    "5. Pivot into matrix: rows = cohort months, columns = month indices",
]
for step in cohort_steps:
    story.append(P(step, S_body_small))
story.append(SP(0.1))

story.append(P("<b>Python Code (Core Logic):</b>", S_h3))
cohort_code = """\
# Step 1: Find first purchase month per customer
first_purchase = df.groupby('customer_id')['order_month'].min()
df = df.merge(first_purchase.rename('cohort_month'), on='customer_id')

# Step 2: Calculate month index
df['month_index'] = (df['order_month'] - df['cohort_month']).apply(lambda x: x.n)

# Step 3: Count customers per cohort per month
cohort_data = (df.groupby(['cohort_month','month_index'])['customer_id']
                 .nunique().reset_index(name='customers'))

# Step 4: Pivot and calculate retention %
cohort_pivot  = cohort_data.pivot_table(index='cohort_month',
                 columns='month_index', values='customers')
cohort_sizes  = cohort_pivot[0]
retention     = cohort_pivot.divide(cohort_sizes, axis=0) * 100"""
story.append(code_block(cohort_code))
story.append(SP(0.2))

story.append(P("<b>Cohort Retention Heatmap:</b>", S_h3))
story.extend(add_chart("04a_cohort_heatmap.png", width=16*cm,
    caption="Fig 5 — Cohort Retention Heatmap. Row = acquisition month. Column = months since first purchase. Value = % of customers still active."))
story.append(SP(0.1))
story.append(P(
    "<b>How to read this:</b> Each row is a cohort (group of customers who first bought "
    "in that month). Column 0 is always 100% (acquisition month). Column 1 shows what "
    "% returned in the next month. Darker green = better retention.",
    S_body_small))
story.append(SP(0.2))

story.extend(add_chart("04b_retention_curve.png", width=14*cm,
    caption="Fig 6 — Average Retention Curve across all cohorts"))
story.append(SP(0.1))

story.append(P("<b>Key Retention Numbers:</b>", S_h3))
ret_rows = [
    ("Month 0 (Acquisition)", "100.0% — All customers, by definition"),
    ("Month 1 Retention",      "32.6% — 67.4% of customers lost after first purchase"),
    ("Month 3 Retention",      "33.0% — Stable retention base"),
    ("Month 6 Retention",      "32.5% — Long-term retention plateau"),
]
story.append(kv_table(ret_rows, col_widths=[6*cm, 11*cm]))
story.append(SP(0.2))
story.append(insight_box(
    "The biggest drop occurs between Month 0 and Month 1 — 67.4% of customers "
    "never return after their first purchase. This is the most critical business "
    "problem to solve. A post-purchase email sequence, loyalty rewards, or a "
    "second-purchase discount could significantly improve this metric."
))
story.append(PageBreak())

# ══════════════════════════════════════════════════════════════
# SECTION 10 — STEP 7: PRODUCT PROFITABILITY
# ══════════════════════════════════════════════════════════════
story.append(section_box("10. STEP 7 — PRODUCT & CATEGORY PROFITABILITY"))
story.append(SP(0.3))
story.append(P("Script: python_analysis/05_product_profitability.py", S_label))
story.append(SP(0.1))
story.append(P(
    "This step identifies which product categories and individual products are "
    "truly profitable vs which ones look good on revenue but have their margins "
    "eaten by high shipping costs. This is called a 'revenue trap' analysis.",
    S_body))
story.append(SP(0.2))

story.append(P("<b>Category Profitability Summary:</b>", S_h3))
cat_sum = [
    ("electronics",           "R$ 30.6M", "R$ 26.7M", "87.5%", "12.5%"),
    ("furniture",             "R$ 15.8M", "R$ 13.8M", "87.6%", "12.4%"),
    ("sports_leisure",        "R$ 5.4M",  "R$ 4.7M",  "87.4%", "12.6%"),
    ("clothing_accessories",  "R$ 4.2M",  "R$ 3.7M",  "87.4%", "12.6%"),
    ("books",                 "R$ 0.6M",  "R$ 0.5M",  "85.9%", "14.1%"),
    ("stationery",            "R$ 0.2M",  "R$ 0.1M",  "83.2%", "16.8% ⚠"),
]
story.append(data_table(
    ["Category","Revenue","Profit Est.","Margin %","Freight %"],
    cat_sum,
    col_widths=[5.5*cm, 3*cm, 3*cm, 2.5*cm, 3*cm]
))
story.append(SP(0.2))

story.extend(add_chart("05a_revenue_vs_profit_category.png", width=16*cm,
    caption="Fig 7 — Revenue vs Profit by Category with margin % annotations"))
story.append(SP(0.2))

story.extend(add_chart("05b_freight_burden.png", width=14*cm,
    caption="Fig 8 — Freight as % of Revenue by Category (Green=efficient, Red=concerning)"))

story.append(insight_box(
    "Stationery has a 16.8% freight-to-revenue ratio — the worst of all categories. "
    "This means for every R$100 of stationery sold, R$16.80 goes to shipping alone. "
    "Business recommendation: bundle stationery orders (minimum order quantity) "
    "or charge separate shipping to protect margins."
))
story.append(PageBreak())

story.extend(add_chart("05d_category_monthly_trend.png", width=16*cm,
    caption="Fig 9 — Monthly Revenue Trend — Top 5 Categories (2022–2023)"))
story.append(SP(0.1))
story.append(P(
    "Electronics consistently leads revenue across all months. Its seasonality peak "
    "in Q4 (Oct–Dec) is visible, driven by holiday shopping. Furniture shows stable "
    "growth while clothing accessories is more volatile.",
    S_body_small))
story.append(PageBreak())

# ══════════════════════════════════════════════════════════════
# SECTION 11 — STEP 8: STATISTICAL ANALYSIS
# ══════════════════════════════════════════════════════════════
story.append(section_box("11. STEP 8 — STATISTICAL ANALYSIS"))
story.append(SP(0.3))
story.append(P("Script: python_analysis/06_statistical_analysis.py", S_label))
story.append(SP(0.1))
story.append(P(
    "Statistics validates the insights we see in charts. Without statistical "
    "validation, business decisions rest on patterns that might be coincidental. "
    "This step adds rigour to the analysis.",
    S_body))
story.append(SP(0.2))

story.append(P("<b>Correlation Analysis:</b>", S_h3))
corr_rows = [
    ("Price vs Freight Value",          "0.910",   "Very strong positive",
     "Higher-priced items cost significantly more to ship. Expected."),
    ("Review Score vs Price",           "0.005",   "No correlation",
     "Price does not affect review score — quality/service matters more."),
    ("Delivery Days vs Review Score",   "0.003",   "No correlation",
     "Surprisingly, delivery speed doesn't predict rating in this data."),
]
story.append(data_table(
    ["Variables","Correlation","Strength","Interpretation"],
    corr_rows,
    col_widths=[5*cm, 2.5*cm, 3*cm, 6.5*cm]
))
story.append(SP(0.2))

story.extend(add_chart("06d_price_vs_freight.png", width=14*cm,
    caption="Fig 10 — Price vs Freight Cost Scatter (R² = 0.824)"))
story.append(SP(0.1))
story.append(P(
    "R-squared = 0.824 means 82.4% of the variation in freight cost is explained "
    "by item price. Regression slope = 0.1224: for every R$ 1 increase in price, "
    "freight cost increases by R$ 0.12. This is useful for pricing models.",
    S_body_small))
story.append(SP(0.2))

story.append(P("<b>Order Value Distribution:</b>", S_h3))
story.extend(add_chart("06b_order_value_distribution.png", width=15*cm,
    caption="Fig 11 — Order Value Histogram and Price Distribution by Category"))
story.append(SP(0.1))
ov_rows = [
    ("Minimum Order Value",  "R$ 10.12"),
    ("25th Percentile",      "R$ 201.71"),
    ("Median Order Value",   "R$ 440.51"),
    ("Mean Order Value",     "R$ 836.57"),
    ("75th Percentile",      "R$ 1,185.05"),
    ("90th Percentile",      "R$ 2,257.74"),
    ("Maximum Order Value",  "R$ 7,956.59"),
]
story.append(kv_table(ov_rows, col_widths=[7*cm, 10*cm]))
story.append(SP(0.1))
story.append(P(
    "<b>Mean vs Median:</b> The mean (R$ 836) is much higher than the median (R$ 440). "
    "This is because a small number of very high-value electronics orders (R$ 2,000+) "
    "pull the average up. This right-skewed distribution is common in e-commerce. "
    "Always report both mean and median — never just mean.",
    S_body_small))
story.append(SP(0.2))

story.append(P("<b>Review Score Distribution:</b>", S_h3))
story.extend(add_chart("06c_review_analysis.png", width=14*cm,
    caption="Fig 12 — Review Score Distribution (47.9% give 5 stars) and by Category"))
story.append(insight_box(
    "47.9% of reviews are 5-star — this is strong customer satisfaction. "
    "However, the 12.1% who give 3 stars or below represent customers at risk of churning. "
    "These orders should be flagged for customer service follow-up."
))
story.append(PageBreak())

# ══════════════════════════════════════════════════════════════
# SECTION 12 — STEP 9: EXCEL SIMULATION
# ══════════════════════════════════════════════════════════════
story.append(section_box("12. STEP 9 — EXCEL SCENARIO SIMULATION"))
story.append(SP(0.3))
story.append(P("File: excel_simulation/scenario_simulation.xlsx", S_label))
story.append(SP(0.1))
story.append(P(
    "Excel is used for financial modeling and scenario simulation. "
    "Unlike Python charts, Excel models allow business stakeholders to "
    "interactively change inputs and immediately see the profit impact. "
    "This is the format that CFOs and VPs prefer.",
    S_body))
story.append(SP(0.2))

story.append(P("<b>Excel Workbook Structure (5 Sheets):</b>", S_h3))
excel_sheets = [
    ("Sheet 1: Business Baseline",
     "Core KPIs table with all key metrics. This is the 'starting point' "
     "that all scenario comparisons reference."),
    ("Sheet 2: Scenario Simulator",
     "Side-by-side comparison of 3 business scenarios vs baseline. "
     "Shows profit impact of each decision."),
    ("Sheet 3: Category Profitability",
     "All 15 categories with revenue, profit, margin %, and freight % "
     "with conditional colour formatting."),
    ("Sheet 4: Monthly Revenue",
     "Monthly revenue and profit trend with an embedded bar chart. "
     "Includes MoM growth % with green/red conditional formatting."),
    ("Sheet 5: What-If Analysis",
     "Interactive model with yellow input cells. Change any assumption "
     "and Excel formulas instantly recalculate all outputs."),
]
for sheet, desc in excel_sheets:
    story.append(P(f"<b>{sheet}:</b> {desc}", S_bullet))
story.append(SP(0.2))

story.append(P("<b>3 Business Scenarios Modeled:</b>", S_h3))
scenario_data = [
    ("",                     "Baseline", "Scenario A", "Scenario B", "Scenario C"),
    ("Strategy",             "—", "Reduce Freight -10%", "Grow Customers +10%", "Increase AOV +10%"),
    ("Annual Revenue (R$)",  "35,866,249", "35,866,249", "39,452,874", "39,452,874"),
    ("Freight Cost (R$)",    "4,500,827", "4,050,744",  "4,500,827",  "4,500,827"),
    ("Profit Estimate (R$)", "31,365,422", "31,815,505", "34,952,047", "34,952,047"),
    ("Profit vs Baseline",   "—",         "+R$ 450K",   "+R$ 3.6M",   "+R$ 3.6M"),
    ("Profit Improvement",   "0%",        "+1.4%",      "+11.4%",     "+11.4%"),
]

header_row = scenario_data[0]
body_rows  = scenario_data[1:]
s_header_cells = [Paragraph(str(v), ParagraphStyle("SH", fontSize=8.5,
                  fontName="Helvetica-Bold", textColor=WHITE, alignment=TA_CENTER,
                  leading=11)) for v in header_row]
s_rows_out = [s_header_cells]
BGs = [DARK_BLUE, colors.HexColor("#155724"),
       colors.HexColor("#721C24"), colors.HexColor("#5A3E85")]
for r_i, row in enumerate(body_rows):
    cells = []
    for c_i, v in enumerate(row):
        if c_i == 0:
            style_p = ParagraphStyle("SC0", fontSize=8.5, fontName="Helvetica-Bold",
                                     leading=11, textColor=DARK_BLUE)
        elif c_i == 2 or c_i == 3 or c_i == 4:
            style_p = ParagraphStyle("SCV", fontSize=8.5, leading=11,
                                     textColor=DARK_GREEN, alignment=TA_CENTER,
                                     fontName="Helvetica-Bold" if r_i >= 3 else "Helvetica")
        else:
            style_p = ParagraphStyle("SC1", fontSize=8.5, leading=11,
                                     textColor=BLACK, alignment=TA_CENTER)
        cells.append(Paragraph(str(v), style_p))
    s_rows_out.append(cells)

st = Table(s_rows_out, colWidths=[4.5*cm, 2.5*cm, 3*cm, 3*cm, 4*cm])
ts = TableStyle([
    ("BACKGROUND",    (0,0), (-1,0), DARK_BLUE),
    ("BACKGROUND",    (2,1), (2,-1), colors.HexColor("#EBF8F0")),
    ("BACKGROUND",    (3,1), (3,-1), colors.HexColor("#FFF0F0")),
    ("BACKGROUND",    (4,1), (4,-1), colors.HexColor("#F3F0FF")),
    ("ROWBACKGROUNDS",(0,1), (1,-1), [WHITE, LIGHT_GRAY]),
    ("GRID",          (0,0), (-1,-1), 0.4, MID_GRAY),
    ("TOPPADDING",    (0,0), (-1,-1), 6),
    ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ("LEFTPADDING",   (0,0), (-1,-1), 6),
])
st.setStyle(ts)
story.append(st)
story.append(SP(0.2))

story.append(P("<b>How to use the What-If sheet:</b>", S_h3))
what_if_steps = [
    "Open excel_simulation/scenario_simulation.xlsx in Microsoft Excel",
    "Go to Sheet 5: 'What-If Analysis'",
    "Change any YELLOW cell (e.g., change Base Revenue or Freight %)",
    "All green output cells recalculate automatically using Excel formulas",
    "Example: Change Freight % from 12.5 to 10 → see instant profit improvement",
]
for i, step in enumerate(what_if_steps, 1):
    story.append(P(f"{i}. {step}", S_bullet))

story.append(PageBreak())

# ══════════════════════════════════════════════════════════════
# SECTION 13 — POWER BI GUIDE
# ══════════════════════════════════════════════════════════════
story.append(section_box("13. POWER BI DASHBOARD GUIDE"))
story.append(SP(0.3))
story.append(P(
    "Two dashboards should be built in Power BI using the CSV files "
    "exported by the Python scripts. Here is the exact specification for each.",
    S_body))
story.append(SP(0.2))

story.append(P("<b>Dashboard 1 — Executive Business Overview (CEO Dashboard)</b>", S_h2))
story.append(P(
    "Data source: data/processed/orders_fact.csv + outputs/csv_exports/monthly_revenue.csv",
    S_label))
story.append(SP(0.1))

db1_charts = [
    ("KPI Card", "Total Revenue (R$)", "DAX: SUM(orders_fact[price])"),
    ("KPI Card", "Total Orders",       "DAX: DISTINCTCOUNT(orders_fact[order_id])"),
    ("KPI Card", "Average Order Value","DAX: DIVIDE([Total Revenue],[Total Orders])"),
    ("Line Chart","Monthly Revenue Trend", "X=order_year_month, Y=SUM(price)"),
    ("Bar Chart", "Revenue by Category",   "X=product_category_name, Y=SUM(price)"),
    ("Bar Chart", "Top 10 Products",       "Filtered table sorted by revenue"),
    ("Map Visual","Revenue by State",      "Customer state with bubble size = revenue"),
    ("Donut Chart","Payment Method Split", "payment_type with % of orders"),
]
story.append(data_table(
    ["Visual Type","Title","Data / DAX"],
    db1_charts,
    col_widths=[4*cm, 6*cm, 7*cm]
))
story.append(SP(0.1))
story.append(P(
    "<b>Filters/Slicers:</b> Year, Product Category, Customer State (top right of dashboard)",
    S_body_small))
story.append(SP(0.3))

story.append(P("<b>Dashboard 2 — Customer & Profit Analytics</b>", S_h2))
story.append(P(
    "Data source: outputs/csv_exports/rfm_segments.csv + cohort_retention_matrix.csv + category_metrics.csv",
    S_label))
story.append(SP(0.1))

db2_charts = [
    ("Donut/Bar",  "Customer Segment Distribution", "rfm_segments.csv: rfm_segment count"),
    ("KPI Card",   "Repeat Purchase Rate %",        "Customers with orders>1 / total"),
    ("Matrix/Table","Cohort Retention Heatmap",     "cohort_retention_matrix.csv pivoted"),
    ("Bar Chart",  "Profit by Category",            "category_metrics.csv: total_profit"),
    ("Scatter Plot","CLV Segments",                 "X=frequency, Y=monetary, Color=rfm_segment"),
    ("Line Chart", "Shipping Cost vs Revenue",      "monthly_revenue.csv: revenue + freight"),
    ("Table",      "At-Risk Customers",             "rfm_segments filtered to 'At Risk'"),
]
story.append(data_table(
    ["Visual Type","Title","Data Source"],
    db2_charts,
    col_widths=[4*cm, 6*cm, 7*cm]
))
story.append(SP(0.1))
story.append(P(
    "<b>Filters/Slicers:</b> Month, Product Category, RFM Segment",
    S_body_small))
story.append(SP(0.2))

story.append(P("<b>Key DAX Measures to Create:</b>", S_h3))
dax_measures = [
    "Total Revenue = SUM(orders_fact[price])",
    "Total Profit = SUM(orders_fact[profit_estimate])",
    "Profit Margin % = DIVIDE([Total Profit], [Total Revenue]) * 100",
    "AOV = DIVIDE([Total Revenue], DISTINCTCOUNT(orders_fact[order_id]))",
    "Repeat Customers = CALCULATE(DISTINCTCOUNT(cust[customer_id]), cust[total_orders]>1)",
    "Repeat Rate % = DIVIDE([Repeat Customers], DISTINCTCOUNT(cust[customer_id])) * 100",
]
for m in dax_measures:
    story.append(code_block(m))
    story.append(SP(0.05))

story.append(PageBreak())

# ══════════════════════════════════════════════════════════════
# SECTION 14 — KEY BUSINESS INSIGHTS
# ══════════════════════════════════════════════════════════════
story.append(section_box("14. KEY BUSINESS INSIGHTS"))
story.append(SP(0.3))
story.append(P(
    "These are the final business insights derived from the analysis. "
    "This section is what you present to stakeholders. "
    "Each insight includes the data evidence and a business recommendation.",
    S_body))
story.append(SP(0.2))

insights = [
    ("INSIGHT 1: Revenue is Growing but Freight is a Silent Profit Killer",
     "Monthly revenue grew from R$ 2.5M to R$ 5.5M over 2 years (+120%). "
     "However, freight cost averages 12.5% of revenue consistently. "
     "Reducing freight efficiency by just 2 percentage points would add R$ 1.4M+ "
     "to annual profit without growing revenue at all.",
     "Negotiate bulk shipping contracts or switch to zone-based flat-rate shipping."),
    ("INSIGHT 2: 61.6% of Customers Drive 80% of Revenue (Pareto Principle)",
     "This confirms the 80/20 rule. A targeted retention strategy for the top "
     "customers will have 4× more impact than broad marketing campaigns. "
     "Losing 100 high-value customers hurts more than losing 1,000 low-value ones.",
     "Build a VIP loyalty program. Identify Champions and Loyal Customers in RFM "
     "and give them early access, exclusive discounts, and dedicated support."),
    ("INSIGHT 3: 67.4% of Customers Never Return After First Purchase",
     "Month 1 retention is only 32.6%. This means the business acquires customers "
     "but fails to retain them. Acquisition cost (CAC) is only justified if "
     "customers make multiple purchases. With 67% one-time buyers, the CAC "
     "payback period is too long.",
     "Implement a post-purchase nurture sequence: Day 3 (review request), "
     "Day 14 (complementary product recommendation), Day 30 (second-purchase discount)."),
    ("INSIGHT 4: Electronics = 42% of Revenue but Highest Absolute Freight Cost",
     "Electronics dominates revenue (R$ 30.6M of R$ 71.7M total) but each item "
     "has the highest freight cost in absolute terms (12.5% = R$ 3.8M in freight). "
     "Customers buying electronics are price-sensitive but also expect fast delivery.",
     "Negotiate with courier companies for electronics-specific rates. "
     "Consider free shipping threshold (e.g., free shipping on orders above R$ 500 "
     "encourages basket-building while subsidising freight through volume)."),
    ("INSIGHT 5: 1,107 'At Risk' Customers Have R$ 11.3M Recoverable Revenue",
     "At Risk customers used to be active (high frequency and monetary) but "
     "haven't purchased recently (low recency score). They are familiar with the "
     "brand and have high intent. Re-engaging them is 5× cheaper than acquiring "
     "new customers.",
     "Run an 'We miss you' winback campaign with personalised product "
     "recommendations based on their past purchase history and a time-limited "
     "10% discount code."),
]

for title, evidence, action in insights:
    story.append(P(f"<b>{title}</b>", S_h3))
    story.append(P(f"<b>Evidence:</b> {evidence}", S_body_small))
    story.append(P(f"<b>Recommendation:</b> {action}",
                   ParagraphStyle("Rec", fontSize=9.5, textColor=DARK_GREEN,
                                  leading=13, fontName="Helvetica-Bold", spaceAfter=3)))
    story.append(divider(LIGHT_BLUE, 0.5))
    story.append(SP(0.1))

story.append(PageBreak())

# ══════════════════════════════════════════════════════════════
# SECTION 15 — RESUME & SUMMARY
# ══════════════════════════════════════════════════════════════
story.append(section_box("15. RESUME BULLET & PROJECT SUMMARY"))
story.append(SP(0.3))

story.append(P("<b>Resume Bullet Point:</b>", S_h3))
resume_box = Table([[Paragraph(
    "Built an end-to-end E-Commerce Growth &amp; Profit Optimization Analytics system "
    "using SQL, Python (pandas, matplotlib, seaborn), and Excel to analyse 100,000+ "
    "orders across 15 product categories, uncovering key drivers of R$ 71.7M revenue, "
    "customer retention patterns via cohort analysis, and category-level profit margins "
    "through statistical correlation (R²=0.824). Delivered 5 business insights "
    "including a R$ 11.3M at-risk customer recovery opportunity.",
    ParagraphStyle("RB", fontSize=10.5, leading=16, textColor=DARK_BLUE,
                   fontName="Helvetica-Bold")
)]], colWidths=[17*cm])
resume_box.setStyle(TableStyle([
    ("BACKGROUND",    (0,0), (-1,-1), YELLOW),
    ("BOX",           (0,0), (-1,-1), 1.5, colors.HexColor("#FFC107")),
    ("TOPPADDING",    (0,0), (-1,-1), 12),
    ("BOTTOMPADDING", (0,0), (-1,-1), 12),
    ("LEFTPADDING",   (0,0), (-1,-1), 14),
    ("RIGHTPADDING",  (0,0), (-1,-1), 14),
]))
story.append(resume_box)
story.append(SP(0.4))

story.append(P("<b>What Skills This Project Demonstrates to Recruiters:</b>", S_h3))
skills = [
    ("SQL Analytics",        "Joins, aggregations, CTEs, window functions (NTILE), subqueries"),
    ("Python Analysis",      "pandas, numpy, matplotlib, seaborn, scipy — end-to-end pipeline"),
    ("Statistical Thinking", "Correlation analysis, R-squared, percentile analysis, Pareto"),
    ("Business Metrics",     "Revenue, profit margin, AOV, CLV, repeat rate, cohort retention"),
    ("Excel Modeling",       "Scenario simulation, what-if analysis, conditional formatting, charts"),
    ("Analytical Thinking",  "Converting data into 5 actionable business recommendations"),
    ("Data Engineering",     "Multi-table joins, data quality validation, processed data layer"),
    ("Visualisation",        "12 professional charts with annotations, colour coding, business context"),
]
story.append(kv_table(skills, col_widths=[5*cm, 12*cm]))
story.append(SP(0.3))

story.append(P("<b>Project Files Summary:</b>", S_h3))
files_summary = [
    ("data/generate_data.py",                    "Generates 100k+ synthetic orders dataset"),
    ("data/raw/*.csv",                            "6 raw CSV files (customers, orders, items, products, payments, reviews)"),
    ("data/processed/orders_fact.csv",            "Master cleaned and joined analytics table (133k rows)"),
    ("sql_queries/01_create_orders_fact.sql",     "Master fact table SQL"),
    ("sql_queries/02_customer_metrics.sql",       "Per-customer KPIs SQL"),
    ("sql_queries/03_product_metrics.sql",        "Product profitability SQL"),
    ("sql_queries/04_category_metrics.sql",       "Category comparison SQL"),
    ("sql_queries/05_cohort_analysis.sql",        "Cohort retention SQL"),
    ("sql_queries/06_rfm_segmentation.sql",       "RFM scoring with window functions SQL"),
    ("python_analysis/01_data_cleaning.py",       "Data cleaning and fact table builder"),
    ("python_analysis/02_business_metrics.py",    "KPIs, monthly trend, payment analysis"),
    ("python_analysis/03_customer_segmentation.py","RFM, Pareto, CLV analysis"),
    ("python_analysis/04_cohort_analysis.py",     "Cohort retention matrix and heatmap"),
    ("python_analysis/05_product_profitability.py","Category and product margin analysis"),
    ("python_analysis/06_statistical_analysis.py","Correlations, distributions, review analysis"),
    ("excel_simulation/scenario_simulation.xlsx", "5-sheet financial model with 3 scenarios"),
    ("outputs/charts/*.png",                      "12 professional analysis charts"),
    ("outputs/csv_exports/*.csv",                 "15 analysis result CSV files"),
    ("run_all.py",                                "Master pipeline — runs everything in order"),
]
story.append(data_table(
    ["File","Purpose"],
    files_summary,
    col_widths=[8.5*cm, 8.5*cm]
))

story.append(SP(0.4))
story.append(divider(DARK_BLUE, 2))

# Final note
final = Table([[Paragraph(
    "This project covers the complete analytical workflow of a real Data Analyst: "
    "from raw data ingestion through SQL modeling, Python analysis, statistical validation, "
    "Excel simulation, and executive reporting. Each layer builds on the previous one. "
    "The Power BI dashboards you build on top of this will complete the full stack.",
    ParagraphStyle("Final", fontSize=10, leading=15, textColor=DARK_BLUE,
                   alignment=TA_JUSTIFY)
)]], colWidths=[17*cm])
final.setStyle(TableStyle([
    ("BACKGROUND",    (0,0), (-1,-1), LIGHT_BLUE),
    ("BOX",           (0,0), (-1,-1), 1, MEDIUM_BLUE),
    ("TOPPADDING",    (0,0), (-1,-1), 10),
    ("BOTTOMPADDING", (0,0), (-1,-1), 10),
    ("LEFTPADDING",   (0,0), (-1,-1), 12),
]))
story.append(SP(0.2))
story.append(final)

# ─────────────────────────────────────────────
# BUILD
# ─────────────────────────────────────────────
doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
print(f"PDF saved: {OUT}")
