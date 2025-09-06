# Create a PDF summary in English describing the Power Query steps
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem, Preformatted
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

file_path = "/mnt/data/PowerQuery_Financials_Summary.pdf"

styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name="TitleBig", parent=styles["Title"], fontSize=22, spaceAfter=16))
styles.add(ParagraphStyle(name="H2", parent=styles["Heading2"], fontSize=14, spaceAfter=8, textColor=colors.HexColor("#0b5394")))
styles.add(ParagraphStyle(name="Body", parent=styles["BodyText"], leading=14, spaceAfter=6))
styles.add(ParagraphStyle(name="Mono", parent=styles["Code"], fontName="Courier", fontSize=9, leading=12, backColor=colors.whitesmoke, borderPadding=6))

doc = SimpleDocTemplate(file_path, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
story = []

story.append(Paragraph("Power Query Processing Summary – Financials_Table", styles["TitleBig"]))
story.append(Paragraph("This document summarizes the data-cleaning and type-handling logic applied in **Power Query** for the *Financial Sample* workbook. All steps are written to be reproducible and robust.", styles["Body"]))

# Section: Objectives
story.append(Paragraph("Objectives", styles["H2"]))
obj_items = [
    "Import the true table from the Excel workbook and avoid accidental sheet imports.",
    "Promote headers only when necessary.",
    "Normalize column names to the canonical schema.",
    "Remove the large block of blank rows that caused an “empty” fact table.",
    "Fix the UnitsID column that turned into zeros.",
    "Cast types safely using Vietnamese culture (vi-VN) so commas are read as decimals.",
]
story.append(ListFlowable([ListItem(Paragraph(i, styles["Body"])) for i in obj_items], bulletType="bullet"))

# Section: Import
story.append(Paragraph("1) Import the correct source", styles["H2"]))
story.append(Paragraph("Open the workbook with Excel.Workbook(...) and fetch the real table named **Financials_Table** (fallback to **Financials** if needed). This prevents the column drift seen when pulling directly from a sheet.", styles["Body"]))

# Section: Promote headers conditionally
story.append(Paragraph("2) Promote headers only when needed", styles["H2"]))
story.append(Paragraph("Detect whether the first row still contains generic headers (e.g., Column1/Column2). Only then call Table.PromoteHeaders to avoid double-promotion.", styles["Body"]))

# Section: Column normalization
story.append(Paragraph("3) Normalize column names", styles["H2"]))
story.append(Paragraph("Trim and remove NBSP (char 160) from all headers so they match the canonical schema below:", styles["Body"]))
canon_cols = [
    "SegmentID, CountryID, ProductID, DiscountID, UnitsID, Manufacturing Price, Sale Price,",
    "Gross Sales, Discounts, Sales, COGS, Profit, Date, DateID, Month Number, Month Name, Year"
]
story.append(Preformatted("\n".join(canon_cols), styles["Mono"]))

# Section: Remove blank rows
story.append(Paragraph("4) Remove blank/noise rows (root cause of “empty” table)", styles["H2"]))
story.append(Paragraph(
    "Keep only records where **Date** and **SegmentID** are not null. "
    "This removes ~1,048,575 blank rows (those with only DateID = 011900), which previously made the table look empty in the model.",
    styles["Body"]
))

# Section: Fix UnitsID
story.append(Paragraph("5) Fix UnitsID becoming 0", styles["H2"]))
story.append(Paragraph(
    "Extract the numeric digits from UnitsID and cast to Int64. Remove/avoid any prior steps that force replacements to 0. "
    "Result: UnitsID values now reflect the original counts instead of all zeros.",
    styles["Body"]
))

# Section: Culture-safe types
story.append(Paragraph("6) Culture-safe type casting (vi-VN)", styles["H2"]))
story.append(Paragraph(
    "Apply types using the Vietnamese culture so comma decimals are parsed correctly. "
    "Numeric columns → number; Date → date; DateID, Month Number, Year → Int64.",
    styles["Body"]
))

# Section: Results
story.append(Paragraph("Results", styles["H2"]))
story.append(ListFlowable([
    ListItem(Paragraph("~700 valid data rows remain after cleaning; numbers and dates parse correctly.", styles["Body"])),
    ListItem(Paragraph("UnitsID is fixed (no longer all zeros).", styles["Body"])),
    ListItem(Paragraph("No phantom rows with DateID 011900.", styles["Body"])),
], bulletType="bullet"))

# Section: Applied Steps
story.append(Paragraph("Applied Steps (as seen in Query Settings)", styles["H2"]))
steps = [
    "Source → Navigation → NeedPromote → Promoted → CleanNames → Filtered (Keep Rows) → TypePairsAll/Cols → TypePairs → Typed"
]
story.append(Preformatted("\n".join(steps), styles["Mono"]))

# Section: Target star schema
story.append(Paragraph("Target star schema", styles["H2"]))
story.append(Paragraph("<b>Fact:</b> Financials_Table", styles["Body"]))
story.append(Paragraph("<b>Dimensions:</b> Segment_Table, Country_Table, Product_Table, Discount_Table, Units_Table, Date_Table (each cleansed to distinct keys, proper types; one-to-many single-direction relationships from dimensions to fact).", styles["Body"]))

# Section: DAX measures
story.append(Paragraph("Suggested DAX measures (starting set)", styles["H2"]))
dax = """Total Sales    = SUM(Financials_Table[Sales])
Total Profit   = SUM(Financials_Table[Profit])
Discount %     = DIVIDE(SUM(Financials_Table[Discounts]), SUM(Financials_Table[Gross Sales]))"""
story.append(Preformatted(dax, styles["Mono"]))

# Footer note
story.append(Spacer(1, 12))
story.append(Paragraph("Note: If you also need the M code skeleton, I can export a separate appendix with the parameterized query.", styles["Body"]))

doc.build(story)

file_path
