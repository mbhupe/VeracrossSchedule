!pip install reportlab
!pip install gdown


import csv
import os
import gdown
import re
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.styles import ParagraphStyle

from reportlab.lib.styles import ParagraphStyle
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.lib.colors import blue
from reportlab.platypus import HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib import colors
import shutil
from google.colab import drive

from reportlab.lib.pagesizes import A4

def add_student_name(canvas, doc, student_display_name):
    canvas.saveState()
    width, height = A4
    canvas.setFont("Helvetica-Oblique", 14)
    # Draw text at top-right corner (margin 40 from right, 820 from bottom)
    canvas.drawRightString(width - 40, height - 40, student_display_name)
    canvas.restoreState()


# Testing with advisory room number
# https://drive.google.com/file/d/1LW-ujoLfAT5B69fWwTSF40FyrR0HVSjK/view?usp=sharing
drive.mount('/content/drive')
# file_id = "1DjVVQS8-FTxZANAWQyel0gMgTBBlzh5S" #without advisory room number
file_id = "1LW-ujoLfAT5B69fWwTSF40FyrR0HVSjK" 

# === Styles ===
styles = getSampleStyleSheet()

# Custom style for welcome paragraph (larger + light blue)
welcome_style = ParagraphStyle(
    name="WelcomeStyle",
    parent=styles["Normal"],
    fontSize=14,
    textColor=colors.lightblue,
    leading=18
)

# Custom centered style for the event date
date_style = ParagraphStyle(
    name="DateStyle",
    parent=styles["Heading3"],
    alignment=TA_CENTER
)

name_style = ParagraphStyle(
    "NameStyle",
    parent=styles["Heading2"],
    alignment=TA_RIGHT,
    fontName="Helvetica-Oblique",
    fontSize=16
)

welcome_style = ParagraphStyle(
    name="WelcomeStyle",
    parent=styles["Normal"],
    fontSize=14,        # larger font
    textColor=blue,
    leading=18           # space between lines
)

class_style = ParagraphStyle(
    'ClassStyle',
    parent=styles["Normal"],
    fontSize=14,
    leading=16,
    spaceAfter=6,
    textColor=colors.black,
    fontName="Helvetica-Bold"  # make it bold
)

# === Configuration ===
url = f"https://drive.google.com/uc?id={file_id}"
inputfile = "open_house.csv"             # your CSV file
gdown.download(url, inputfile, quiet=False)


input_csv = inputfile
# output_folder = "open_house_pdfs"        # folder to save PDFs
# output_folder = "/content/drive/My Drive/open_house_pdfs"
output_folder = "/content/drive/My Drive/open_houseAdv_pdfs"
logo_file = "/content/drive/My Drive/sfp_logo.png"               # your logo file (PNG)
os.makedirs(output_folder, exist_ok=True)
# Remove old folder if it exists, then create a new one
if os.path.exists(output_folder):
    shutil.rmtree(output_folder)
os.makedirs(output_folder)
styles = getSampleStyleSheet()

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors

def add_page_decorations(canvas, doc, student_display_name):
    canvas.saveState()
    width, height = A4

    # === Top color band ===
    band_height = 10  # thickness of the band
    canvas.setFillColor(colors.blue)   # change color here
    canvas.rect(0, height - band_height, width, band_height, fill=1, stroke=0)

    # === Student name (top-right) ===
    canvas.setFont("Helvetica-Oblique", 14)
    canvas.setFillColor(colors.black)
    canvas.drawRightString(width - 40, height - 35, student_display_name)

    canvas.restoreState()


# === Read CSV ===
with open(input_csv, newline="", encoding="utf-8") as f:
    reader = csv.DictReader(f)
    
    for row in reader:
        raw_name = row["Full Name"].strip()
       # Fix order if it’s "Last, First"
        if "," in raw_name:
          last, first = [x.strip() for x in raw_name.split(",", 1)]
          student_display_name = f"{first} {last}"
        else:
          student_display_name = raw_name  
        grade = row["Current Grade"].strip()
        advisor = row["Advisor"].strip()
        advisory_raw = row.get("Period 8", "").strip()
        match = re.search(r'Room\s+(\d+)', advisory_raw, re.IGNORECASE)
        advisory_number = match.group(1) if match else ""

        advisor_text = f"<b>Advisor:</b> {advisor}"
        if advisory_number:
          advisor_text += f" (Room {advisory_number})"

        

        pdf_file = os.path.join(output_folder, f"{student_display_name.replace(' ', '_')}_schedule.pdf")
        doc = SimpleDocTemplate(
            pdf_file,
            pagesize=A4,
            topMargin=10,    # adjust this number to move header higher/lower
            bottomMargin=10,
            leftMargin=20,
            rightMargin=20
        )
        elements = []

        # === Student Name (top-right corner) ===
        elements.append(Paragraph(student_display_name, name_style))
        elements.append(Spacer(1, 12))
        pdf_file = os.path.join(output_folder, f"{student_display_name.replace(' ', '_')}_schedule.pdf")
        doc = SimpleDocTemplate(
            pdf_file,
            pagesize=A4,
            topMargin=10,    # adjust this number to move header higher/lower
            bottomMargin=10,
            leftMargin=20,
            rightMargin=20
        )
        elements = []

        # === Add Logo ===
        if os.path.exists(logo_file):
            logo = Image(logo_file)
            logo.drawHeight = 30*mm
            logo.drawWidth = 50*mm
            logo.hAlign = 'CENTER'
            elements.append(logo)
            elements.append(Spacer(1, 6))

            # Horizontal line
            elements.append(HRFlowable(
                width="100%", 
                thickness=1, 
                lineCap='round', 
                color=colors.black, 
                spaceBefore=3, 
                spaceAfter=3
            ))

        # # === Add School Name ===
        # elements.append(Paragraph("<b>Santa Fe Prep</b>", styles['Title']))
        # elements.append(Spacer(1, 6))
        
        event_date = "September 27, 2025"
        elements.append(Paragraph(event_date, date_style))
        elements.append(Spacer(1, 6))

        # === Welcome Paragraph ===
        welcome_text = f"Greetings parents! We are excited to have you visit {student_display_name}'s classes during Open House. Here is your child's schedule for the day. If there is a free period on your child's schedule, please go to the US Quad and mingle with other parents."
        elements.append(Paragraph(welcome_text, welcome_style))
        elements.append(Spacer(1, 12))
        
        # Title for Student
        elements.append(Paragraph(f"<b>Open House Schedule for {student_display_name}</b>", styles['Heading2']))
        elements.append(Spacer(1, 6))

        # Student Info
        elements.append(Paragraph(f"<b>Grade:</b> {grade}",welcome_style ))
        elements.append(Paragraph(advisor_text, welcome_style))
        elements.append(Spacer(1, 12))
        
        # === Add Pre-Schedule Events ===
        pre_events = [
            "8:30 AM – Check-In, Enjoy Coffee and Mingle in Upper School Quad",
            "9:00 AM – Head of School Welcome in Upper School Quad",
            "9:25 AM – Meet with Advisors in Advisory location"
        ]

        for event in pre_events:
            elements.append(Paragraph(event,welcome_style ))
            elements.append(Spacer(1, 6))
        elements.append(Spacer(1, 12))

        # === Table: Periods with Times ===
        table_data = [["Period", "Time", "Class"]]

        start_time = datetime.strptime("09:45", "%H:%M")
        period_duration = timedelta(minutes=15)
        break_duration = timedelta(minutes=5)
        current_time = start_time

        for i in range(1, 8):
            # period_col = f"Period {i}"
            # class_text = str(row.get(period_col, "") or "").strip()

            # # # Skip blank/free periods
            # if class_text == "":
            #     class_text = "Free Period"

            # # Format time range
            # period_start = current_time
            # period_end = period_start + period_duration
            # # Only add AM/PM at the end
            # time_str = f"{period_start.strftime('%I:%M')} - {period_end.strftime('%I:%M %p')}"


            # # Replace line breaks with colon
            # class_text = class_text.replace("\n", " : ")

            #Room change attempt:
            period_col = f"Period {i}"
            class_text = row.get(period_col, "") or ""
            class_text = class_text.replace("\n", " : ")

            # 🧹 Normalize "Room" format
            match = re.search(r"Room\s*:?\s*Room\s*(\d+)", class_text)
            if match:
            # Replace with "Room <number>"
              room_number = match.group(1)
              class_text = re.sub(r"Room\s*:?\s*Room\s*\d+", f"Room {room_number}", class_text)

            if not class_text.strip():
              class_text = "Free Period (Meet at Advisory Location)"
           #Room change attempt


            table_data.append([f"Period {i}", time_str, Paragraph(class_text, class_style)])

            # Update current time: period + break
            current_time = period_end + break_duration

        # Create Table
        table = Table(table_data, colWidths=[80, 120, 300])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.cyan),
            ('TEXTCOLOR',(0,0),(-1,0),colors.black),
            ('ALIGN',(0,0),(-1,-1),'CENTER'),
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 14),
            ('FONTSIZE', (0,1), (-1,-1), 14),
            ('FONTSIZE', (0,1), (-1,-1), 18),
            ('FONTSIZE', (0,1), (-1,-1), 14),
            ('LEFTPADDING', (0,0), (-1,-1), 6),
            ('RIGHTPADDING', (0,0), (-1,-1), 6),
            ('TOPPADDING', (0,0), (-1,-1), 10),
            ('BOTTOMPADDING', (0,0), (-1,-1), 10),
            # Only alternate background if cell has content
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.whitesmoke, colors.lightgreen])
        ]))

        # === Highlight Free Period Rows ===
        for row_idx, row_data in enumerate(table_data[1:], start=1):  # skip header
            if isinstance(row_data[2], Paragraph) and "Free Period" in row_data[2].getPlainText():
                table.setStyle([
                    ('BACKGROUND', (0, row_idx), (-1, row_idx), colors.yellow),
                    ('TEXTCOLOR', (0, row_idx), (-1, row_idx), colors.black),
                ])


        elements.append(table)
        elements.append(Spacer(1, 12))
        sporting_style = ParagraphStyle(
            name="SportingStyle",
            parent=styles["Normal"],
            fontSize=14,
            textColor=colors.darkred,
            leading=16
        )

        sporting_event = "***1:00 pm: On-campus Sporting Events: Varsity Girls Soccer vs Monte del Sol @ Sun Mountain Field"
        elements.append(Paragraph(sporting_event, sporting_style))
        # Build PDF
        # doc.build(elements)
        doc.build(elements, onFirstPage=lambda canvas, doc: add_page_decorations(canvas, doc, student_display_name))

        # print(f"✅ Saved PDF: {pdf_file}")
        if reader.line_num == 2:
          open_pdf(pdf_file)

# print("🎉 All Open House PDFs with header, logo, and times generated successfully!")
