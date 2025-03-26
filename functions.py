import re
from pypdf import PdfReader, PdfWriter
import os
import subprocess
from docxtpl import DocxTemplate
import pandas as pd
from fuzzywuzzy import fuzz
import streamlit as st
import io
import shutil
import pypandoc

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

# -----------------------------
# Credential and Conversion Functions
# -----------------------------
def get_credentials():
    # Load credentials from Streamlit secrets
    credentials_dict = st.secrets["google_service_account"]
    creds = service_account.Credentials.from_service_account_info(
        credentials_dict,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    return creds

def convert_docx_to_pdf_gdrive(docx_path, output_pdf_path):
    """
    Converts a DOCX file to PDF using the Google Drive API by first converting
    the DOCX file into a native Google Docs file (which can then be exported).
    
    Parameters:
      docx_path (str): Local path to the input DOCX file.
      output_pdf_path (str): Local path where the output PDF will be saved.
    
    Returns:
      str: The path to the converted PDF file.
    """
    # Load credentials using st.secrets
    creds = get_credentials()
    
    # Build the Drive API client
    service = build('drive', 'v3', credentials=creds)
    
    # Upload the DOCX file and force conversion to a Google Docs file
    file_metadata = {
        'name': os.path.basename(docx_path),
        # Force conversion to Google Docs
        'mimeType': 'application/vnd.google-apps.document'
    }
    media = MediaFileUpload(
        docx_path,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
    file = service.files().create(
        body=file_metadata, 
        media_body=media, 
        fields='id'
    ).execute()
    file_id = file.get('id')
    st.write(f"Uploaded file ID: {file_id}")
    
    # Export the newly created Google Docs file as a PDF
    request = service.files().export_media(
        fileId=file_id,
        mimeType='application/pdf'
    )
    with io.FileIO(output_pdf_path, 'wb') as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            if status:
                st.write(f"Download {int(status.progress() * 100)}%.")
    
    st.write(f"Converted PDF saved as: {output_pdf_path}")
    
    # Optionally, delete the file from Drive to clean up
    service.files().delete(fileId=file_id).execute()
    st.write("Temporary file deleted from Google Drive.")
    
    return output_pdf_path

def is_name_match(name1, name2, threshold=80):
    """
    Compare two names using fuzzy matching.
    Returns True if the similarity score is above the threshold.
    """
    return fuzz.ratio(name1, name2) >= threshold

SCORE_MAP = {
    "Rarely": 1,
    "Sometimes": 2,
    "Often": 3,
    "Always": 4
}

QUESTION_CATEGORIES = {
    "I discuss issues with others to try to find solutions that meet everyone's needs.": "Collaborating",
    "I try to negotiate and use a give-and-take approach to problem situations.": "Compromising",
    "I try to meet the expectations of others.": "Accommodating",
    "I would argue my case and insist on the advantages of my point of view.": "Competing",
    "When there is a disagreement, I gather as much information as I can and keep the lines of communication open.": "Collaborating",
    "When I find myself in an argument, I usually say very little and try to leave as soon as possible.": "Avoiding",
    "I try to see conflicts from both sides. What do I need? What does the other person need? What are the issues involved?": "Collaborating",
    "I prefer to compromise when solving problems and just move on.": "Compromising",
    "I find conflicts exhilarating; I enjoy the battle of wits that usually follows.": "Competing",
    "Being in a disagreement with other people makes me feel uncomfortable and anxious.": "Avoiding",
    "I try to meet the wishes of my friends and family.": "Accommodating",
    "I can figure out what needs to be done and I am usually right.": "Competing",
    "To break deadlocks, I would meet people halfway.": "Compromising",
    "I may not get what I want but its a small price to pay for keeping the peace.": "Accommodating",
    "I avoid hard feelings by keeping my disagreements with others to myself.": "Avoiding",
}

# -----------------------------
# Strength Data (from strengths_data.py)
# -----------------------------
STRENGTH_DATA = {
    "Spirituality": {
        "underuse": "lack of purpose; disconnected from sacred",
        "optimal": "finding purpose; pursuing life meaning/connecting with sacred",
        "overuse": "fanatical; preachy/rigid beliefs"
    },
    "Gratitude": {
        "underuse": "entitled; unappreciative",
        "optimal": "positive expectations; optimistic",
        "overuse": "dependent; blind acceptance/loss of individuality"
    },
    "Hope": {
        "underuse": "apathy; pessimistic despair",
        "optimal": "positive expectations; optimistic",
        "overuse": "delusional positivity; ignoring problems"
    },
    "Humor": {
        "underuse": "grim; unapproachable",
        "optimal": "healthy levity; group-oriented",
        "overuse": "excessive teasing; belittling"
    },
    "Kindness": {
        "underuse": "aloof; selfish",
        "optimal": "compassion; empathy in action",
        "overuse": "martyrdom; compassion fatigue"
    },
    "Love": {
        "underuse": "disconnected; lonely",
        "optimal": "warmth and closeness with others",
        "overuse": "clinging; ignoring personal boundaries"
    },
    "Bravery": {
        "underuse": "fear-driven; easily intimidated",
        "optimal": "standing up for beliefs; persevering through adversity",
        "overuse": "reckless risk-taking"
    },
    "Curiosity": {
        "underuse": "uninterested; apathetic",
        "optimal": "information seeking; exploration",
        "overuse": "scattered focus; superficial dabbling"
    },
    "Love Of Learning": {
        "underuse": "disengaged with knowledge",
        "optimal": "intentional learning; open minded",
        "overuse": "analysis paralysis; ignoring practicality"
    },
    "Perspective": {
        "underuse": "unaware; limited viewpoint",
        "optimal": "wisdom-based insight; broad perspective",
        "overuse": "overthinking; constant re-evaluation"
    },
    "Creativity": {
        "underuse": "uninspired; stuck thinking",
        "optimal": "imaginative solutions; innovative",
        "overuse": "unrealistic; ignoring constraints"
    },
    "Judgment": {
        "underuse": "uncritical acceptance; naive",
        "optimal": "thoughtful consideration; balanced reasoning",
        "overuse": "hypercritical; indecisive"
    },
    "Zest": {
        "underuse": "low energy; indifferent",
        "optimal": "enthusiasm; active engagement",
        "overuse": "impulsivity; burnout from overcommitment"
    },
    "Perseverance": {
        "underuse": "easily give up; no follow-through",
        "optimal": "steadfast pursuit of goals; resilience",
        "overuse": "stubbornness; ignoring diminishing returns"
    },
    "Honesty": {
        "underuse": "deception; lack of authenticity",
        "optimal": "authentic self-expression; responsibility",
        "overuse": "bluntness; ignoring tact or empathy"
    },
    "Leadership": {
        "underuse": "lack of direction; passive group involvement",
        "optimal": "guiding vision; collaborative organization",
        "overuse": "domineering; micromanagement"
    },
    "Teamwork": {
        "underuse": "isolated; lacking group synergy",
        "optimal": "cooperative effort; shared goals",
        "overuse": "groupthink; conformity"
    },
    "Fairness": {
        "underuse": "bias; partial treatment",
        "optimal": "equitable decisions; impartial justice",
        "overuse": "inflexible adherence to rules over context"
    },
    "Forgiveness": {
        "underuse": "resentful; vengeful",
        "optimal": "letting go of grudges; understanding",
        "overuse": "enabling repeated harm; ignoring boundaries"
    },
    "Humility": {
        "underuse": "arrogance; self-centeredness",
        "optimal": "accurate self-view; respectful of others",
        "overuse": "self-effacing; inability to accept credit"
    },
    "Prudence": {
        "underuse": "impulsive; risky decisions",
        "optimal": "thoughtful planning; caution",
        "overuse": "overly cautious; fear of risk"
    },
    "Self-Regulation": {
        "underuse": "indulgent; lacking discipline",
        "optimal": "balanced habits; emotional control",
        "overuse": "rigidity; denying basic needs"
    },
    "Appreciation Of Beauty & Excellence": {
        "underuse": "oblivious; uninterested in excellence",
        "optimal": "valuing artistry, skill, or beauty",
        "overuse": "hyperfocus on perfection; aesthetic snobbery"
    },
    "Social Intelligence": {
        "underuse": "clueless about social cues; insensitive",
        "optimal": "aware of social dynamics; empathetic communication",
        "overuse": "manipulative; overthinking interactions"
    }
}

# -----------------------------
# PDF Parsing and Overlay Functions
# -----------------------------
from pdfminer.high_level import extract_text
import fitz  # PyMuPDF

def parse_via_pdf(pdf_path):
    print(f"Reading PDF using PyMuPDF from: {pdf_path}")
    doc = fitz.open(pdf_path)
    full_text = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text()
        print(f"--- Page {page_num + 1} ---")
        print(text)
        full_text += text + "\n"
    doc.close()
    print("\n=== Full Extracted Text ===")
    print(full_text)
    print("===========================\n")
    name_match = re.search(r"^(.*?)\nVIA Character Strengths Profile", full_text, re.MULTILINE)
    if name_match:
        person_name = name_match.group(1).strip()
        person_name = re.sub(r'\s+', ' ', person_name)
    else:
        person_name = "Unknown"
    pattern = re.compile(r"(\d+)\.\s+(.+)")
    matches = pattern.findall(full_text)
    results = [(int(rank), strength.strip()) for rank, strength in matches]
    print(f"Extracted Name: {person_name}")
    print("Extracted Strengths:")
    for rank, strength in results:
        print(f"{rank}: {strength}")
    return person_name, results

def create_page_number_overlay(page_width, page_height, page_number, margin=36, text_color="black"):
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))
    c.setFont("Times-Roman", 10)
    if text_color.lower() == "white":
        c.setFillColor(white)
    else:
        c.setFillColor(black)
    text = str(page_number)
    text_width = c.stringWidth(text, "Times-Roman", 10)
    x = page_width - margin - text_width
    y = margin
    c.drawString(x, y, text)
    c.save()
    packet.seek(0)
    overlay_reader = PdfReader(packet)
    return overlay_reader.pages[0]

def paginate_pdf(input_pdf, output_pdf, start_page_index=3, start_page_number=3):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    num_pages = len(reader.pages)
    for i in range(num_pages):
        page = reader.pages[i]
        if i >= start_page_index:
            page_number = start_page_number + (i - start_page_index)
            page_width = float(page.mediabox.upper_right[0])
            page_height = float(page.mediabox.upper_right[1])
            text_color = "white" if page_width > page_height else "black"
            overlay = create_page_number_overlay(page_width, page_height, page_number, text_color=text_color)
            page.merge_page(overlay)
        writer.add_page(page)
    with open(output_pdf, "wb") as f:
        writer.write(f)
    print(f"Paginated PDF saved as: {output_pdf}")

# -----------------------------
# DOCX Template Processing Functions
# -----------------------------
def fill_template(parsed_strengths, strength_data, person_name, template_path, output_docx_path):
    context = {}
    context["name"] = person_name
    for i in range(24):
        placeholder_index = i + 1
        if i < len(parsed_strengths):
            _, strength = parsed_strengths[i]
            strength_title = strength.title()
            context[f"strength{placeholder_index}"] = strength_title
            if strength_title in strength_data:
                context[f"underuse{placeholder_index}"] = strength_data[strength_title]["underuse"]
                context[f"optimal{placeholder_index}"] = strength_data[strength_title]["optimal"]
                context[f"overuse{placeholder_index}"] = strength_data[strength_title]["overuse"]
            else:
                context[f"underuse{placeholder_index}"] = ""
                context[f"optimal{placeholder_index}"] = ""
                context[f"overuse{placeholder_index}"] = ""
        else:
            context[f"strength{placeholder_index}"] = ""
            context[f"underuse{placeholder_index}"] = ""
            context[f"optimal{placeholder_index}"] = ""
            context[f"overuse{placeholder_index}"] = ""
    doc = DocxTemplate(template_path)
    doc.render(context)
    doc.save(output_docx_path)
    print(f"Template has been filled and saved as: {output_docx_path}")
    pdf_output_path = os.path.splitext(output_docx_path)[0] + ".pdf"
    pdf_output_path = convert_docx_to_pdf_gdrive(output_docx_path, pdf_output_path)
    print(f"Converted to PDF: {pdf_output_path}")
    return pdf_output_path

def fill_conflict_docs(csv_path, template_path, output_dir="."):
    df = pd.read_csv(csv_path)
    participant_names = []
    for idx, row in df.iterrows():
        full_name = str(row["First and Last Name"]).strip()
        if pd.isna(full_name) or full_name == "":
            continue
        participant_names.append(full_name)
        category_scores = {category: 0 for category in QUESTION_CATEGORIES.values()}
        for question_col, category in QUESTION_CATEGORIES.items():
            if question_col in df.columns:
                answer_text = str(row[question_col]).strip()
                numeric_score = SCORE_MAP.get(answer_text, 0)
                category_scores[category] += numeric_score
        context = {
            "name": full_name,
            "Col": category_scores["Collaborating"],
            "Com": category_scores["Competing"],
            "Avo": category_scores["Avoiding"],
            "Acc": category_scores["Accommodating"],
            "Co2": category_scores["Compromising"],
        }
        safe_name = full_name.replace(" ", "_")
        output_filename = f"{safe_name}_ConflictStyle3.docx"
        output_path = os.path.join(output_dir, output_filename)
        doc = DocxTemplate(template_path)
        doc.render(context)
        doc.save(output_path)
        # Compute PDF path and convert
        pdf_output_path = os.path.splitext(output_path)[0] + ".pdf"
        pdf_output_path = convert_docx_to_pdf_gdrive(output_path, pdf_output_path)
        os.remove(output_path)
    return participant_names

def fill_conflict_docs_for_one(csv_path, template_path, output_dir, participant_name):
    import os
    import pandas as pd
    from docxtpl import DocxTemplate
    df = pd.read_csv(csv_path)
    filtered_df = df[df["First and Last Name"] == participant_name]
    if filtered_df.empty:
        print(f"No responses found for {participant_name} in {csv_path}")
        return
    row = filtered_df.iloc[0]
    full_name = str(row["First and Last Name"]).strip()
    category_scores = {
        "Collaborating": 0,
        "Compromising": 0,
        "Avoiding": 0,
        "Accommodating": 0,
        "Competing": 0
    }
    for question_col, category in QUESTION_CATEGORIES.items():
        if question_col in df.columns:
            answer_text = str(row[question_col]).strip()
            numeric_score = SCORE_MAP.get(answer_text, 0)
            category_scores[category] += numeric_score
        else:
            print(f"Warning: '{question_col}' not found in CSV columns.")
    context = {
        "name": full_name,
        "Col": category_scores["Collaborating"],
        "Com": category_scores["Competing"],
        "Avo": category_scores["Avoiding"],
        "Acc": category_scores["Accommodating"],
        "Co2": category_scores["Compromising"],
    }
    doc = DocxTemplate(template_path)
    doc.render(context)
    safe_name = full_name.replace(" ", "_")
    output_filename = f"{safe_name}_ConflictStyle3.docx"
    output_path = os.path.join(output_dir, output_filename)
    doc.save(output_path)
    print(f"Saved DOCX: {output_path}")
    pdf_output_path = os.path.splitext(output_path)[0] + ".pdf"
    pdf_output_path = convert_docx_to_pdf_gdrive(output_path, pdf_output_path)
    print(f"Converted to PDF: {pdf_output_path}")
    os.remove(output_path)
    return pdf_output_path

# -----------------------------
# PDF Merging and Pagination Functions
# -----------------------------
def merge_custom_pages_by_index(template_pdf, cover_pdf, via_pdf, sweet_pdf, conflict_pdf, output_pdf):
    writer = PdfWriter()
    template_reader = PdfReader(template_pdf)
    cover_reader = PdfReader(cover_pdf)
    via_reader = PdfReader(via_pdf)
    sweet_reader = PdfReader(sweet_pdf)
    conflict_reader = PdfReader(conflict_pdf)
    for i in range(len(template_reader.pages)):
        if i == 0:
            for cp in cover_reader.pages:
                writer.add_page(cp)
        elif i == 4:
            for vp in via_reader.pages:
                writer.add_page(vp)
        elif i == 8:
            for sp in sweet_reader.pages:
                writer.add_page(sp)
        elif i == 11:
            for cr in conflict_reader.pages:
                writer.add_page(cr)
        else:
            writer.add_page(template_reader.pages[i])
    with open(output_pdf, "wb") as out:
        writer.write(out)
    print(f"Merged PDF created: {output_pdf}")

def create_page_number_overlay(page_width, page_height, page_number, margin=36, text_color="black"):
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))
    c.setFont("Times-Roman", 10)
    if text_color.lower() == "white":
        c.setFillColor(white)
    else:
        c.setFillColor(black)
    text = str(page_number)
    text_width = c.stringWidth(text, "Times-Roman", 10)
    x = page_width - margin - text_width
    y = margin
    c.drawString(x, y, text)
    c.save()
    packet.seek(0)
    overlay_reader = PdfReader(packet)
    return overlay_reader.pages[0]

def paginate_pdf(input_pdf, output_pdf, start_page_index=3, start_page_number=3):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    num_pages = len(reader.pages)
    for i in range(num_pages):
        page = reader.pages[i]
        if i >= start_page_index:
            page_number = start_page_number + (i - start_page_index)
            page_width = float(page.mediabox.upper_right[0])
            page_height = float(page.mediabox.upper_right[1])
            text_color = "white" if page_width > page_height else "black"
            overlay = create_page_number_overlay(page_width, page_height, page_number, text_color=text_color)
            page.merge_page(overlay)
        writer.add_page(page)
    with open(output_pdf, "wb") as f:
        writer.write(f)
    print(f"Paginated PDF saved as: {output_pdf}")
