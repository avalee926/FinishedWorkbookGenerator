import re
from pypdf import PdfReader
import os
import subprocess
from docxtpl import DocxTemplate
import pandas as pd
from fuzzywuzzy import fuzz
import streamlit as st


import os
import subprocess
import pypandoc

import os
import subprocess

import os
import subprocess
import shutil
import pypandoc

import os
import io
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

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
        # This MIME type instructs Drive to convert the file to a native Docs format.
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


# strengths_data.py

STRENGTH_DATA = {
    "Spirituality": {
        "underuse": "lack of purpose; disconnected from sacred",
        "optimal": "finding purpose; pursuing life meaning/connecting with sacred",
        "overuse": "fanatical; preachy; rigid values"
    },
    "Gratitude": {
        "underuse": "entitled; unappreciative",
        "optimal": "attitude of thankfulness",
        "overuse": "ingratiation; profuse; repetitive"
    },
    "Hope": {
        "underuse": "negative; pessimistic; despair",
        "optimal": "positive expectations; optimistic",
        "overuse": "blind optimism, unrealistic"
    },
    "Humor": {
        "underuse": "overly serious",
        "optimal": "offering laughter to others; playful",
        "overuse": "giddy; tasteless/offensive"
    },
    "Kindness": {
        "underuse": "indifferent; selfish; mean",
        "optimal": "compassionate; doing for others",
        "overuse": "intrusive; compassion-fatigue"
    },
    "Love": {
        "underuse": "afraid to care; not relating",
        "optimal": "genuine, reciprocal warmth",
        "overuse": "sugary sweet; touchy-feely"
    },
    "Bravery": {
        "underuse": "cowardly; unwilling to be vulnerable",
        "optimal": "facing fears, confronting adversity",
        "overuse": "foolish, overconfident"
    },
    "Curiosity": {
        "underuse": "uninterested; self-involved",
        "optimal": "explorer; novelty-seeker; open",
        "overuse": "nosy; self-serving"
    },
    "Love Of Learning": {
        "underuse": "complacent with knowledge",
        "optimal": "information-seeking",
        "overuse": "know-it-all"
    },
    "Perspective": {
        "underuse": "shallow; superficial",
        "optimal": "sees/offers the wider view; wise",
        "overuse": "overbearing; arrogant"
    },
    "Creativity": {
        "underuse": "plain/dull; unimaginative",
        "optimal": "original; clever; imaginative",
        "overuse": "eccentric; odd; scattered"
    },
    "Judgment": {
        "underuse": "illogical; unreflective; closed-minded",
        "optimal": "analytical; rational; open-minded; logical",
        "overuse": "narrow-minded; rigid; indecisive"
    },
    "Zest": {
        "underuse": "sedentary; passive; tired",
        "optimal": "enthusiasm for life; happy; active",
        "overuse": "hyper; overactive; annoying"
    },
    "Perseverance": {
        "underuse": "lazy; helpless; giving up",
        "optimal": "persistent; overcomes all obstacles",
        "overuse": "obsessive; stubborn"
    },
    "Honesty": {
        "underuse": "phony; dishonest; inauthentic",
        "optimal": "authentic; truth sharer and seeker",
        "overuse": "self-righteous; rude"
    },
    "Leadership": {
        "underuse": "compliant; follower; passive",
        "optimal": "positively influencing others",
        "overuse": "bossy; controlling; authoritarian"
    },
    "Teamwork": {
        "underuse": "self-serving; individualistic",
        "optimal": "collaborative; loyal; socially responsible;",
        "overuse": "dependent; blind obedience; loss of individuality"
    },
    "Fairness": {
        "underuse": "bias; partial treatment",
        "optimal": "equitable decisions; impartial justice",
        "overuse": "indecisive on justice issues"
    },
    "Forgiveness": {
        "underuse": "merciless; vengeful",
        "optimal": "letting go of hurt when wronged",
        "overuse": "permissive; too lenient or soft"
    },
    "Humility": {
        "underuse": "arrogant; self-focused",
        "optimal": "others-focused; modest",
        "overuse": "limited self-image; subservient"
    },
    "Prudence": {
        "underuse": "reckless; acting before thinking",
        "optimal": "wisely cautious; planful",
        "overuse": "stuffy; rigid"
    },
    "Self-Regulation": {
        "underuse": "self-indulgent; undisciplined",
        "optimal": "self-manager of vices",
        "overuse": "inhibited; tightly wound"
    },
    "Appreciation Of Beauty & Excellence": {
        "underuse": "oblivious; mindlessness",
        "optimal": "seeing the life behind things; awe/wonder in presence of beauty",
        "overuse": "snobbery or perfectionistic; unrelenting standards"
    },
    "Social Intelligence": {
        "underuse": "clueless; insensitive",
        "optimal": "empathic; tuned in, then savvy",
        "overuse": "over-analytical; overly sensitive"
    }
}


from pdfminer.high_level import extract_text
import re
import fitz  # PyMuPDF
import re

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

    # Extract participant name
    name_match = re.search(r"^(.*?)\nVIA Character Strengths Profile", full_text, re.MULTILINE)
    name_match = re.search(r"^(.*?)\nVIA Character Strengths Profile", full_text, re.MULTILINE)
    if name_match:
        person_name = name_match.group(1).strip()
        # Replace multiple whitespace characters with a single space
        person_name = re.sub(r'\s+', ' ', person_name)
    else:
        person_name = "Unknown"
    # Extract strengths (e.g., "1. Humor")
    pattern = re.compile(r"(\d+)\.\s+(.+)")
    matches = pattern.findall(full_text)

    results = [(int(rank), strength.strip()) for rank, strength in matches]

    print(f"Extracted Name: {person_name}")
    print("Extracted Strengths:")
    for rank, strength in results:
        print(f"{rank}: {strength}")

    return person_name, results









def fill_template(parsed_strengths, strength_data, person_name, template_path, output_docx_path):
    """
    Fills the Sweet Spot Template with the parsed strengths and their corresponding definitions,
    then converts the filled DOCX to a PDF.

    Parameters:
      parsed_strengths: A list of tuples (rank, strength_name), sorted by rank.
      strength_data: A dictionary mapping strength names (Title Case) to a dict with keys "underuse", "optimal", "overuse".
      person_name: The name of the individual (to fill the {{ name }} placeholder).
      template_path: Path to the template DOCX file.
      output_docx_path: Path where the filled DOCX file will be saved.

    After saving the DOCX, the function converts it to PDF (by replacing the .docx extension with .pdf)
    using the Google Drive conversion function.
    """
    context = {}
    # Set the person's name in the template.
    context["name"] = person_name

    # Loop through 24 positions (template expects 24 rows)
    for i in range(24):
        placeholder_index = i + 1  # placeholders are 1-indexed: strength1, underuse1, etc.
        if i < len(parsed_strengths):
            # Get the strength name from parsed results
            _, strength = parsed_strengths[i]
            # Ensure the strength name is in Title Case to match the dictionary keys
            strength_title = strength.title()
            context[f"strength{placeholder_index}"] = strength_title
            # Look up the definitions from the dictionary
            if strength_title in strength_data:
                context[f"underuse{placeholder_index}"] = strength_data[strength_title]["underuse"]
                context[f"optimal{placeholder_index}"] = strength_data[strength_title]["optimal"]
                context[f"overuse{placeholder_index}"] = strength_data[strength_title]["overuse"]
            else:
                # If the strength isn't found, leave the fields blank
                context[f"underuse{placeholder_index}"] = ""
                context[f"optimal{placeholder_index}"] = ""
                context[f"overuse{placeholder_index}"] = ""
        else:
            # For any positions beyond the parsed list, leave the placeholders blank
            context[f"strength{placeholder_index}"] = ""
            context[f"underuse{placeholder_index}"] = ""
            context[f"optimal{placeholder_index}"] = ""
            context[f"overuse{placeholder_index}"] = ""

    # Load the template, render the context, and save the output DOCX.
    doc = DocxTemplate(template_path)
    doc.render(context)
    doc.save(output_docx_path)
    print(f"Template has been filled and saved as: {output_docx_path}")

    # Compute the PDF output path by replacing the .docx extension with .pdf.
    pdf_output_path = os.path.splitext(output_docx_path)[0] + ".pdf"
    
    # Convert the DOCX to PDF using the Google Drive conversion function.
    pdf_output_path = convert_docx_to_pdf_gdrive(output_docx_path, pdf_output_path)
    print(f"Converted to PDF: {pdf_output_path}")
    return pdf_output_path


import os
import pandas as pd
from docxtpl import DocxTemplate
import subprocess

# Assuming SCORE_MAP and QUESTION_CATEGORIES are defined elsewhere in your module.
# Also assuming convert_to_pdf_via_libreoffice is defined as follows:
def fill_conflict_docs(csv_path, template_path, output_dir="."):
    """
    1) Reads survey responses from `csv_path`.
    2) Converts textual answers (Rarely, Sometimes, etc.) to numeric scores using SCORE_MAP.
    3) Sums scores by category (Collaborating, Avoiding, etc.) based on QUESTION_CATEGORIES.
    4) Fills a Word template for each respondent and saves the .docx file.
    5) Converts the saved DOCX to a PDF.
    6) **Returns a list of participant names**.

    Expects a single column "First and Last Name" in the CSV.
    """
    df = pd.read_csv(csv_path)
    participant_names = []  # Store participant names

    for idx, row in df.iterrows():
        full_name = str(row["First and Last Name"]).strip()
        if pd.isna(full_name) or full_name == "":
            continue  # Skip empty names

        participant_names.append(full_name)  # Collect valid names

        # Initialize category scores
        category_scores = {category: 0 for category in QUESTION_CATEGORIES.values()}

        for question_col, category in QUESTION_CATEGORIES.items():
            if question_col in df.columns:
                answer_text = str(row[question_col]).strip()
                numeric_score = SCORE_MAP.get(answer_text, 0)
                category_scores[category] += numeric_score

        # Build template context
        context = {
            "name": full_name,
            "Col": category_scores["Collaborating"],
            "Com": category_scores["Competing"],
            "Avo": category_scores["Avoiding"],
            "Acc": category_scores["Accommodating"],
            "Co2": category_scores["Compromising"],
        }

        # Save DOCX
        safe_name = full_name.replace(" ", "_")
        output_filename = f"{safe_name}_ConflictStyle3.docx"
        output_path = os.path.join(output_dir, output_filename)
        doc = DocxTemplate(template_path)
        doc.render(context)
        doc.save(output_path)
        print(f"Saved DOCX: {output_path}")

        # Compute PDF output path by replacing .docx with .pdf
        pdf_output_path = os.path.splitext(output_path)[0] + ".pdf"
        pdf_output_path = convert_docx_to_pdf_gdrive(output_path, pdf_output_path)
        print(f"Converted to PDF: {pdf_output_path}")

        # Remove the intermediate DOCX after conversion
        os.remove(output_path)

    return participant_names  # Return the list of names


def fill_conflict_docs_for_one(csv_path, template_path, output_dir, participant_name):
    """
    Reads survey responses from `csv_path`, filters for a single participant, converts textual answers
    to numeric scores using SCORE_MAP, sums scores by category based on QUESTION_CATEGORIES, and fills a
    Word template for that participant. Saves the DOCX file to output_dir and then converts it to a PDF.

    Expects a column "First and Last Name" in the CSV.
    """
    import os
    import pandas as pd
    from docxtpl import DocxTemplate

    # Read the CSV into a DataFrame
    df = pd.read_csv(csv_path)

    # Filter for the specified participant
    filtered_df = df[df["First and Last Name"] == participant_name]
    if filtered_df.empty:
        print(f"No responses found for {participant_name} in {csv_path}")
        return

    # Process only the first matching row
    row = filtered_df.iloc[0]
    full_name = str(row["First and Last Name"]).strip()

    # Initialize category scores
    category_scores = {
        "Collaborating": 0,
        "Compromising": 0,
        "Avoiding": 0,
        "Accommodating": 0,
        "Competing": 0
    }

    # For each question column mapped in QUESTION_CATEGORIES, convert response to a number and sum by category.
    for question_col, category in QUESTION_CATEGORIES.items():
        if question_col in df.columns:
            answer_text = str(row[question_col]).strip()
            numeric_score = SCORE_MAP.get(answer_text, 0)
            category_scores[category] += numeric_score
        else:
            print(f"Warning: '{question_col}' not found in CSV columns.")

    # Build context for the DOCX template
    context = {
        "name": full_name,
        "Col": category_scores["Collaborating"],
        "Com": category_scores["Competing"],
        "Avo": category_scores["Avoiding"],
        "Acc": category_scores["Accommodating"],
        "Co2": category_scores["Compromising"],
    }

    # Load the Word template and render the context
    doc = DocxTemplate(template_path)
    doc.render(context)

    safe_name = full_name.replace(" ", "_")
    output_filename = f"{safe_name}_ConflictStyle3.docx"
    output_path = os.path.join(output_dir, output_filename)

    # Save the filled DOCX
    doc.save(output_path)
    print(f"Saved DOCX: {output_path}")

    # Compute the PDF output path by replacing the .docx extension with .pdf
    pdf_output_path = os.path.splitext(output_path)[0] + ".pdf"

    # Convert the DOCX to PDF using your helper function
    pdf_output_path = convert_docx_to_pdf_gdrive(output_path, pdf_output_path)
    print(f"Converted to PDF: {pdf_output_path}")

    # Optionally, delete the intermediate DOCX:
    os.remove(output_path)

    return pdf_output_path



from pypdf import PdfReader, PdfWriter

def merge_custom_pages_by_index(
    template_pdf,
    cover_pdf,
    via_pdf,
    sweet_pdf,
    conflict_pdf,
    output_pdf
):
    """
    Replaces specific pages (by index) in the template PDF with entire custom PDFs.
    - Page 0 -> cover_pdf
    - Page 4 -> via_pdf
    - Page 8 -> sweet_pdf
    - Page 11 -> conflict_pdf
    - All other pages remain as-is.
    """

    writer = PdfWriter()

    # Read all PDFs
    template_reader = PdfReader(template_pdf)
    cover_reader    = PdfReader(cover_pdf)
    via_reader      = PdfReader(via_pdf)
    sweet_reader    = PdfReader(sweet_pdf)
    conflict_reader = PdfReader(conflict_pdf)

    # Loop through every page in the template
    for i in range(len(template_reader.pages)):
        if i == 0:
            # Insert all pages from cover_pdf
            for cp in cover_reader.pages:
                writer.add_page(cp)
        elif i == 4:
            # Insert all pages from via_pdf
            for vp in via_reader.pages:
                writer.add_page(vp)
        elif i == 8:
            # Insert all pages from sweet_pdf
            for sp in sweet_reader.pages:
                writer.add_page(sp)
        elif i == 11:
            # Insert all pages from conflict_pdf
            for cr in conflict_reader.pages:
                writer.add_page(cr)
        else:
            # Keep the original page from the template
            writer.add_page(template_reader.pages[i])

    # Write out the merged PDF
    with open(output_pdf, "wb") as out:
        writer.write(out)

    print(f"Merged PDF created: {output_pdf}")
from pypdf import PdfReader, PdfWriter
from io import BytesIO
from reportlab.pdfgen import canvas

from reportlab.pdfgen import canvas
from reportlab.lib.colors import black, white
from io import BytesIO
from pypdf import PdfReader

from reportlab.pdfgen import canvas
from reportlab.lib.colors import black, white
from io import BytesIO
from pypdf import PdfReader

def create_page_number_overlay(page_width, page_height, page_number, margin=36, text_color="black"):
    """
    Creates a PDF overlay with the page number in Times New Roman 10 pt 
    at the lower right corner with a given margin (36 pts ~ 0.5 inch).
    
    The text color can be specified with the text_color parameter.
    """
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))
    c.setFont("Times-Roman", 10)

    # Set fill color based on text_color
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
    """
    Adds page numbers to the PDF starting at the given page index.
    
    - Pages with index less than start_page_index are left unnumbered.
    - The first numbered page (index start_page_index) is assigned the page number start_page_number.
    - The number is placed in the lower right footer in Times New Roman 10 pt.
    - If a page is horizontal (landscape), the page number is written in white text.
    """
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    num_pages = len(reader.pages)

    for i in range(num_pages):
        page = reader.pages[i]
        if i >= start_page_index:
            # Compute page number
            page_number = start_page_number + (i - start_page_index)
            # Get page dimensions from the media box
            page_width = float(page.mediabox.upper_right[0])
            page_height = float(page.mediabox.upper_right[1])
            # Determine text color: white for landscape, black for portrait
            text_color = "white" if page_width > page_height else "black"
            # Pass text_color as a keyword argument so margin stays at its default value
            overlay = create_page_number_overlay(page_width, page_height, page_number, text_color=text_color)
            page.merge_page(overlay)
        writer.add_page(page)

    with open(output_pdf, "wb") as f:
        writer.write(f)
    print(f"Paginated PDF saved as: {output_pdf}")


import os
from docxtpl import DocxTemplate
import subprocess


def generate_cover_pdf(participant_name, date, cohort, output_folder="."):
    """
    Generates a customized cover page PDF using a DOCX cover template.

    Parameters:
      participant_name (str): The participant's name.
      date (str): The term/date.
      cohort (str): The cohort identifier.
      output_folder (str): The folder where files will be saved.

    Returns:
      str: The path to the generated cover PDF.
    """
    # Define the path to your cover template DOCX
    cover_template_path = os.path.join("resources", "coverTemplate.docx")
    
    # Create a safe filename based on the participant name
    safe_name = participant_name.replace(" ", "_")
    
    # Define the path for the intermediate DOCX file
    output_docx_path = os.path.join(output_folder, f"{safe_name}_Cover.docx")
    
    # Build the context for rendering the template
    context = {
        "name": participant_name,
        "date": date,
        "cohort": cohort
    }
    
    # Render the DOCX template with the context and save it
    doc = DocxTemplate(cover_template_path)
    doc.render(context)
    doc.save(output_docx_path)
    print(f"Cover DOCX saved as: {output_docx_path}")
    
    # Create a proper PDF filename from the DOCX filename
    pdf_filename = os.path.splitext(os.path.basename(output_docx_path))[0] + ".pdf"
    output_pdf_path = os.path.join(output_folder, pdf_filename)
    
    # Convert the DOCX to PDF using your conversion function
    cover_pdf = convert_docx_to_pdf_gdrive(output_docx_path, output_pdf_path)
    print(f"Cover PDF saved as: {cover_pdf}")
    
    # Remove the intermediate DOCX file
    os.remove(output_docx_path)
    print(f"Intermediate DOCX file {output_docx_path} deleted.")
    
    return cover_pdf


def process_via_survey(pdf_path, strength_data, template_path, output_folder):
    """
    Processes the VIA survey PDF by:
      1. Extracting the participant's name and strengths using parse_via_pdf.
      2. Generating a customized Sweet Spot page by filling the template.
      3. Converting the filled template DOCX to PDF.

    Parameters:
      pdf_path: Path to the VIA survey PDF.
      strength_data: Dictionary of strengths definitions (e.g., STRENGTH_DATA).
      template_path: Path to the Sweet Spot template DOCX.
      output_folder: Folder where the generated files will be saved.

    Returns:
      The file path to the generated Sweet Spot PDF.
    """
    # Step 1: Parse the VIA PDF to get the participant's name and strengths.
    person_name, parsed_strengths = parse_via_pdf(pdf_path)

    # Use the participant's name (cleaned) to build an output DOCX path.
    safe_name = person_name.replace(" ", "_")
    output_docx_path = os.path.join(output_folder, f"{safe_name}_SweetSpot.docx")

    # Step 2: Fill the template with the parsed strengths.
    # This function renders the Sweet Spot DOCX and converts it to PDF.
    sweet_spot_pdf = fill_template(parsed_strengths, strength_data, person_name, template_path, output_docx_path)

    return sweet_spot_pdf

