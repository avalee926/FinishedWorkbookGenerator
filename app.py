import streamlit as st
import os
import pandas as pd
import zipfile
from io import BytesIO
from urllib.parse import quote

#Comment to update 

import subprocess
try:
    subprocess.run(["libreoffice", "--version"], check=True)
    print("✅ LibreOffice is installed!")
except:
    print("❌ LibreOffice is missing!")

    
# Import your processing functions and constants
from functions import (
    generate_cover_pdf,
    parse_via_pdf,
    fill_template,
    fill_conflict_docs_for_one,
    merge_custom_pages_by_index,
    paginate_pdf,
    is_name_match,
    STRENGTH_DATA
)

def normalize_spaces(s: str) -> str:
    return " ".join(s.split())

def split_first_last(person_name: str):
    """
    Split a full name into (first_name, last_name) for spreadsheet export.

    Heuristics:
    - Handles 'Last, First Middle' -> ('First', 'Last')
    - Handles suffixes (e.g., Jr., Sr., II, III) by dropping them from the tail
    - If no last name present, last_name = '' (keeps it safe for spreadsheets)
    """
    if not person_name:
        return "", ""

    name = normalize_spaces(person_name)

    # Remove 'VIA' artifacts just in case (defensive; parse_via_pdf already strips)
    name = name.replace("VIA Character Strengths Profile", "").strip()

    # Common suffixes to ignore at the end
    suffixes = {"Jr", "Jr.", "Sr", "Sr.", "II", "III", "IV", "V"}

    if "," in name:
        # Format: "Last, First Middle ..."
        last, first_part = [normalize_spaces(x) for x in name.split(",", 1)]
        first_tokens = first_part.split()
        if first_tokens and first_tokens[-1] in suffixes:
            first_tokens = first_tokens[:-1]
        first = first_tokens[0] if first_tokens else ""
        last_tokens = last.split()
        if last_tokens and last_tokens[-1] in suffixes:
            last_tokens = last_tokens[:-1]
        last = " ".join(last_tokens)
        return first, last

    # Format: "First Middle Last [Suffix]"
    tokens = name.split()
    if tokens and tokens[-1] in suffixes:
        tokens = tokens[:-1]

    if len(tokens) == 1:
        return tokens[0], ""
    else:
        # First token is first name, everything after the first token collapsed into last name
        first = tokens[0]
        last = " ".join(tokens[1:])
        return first, last

def strengths_to_row(results, top_n=24):
    """Convert parse_via_pdf results -> list of strength names ordered by rank, clipped/padded to top_n."""
    strengths = [s for (_, s) in sorted(results, key=lambda x: x[0])]
    strengths = strengths[:top_n]
    if len(strengths) < top_n:
        strengths += [""] * (top_n - len(strengths))
    return strengths




# Define resource paths
BIG_TEMPLATE_PDF = os.path.join("resources", "bigTemplate.pdf")
TEAM_TEMPLATE_PDF = os.path.join("resources", "teamTemplate.pdf")
TINY_TEMPLATE_PDF = os.path.join("resources", "tinyTemplate.pdf")
CONFLICT_TEMPLATE_DOCX = os.path.join("resources", "Conflict_Template.docx")
SWEET_SPOT_TEMPLATE_DOCX = os.path.join("resources", "Sweet_Spot_Template.docx")

# Create output folder if it doesn't exist
OUTPUT_FOLDER = "output"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)  # No error if it already exists

st.title("Automated Workbook Creator")

# Sidebar: choose mode and template
mode = st.sidebar.radio("Select Mode", ["Individual", "Batch", "VIA → Spreadsheet"])
template_version = st.sidebar.selectbox("Select Template", ["Open", "Team", "Tiny"])

if template_version == "Open":
    template_pdf = BIG_TEMPLATE_PDF
elif template_version == "Team":
    template_pdf = TEAM_TEMPLATE_PDF
elif template_version == "Tiny":
    template_pdf = TINY_TEMPLATE_PDF

# -------------------------------
# INDIVIDUAL MODE
# -------------------------------
if mode == "Individual":
    st.header("Individual Mode")
    participant_name = st.text_input("Participant Name")
    term = st.text_input("Term (Date)")
    cohort = st.text_input("Cohort")
    via_file = st.file_uploader("Upload VIA File (PDF)", type=["pdf"])
    conflict_csv = st.file_uploader("Upload Conflict CSV File", type=["csv"])
    
    if st.button("Generate Workbook"):
        if participant_name and term and cohort and via_file is not None and conflict_csv is not None:
            # Save the uploaded files to disk
            via_filepath = os.path.join(OUTPUT_FOLDER, f"{participant_name}_via.pdf")
            conflict_csv_path = os.path.join(OUTPUT_FOLDER, f"{participant_name}_conflict.csv")
            with open(via_filepath, "wb") as f:
                f.write(via_file.read())
            with open(conflict_csv_path, "wb") as f:
                f.write(conflict_csv.read())
            
            # 1. Generate Cover Page
            cover_pdf = generate_cover_pdf(participant_name, term, cohort, OUTPUT_FOLDER)
            
            # 2. Parse VIA Survey
            parsed_name, results = parse_via_pdf(via_filepath)
            final_name = participant_name  # or use parsed_name if needed
            
            # 3. Fill Sweet Spot Template
            sweet_output_docx = os.path.join(OUTPUT_FOLDER, f"{final_name}_SweetSpot.docx")
            sweet_pdf = fill_template(results, STRENGTH_DATA, final_name, SWEET_SPOT_TEMPLATE_DOCX, sweet_output_docx)
            
            # 4. Process Conflict Resolution
            conflict_pdf = fill_conflict_docs_for_one(conflict_csv_path, CONFLICT_TEMPLATE_DOCX, OUTPUT_FOLDER, final_name)
            
            # 5. Merge PDFs using the selected template
            merged_pdf = os.path.join(OUTPUT_FOLDER, f"{final_name.replace(' ', '_')}_merged.pdf")
            merge_custom_pages_by_index(
                template_pdf=template_pdf,
                cover_pdf=cover_pdf,
                via_pdf=via_filepath,
                sweet_pdf=sweet_pdf,
                conflict_pdf=conflict_pdf,
                output_pdf=merged_pdf
            )
            
            # 6. Paginate the Merged PDF
            final_workbook_pdf = os.path.join(OUTPUT_FOLDER, f"{final_name.replace(' ', '_')}_workbook.pdf")
            paginate_pdf(merged_pdf, final_workbook_pdf, start_page_index=3, start_page_number=3)
            
            st.success(f"Workbook for {participant_name} generated successfully!")
            
            # Provide a download button for the generated workbook
            with open(final_workbook_pdf, "rb") as f:
                workbook_bytes = f.read()
            st.download_button("Download Workbook", workbook_bytes, file_name=os.path.basename(final_workbook_pdf), mime="application/pdf")
        else:
            st.error("Please provide all required inputs and files.")

# -------------------------------
# BATCH MODE
# -------------------------------
elif mode == "Batch":
    st.header("Batch Mode")
    term = st.text_input("Term (Date)", key="batch_term")
    cohort = st.text_input("Cohort", key="batch_cohort")
    via_files = st.file_uploader("Upload VIA Files (PDFs)", type=["pdf"], accept_multiple_files=True)
    conflict_csv_batch = st.file_uploader("Upload Conflict CSV File for Batch", type=["csv"])
    
    if st.button("Generate Batch Workbooks"):
        if term and cohort and via_files and conflict_csv_batch is not None:
            # Save the conflict CSV file
            conflict_csv_path = os.path.join(OUTPUT_FOLDER, "batch_conflict.csv")
            with open(conflict_csv_path, "wb") as f:
                f.write(conflict_csv_batch.read())
            
            # Parse the CSV for participant names
            df = pd.read_csv(conflict_csv_path)
            csv_names = set(df["First and Last Name"].str.strip().dropna().unique())
            
            # Save VIA PDFs and parse names from each
            pdf_names = {}
            for via_file in via_files:
                via_filepath = os.path.join(OUTPUT_FOLDER, via_file.name)
                with open(via_filepath, "wb") as f:
                    f.write(via_file.read())
                participant_name, _ = parse_via_pdf(via_filepath)
                pdf_names[via_file.name] = participant_name
            
            # Matching logic
            matched_pairs = []
            missing_pdf = []
            missing_csv = []
            name_mismatches = []
            
            for csv_name in csv_names:
                matched = False
                for pdf_filename, pdf_name in pdf_names.items():
                    if is_name_match(csv_name, pdf_name):
                        matched_pairs.append((csv_name, pdf_name, pdf_filename))
                        matched = True
                        break
                if not matched:
                    missing_pdf.append(csv_name)
            
            for pdf_filename, pdf_name in pdf_names.items():
                matched = False
                for csv_name in csv_names:
                    if is_name_match(csv_name, pdf_name):
                        matched = True
                        break
                if not matched:
                    missing_csv.append(pdf_name)
            
            generated_files = []
            # Process each matched pair to generate workbooks
            for csv_name, pdf_name, pdf_filename in matched_pairs:
                via_filepath = os.path.join(OUTPUT_FOLDER, pdf_filename)
                conflict_pdf = fill_conflict_docs_for_one(conflict_csv_path, CONFLICT_TEMPLATE_DOCX, OUTPUT_FOLDER, csv_name)
                if not conflict_pdf:
                    name_mismatches.append((csv_name, pdf_name))
                    continue
                
                # Generate cover page
                cover_pdf = generate_cover_pdf(csv_name, term, cohort, OUTPUT_FOLDER)
                
                # Parse VIA PDF
                parsed_name, results = parse_via_pdf(via_filepath)
                
                # Fill Sweet Spot Template
                sweet_output_docx = os.path.join(OUTPUT_FOLDER, f"{csv_name.replace(' ', '_')}_SweetSpot.docx")
                sweet_pdf = fill_template(results, STRENGTH_DATA, csv_name, SWEET_SPOT_TEMPLATE_DOCX, sweet_output_docx)
                
                # Merge PDFs
                merged_pdf = os.path.join(OUTPUT_FOLDER, f"{csv_name.replace(' ', '_')}_merged.pdf")
                merge_custom_pages_by_index(
                    template_pdf=template_pdf,
                    cover_pdf=cover_pdf,
                    via_pdf=via_filepath,
                    sweet_pdf=sweet_pdf,
                    conflict_pdf=conflict_pdf,
                    output_pdf=merged_pdf
                )
                
                # Paginate the merged PDF
                final_workbook_pdf = os.path.join(OUTPUT_FOLDER, f"{csv_name.replace(' ', '_')}_workbook.pdf")
                paginate_pdf(merged_pdf, final_workbook_pdf, start_page_index=3, start_page_number=3)
                
                generated_files.append(final_workbook_pdf)
            
            st.success("Batch processing complete!")
            st.subheader("Report Summary")
            if matched_pairs:
                st.markdown("**Successfully Generated Workbooks:**")
                for csv_name, pdf_name, _ in matched_pairs:
                    st.markdown(f"- {csv_name} (matched with {pdf_name})")
            if missing_pdf:
                st.markdown("**Participants Missing PDFs:**")
                for name in missing_pdf:
                    st.markdown(f"- {name}")
            if missing_csv:
                st.markdown("**PDFs Missing CSV Entries:**")
                for name in missing_csv:
                    st.markdown(f"- {name}")
            if name_mismatches:
                st.markdown("**Name Mismatches:**")
                for csv_name, pdf_name in name_mismatches:
                    st.markdown(f"- {csv_name} (CSV) vs. {pdf_name} (PDF)")
            
            # Create a ZIP archive in memory with all generated workbooks
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for file_path in generated_files:
                    if os.path.exists(file_path):
                        zip_file.write(file_path, arcname=os.path.basename(file_path))
            zip_buffer.seek(0)
            
            st.download_button("Download All Workbooks as ZIP", zip_buffer.getvalue(), file_name="workbooks.zip", mime="application/zip")
        else:
            st.error("Please provide all required inputs and files for batch processing.")

elif mode == "VIA → Spreadsheet":
    st.header("VIA → Spreadsheet")
    st.caption("Upload a folder’s worth of VIA PDFs and get a copy-paste block ready for Google Sheets.")
    via_files = st.file_uploader("Upload VIA Files (PDFs)", type=["pdf"], accept_multiple_files=True)


    if st.button("Extract to Table"):
        if not via_files:
            st.error("Please upload at least one VIA PDF.")
        else:
            rows = []
            failed = []

            for f in via_files:
                # Save temporarily (parse_via_pdf expects a path)
                tmp_path = os.path.join(OUTPUT_FOLDER, f.name)
                with open(tmp_path, "wb") as out:
                    out.write(f.read())

                try:
                    person_name, results = parse_via_pdf(tmp_path)
                    first, last = split_first_last(person_name)
                    strengths = strengths_to_row(results, top_n=24)  # always 24
                    rows.append([first, last] + strengths)
                except Exception as e:
                    failed.append((f.name, str(e)))

            # Build DataFrame (with headers for display / download)
            columns = ["First Name", "Last Name"] + [f"Strength {i}" for i in range(1, 25)]
            df = pd.DataFrame(rows, columns=columns)

            st.success("Extraction complete.")
            st.dataframe(df, use_container_width=True)

            # ✅ Copy-paste block (NO headers, tab-delimited)
            tsv_text_no_header = df.to_csv(index=False, sep="\t", header=False)
            st.markdown("**Copy-Paste (Google Sheets Ready — no headers):**")
            st.code(tsv_text_no_header, language="text")

            # ✅ Download CSV (keeps headers for safer record-keeping)
            csv_text = df.to_csv(index=False)
            st.download_button(
                "Download as CSV (with headers)",
                data=csv_text.encode("utf-8"),
                file_name="via_strengths_export.csv",
                mime="text/csv",
            )

            # Any failures?
            if failed:
                st.warning("Some files could not be parsed:")
                for fname, err in failed:
                    st.write(f"- {fname}: {err}")
