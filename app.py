import streamlit as st
import pandas as pd
import io
import re
import ssl
import os
from Bio import Entrez, Medline
from docx import Document
from difflib import get_close_matches

if (not os.environ.get('PYTHONHTTPSVERIFY', '') and
        getattr(ssl, '_create_unverified_context', None)):
    ssl._create_default_https_context = ssl._create_unverified_context

st.set_page_config(page_title="PubMed Web Scanner by RTOmega", page_icon="🧬", layout="wide")

STUDY_TYPES = [
    "Adaptive Clinical Trial", "Address", "Autobiography", "Bibliography", "Biography", 
    "Books and Documents", "Case Reports", "Classical Article", "Clinical Conference", 
    "Clinical Study", "Clinical Trial", "Clinical Trial Protocol", "Clinical Trial, Phase I", 
    "Clinical Trial, Phase II", "Clinical Trial, Phase III", "Clinical Trial, Phase IV", 
    "Clinical Trial, Veterinary", "Collected Work", "Comment", "Comparative Study", 
    "Congress", "Consensus Development Conference", "Consensus Development Conference, NIH", 
    "Controlled Clinical Trial", "Corrected and Republished Article", "Dataset", "Dictionary", 
    "Directory", "Duplicate Publication", "Editorial", "Electronic Supplementary Materials", 
    "English Abstract", "Equivalence Trial", "Evaluation Study", "Expression of Concern", 
    "Festschrift", "Government Publication", "Guideline", "Historical Article", 
    "Interactive Tutorial", "Interview", "Introductory Journal Article", "Lecture", 
    "Legal Case", "Legislation", "Letter", "Meta-Analysis", "Multicenter Study", 
    "Network Meta-Analysis", "News", "Newspaper Article", "Observational Study", 
    "Observational Study, Veterinary", "Overall", "Patient Education Handout", 
    "Periodical Index", "Personal Narrative", "Portrait", "Practice Guideline", 
    "Pragmatic Clinical Trial", "Preprint", "Published Erratum", "Randomized Controlled Trial", 
    "Randomized Controlled Trial, Veterinary", "Research Support, American Recovery and Reinvestment Act", 
    "Research Support, N.I.H., Extramural", "Research Support, N.I.H., Intramural", 
    "Research Support, Non-U.S. Gov't", "Research Support, U.S. Gov't, Non-P.H.S.", 
    "Research Support, U.S. Gov't, P.H.S.", "Research Support, U.S. Gov't", 
    "Retracted Publication", "Retraction of Publication", "Review", "Scientific Integrity Review", 
    "Scoping Review", "Systematic Review", "Technical Report", "Twin Study", 
    "Validation Study", "Video-Audio Media", "Webcast"
]

def normalize_journal_name(name):
    if not isinstance(name, str):
        return ""
    name = name.lower()
    name = re.sub(r'\bthe\b', '', name)
    name = re.sub(r'[^a-z0-9 ]+', ' ', name)
    return name.strip()

def search_pubmed(query, max_results):
    Entrez.email = "pubmed_tool_web@example.com"
    try:
        handle = Entrez.esearch(db="pubmed", term=query, retmax=max_results)
        record = Entrez.read(handle)
        handle.close()
        
        ids = record["IdList"]
        if not ids:
            return pd.DataFrame()

        handle = Entrez.efetch(db="pubmed", id=",".join(ids), rettype="medline", retmode="text")
        records = Medline.parse(handle)
        
        articles = []
        for r in records:
            # Handle DOI
            doi_raw = r.get("LID", r.get("AID", ""))
            doi_link = ""
            if doi_raw and "[doi]" in doi_raw:
                clean_doi = doi_raw.split(' ')[0]
                doi_link = f"https://doi.org/{clean_doi}"

            # Handle PMID (Convert to FULL URL for LinkColumn to work)
            pmid = r.get("PMID", "")
            pmid_link = f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/" if pmid else ""

            articles.append({
                "Select": False,
                "PMID": pmid_link, # Store URL, display ID later
                "Title": r.get("TI", ""),
                "First Author": r.get("AU", ["N/A"])[0],
                "Journal": r.get("JT", ""),
                "Year": r.get("DP", "N/A")[:4],
                "DOI": doi_link,
                "Article Type": "; ".join(r.get("PT", []))
            })
        
        return pd.DataFrame(articles)
    except Exception as e:
        st.error(f"Error connecting to PubMed: {e}")
        return pd.DataFrame()

def process_quartiles(df, file_source):
    if file_source is None:
        df["Quartile"] = "Unknown (No File)"
        return df

    try:
        sjr = pd.read_csv(file_source, sep=';', quotechar='"', on_bad_lines='warn')
        
        title_col = next((c for c in sjr.columns if c.lower() in ["title", "journal title", "source title"]), None)
        quartile_col = next((c for c in sjr.columns if "quartile" in c.lower()), None)

        if not title_col or not quartile_col:
            df["Quartile"] = "Unknown (Column Error)"
            return df

        sjr["norm_title"] = sjr[title_col].apply(normalize_journal_name)
        quartile_map = dict(zip(sjr["norm_title"], sjr[quartile_col]))
        
        journal_names_norm = df["Journal"].apply(normalize_journal_name)
        quartiles = []
        
        for norm_name in journal_names_norm:
            if norm_name in quartile_map:
                quartiles.append(quartile_map[norm_name])
            else:
                close = get_close_matches(norm_name, quartile_map.keys(), n=1, cutoff=0.8)
                quartiles.append(quartile_map[close[0]] if close else "Unknown")

        df["Quartile"] = quartiles
        return df
    except Exception as e:
        st.warning(f"Error processing Quartile file: {e}")
        df["Quartile"] = "Error"
        return df

def to_excel(df):
    output = io.BytesIO()
    export_df = df.drop(columns=['Select'], errors='ignore')
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        export_df.to_excel(writer, index=False, sheet_name='Results')
        workbook = writer.book
        worksheet = writer.sheets['Results']
        
        link_fmt = workbook.add_format({'font_color': 'blue', 'underline': 1})
        
        pmid_idx = export_df.columns.get_loc("PMID")
        doi_idx = export_df.columns.get_loc("DOI")
        
        for row_num, (pmid_url, doi_url) in enumerate(zip(export_df['PMID'], export_df['DOI']), start=1):
            if pmid_url:
                try:
                    display_id = pmid_url.strip("/").split("/")[-1]
                except:
                    display_id = "Link"
                worksheet.write_url(row_num, pmid_idx, pmid_url, string=display_id, cell_format=link_fmt)
            
            if doi_url:
                display_doi = doi_url.replace("https://doi.org/", "")
                worksheet.write_url(row_num, doi_idx, doi_url, string=display_doi, cell_format=link_fmt)
                
        worksheet.autofit()
        
    return output.getvalue()

def generate_word_summary(pmid_urls):
    """Fetches abstracts for selected PMIDs (passed as URLs) and creates a Word doc."""
    Entrez.email = "pubmed_tool_web@example.com"
    doc = Document()
    doc.add_heading('PubMed Article Summaries', level=0)
    
    # Extract IDs from URLs
    clean_ids = []
    for url in pmid_urls:
        if url and "pubmed" in url:
            parts = url.strip("/").split("/")
            if parts:
                clean_ids.append(parts[-1])
    
    if not clean_ids:
        return None

    try:
        handle = Entrez.efetch(db="pubmed", id=",".join(clean_ids), rettype="abstract", retmode="xml")
        records = Entrez.read(handle)
        handle.close()
        
        for record in records.get("PubmedArticle", []):
            citation = record.get("MedlineCitation", {})
            article = citation.get("Article", {})
            pmid = citation.get("PMID", "N/A")
            title = article.get("ArticleTitle", "No title")
            
            authors_list = []
            for author in article.get("AuthorList", []):
                if "LastName" in author and "ForeName" in author:
                    authors_list.append(f"{author['ForeName']} {author['LastName']}")
            authors_str = ", ".join(authors_list) if authors_list else "No authors listed"

            journal = article.get("Journal", {}).get("Title", "No journal")
            
            abstract_parts = article.get("Abstract", {}).get("AbstractText", [])
            abstract_text = " ".join(abstract_parts) if abstract_parts else "No abstract found."

            doc.add_heading(f"PMID: {pmid}", level=2)
            
            p = doc.add_paragraph()
            p.add_run("Title: ").bold = True
            p.add_run(title)
            
            p = doc.add_paragraph()
            p.add_run("Authors: ").bold = True
            p.add_run(authors_str)
            
            p = doc.add_paragraph()
            p.add_run("Journal: ").bold = True
            p.add_run(journal)
            
            p = doc.add_paragraph()
            p.add_run("Abstract: ").bold = True
            p.add_run(abstract_text)
            
            doc.add_paragraph("_" * 50) 

        doc_buffer = io.BytesIO()
        doc.save(doc_buffer)
        return doc_buffer.getvalue()

    except Exception as e:
        print(e)
        return None

col_header, col_tutorial = st.columns([7, 1]) # Adjust ratio to move button

with col_header:
    st.title("🧬 PubMed Research Scanner")
    st.markdown("Search PubMed, select articles, and download Excel lists or Word summaries.")

with col_tutorial:
    st.link_button("Turkish Video Tutorial", "https://www.youtube.com/watch?v=KvsBj1QGqso")

with st.sidebar:
    st.header("Configuration")
    uploaded_scimago = st.file_uploader("Upload Scimago CSV (Optional)", type=["csv"])
    
    scimago_source = None
    default_filename = "scimago.csv"
    
    if uploaded_scimago is not None:
        scimago_source = uploaded_scimago
        st.success("✅ Using your uploaded CSV.")
    elif os.path.exists(default_filename):
        scimago_source = default_filename
        st.info("ℹ️ Using default 'scimago.csv'.")
    else:
        st.warning("⚠️ No Scimago file found.")

col1, col2 = st.columns(2)
with col1:
    kw_or = st.text_input("OR Keywords (e.g. lung cancer, nsclc)")
    kw_and = st.text_input("AND Keywords (e.g. biomarker)")
    study_type = st.multiselect("Study Types", STUDY_TYPES)
with col2:
    start_year = st.text_input("Start Year", value="2020")
    end_year = st.text_input("End Year", value="2025")
    max_results = st.number_input("Max Results", min_value=10, max_value=5000, value=50)

k_or_list = [x.strip() for x in kw_or.split(",") if x.strip()]
k_and_list = [x.strip() for x in kw_and.split(",") if x.strip()]
or_part = " OR ".join([f'"{kw}"[Title/Abstract]' for kw in k_or_list])
and_part = " AND ".join([f'"{kw}"[Title/Abstract]' for kw in k_and_list])
type_part = " OR ".join([f'"{t}"[Publication Type]' for t in study_type])
date_part = f'("{start_year}"[Date - Publication] : "{end_year}"[Date - Publication])'

parts = []
if or_part: parts.append(f"({or_part})")
if and_part: parts.append(f"({and_part})")
if type_part: parts.append(f"({type_part})")
parts.append(date_part)
final_query = " AND ".join(parts)

if 'search_results' not in st.session_state:
    st.session_state.search_results = pd.DataFrame()

if st.button("🔎 Start Search", type="primary"):
    with st.spinner("Searching PubMed..."):
        df = search_pubmed(final_query, max_results)
        
        if df.empty:
            st.warning("No results found.")
            st.session_state.search_results = pd.DataFrame()
        else:
            if scimago_source:
                df = process_quartiles(df, scimago_source)
            else:
                df["Quartile"] = "Unknown (No File)"
            
            cols = ["Select", "PMID", "Quartile", "Title", "First Author", "Journal", "Year", "DOI", "Article Type"]
            df = df[cols]
            st.session_state.search_results = df

if not st.session_state.search_results.empty:
    st.divider()
    st.subheader("Search Results")
    st.caption("Select rows to generate a Word summary.")

    edited_df = st.data_editor(
        st.session_state.search_results,
        column_config={
            "Select": st.column_config.CheckboxColumn(
                "Select",
                help="Select to include in Word Summary",
                default=False,
            ),
            "PMID": st.column_config.LinkColumn(
                label="PMID",
                display_text=r"https://pubmed\.ncbi\.nlm\.nih\.gov/(.*?)/"
            ),
            "DOI": st.column_config.LinkColumn(
                label="DOI",
                display_text=r"https://doi\.org/(.*)"
            )
        },
        disabled=["PMID", "Quartile", "Title", "First Author", "Journal", "Year", "DOI", "Article Type"],
        hide_index=True,
        use_container_width=True
    )

    col_d1, col_d2 = st.columns([1, 1])

    with col_d1:
        excel_data = to_excel(edited_df)
        st.download_button(
            label="📥 Download Excel List",
            data=excel_data,
            file_name="PubMed_List.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col_d2:
        selected_rows = edited_df[edited_df["Select"] == True]
        
        if not selected_rows.empty:
            if st.button("📄 Generate Word Summary for Selected"):
                with st.spinner("Fetching abstracts and generating Word doc..."):
                    pmid_urls = selected_rows["PMID"].astype(str).tolist()
                    word_data = generate_word_summary(pmid_urls)
                    
                    if word_data:
                        st.download_button(
                            label="⬇️ Download Word Summary (.docx)",
                            data=word_data,
                            file_name="PubMed_Abstracts_Summary.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    else:
                        st.error("Failed to generate Word document.")
        else:
            st.info("Select checkboxes above to enable Word summary generation.")
