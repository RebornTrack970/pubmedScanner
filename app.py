import streamlit as st
import pandas as pd
import io
import re
import ssl
import os
from Bio import Entrez, Medline

if (not os.environ.get('PYTHONHTTPSVERIFY', '') and
        getattr(ssl, '_create_unverified_context', None)):
    ssl._create_default_https_context = ssl._create_unverified_context

st.set_page_config(page_title="PubMed Web Scanner", page_icon="🧬", layout="wide")

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
            doi_raw = r.get("LID", r.get("AID", ""))
            doi_link = ""
            if doi_raw and "[doi]" in doi_raw:
                clean_doi = doi_raw.split(' ')[0]
                doi_link = f"https://doi.org/{clean_doi}"

            articles.append({
                "PMID": r.get("PMID", ""),
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

def process_quartiles(df, uploaded_file):
    if uploaded_file is None:
        df["Quartile"] = "Unknown (No File)"
        return df

    try:
        sjr = pd.read_csv(uploaded_file, sep=';', quotechar='"', on_bad_lines='warn')
        
        title_col = next((c for c in sjr.columns if c.lower() in ["title", "journal title", "source title"]), None)
        quartile_col = next((c for c in sjr.columns if "quartile" in c.lower()), None)

        if not title_col or not quartile_col:
            df["Quartile"] = "Unknown (Column Error)"
            return df

        sjr["norm_title"] = sjr[title_col].apply(normalize_journal_name)
        quartile_map = dict(zip(sjr["norm_title"], sjr[quartile_col]))

        from difflib import get_close_matches
        
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
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Results')
        workbook = writer.book
        worksheet = writer.sheets['Results']
        
        link_fmt = workbook.add_format({'font_color': 'blue', 'underline': 1})
        
        # Add clickable links
        pmid_idx = df.columns.get_loc("PMID")
        doi_idx = df.columns.get_loc("DOI")
        
        for row_num, (pmid, doi) in enumerate(zip(df['PMID'], df['DOI']), start=1):
            if pmid:
                worksheet.write_url(row_num, pmid_idx, f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/", string=str(pmid), cell_format=link_fmt)
            if doi:
                worksheet.write_url(row_num, doi_idx, doi, string="DOI Link", cell_format=link_fmt)
                
        worksheet.autofit()
        
    return output.getvalue()

st.title("🧬 PubMed Research Scanner")
st.markdown("Search PubMed, match Journal Quartiles, and download formatted Excel reports.")

with st.sidebar:
    st.header("Configuration")
    scimago_file = st.file_uploader("Upload Scimago CSV (Optional)", type=["csv"])
    st.info("Download Scimago data from [scimagojr.com](https://www.scimagojr.com/journalrank.php)")

# Main Search Area
col1, col2 = st.columns(2)

with col1:
    kw_or = st.text_input("OR Keywords (e.g. lung cancer, nsclc)")
    kw_and = st.text_input("AND Keywords (e.g. biomarker)")
    study_type = st.text_input("Study Types (e.g. Clinical Trial)")

with col2:
    start_year = st.text_input("Start Year", value="2020")
    end_year = st.text_input("End Year", value="2025")
    max_results = st.number_input("Max Results", min_value=10, max_value=5000, value=50)

# Build Query
k_or_list = [x.strip() for x in kw_or.split(",") if x.strip()]
k_and_list = [x.strip() for x in kw_and.split(",") if x.strip()]
s_type_list = [x.strip() for x in study_type.split(",") if x.strip()]

or_part = " OR ".join([f'"{kw}"[Title/Abstract]' for kw in k_or_list])
and_part = " AND ".join([f'"{kw}"[Title/Abstract]' for kw in k_and_list])
type_part = " OR ".join([f'"{t}"[Publication Type]' for t in s_type_list])
date_part = f'("{start_year}"[Date - Publication] : "{end_year}"[Date - Publication])'

parts = []
if or_part: parts.append(f"({or_part})")
if and_part: parts.append(f"({and_part})")
if type_part: parts.append(f"({type_part})")
parts.append(date_part)
final_query = " AND ".join(parts)

# Action Button
if st.button("🚀 Start Search", type="primary"):
    with st.spinner("Searching PubMed..."):
        df = search_pubmed(final_query, max_results)
        
        if df.empty:
            st.warning("No results found. Try broadening your keywords.")
        else:
            st.success(f"Found {len(df)} articles.")
            
            if scimago_file:
                with st.spinner("Matching Quartiles..."):
                    df = process_quartiles(df, scimago_file)
            else:
                df["Quartile"] = "Unknown (No File)"

            cols = ["PMID", "Quartile", "Title", "First Author", "Journal", "Year", "DOI", "Article Type"]
            df = df[cols]
            st.dataframe(df, use_container_width=True)
            excel_data = to_excel(df)
          
            st.download_button(
                label="📥 Download Excel Report",
                data=excel_data,
                file_name="PubMed_Results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
