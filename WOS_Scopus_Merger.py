import pandas as pd
import re

# File paths
wos_file = "WOS.xlsx"
scopus_file = "scopus.csv"
output_file = "merged_WoS_format.xlsx"

# Define WoS headers
wos_headers = [
    "Publication Type", "Authors", "Book Authors", "Book Editors", "Book Group Authors", "Author Full Names",
    "Book Author Full Names", "Group Authors", "Article Title", "Source Title", "Book Series Title",
    "Book Series Subtitle", "Language", "Document Type", "Conference Title", "Conference Date",
    "Conference Location", "Conference Sponsor", "Conference Host", "Author Keywords", "Keywords Plus",
    "Abstract", "Addresses", "Affiliations", "Reprint Addresses", "Email Addresses", "Researcher Ids", "ORCIDs",
    "Funding Orgs", "Funding Name Preferred", "Funding Text", "Cited References", "Cited Reference Count",
    "Times Cited, WoS Core", "Times Cited, All Databases", "180 Day Usage Count", "Since 2013 Usage Count",
    "Publisher", "Publisher City", "Publisher Address", "ISSN", "eISSN", "ISBN", "Journal Abbreviation",
    "Journal ISO Abbreviation", "Publication Date", "Publication Year", "Volume", "Issue", "Part Number",
    "Supplement", "Special Issue", "Meeting Abstract", "Start Page", "End Page", "Article Number", "DOI",
    "DOI Link", "Book DOI", "Early Access Date", "Number of Pages", "WoS Categories", "Web of Science Index",
    "Research Areas", "IDS Number", "Pubmed Id", "Open Access Designations", "Highly Cited Status",
    "Hot Paper Status", "Date of Export", "UT (Unique WOS ID)", "Web of Science Record"
]

# Scopus to WoS column mapping
column_mapping = {
    "Authors": "Authors",
    "Author(s) ID": "Researcher Ids",
    "Author full names": "Author Full Names",
    "Title": "Article Title",
    "Source title": "Source Title",
    "Volume": "Volume",
    "Issue": "Issue",
    "Art. No.": "Article Number",
    "Page start": "Start Page",
    "Page end": "End Page",
    "DOI": "DOI",
    "Abstract": "Abstract",
    "Author Keywords": "Author Keywords",
    "Index Keywords": "Keywords Plus",
    "Authors with affiliations": "Affiliations",
    "Publisher": "Publisher",
    "Conference name": "Conference Title",
    "Conference date": "Conference Date",
    "Conference location": "Conference Location",
    "ISSN": "ISSN",
    "ISBN": "ISBN",
    "PubMed ID": "Pubmed Id",
    "Document Type": "Document Type",
    "Open Access": "Open Access Designations",
    "Funding Details": "Funding Orgs",
    "Funding Text": "Funding Text",
    "References": "Cited References",
    "Cited by": "Times Cited, WoS Core",
    "Language of Original Document": "Language",
    "Year": "Publication Year"
}

def normalize_title(title):
    """Normalize article title for consistent deduplication."""
    if pd.isna(title):
        return ""
    title = str(title).lower().strip()
    # Remove punctuation, extra spaces, and special characters
    title = re.sub(r'[^a-z0-9\s]', '', title)
    title = re.sub(r'\s+', ' ', title)
    return title

try:
    # Load data
    print("Loading WoS and Scopus files...")
    df_wos = pd.read_excel(wos_file, sheet_name=0)
    df_scopus = pd.read_csv(scopus_file, encoding="utf-8", low_memory=False)

    print(f"WoS samples: {len(df_wos)}")
    print(f"Scopus samples: {len(df_scopus)}")

    # Clean column names
    df_wos.columns = df_wos.columns.str.strip().astype(str)
    df_scopus.columns = df_scopus.columns.str.strip().astype(str)

    # Rename Scopus columns to match WoS
    df_scopus = df_scopus.rename(columns=column_mapping)

    # Ensure all WoS headers exist in both dataframes
    for col in wos_headers:
        if col not in df_wos.columns:
            df_wos[col] = " "
        if col not in df_scopus.columns:
            df_scopus[col] = " "

    # Reorder columns to match WoS headers
    df_wos = df_wos[wos_headers]
    df_scopus = df_scopus[wos_headers]

    # Remove any duplicate columns
    df_wos = df_wos.loc[:, ~df_wos.columns.duplicated()]
    df_scopus = df_scopus.loc[:, ~df_scopus.columns.duplicated()]

    # Add source identifier
    df_wos['Source'] = 'WoS'
    df_scopus['Source'] = 'Scopus'

    # Check for duplicates within each dataset
    wos_titles = set(df_wos["Article Title"].apply(normalize_title))
    scopus_titles = set(df_scopus["Article Title"].apply(normalize_title))
    print(f"Unique titles in WoS: {len(wos_titles)}")
    print(f"Unique titles in Scopus: {len(scopus_titles)}")
    overlap = wos_titles.intersection(scopus_titles)
    print(f"Overlapping titles between WoS and Scopus: {len(overlap)}")
    print(f"Titles unique to Scopus: {len(scopus_titles - wos_titles)}")

    # Merge data, prioritizing WoS by placing it first
    df_combined = pd.concat([df_wos, df_scopus], ignore_index=True)

    # Normalize titles for deduplication
    df_combined["Normalized Title"] = df_combined["Article Title"].apply(normalize_title)

    # Remove duplicates based on normalized title, keeping first (WoS prioritized)
    df_cleaned = df_combined.drop_duplicates(subset=["Normalized Title"], keep="first")

    # Drop the temporary normalized title column
    df_cleaned = df_cleaned.drop(columns=["Normalized Title"])

    print(f"Final merged samples: {len(df_cleaned)}")
    print(f"WoS samples in final: {len(df_cleaned[df_cleaned['Source'] == 'WoS'])}")
    print(f"Scopus samples in final: {len(df_cleaned[df_cleaned['Source'] == 'Scopus'])}")

    # Save to Excel
    df_cleaned.to_excel(output_file, index=False)
    print(f"✅ Merge completed! File saved as '{output_file}'. WoS data was prioritized for duplicates.")

except FileNotFoundError as e:
    print(f"Error: File not found - {e}")
except Exception as e:
    print(f"Error occurred: {e}")