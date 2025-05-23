import pandas as pd
import re

# ── CONFIG ─────────────────────────────────────────────────────────────────────
wos_file     = "WOS_base.xlsx"
scopus_file  = "scopus_base.csv"
output_file  = "merged_WoS_format.xlsx"
# ────────────────────────────────────────────────────────────────────────────────

# 1) Full WoS headers
wos_headers = [
    "Publication Type","Authors","Book Authors","Book Editors","Book Group Authors","Author Full Names",
    "Book Author Full Names","Group Authors","Article Title","Source Title","Book Series Title",
    "Book Series Subtitle","Language","Document Type","Conference Title","Conference Date",
    "Conference Location","Conference Sponsor","Conference Host","Author Keywords","Keywords Plus",
    "Abstract","Addresses","Affiliations","Reprint Addresses","Email Addresses","Researcher Ids","ORCIDs",
    "Funding Orgs","Funding Name Preferred","Funding Text","Cited References","Cited Reference Count",
    "Times Cited, WoS Core","Times Cited, All Databases","180 Day Usage Count","Since 2013 Usage Count",
    "Publisher","Publisher City","Publisher Address","ISSN","eISSN","ISBN","Journal Abbreviation",
    "Journal ISO Abbreviation","Publication Date","Publication Year","Volume","Issue","Part Number",
    "Supplement","Special Issue","Meeting Abstract","Start Page","End Page","Article Number","DOI",
    "DOI Link","Book DOI","Early Access Date","Number of Pages","WoS Categories","Web of Science Index",
    "Research Areas","IDS Number","Pubmed Id","Open Access Designations","Highly Cited Status",
    "Hot Paper Status","Date of Export","UT (Unique WOS ID)","Web of Science Record"
]

# 2) Scopus → WoS mapping (including Year, References, Language)
column_mapping = {
    "Authors":"Authors",
    "Author(s) ID":"Researcher Ids",
    "Author full names":"Author Full Names",
    "Title":"Article Title",
    "Source title":"Source Title",
    "Volume":"Volume",
    "Issue":"Issue",
    "Art. No.":"Article Number",
    "Page start":"Start Page",
    "Page end":"End Page",
    "DOI":"DOI",
    "Abstract":"Abstract",
    "Author Keywords":"Author Keywords",
    "Index Keywords":"Keywords Plus",
    "Authors with affiliations":"Affiliations",
    "Publisher":"Publisher",
    "Conference name":"Conference Title",
    "Conference date":"Conference Date",
    "Conference location":"Conference Location",
    "ISSN":"ISSN",
    "ISBN":"ISBN",
    "PubMed ID":"Pubmed Id",
    "Document Type":"Document Type",
    "Open Access":"Open Access Designations",
    "Funding Details":"Funding Orgs",
    "Funding Text":"Funding Text",
    "References":"Cited References",
    "Cited by":"Times Cited, WoS Core",
    "Year":"Publication Year",
    "Language of Original Document":"Language"
}

# 3) Helpers
def normalize_title(t):
    if pd.isna(t):
        return ""
    s = re.sub(r'[^A-Za-z0-9 ]', ' ', str(t).lower())
    return re.sub(r'\s+', ' ', s).strip()

country_map = {
    r"peoples\s*r\s*china":          "China",
    r"\bprc\b":                      "China",
    r"\bchina\b":                    "China",
    r"u\.?s\.?a\.?":                 "United States",
    r"united\s+states":              "United States",
    r"\buae\b":                      "United Arab Emirates",
    r"united\s+arab\s+emirates":     "United Arab Emirates"
}

def canonical_country(tok):
    key = tok.strip().lower()
    for pat, canon in country_map.items():
        if re.search(pat, key):
            return canon
    return tok.strip().title()

def normalize_addresses(addr):
    if not isinstance(addr, str) or not addr.strip():
        return ""
    authors = []
    for chunk in addr.split(";"):
        p = chunk.strip()
        if not p: continue
        parts = [x.strip() for x in p.rsplit(",", 5)]
        raw = parts[-1]
        parts[-1] = canonical_country(raw)
        authors.append(", ".join(parts))
    return "; ".join(authors)

# ────────────────────────────────────────────────────────────────────────────────

try:
    print("🔄 Loading WOS and Scopus...")
    df_wos    = pd.read_excel(wos_file, dtype=str)
    df_scopus = pd.read_csv(scopus_file, dtype=str, encoding="utf-8", low_memory=False)
    print(f"WoS samples: {len(df_wos)}")
    print(f"Scopus samples: {len(df_scopus)}")

    # Clean headers
    df_wos.columns    = df_wos.columns.str.strip()
    df_scopus.columns = df_scopus.columns.str.strip()

    # Rename scopus → WoS
    df_scopus.rename(columns=column_mapping, inplace=True)

    # Ensure all headers
    for col in wos_headers:
        if col not in df_wos:    df_wos[col]    = ""
        if col not in df_scopus: df_scopus[col] = ""

    # Reorder & drop dup cols
    df_wos    = df_wos.loc[:, ~df_wos.columns.duplicated()][wos_headers]
    df_scopus = df_scopus.loc[:, ~df_scopus.columns.duplicated()][wos_headers]

    # Fill cited references
    df_scopus["Cited References"] = df_scopus["Cited References"].fillna("")

    # Default Scopus language to English if still blank
    df_scopus["Language"] = df_scopus["Language"].fillna("").replace("", "English")

    # Copy Affiliations → Addresses if empty
    mask = df_scopus["Addresses"].isna() | (df_scopus["Addresses"] == "")
    df_scopus.loc[mask, "Addresses"] = df_scopus.loc[mask, "Affiliations"]

    # Label
    df_wos["Source"]    = "WoS"
    df_scopus["Source"] = "Scopus"

    # Pre-merge stats
    wos_titles    = set(df_wos["Article Title"].apply(normalize_title))
    scopus_titles = set(df_scopus["Article Title"].apply(normalize_title))
    print(f"Unique titles in WOS: {len(wos_titles)}")
    print(f"Unique titles in Scopus: {len(scopus_titles)}")
    overlap = wos_titles.intersection(scopus_titles)
    print(f"Overlapping titles: {len(overlap)}")
    print(f"Titles unique to Scopus: {len(scopus_titles - wos_titles)}")

    # Merge & dedupe
    df_all    = pd.concat([df_wos, df_scopus], ignore_index=True)
    df_all["_norm"] = df_all["Article Title"].apply(normalize_title)
    df_merged = df_all.drop_duplicates(subset="_norm", keep="first").drop(columns="_norm")

    # Post-merge stats
    total     = len(df_merged)
    wos_n     = (df_merged["Source"]=="WoS").sum()
    scopus_n  = (df_merged["Source"]=="Scopus").sum()
    print(f"Final merged samples: {total}")
    print(f"WoS samples in final: {wos_n}")
    print(f"Scopus samples in final: {scopus_n}")

    # Normalize country tokens
    df_merged["Addresses"] = df_merged["Addresses"].apply(normalize_addresses)

    # Save
    df_merged.to_excel(output_file, index=False)
    print(f"✅ Merge completed! File saved to '{output_file}'")
except FileNotFoundError as e:
    print(f"Error: File not found - {e}")
except Exception as e:
    print(f"Error occurred: {e}")
