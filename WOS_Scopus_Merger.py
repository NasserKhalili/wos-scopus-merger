import pandas as pd
import re

# ── CONFIG ─────────────────────────────────────────────────────────────────────
wos_file     = "WOS-Filtered.xlsx"
scopus_file  = "scopus-Filtered.csv"
output_file_cocitation = "merged_WOS_format.xlsx"
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

# 2) Scopus → WoS mapping
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

# ── HELPERS ─────────────────────────────────────────────────────────────────────
def normalize_title(s):
    s = str(s).lower().strip()
    s = re.sub(r'[^\w\s]+', ' ', s)
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

def get_dup_titles(df):
    norm = df["Article Title"].map(normalize_title)
    mask = norm.duplicated(keep=False)
    return df.loc[mask, "Article Title"].drop_duplicates().tolist()

def print_dup_info(name, titles):
    n = len(titles)
    if n == 0:
        print(f"{name}: no duplicates")
    elif n <= 5:
        print(f"{name} duplicates ({n}):")
        for t in titles:
            print("  ", t)
    else:
        print(f"{name} duplicates (showing first 5 of {n}):")
        for t in titles[:5]:
            print("  ", t)

def fix_authors(a):
    """
    Convert "Last, F." or "Last F." into "Last, F"
    and join multiple authors with semicolon.
    """
    if pd.isna(a) or not str(a).strip():
        return ""
    parts = [p.strip() for p in str(a).split(";") if p.strip()]
    out = []
    for p in parts:
        p2 = re.sub(r'\([^)]*\)', '', p)
        p2 = re.sub(r'\d+', '', p2).strip()
        if "," in p2:
            last, rest = [x.strip() for x in p2.split(",", 1)]
            initials = re.sub(r'[^A-Z\.]', '', rest).strip()
            out.append(f"{last}, {initials}")
        else:
            tokens = p2.split()
            if len(tokens) >= 2:
                last = tokens[0]
                initials = "".join(re.findall(r'[A-Z]', " ".join(tokens[1:])))
                out.append(f"{last}, {initials}.")
            else:
                out.append(p2)
    return "; ".join(out)

def clean_fullnames(fn):
    """
    Convert "Last, First M." or "Last, First" into "First Last"
    and remove IDs in parentheses.
    """
    if pd.isna(fn) or not str(fn).strip():
        return ""
    parts = [p.strip() for p in str(fn).split(";") if p.strip()]
    out = []
    for p in parts:
        p2 = re.sub(r'\([^)]*\)', '', p).strip()
        if "," in p2:
            last, first = [x.strip() for x in p2.split(",", 1)]
            out.append(f"{first} {last}")
        else:
            out.append(p2)
    return "; ".join(out)

country_map = {
    r"peoples\s*r\s*china": "China",
    r"\bprc\b":             "China",
    r"\bchina\b":           "China",
    r"u\.?s\.?a\.?":        "United States",
    r"united\s+states":     "United States",
    r"\buae\b":             "United Arab Emirates",
    r"united\s+arab\s+emirates": "United Arab Emirates"
}
def canonical_country(tok):
    key = tok.strip().lower()
    for pat, canon in country_map.items():
        if re.search(pat, key):
            return canon
    return tok.strip().title()

def normalize_addresses(addr):
    """
    For each address chunk, split off country, map it to canonical form,
    then rejoin. Keeps institution+city together.
    """
    if not isinstance(addr, str) or not addr.strip():
        return ""
    out = []
    for chunk in addr.split(";"):
        p = chunk.strip()
        if not p:
            continue
        seg = [x.strip() for x in p.rsplit(",", 5)]
        raw_country = seg[-1]
        seg[-1] = canonical_country(raw_country)
        out.append(", ".join(seg))
    return "; ".join(out)

def normalize_journal_name(journal):
    """
    Upper-case, remove dots, condense spaces.
    """
    j = journal.strip().upper()
    j = re.sub(r'\.', ' ', j)
    j = re.sub(r'\s+', ' ', j)
    return j.strip()

def normalize_author_name(author):
    """
    Ensure "First Last" or "Last, First" both become "Last FirstInitials" (no comma).
    """
    author = author.strip()
    if "," in author:
        last, first = [x.strip() for x in author.split(",", 1)]
        initials = "".join(re.findall(r'[A-Z]', first))
        return f"{last.lower()} {initials.lower()}"
    tokens = author.split()
    if len(tokens) >= 2:
        last = tokens[0]
        initials = "".join(re.findall(r'[A-Z]', " ".join(tokens[1:])))
        return f"{last.lower()} {initials.lower()}"
    return author.lower()

def parse_ref(ref):
    """
    Parse a reference into "author, year, journal" format (excluding volume and page).
    WoS: "Smith J., 2019, J NAME, V12, P123" -> "smith j, 2019, J NAME"
    Scopus: "Smith J, Doe A, Title, J NAME, 12(3), pp 123-130, 2019" -> "smith j, 2019, J NAME"
    """
    ref = ref.strip()
    # Step 1: Split carefully, preserving commas in parentheses
    parts = []
    current = ""
    paren_level = 0
    for char in ref:
        if char == "," and paren_level == 0:
            if current.strip():
                parts.append(current.strip())
            current = ""
        else:
            current += char
            if char == "(":
                paren_level += 1
            elif char == ")":
                paren_level -= 1
    if current.strip():
        parts.append(current.strip())

    if len(parts) < 2:
        return ""

    # Step 2: Extract author
    author = normalize_author_name(parts[0])
    if not author:
        return ""

    # Step 3: Find year
    year = None
    year_idx = -1
    for i, part in enumerate(parts):
        match = re.search(r'\b(\d{4})\b|\((\d{4})\)', part)
        if match:
            year = match.group(1) or match.group(2)
            year_idx = i
            break
    if not year:
        return ""

    # Step 4: Determine format and extract journal
    journal = ""
    if year_idx == 1:  # WoS format: "author, year, journal, volume, page"
        journal_start = 2
        i = journal_start
        while i < len(parts):
            part = parts[i].strip().lower()
            # Stop at volume (e.g., "v12", "12")
            if re.match(r"v?\d+", part) or re.match(r"\d+\(?\d*\)?", part):
                journal = normalize_journal_name(" ".join(parts[journal_start:i]))
                break
            # Stop at DOI or other trailing info
            elif "doi" in part.lower():
                journal = normalize_journal_name(" ".join(parts[journal_start:i]))
                break
            i += 1
        if not journal:  # If no volume or DOI found, treat remaining as journal
            journal = normalize_journal_name(" ".join(parts[journal_start:i]))
    else:  # Scopus format: "author, title, journal, volume(issue), pages, year"
        # Skip the title by looking for the first part that could be a journal
        i = 1
        journal_start = i
        while i < len(parts):
            part = parts[i].strip().lower()
            # Stop at volume (e.g., "12", "12(3)") or pages (e.g., "pp. 123-130")
            if re.match(r"\d+\(?\d*\)?", part):
                journal = normalize_journal_name(" ".join(parts[journal_start:i]))
                break
            elif re.match(r"pp\.?\s*\d+-\d+|\d+-\d+", part):
                journal = normalize_journal_name(" ".join(parts[journal_start:i]))
                break
            elif "doi" in part.lower() or re.search(r'\b\d{4}\b|\(\d{4}\)', part):
                journal = normalize_journal_name(" ".join(parts[journal_start:i]))
                break
            i += 1
            journal_start = i
        if not journal and journal_start < len(parts):
            journal = normalize_journal_name(" ".join(parts[journal_start:i]))

    # Step 5: Validate and build the reference
    if not journal or re.search(r'\d', journal):  # Skip if journal contains numbers
        return ""
    ref_parts = [f"{author}, {year}, {journal}"]
    return ", ".join(ref_parts)

def normalize_cr_cocitation(cr, source):
    """
    Normalize cited references for co-citation analysis.
    """
    if pd.isna(cr) or not str(cr).strip():
        return ""
    items = [r.strip() for r in str(cr).split(";") if r.strip()]
    out = []
    for item in items:
        parsed = parse_ref(item)
        if parsed:
            out.append(parsed)
    return "; ".join(sorted(set(out)))

def normalize_cr_citation(cr, source):
    """
    Preserve original cited references for citation analysis.
    """
    if pd.isna(cr) or not str(cr).strip():
        return ""
    items = [r.strip() for r in str(cr).split(";") if r.strip()]
    return "; ".join(items)

# ── MAIN ───────────────────────────────────────────────────────────────────────

# 1) Load
df_wos = pd.read_excel(wos_file, dtype=str)
df_scopus = pd.read_csv(scopus_file, dtype=str, encoding="utf-8", low_memory=False)

# 2) Validate columns
if "Article Title" not in df_wos.columns:
    raise KeyError("'Article Title' not found in WoS data. Check column names.")
title_cols = [c for c in df_scopus.columns if c.strip().lower() in {"title", "document title"}]
if not title_cols:
    raise KeyError("'Title' or 'Document Title' not found in Scopus data. Check column names.")
scopus_title_col = title_cols[0]
df_scopus.rename(columns={scopus_title_col: "Article Title"}, inplace=True)

# 3) Rename & ensure headers
df_wos.columns = df_wos.columns.str.strip()
df_scopus.columns = df_scopus.columns.str.strip()
df_scopus.rename(columns=column_mapping, inplace=True)

for c in wos_headers:
    if c not in df_wos:
        df_wos[c] = ""
    if c not in df_scopus:
        df_scopus[c] = ""

df_wos["Source"] = "WoS"
df_scopus["Source"] = "Scopus"
cols = wos_headers + ["Source"]
df_wos = df_wos.loc[:, ~df_wos.columns.duplicated()].reindex(columns=cols)
df_scopus = df_scopus.loc[:, ~df_scopus.columns.duplicated()].reindex(columns=cols)

# 4) Filter empty titles
df_wos = df_wos[df_wos["Article Title"].notna() & (df_wos["Article Title"].str.strip() != "")]
df_scopus = df_scopus[df_scopus["Article Title"].notna() & (df_scopus["Article Title"].str.strip() != "")]

# 5) Pre-merge counts
wos_records = len(df_wos)
scopus_records = len(df_scopus)
total_pre = wos_records + scopus_records

# 6) Pre-merge duplicate info
wos_dupes = get_dup_titles(df_wos)
scopus_dupes = get_dup_titles(df_scopus)

print_dup_info("WOS", wos_dupes)
print()
print_dup_info("Scopus", scopus_dupes)
print()

print(f"WOS Records = {wos_records}")
print(f"Scopus Records = {scopus_records}")
print(f"Total Records (pre-merge) = {total_pre}\n")

print(f"WOS Duplicates = {len(wos_dupes)/2}")
print(f"Scopus Duplicates = {len(scopus_dupes)/2}\n")

# 7) Clean fields
for D in (df_wos, df_scopus):
    D["Authors"] = D["Authors"].apply(fix_authors)
    D["Author Full Names"] = D["Author Full Names"].apply(clean_fullnames)
    D["Addresses"] = D["Addresses"].apply(normalize_addresses)
    D["Affiliations"] = D["Affiliations"].apply(normalize_addresses)
    D["Cited References (Co-Citation)"] = D.apply(lambda row: normalize_cr_cocitation(row["Cited References"], row["Source"]), axis=1)
    D["Cited References (Citation)"] = D.apply(lambda row: normalize_cr_citation(row["Cited References"], row["Source"]), axis=1)

# 8) Overlap before merge
norm_w = set(df_wos["Article Title"].map(normalize_title))
norm_s = set(df_scopus["Article Title"].map(normalize_title))
overlap = len(norm_w & norm_s)
print(f"Total Overlap After Merging = {overlap}\n")

# 9) Merge & post-merge, preserve unique CR entries
df_all = pd.concat([df_wos, df_scopus], ignore_index=True)
df_all["_norm"] = df_all["Article Title"].map(normalize_title)

# Combine CR fields for overlapping records
def combine_cr_cocitation(group):
    cr_list = [cr for cr in group["Cited References (Co-Citation)"] if pd.notna(cr) and cr.strip()]
    all_refs = []
    for cr in cr_list:
        refs = [r.strip() for r in cr.split(";") if r.strip()]
        all_refs.extend(refs)
    return "; ".join(sorted(set(all_refs)))

def combine_cr_citation(group):
    cr_list = [cr for cr in group["Cited References (Citation)"] if pd.notna(cr) and cr.strip()]
    all_refs = []
    for cr in cr_list:
        refs = [r.strip() for r in cr.split(";") if r.strip()]
        all_refs.extend(refs)
    return "; ".join(sorted(set(all_refs)))

if overlap > 0:
    grouped = df_all.groupby("_norm")
    df_all = grouped.apply(lambda x: pd.Series({
        "_norm": x.name, 
        **x.iloc[0].to_dict(),
        "Cited References (Co-Citation)": combine_cr_cocitation(x),
        "Cited References (Citation)": combine_cr_citation(x),
        "Source": " ".join(set(x["Source"]))
    }), include_groups=False).reset_index(drop=True)

# Log dropped records
dropped = df_all[df_all.duplicated(subset=["_norm", "Authors", "Publication Year"], keep=False)]
if not dropped.empty:
    dropped[["Article Title", "Source", "DOI", "_norm", "Authors", "Publication Year"]].to_excel("dropped_records.xlsx", index=False)

df_final_cocitation = df_all.drop_duplicates(subset=["_norm", "Authors", "Publication Year"], keep="first").drop(columns=["Cited References (Citation)", "_norm"])
df_final_citation = df_all.drop_duplicates(subset=["_norm", "Authors", "Publication Year"], keep="first").drop(columns=["Cited References (Co-Citation)", "_norm"])

wos_final = df_final_cocitation.Source.str.contains("WoS").sum()
scopus_final = df_final_cocitation.Source.str.contains("Scopus").sum()

print(f"Total Records Written: {len(df_final_cocitation)}  (WOS: {wos_final}, Scopus: {scopus_final - overlap})")
