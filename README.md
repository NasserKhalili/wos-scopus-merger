# WoS + Scopus Merger → Unified WoS Format

**Description:**  
A Python tool to merge Web of Science (`.xlsx`) and Scopus (`.csv`) exports into a single, deduplicated Web of Science–style Excel file. Ideal for creating a comprehensive bibliographic dataset before downstream bibliometric analysis.

---

## 🔧 Script Overview

### `merge_wos_scopus.py`
- **Input:**  
  - `WOS_New.xlsx` — Web of Science export  
  - `scopus_New.csv` — Scopus export  
- **Process:**  
  1. Normalize column names and article titles (lowercase, strip punctuation)  
  2. Map Scopus fields onto the full WoS schema, filling missing columns with blanks  
  3. Concatenate both datasets and remove duplicates (WoS records preferred)  
- **Output:**  
  - `merged_WoS_format.xlsx` — unified, deduplicated dataset in the full WoS column layout  

---

## 💡 How to Use

1. **Clone** this repository:  
   ```bash
   git clone https://github.com/yourusername/wos-scopus-merger.git
   cd wos-scopus-merger

2. **Install** dependencies:
   pip install pandas openpyxl

3. **Place** your export files in the project directory:
   WOS.xlsx
  scopus.csv

4. **Run** the merge script:
   python merge_wos_scopus.py

**Find** the result in:
  merged_WoS_format.xlsx

## 📂 Input / Output Examples

| File                         | Description                                                    |
|------------------------------|----------------------------------------------------------------|
| `WOS.xlsx`                   | Original Web of Science Excel export                           |
| `scopus.csv`                 | Original Scopus CSV export                                     |
| `merged_WoS_format.xlsx`     | Unified, deduplicated dataset in full Web of Science schema    |

## 🔗 Next Steps

1. **Filter** the merged workbook (e.g., apply PRISMA guidelines, remove irrelevant rows).  
2. **Convert** your filtered `merged_WoS_format.xlsx` into analysis‐ready formats using the [wos-format-converter](https://github.com/NasserKhalili/wos-format-converter) scripts:  
   - `TabDelimited.txt` for **VOSviewer**  
   - `PlainText.txt` for **Bibliometrix/Biblioshiny**  



## 📄 License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.

## ✍️ Author

**Nasser Khalili**  
GitHub: [@nasserkhalili](https://github.com/nasserkhalili)
