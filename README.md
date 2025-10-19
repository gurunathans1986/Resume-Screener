# 📄 Resume Evaluation with Gemini

This project automates **resume evaluation** against multiple **job descriptions (JDs)** using **Google Gemini**.  
It scores resumes, generates structured reasoning, and produces both per-role result files and a master results file.  

---

## 🚀 Features
- Handles multiple **job descriptions** and **roles**.  
- Processes **daily batches of resumes** by date.  
- Uses **Gemini LLM** to evaluate resumes.  
- Produces structured results with:  
  - Score (0–100)  
  - Shortlisted flag (Yes/No, score ≥ 70)  
  - Tech Stack Match  
  - Years of Experience Match  
  - Domain Match  
  - Leadership  
  - Red Flags  
  - Reasoning (3–4 lines summary)  
  - Sent Codility Test (manual Yes/No field)  
- Exports results to **Excel (.xlsx)** with styling:  
  - Bold headers  
  - Green fill for `Yes`, Red fill for `No`  

---

## 📂 Folder Structure

```bash
project_root/
│
├── job_descriptions/             
│   ├── Petroleum_Engineer_(QA_Specialist).txt
│   ├── data_scientist.txt
│   └── ...
│
├── resumes/                      
│   ├── Petroleum_Engineer_(QA_Specialist)/
│   │   ├── 2025-09-03/
│   │   │   ├── resume1.pdf
│   │   │   └── resume2.docx
│   │   └── ...
│   └── data_scientist/
│       └── 2025-09-03/
│           └── resume3.pdf
│
├── results/                      
│   ├── Petroleum_Engineer_(QA_Specialist)/
│   │   └── 2025-09-03_results.xlsx
│   ├── data_scientist/
│   │   └── 2025-09-03_results.xlsx
│   └── master_results.xlsx   <-- all roles, all dates
│
├── main.py                       
└── README.md
```

---

## ⚙️ Setup

1. **Clone this repo** or copy the files locally.  

2. **Install dependencies**:  
   ```bash
   pip install google-generativeai pandas docx2txt PyPDF2 openpyxl
   ```

3. **Set your Gemini API key**:  
   ```bash
   export GOOGLE_API_KEY="your_api_key_here"   # Mac/Linux
   setx GOOGLE_API_KEY "your_api_key_here"     # Windows
   ```

4. **Prepare JDs**:  
   - Save each JD as a `.txt` file inside `job_descriptions/`.  
   - File name = role name (e.g., `Petroleum_Engineer_(QA_Specialist).txt`).  

5. **Organize resumes**:  
   - Place resumes under `resumes/<role>/<date>/`.  
   - Example:  
     ```
     resumes/Petroleum_Engineer_(QA_Specialist)/2025-09-03/resume1.pdf
     ```

---

## ▶️ Running the Script

Process all roles for a given date:

```bash
python main.py --date YYYY-MM-DD
```

Example:

```bash
python main.py --date 2025-09-03
```

---

## 📊 Output

1. **Per-role results** → saved in:  
   ```
   results/<role>/<YYYY-MM-DD>_results.xlsx
   ```

2. **Master results file** → appended indefinitely:  
   ```
   results/master_results.xlsx
   ```

Both outputs are Excel files with styled headers and conditional formatting for Yes/No fields.

---

## 📋 Result File Columns

| Date       | Role   | Resume Name | Score | Shortlisted | Tech Stack Match | Years of Exp Match | Domain Match | Leadership | Red Flags | Reasoning | Sent Codility Test |
|------------|--------|-------------|-------|-------------|------------------|--------------------|--------------|------------|-----------|-----------|--------------------|

---

## ✅ Example

```bash
python main.py --date 2025-09-03
```

Output → `results/master_results.xlsx`:

| Date       | Role                           | Resume Name   | Score | Shortlisted | Tech Stack Match           | Years of Exp Match | Domain Match | Leadership      | Red Flags       | Reasoning                                                                 | Sent Codility Test |
|------------|--------------------------------|---------------|-------|-------------|----------------------------|--------------------|--------------|-----------------|-----------------|---------------------------------------------------------------------------|--------------------|
| 2025-09-03 | Petroleum_Engineer_(QA_Specialist) | resume1.pdf   | 85    | Yes         | Strong QA, petroleum exp   | 5 yrs vs 4 req     | Oil & Gas    | Mentored juniors| None            | Candidate has strong QA and domain knowledge. Good technical match, fits role. | Yes                |

---

## ⚡ Notes
- Gemini is instructed to return **strict JSON**.  
- If parsing ever fails, the script prints the raw response and marks the row as “Parsing failed”.  
- Currently supports **PDF, DOCX, and TXT resumes**.  
- `Sent Codility Test` column is left blank initially; fill it manually after reviewing candidates.  
