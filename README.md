# Bank Statement Collector App

A Streamlit-based web app for parsing and managing **bank statements** from CSV or Excel files.  
Features:
- Inline editing with **AgGrid** (no full reruns while typing).
- CSV + Excel upload (auto detects bank name).
- Export to **PDF (with totals row)**, **CSV**, and **Excel**.
- SQLite persistence for items and user progress.
- Manage Item choices (add/remove dynamically).
- Styled inline editing (Cr/Dr highlighting in Type column).

---

## 🚀 Run Locally

### 1. Clone the repo
```bash
git clone https://github.com/yourname/bank-statement-collector.git
cd bank-statement-collector
```
### 2. Create virtual environment
```bash
python -m venv .venv
source .venv/bin/activate   # Linux / macOS
.venv\Scripts\activate 
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```
### 4. Run the app
```bash 
streamlit run app.py
```

The app will start at:👉 http://localhost:8501


📦 Project Structure
```bash
.
├── app.py                  # Main Streamlit app (rename your latest version to app.py before docker build)
├── requirements.txt        # Python dependencies
├── Dockerfile              # Docker build file
├── statements_app.db       # SQLite DB (created automatically)
└── README.md               # Documentation
```