# TSO PDF → Excel · Streamlit Web App

Converts M&M TSO PDF documents into fully-populated Excel workbooks
matching the official TSO Download format. 100% local or cloud — no AI, no API key.

---

## Deploy to the Web (Free — Streamlit Community Cloud)

### Step 1 — Push to GitHub

1. Create a free account at github.com
2. Click New repository → name it `tso-pdf-excel` → set to Public → Create
3. Upload these files into the repo (drag & drop in GitHub web UI):
   - app.py
   - tso_converter.py
   - requirements.txt
   - .streamlit/config.toml
   - README.md

### Step 2 — Deploy on Streamlit Cloud

1. Go to share.streamlit.io → sign in with GitHub
2. Click New app
3. Fill in:
   - Repository: your-github-username/tso-pdf-excel
   - Branch: main
   - Main file path: app.py
4. Click Deploy!

Your app will be live at:
https://your-username-tso-pdf-excel-app-xxxxx.streamlit.app

---

## Run Locally

pip install -r requirements.txt
streamlit run app.py
Open http://localhost:8501

---

## File Structure

tso-pdf-excel/
├── app.py
├── tso_converter.py
├── requirements.txt
├── .streamlit/
│   └── config.toml
└── README.md
