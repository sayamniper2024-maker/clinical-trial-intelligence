# 🧬 Clinical Trial Intelligence Agent

An AI-powered competitive intelligence platform that analyzes
clinical trial landscapes in real time.

## 🚀 Live Demo
[Click here to open the app](https://your-app-url.streamlit.app)
*(update this link after deploying)*

## What It Does

Type any disease or therapy area — the agent fetches live data
from ClinicalTrials.gov, runs competitive intelligence analysis,
generates an AI consulting report, and produces downloadable
PowerPoint and Excel outputs.

## Features

- **Live data** — pulls up to 1,000 trials per query from ClinicalTrials.gov API
- **7 interactive charts** — phase distribution, sponsor landscape, geographic
  heatmap, timeline, dropout analysis, enrollment distribution
- **AI analysis** — Groq Llama 3.1 generates a 4-section consulting report
- **Sector filters** — Pharma, Biotech, MedTech, Biologics, Nutraceuticals
- **Dashboard filters** — filter by phase, status, sponsor type, year, enrollment
- **Export** — one-click PowerPoint deck, multi-sheet Excel, JSON summary

## Tech Stack

| Layer | Technology |
|-------|-----------|
| UI | Streamlit |
| Data | ClinicalTrials.gov REST API v2 |
| Analysis | Python, Pandas, SQLite |
| Charts | Plotly |
| AI | Groq API — Llama 3.1 8B |
| Export | python-pptx, openpyxl |
| Deploy | Streamlit Cloud |

## How to Run Locally

```bash
# Clone the repo
git clone https://github.com/YOUR_USERNAME/clinical-trial-intelligence.git
cd clinical-trial-intelligence

# Install dependencies
pip install -r requirements.txt

# Add your API key
# Create .streamlit/secrets.toml with:
# GROQ_API_KEY = "your_groq_key_here"

# Run
streamlit run app.py
```

## Project Structure
clinical-trial-intelligence/

├── app.py                  # Main Streamlit application

├── requirements.txt        # Python dependencies

├── .streamlit/

│   └── config.toml         # UI theme configuration

└── README.md

## Data Source

All trial data is fetched live from
[ClinicalTrials.gov](https://clinicaltrials.gov) —
the official US registry of clinical studies.

## Author

Built as a portfolio project demonstrating:
- Real-world REST API integration
- Data pipeline engineering
- AI agent development
- Interactive dashboard design
