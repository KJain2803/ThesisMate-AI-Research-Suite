# 🧠 ThesisMate AI Research Suite

## End-to-End Dissertation Intelligence • Automated Biostatistics • AI Scientific Writing Engine

ThesisMate AI Research Suite is an integrated AI-powered research intelligence platform designed specifically for postgraduate dissertation scholars, biotechnology researchers and life science students who struggle with fragmented data analysis, scientific interpretation and thesis writing workflows.

This platform transforms raw experimental spreadsheets and dissertation drafts into:

- 📊 Automated Statistical Analytics
- 🧪 Instant ANOVA / T-Test / Correlation / Regression Outputs
- ✍ AI-Powered Scientific Result Interpretation
- 📄 Downloadable Executive Scientific PDF Dossiers
- 🎓 Dissertation NLP Review & Quality Assessment
- 📈 Smart Graphical Scientific Evidence Generation

---

## 🚀 Live Deployed Application

**Public Demo URL:**  
[https://thesismate-ai-research-suite-w4iqwcng6r4au8uvqyamg9.streamlit.app](https://thesismate-ai-research-suite-w4iqwcng6r4au8uvqyamg9.streamlit.app)

---

## 🛠 Setup

### Prerequisites

- Python 3.14 recommended. The local app was verified with Python 3.14.4.
- `pip`
- A terminal or shell environment

### Create a Virtual Environment

```bash
python -m venv .venv
source .venv/bin/activate
```

On Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### Install Dependencies

```bash
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

### Install Chrome for Plotly PDF Image Export

Plotly static chart export in the generated PDF reports uses Kaleido. Install its tested Chrome runtime after installing dependencies:

```bash
choreo_get_chrome
```

If the virtual environment is not activated, run:

```bash
.venv/bin/choreo_get_chrome
```

On Windows:

```powershell
.\.venv\Scripts\choreo_get_chrome.exe
```

---

## ▶ Run Locally

Start the Streamlit app:

```bash
streamlit run app.py
```

Or run it through the virtual environment explicitly:

```bash
.venv/bin/python -m streamlit run app.py
```

Open the local URL shown by Streamlit, usually:

```text
http://localhost:8501
```

You can use `sample_data.xlsx` to test the statistical analytics workflow.

---

## 🔥 Core Problem Solved

Research scholars often spend:

- weeks performing manual biostatistics,
- days writing result interpretations,
- and significant money on fragmented statistical/software tools.

ThesisMate compresses this complete workflow into a single AI-assisted scientific dashboard where a user uploads thesis drafts and experimental Excel files to receive instant publishable scientific insights.

---

## ⚙ Key Modules

### 1. Dissertation NLP Review Engine

Analyzes uploaded thesis/dissertation files and generates:

- writing clarity score
- methodology depth score
- citation strength estimate
- AI academic review comments

### 2. Smart Statistical Analytics Engine

Processes uploaded Excel datasets and computes:

- descriptive statistics
- coefficient of variation
- Pearson correlation
- linear regression
- ANOVA / T-test outputs
- treatment ranking intelligence

### 3. AI Scientific Interpretation Writer

Automatically generates dissertation-grade Results & Discussion style paragraphs from raw statistical outputs.

### 4. Export Scientific Reports

Generates downloadable:

- AI Executive Research Dossier PDF
- Scientific presentation PPT outputs

---

## 🛠 Technology Stack

- Python
- Streamlit
- Pandas
- NumPy
- SciPy
- Plotly
- Matplotlib
- ReportLab
- python-docx
- python-pptx

---

## 🎯 Innovation Highlights

- Combines dissertation review + biostatistics + AI writing in one pipeline
- Designed specifically for life science and biotechnology thesis workflows
- Generates executive-grade scientific PDF dossiers
- Public cloud deployed as accessible SaaS prototype
- Reduces analytical turnaround from days to minutes

---

## 📌 Future Scope

- journal manuscript auto-generation
- plagiarism-linked literature suggestions
- advanced multivariate statistics
- ML-based predictive tissue culture modelling
- institutional research dashboard deployment
