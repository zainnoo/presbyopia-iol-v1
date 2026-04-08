# Presbyopia IOL Selection Software

A clinical decision-support web app for intraocular lens (IOL) selection, converted from the **PRESBYOPIA-IOL-SOFTWARE-v2.xlsx** workbook.

Built with [Streamlit](https://streamlit.io/) — run locally or deploy for free via Streamlit Community Cloud.

---

## Overview

The app evaluates **9 IOLs across 3 categories** based on patient biometry and ocular parameters:

| Category | Lenses |
|---|---|
| **Monofocal** | Spherical Monofocal, Zero Aspheric Monofocal, Negative Aspheric Monofocal |
| **EDOF / Extended** | Eyhance, Rayner EMV, Vivity |
| **Diffractive / Trifocal** | PanOptix, Gemetric, Atlisa Tri |

For each lens the app computes:
- **Predicted Depth of Focus (DOF)** in dioptres
- **Predicted MTF max** (contrast sensitivity proxy)
- **Dysphotopsia Score** (0–10)
- **Refractive Stability** prediction (EDOF lenses)
- **Patient Satisfaction** estimate (%)

---

## Inputs

| Parameter | Description |
|---|---|
| Age | Patient age in years |
| Photopic Pupil | Bright-light pupil diameter (mm) |
| Scotopic Pupil | Dim-light pupil diameter (mm) |
| Corneal SA Z40 | Corneal spherical aberration at 6 mm (µm) |
| Vertical Coma | Corneal vertical coma (µm) |
| Horizontal Coma | Corneal horizontal coma (µm) |
| Angle Alpha | Angle alpha in degrees |
| Angle Kappa | Angle kappa in degrees |
| Ocular Pathologies | Dry Eye, Macular Pathology, Diabetic Retinopathy, Glaucoma, Corneal Endothelial Disease (grades 1–3) |

---

## Running Locally

### Prerequisites
- Python 3.9 or newer

### Steps

```bash
# 1. Clone the repository
git clone https://github.com/YOUR_USERNAME/presbyopia-iol-software.git
cd presbyopia-iol-software

# 2. (Optional) Create a virtual environment
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Launch the app
streamlit run app.py
```

The app will open automatically at `http://localhost:8501`.

---

## Deploying to Streamlit Community Cloud (Free)

1. Push this repository to GitHub (public or private).
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
3. Click **New app** → select your repo, branch `main`, and entry point `app.py`.
4. Click **Deploy** — your app will be live in minutes at a public URL.

---

## Repository Structure

```
presbyopia-iol-software/
├── app.py               # Main Streamlit application
├── requirements.txt     # Python dependencies
└── README.md            # This file
```

---

## Disclaimer

> ⚠️ This software is a **clinical decision-support tool only**.  
> All recommendations must be interpreted by a qualified ophthalmologist in the context of a full clinical assessment.  
> Scores are predictive models derived from parametric formulas and do not substitute clinical judgment.

---

## Source

Converted from **PRESBYOPIA-IOL-SOFTWARE-v2.xlsx** — all lookup tables, cross-sheet formula logic, and calculation chains have been faithfully translated to Python.
