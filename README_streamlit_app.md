# PRESBYOPIA IOL SOFTWARE - Streamlit App

Files included:
- `app.py` - Streamlit application
- `PRESBYOPIA IOL SOFTWARE v2.xlsx` - original Excel workbook used as the calculation engine
- `requirements.txt` - Python dependencies

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Current scope

- Uses the same Excel workbook formulas, parameters, outcomes, and wording
- No logic changes added
- Only extra feature: `PATIENT ID` with save/load for multiple patients

## Important

This app recalculates by writing the entered inputs into the workbook and evaluating the workbook formulas through Python, so the Excel logic remains the source of truth.
