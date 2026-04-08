from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

import streamlit as st

from engine import INPUT_CELLS, calculate_outputs, get_template_metadata

APP_DIR = Path(__file__).resolve().parent
PATIENT_DB_PATH = APP_DIR / "patient_records.json"

st.set_page_config(page_title="PRESBYOPIA IOL SOFTWARE", layout="wide")


def load_patient_db() -> dict[str, Any]:
    if not PATIENT_DB_PATH.exists():
        return {}
    try:
        return json.loads(PATIENT_DB_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_patient_db(db: dict[str, Any]) -> None:
    PATIENT_DB_PATH.write_text(json.dumps(db, indent=2), encoding="utf-8")


def get_initial_inputs() -> dict[str, Any]:
    metadata = get_template_metadata()
    defaults = metadata["defaults"].copy()
    loaded_id = st.session_state.get("loaded_patient_id")
    if loaded_id:
        db = load_patient_db()
        patient = db.get(loaded_id)
        if patient:
            defaults.update(patient["inputs"])
    return defaults


st.title("PRESBYOPIA IOL SOFTWARE")
st.caption("Excel-driven Streamlit version using the same workbook formulas, parameters, outcomes and wordings.")

patient_db = load_patient_db()
patient_ids = sorted(patient_db.keys())

with st.sidebar:
    st.header("PATIENT RECORDS")
    selected_patient = st.selectbox("Load saved patient", options=[""] + patient_ids, index=0)
    if st.button("Load patient", use_container_width=True):
        if selected_patient:
            st.session_state["loaded_patient_id"] = selected_patient
            st.rerun()

initial_inputs = get_initial_inputs()

with st.form("input_form"):
    st.subheader("INPUT DATA")
    patient_id = st.text_input("PATIENT ID", value=st.session_state.get("loaded_patient_id", ""))

    c1, c2, c3 = st.columns(3)
    with c1:
        age = st.number_input("AGE", min_value=0, max_value=120, value=int(initial_inputs["AGE"]), step=1)
        photopic = st.number_input("PHOTOPIC PUPIL", value=float(initial_inputs["PHOTOPIC PUPIL"]), step=0.01, format="%.2f")
        scotopic = st.number_input("SCOTOPIC PUPIL", value=float(initial_inputs["SCOTOPIC PUPIL"]), step=0.01, format="%.2f")
    with c2:
        angle_alpha = st.number_input("ANGLE ALPHA", value=float(initial_inputs["ANGLE ALPHA"]), step=0.01, format="%.2f")
        angle_kappa = st.number_input("ANGLE KAPPA", value=float(initial_inputs["ANGLE KAPPA"]), step=0.01, format="%.2f")
        spherical_aberration = st.number_input("CORNEAL SPHERICAL ABERRATION 6mm (Z40)", value=float(initial_inputs["CORNEAL SPHERICAL ABERRATION 6mm (Z40)"]), step=0.01, format="%.2f")
    with c3:
        vertical_coma = st.number_input("CORNEAL VERTICAL COMA 6mm", value=float(initial_inputs["CORNEAL VERTICAL COMA 6mm"]), step=0.01, format="%.2f")
        horizontal_coma = st.number_input("CORNEAL HORIZONTAL COMA 6mm", value=float(initial_inputs["CORNEAL HORIZONTAL COMA 6mm"]), step=0.01, format="%.2f")

    st.markdown("### OCULAR PATHOLOGIES")
    p1, p2, p3, p4, p5 = st.columns(5)
    with p1:
        corneal_endothelial = st.number_input("CORNEAL ENDOTHELIAL DISEASE", min_value=0, max_value=3, value=int(initial_inputs["CORNEAL ENDOTHELIAL DISEASE"]), step=1)
    with p2:
        dry_eye = st.number_input("DRY EYE DISEASE", min_value=0, max_value=3, value=int(initial_inputs["DRY EYE DISEASE"]), step=1)
    with p3:
        diabetic_retinopathy = st.number_input("DIABETIC RETINOPATHY", min_value=0, max_value=3, value=int(initial_inputs["DIABETIC RETINOPATHY"]), step=1)
    with p4:
        macular_pathology = st.number_input("MACULAR PATHOLOGY", min_value=0, max_value=3, value=int(initial_inputs["MACULAR PATHOLOGY"]), step=1)
    with p5:
        glaucoma = st.number_input("GLAUCOMA", min_value=0, max_value=3, value=int(initial_inputs["GLAUCOMA"]), step=1)

    submitted = st.form_submit_button("Calculate", use_container_width=True)

if submitted:
    current_inputs = {
        "AGE": int(age),
        "PHOTOPIC PUPIL": float(photopic),
        "SCOTOPIC PUPIL": float(scotopic),
        "ANGLE ALPHA": float(angle_alpha),
        "ANGLE KAPPA": float(angle_kappa),
        "CORNEAL SPHERICAL ABERRATION 6mm (Z40)": float(spherical_aberration),
        "CORNEAL VERTICAL COMA 6mm": float(vertical_coma),
        "CORNEAL HORIZONTAL COMA 6mm": float(horizontal_coma),
        "CORNEAL ENDOTHELIAL DISEASE": int(corneal_endothelial),
        "DRY EYE DISEASE": int(dry_eye),
        "DIABETIC RETINOPATHY": int(diabetic_retinopathy),
        "MACULAR PATHOLOGY": int(macular_pathology),
        "GLAUCOMA": int(glaucoma),
    }
    with st.spinner("Calculating from Excel logic..."):
        results = calculate_outputs(current_inputs)
    st.session_state["last_inputs"] = current_inputs
    st.session_state["last_results"] = results
    if patient_id:
        st.session_state["current_patient_id"] = patient_id

if "last_results" in st.session_state:
    st.divider()
    st.subheader("RESULTS")
    for section in st.session_state["last_results"]["tables"]:
        st.markdown(f"### {section['title']}")
        st.dataframe(section["data"], use_container_width=True, hide_index=True)

    csave, cdelete = st.columns(2)
    with csave:
        if st.button("Save current patient", use_container_width=True):
            current_patient_id = st.session_state.get("current_patient_id") or patient_id
            if not current_patient_id:
                st.warning("Enter a PATIENT ID before saving.")
            else:
                patient_db[current_patient_id] = {
                    "saved_at": datetime.now().isoformat(timespec="seconds"),
                    "inputs": st.session_state["last_inputs"],
                    "results": {section["title"]: section["data"].to_dict(orient="records") for section in st.session_state["last_results"]["tables"]},
                }
                save_patient_db(patient_db)
                st.session_state["loaded_patient_id"] = current_patient_id
                st.success(f"Saved patient: {current_patient_id}")
    with cdelete:
        if st.button("Delete loaded patient", use_container_width=True):
            loaded_id = st.session_state.get("loaded_patient_id")
            if loaded_id and loaded_id in patient_db:
                patient_db.pop(loaded_id, None)
                save_patient_db(patient_db)
                st.session_state.pop("loaded_patient_id", None)
                st.success(f"Deleted patient: {loaded_id}")
                st.rerun()

st.divider()
st.markdown("**Note:** The software is intentionally restricted to the same workbook logic. The only extra addition is PATIENT ID with save/load support for multiple patients.")
