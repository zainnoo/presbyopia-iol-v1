"""
Presbyopia IOL Selection Software
Converted from PRESBYOPIA-IOL-SOFTWARE-v2.xlsx
Developed by Dr Zain Khatib | zainkhatib89@gmail.com
"""

import streamlit as st
import math
import bisect
import io
import datetime
import json
import os
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Presbyopia IOL Software",
    page_icon="👁️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# LOOKUP TABLES
# ─────────────────────────────────────────────────────────────────────────────

SA_COLS = [-0.5, -0.4, -0.3, -0.2, -0.1, 0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 1.0]

# ── SPHERICAL MONOFOCAL ──────────────────────────────────────────────────────
SPH_DOF = [
    [1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5],
    [1.25,1.25,1.25,1.25,1.25,1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5],
    [1.25,1.25,1.25,1.25,1.25,1.25,1.25,1.25,1.25,1.25,1.25,1.25,1.25],
    [1.25,1.25,1.25,1.0, 1.0, 1.25,1.25,1.0, 0.75,0.75,0.5, 0.25,0.25],
    [0.25,0.25,0.75,0.75,0.75,0.75,0.75,0.5, 0.25,0.25,0.25,0.0, 0.0 ],
]
SPH_MTF = [
    [8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.0, 8.0, 8.0, 8.0, 7.5, 7.5, 7.0],
    [8.0, 8.5, 8.5, 8.5, 9.0, 8.0, 6.5, 6.0, 6.0, 6.0, 6.0, 6.0, 6.0],
    [6.0, 6.5, 7.0, 9.0, 8.5, 7.0, 6.5, 6.0, 5.5, 5.0, 4.0, 4.0, 4.0],
    [4.5, 5.5, 6.0, 9.0, 7.5, 6.5, 5.5, 5.0, 4.0, 4.0, 4.0, 3.0, 3.0],
]

# ── ZERO ASPHERIC MONOFOCAL ──────────────────────────────────────────────────
ZA_DOF = [
    [1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5],
    [1.5, 1.25,1.25,1.25,1.25,1.25,1.25,1.5, 1.5, 1.5, 1.5, 1.5, 1.5],
    [1.0, 1.0, 1.0, 1.0, 1.0, 1.25,1.25,1.25,1.25,1.25,1.25,1.25,1.25],
    [1.0, 1.0, 1.0, 1.0, 0.75,0.75,1.25,1.0, 0.75,0.75,0.75,0.5, 0.5 ],
    [0.0, 0.25,0.5, 0.75,0.75,0.75,0.75,0.75,0.5, 0.5, 0.25,0.0, 0.0 ],
]
ZA_MTF = [
    [8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5],
    [6.5, 7.0, 8.0, 8.5, 8.5, 9.0, 8.0, 7.5, 7.5, 7.0, 6.5, 6.0, 5.5],
    [4.5, 5.0, 6.5, 7.5, 8.0, 9.0, 7.5, 7.0, 6.0, 5.5, 5.0, 4.5, 4.5],
    [3.5, 4.5, 6.5, 7.0, 8.0, 8.0, 7.0, 6.0, 6.0, 5.5, 4.5, 4.5, 3.0],
]

# ── NEGATIVE ASPHERIC MONOFOCAL ──────────────────────────────────────────────
NA_DOF = [
    [1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5],
    [1.5, 1.5, 1.5, 1.25,1.25,1.25,1.25,1.25,1.25,1.25,1.5, 1.5, 1.5],
    [1.0, 1.0, 1.0, 0.75,0.75,1.0, 1.0, 1.25,1.25,1.25,1.25,1.25,1.25],
    [0.25,0.25,0.75,1.0, 1.0, 1.0, 1.0, 0.75,1.0, 1.25,1.0, 0.75,0.5 ],
    [0.0, 0.0, 0.25,0.75,0.75,0.75,0.75,0.75,0.75,0.75,0.5, 0.0, 0.0 ],
]
NA_MTF = [
    [8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5, 8.5],
    [5.5, 6.5, 7.5, 8.0, 8.5, 8.5, 9.0, 9.0, 8.5, 7.5, 7.0, 6.5, 6.5],
    [3.5, 4.0, 4.5, 5.5, 6.5, 8.5, 9.0, 9.0, 7.5, 7.5, 6.5, 5.5, 4.5],
    [2.5, 3.0, 3.5, 4.5, 5.5, 7.0, 9.0, 8.5, 7.5, 6.0, 4.5, 4.5, 3.5],
]

# ── EYHANCE ──────────────────────────────────────────────────────────────────
EYHANCE_PUPIL_ROWS = [2, 3, 4, 5]
_eyh_r2_dof = [1.9,1.9,1.9,1.9,1.9,2.0,2.0,2.0,2.0,2.0,1.9,1.9,1.9]
_eyh_r3_dof = [1.9,1.9,1.9,1.8,1.8,1.8,1.7,1.7,1.7,1.7,1.7,1.7,1.7]
_eyh_r5_dof = [0.75,0.75,0.75,0.75,0.75,0.75,0.75,0.85,0.85,1.0,0.75,0.75,0.75]
_eyh_r4_dof = []
for _i in range(13):
    if _i == 9:    _eyh_r4_dof.append(1.4)
    elif _i == 10: _eyh_r4_dof.append(1.3)
    else:          _eyh_r4_dof.append(round((_eyh_r3_dof[_i] + _eyh_r5_dof[_i]) / 2, 1))
EYHANCE_DOF = [_eyh_r2_dof, _eyh_r3_dof, _eyh_r4_dof, _eyh_r5_dof]

EYHANCE_SCOT_ROWS = [2, 3, 4, 5]
EYHANCE_MTF = [
    [6.0,6.0,6.3,6.4,6.5,6.6,6.7,6.8,6.9,6.8,6.8,6.8,6.8],
    [4.5,4.5,4.5,4.8,5.0,5.2,5.8,6.2,6.6,6.6,6.2,6.0,5.6],
    [3.8,3.8,3.8,4.0,4.5,5.5,6.0,6.5,6.7,5.8,4.8,3.8,3.8],
    [2.0,2.4,2.8,3.4,3.5,3.8,4.3,5.6,5.0,4.2,3.7,2.8,2.8],
]

# ── RAYNER EMV ───────────────────────────────────────────────────────────────
RAYNER_PUPIL_ROWS = [2, 3, 4, 5]
RAYNER_DOF = [
    [1.8, 1.8, 1.8, 1.8, 1.8, 1.8, 1.8, 1.8, 1.8, 1.8, 1.7, 1.7, 1.7],
    [1.25,1.1, 0.75,0.75,1.0, 1.2, 1.4, 1.7, 1.9, 2.0, 1.9, 1.9, 1.9],
    [0.5, 0.75,1.0, 1.1, 1.2, 1.4, 1.4, 1.4, 1.5, 1.3, 1.2, 1.0, 0.75],
    [0.5, 0.5, 0.5, 0.5, 0.5, 0.75,0.75,0.75,1.0, 1.0, 0.75,0.75,0.75],
]
RAYNER_SCOT_ROWS = [2, 3, 4, 5]
RAYNER_MTF = [
    [6.5,6.5,6.5,7.0,7.0,7.5,7.5,7.5,7.5,7.5,7.0,7.0,7.0],
    [5.5,5.5,6.0,6.5,7.0,7.5,7.0,6.3,5.7,5.0,4.5,4.0,4.0],
    [3.8,3.8,3.8,4.0,4.0,4.5,4.2,4.0,3.5,3.2,3.0,2.8,2.8],
    [2.0,2.4,2.8,3.0,3.2,4.0,3.5,3.2,3.0,2.5,2.0,1.8,1.8],
]

# ── VIVITY ───────────────────────────────────────────────────────────────────
VIVITY_PUPIL_ROWS = [2, 3, 4, 5]
VIVITY_DOF = [
    [2.25,2.25,2.25,2.25,2.25,2.5, 2.2, 2.5, 2.5, 2.5, 2.25,2.25,2.25],
    [2.0, 2.0, 1.9, 1.8, 1.7, 1.8, 1.8, 1.9, 1.9, 1.8, 1.7, 1.6, 1.5],
    [0.5, 0.75,1.0, 1.1, 1.2, 1.0, 1.2, 1.3, 1.4, 1.5, 1.2, 1.0, 0.75],
    [0.5, 0.5, 0.5, 0.5, 0.5, 0.75,0.75,0.75,1.0, 1.0, 0.75,0.75,0.75],
]
VIVITY_SCOT_ROWS = [2, 3, 4, 5]
VIVITY_MTF = [
    [4.8,4.8,4.8,4.8,4.8,5.0,5.0,5.0,5.0,5.0,4.8,4.8,4.8],
    [3.0,3.5,3.8,4.0,4.2,4.2,4.0,3.9,3.7,3.5,3.5,3.5,3.5],
    [4.0,4.0,4.5,4.5,4.5,5.0,5.0,5.0,4.5,3.5,3.5,3.5,3.5],
    [2.0,2.5,3.0,3.5,4.0,4.5,4.5,3.2,3.0,3.0,2.5,2.5,2.0],
]

# ── PANOPTIX ─────────────────────────────────────────────────────────────────
PAN_DOF_R3  = [2.9,2.9,3.0,3.0,3.0,3.0,3.0,3.0,3.0,3.0,3.0,3.0,3.0]
PAN_DOF_R45 = [1.0,1.0,1.0,1.0,1.1,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0]
PAN_MTF_R3  = [4.8,5.0,5.1,5.3,5.5,5.5,5.7,5.7,5.8,5.8,5.8,5.8,5.8]
PAN_MTF_R45 = [3.0,3.1,3.3,3.6,3.8,4.2,4.6,5.0,4.5,3.5,3.5,3.0,2.5]

# ── GEMETRIC ─────────────────────────────────────────────────────────────────
GEM_DOF_R3  = [2.5,2.4,2.3,2.3,2.3,2.3,2.3,2.3,2.3,2.3,2.3,2.3,2.4]
GEM_DOF_R45 = [1.0]*13
GEM_MTF_R3  = [4.7,4.8,4.9,5.1,5.2,5.2,5.3,5.4,5.4,5.4,5.4,5.4,5.4]
GEM_MTF_R45 = [3.5,3.7,3.9,4.2,4.5,4.9,5.1,5.1,5.0,4.5,4.0,3.5,3.0]

# ── ATLISA TRI ───────────────────────────────────────────────────────────────
ATL_DOF_R3  = [2.6,2.6,2.6,2.6,2.6,2.6,2.5,2.5,2.5,2.5,2.5,2.5,2.6]
ATL_DOF_R45 = [1.2,1.2,1.2,1.1,1.1,1.1,1.1,1.1,1.0,1.0,1.0,1.0,1.0]
ATL_MTF_R3  = [4.7,4.9,5.1,5.3,5.5,5.6,5.7,5.8,5.8,5.8,5.8,5.8,5.8]
ATL_MTF_R45 = [2.0,2.5,2.5,3.5,3.5,4.3,4.7,5.1,5.4,3.5,3.5,3.0,2.5]

# ── REFRACTIVE STABILITY TABLES ───────────────────────────────────────────────
EYHANCE_REF = {
    2: [-0.5,  -0.75, -0.75],
    3: [ 0.0,  -0.25, -0.5 ],
    4: [ 0.5,   0.0,  -0.5 ],
    5: [ 0.5,   0.0,  -0.5 ],
}
RAYNER_REF = {
    2: [0.5,  0.5,  0.5],
    3: [0.25, 0.0,  0.0],
    4: [0.0,  0.0,  0.0],
    5: [0.0,  0.0,  0.0],
}
VIVITY_REF = {
    2: [0.0,   0.0,   0.0 ],
    3: [0.0,  -0.25, -0.5 ],
    4: [0.0,  -0.37, -0.75],
    5: [0.0,  -0.37, -0.75],
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def sa_col_exact(sa):
    sa_r = round(min(1.0, max(-0.5, sa)), 1)
    sa_r = round(sa_r, 10)
    return min(range(len(SA_COLS)), key=lambda i: abs(SA_COLS[i] - sa_r))

def sa_col_le(sa):
    sa_c = min(1.0, max(-0.5, sa))
    idx = bisect.bisect_right(SA_COLS, sa_c) - 1
    return max(0, min(idx, len(SA_COLS) - 1))

def row_le(val, sorted_rows):
    idx = bisect.bisect_right(sorted_rows, val) - 1
    return max(0, min(idx, len(sorted_rows) - 1))

def mono_dof_row(photopic):
    if photopic <= 2.0:   return 0
    elif photopic <= 3.0: return 1
    elif photopic <= 3.5: return 2
    elif photopic <= 4.5: return 3
    else:                 return 4

def mono_mtf_row(scotopic):
    if scotopic < 3.0:    return 0
    elif scotopic <= 4.5: return 1
    elif scotopic <= 5.5: return 2
    else:                 return 3

def pathology_products(pathologies):
    MTF_S = {
        'DRY EYE DISEASE':           [1, 0.95, 0.85, 0.70],
        'MACULAR PATHOLOGY':         [1, 0.90, 0.75, 0.55],
        'DIABETIC RETINOPATHY':      [1, 0.92, 0.80, 0.60],
        'GLAUCOMA':                  [1, 0.93, 0.82, 0.65],
        'CORNEAL ENDOTHELIAL DISEASE':[1, 0.90, 0.72, 0.50],
    }
    DOF_S = {
        'DRY EYE DISEASE':           [1, 0.97, 0.90, 0.75],
        'MACULAR PATHOLOGY':         [1, 0.92, 0.80, 0.60],
        'DIABETIC RETINOPATHY':      [1, 0.94, 0.85, 0.68],
        'GLAUCOMA':                  [1, 0.95, 0.86, 0.70],
        'CORNEAL ENDOTHELIAL DISEASE':[1, 0.94, 0.82, 0.60],
    }
    mp, dp = 1.0, 1.0
    for name, grade in pathologies.items():
        g = min(3, max(0, int(grade)))
        if name in MTF_S:
            mp *= MTF_S[name][g]
            dp *= DOF_S[name][g]
    return mp, dp

def apply_pathology(raw_mtf, raw_dof, pathologies, mtf_scale, dof_scale):
    mp, dp = pathology_products(pathologies)
    w47 = round(raw_mtf * (1 - (1 - mp) * mtf_scale), 1)
    w48 = round(raw_dof * (1 - (1 - dp) * dof_scale), 2)
    return w47, w48

def dof_grade(w48):
    if w48 < 0.75:  return "Distance only / No functional near"
    if w48 < 1.25:  return "Weak intermediate, poor near"
    if w48 < 1.75:  return "Functional intermediate, limited near"
    if w48 < 2.25:  return "Good intermediate and near"
    return "Excellent near performance"

def vq_grade(w47):
    if w47 >= 6: return "Excellent visual quality across all lighting conditions"
    if w47 >= 5: return "Excellent in normal lighting; mild reduction in very low-light"
    if w47 >= 4: return "Very good in bright lighting; noticeable decline in low-light"
    if w47 >= 3: return "Functional in good lighting; significant deterioration in dim light"
    if w47 >= 2: return "Suboptimal even in bright conditions; poor in low light"
    return "Poor overall visual quality across all lighting conditions"

def coma_magnitude(vc, hc):
    return math.sqrt(vc**2 + (0.7 * hc)**2)

def dysphotopsia_score(iol_key, corneal_sa, scotopic_pupil, vc, hc, alpha, kappa, pathologies):
    IOL_MULT = {
        'NEG ASP': 0.75, 'ZERO ASP': 0.78, 'SPH': 0.80,
        'EMV': 0.85, 'EYHANCE': 0.90, 'VIVITY': 1.00,
        'GEME': 1.10, 'ATLISA': 1.15, 'PANOP': 1.20,
    }
    PATHO_W = {
        'DRY EYE DISEASE':            (0.6, [0, 0.5, 1.5, 3.0]),
        'MACULAR PATHOLOGY':          (0.5, [0, 0.3, 0.8, 1.5]),
        'DIABETIC RETINOPATHY':       (0.5, [0, 0.2, 0.5, 1.0]),
        'GLAUCOMA':                   (0.5, [0, 0.2, 0.5, 1.0]),
        'CORNEAL ENDOTHELIAL DISEASE':(0.6, [0, 0.8, 2.0, 4.0]),
    }
    total_hoa = round(math.sqrt((1.5 * corneal_sa)**2 + vc**2 + hc**2), 3)

    if iol_key == 'NEG ASP':   diff = abs(corneal_sa - 0.15)
    elif iol_key == 'ZERO ASP': diff = abs(corneal_sa)
    elif iol_key == 'SPH':      diff = abs(corneal_sa + 0.15)
    else:
        fixed = {'EMV': 2.0, 'EYHANCE': 2.2, 'VIVITY': 3.4,
                 'GEME': 4.6, 'ATLISA': 4.9, 'PANOP': 5.1}
        diff = None
        basic = fixed.get(iol_key, 5.0)

    if diff is not None:
        basic = 0.0 if diff <= 0.1 else (0.4 if diff <= 0.25 else (0.8 if diff <= 0.5 else 1.5))

    pupil_mod = round(max(0, ((scotopic_pupil - 3)**2) * 0.08 * basic), 2) + basic
    e_score = round(min(10, pupil_mod +
                  ((0.6 * total_hoa) + (1.2 * max(0, alpha - 0.2)) + (0.6 * max(0, kappa - 0.2)))
                  * (0.5 + 0.05 * pupil_mod)), 1)

    dry_g  = min(3, pathologies.get('DRY EYE DISEASE', 0))
    corn_g = min(3, pathologies.get('CORNEAL ENDOTHELIAL DISEASE', 0))
    patho_total = 0.0
    for name, (wt, vals) in PATHO_W.items():
        g = min(3, pathologies.get(name, 0))
        patho_total += wt * vals[g]
    if dry_g >= 2 and corn_g >= 2:
        patho_total += 0.5

    mult = IOL_MULT.get(iol_key, 1.0)
    return round(min(10, e_score + patho_total * mult), 1)

def refractive_stability_text(iol_name, photopic, scotopic, corneal_sa):
    TABLES = {'EYHANCE': EYHANCE_REF, 'RAYNER EMV': RAYNER_REF, 'VIVITY': VIVITY_REF}
    if iol_name not in TABLES:
        return "No clinically significant pupil-dependent refractive shift expected"
    tbl = TABLES[iol_name]

    def sa_cat(sa):
        return '0 AND LESS' if sa <= 0 else ('0.1 TO 0.3' if sa < 0.4 else '0.4 AND ABOVE')
    def pupil_bin(p):
        return 2 if p < 2.5 else (3 if p < 3.5 else (4 if p < 4.5 else 5))

    sa_idx = {'0 AND LESS': 0, '0.1 TO 0.3': 1, '0.4 AND ABOVE': 2}[sa_cat(corneal_sa)]
    j = tbl[pupil_bin(photopic)][sa_idx]
    k = tbl[pupil_bin(scotopic)][sa_idx]

    if abs(j) < 0.5 and abs(k) < 0.5:
        return "No clinically significant pupil-dependent refractive shift expected"
    if j <= -0.5 and abs(k) < 0.5:
        return "Beware of myopic shift for distance in bright light / small pupil conditions"
    if j >= 0.5 and abs(k) < 0.5:
        return "Possible hyperopic shift for distance in bright light / small pupil conditions"
    if abs(j) < 0.5 and k <= -0.5:
        return "Chances of low-light myopia / night myopia likely"
    if abs(j) < 0.5 and k >= 0.5:
        return "Possible hyperopic shift in dim light / larger pupil conditions"
    if j <= -0.5 and k <= -0.5:
        return "Consistent myopic shift tendency across lighting conditions"
    if j >= 0.5 and k >= 0.5:
        return "Consistent hyperopic shift tendency across lighting conditions"
    if j <= -0.5 and k >= 0.5:
        return "Lighting-dependent refractive behavior: myopic in bright light, hyperopic in dim light"
    if j >= 0.5 and k <= -0.5:
        return "Lighting-dependent refractive behavior: hyperopic in bright light, night myopia in dim light"
    return "No clinically significant pupil-dependent refractive shift expected"

def ref_stability_factor(text):
    if "No clinically significant" in text:   return 1.00
    if "Beware of myopic shift" in text:       return 0.92
    if "hyperopic shift for distance" in text: return 0.92
    if "low-light myopia" in text:             return 0.88
    if "hyperopic shift in dim" in text:       return 0.89
    if "Consistent myopic" in text:            return 0.84
    if "Consistent hyperopic" in text:         return 0.84
    if "Lighting-dependent" in text:           return 0.80
    return 1.0

# ─────────────────────────────────────────────────────────────────────────────
# IOL CALCULATION FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def calc_monofocal(dof_tbl, mtf_tbl, ang_max, ang_mult, patho_scale,
                   ph_p, sc_p, sa, vc, hc, alpha, pathologies):
    sa_c  = sa_col_exact(sa)
    dof_r = mono_dof_row(ph_p)
    mtf_r = mono_mtf_row(sc_p)
    s13   = dof_tbl[dof_r][sa_c]
    s14   = mtf_tbl[mtf_r][sa_c]
    cm    = coma_magnitude(vc, hc)
    v12   = round(s13 + min(0.45, 1.0 * cm * (ph_p / 6)**3), 2)
    v21   = round(max(0, s14 - min(2.5, 5.0 * cm * (sc_p / 6)**3)), 1)
    u24   = round(min(ang_max, max(0, (alpha - 0.3) * ang_mult)), 1)
    v24   = round(max(0, v21 - u24), 1)
    w47, w48 = apply_pathology(v24, v12, pathologies, patho_scale, patho_scale)
    return dict(dof_lookup=s13, mtf_lookup=s14,
                dof_raw=v12, mtf_raw=v24,
                mtf_adj=w47, dof_adj=w48,
                dof_grade=dof_grade(w48), vq_grade=vq_grade(w47))

def calc_edof(dof_tbl, dof_rows, mtf_tbl, mtf_rows, patho_scale,
              ph_p, sc_p, sa, vc, hc, alpha, pathologies):
    sa_c  = sa_col_le(sa)
    dof_r = row_le(max(2, min(5, ph_p)), dof_rows)
    mtf_r = row_le(max(2, min(5, sc_p)), mtf_rows)
    s13   = dof_tbl[dof_r][sa_c]
    s14   = mtf_tbl[mtf_r][sa_c]
    cm    = coma_magnitude(vc, hc)
    if alpha <= 0.25:   w12 = 1.0
    elif alpha <= 0.35: w12 = 1.0 - ((alpha - 0.25) / 0.1) * 0.08
    elif alpha <= 0.4:  w12 = 0.92 - ((alpha - 0.35) / 0.05) * 0.12
    else:               w12 = 0.8
    v12  = round((s13 + min(0.45, cm * (ph_p / 6)**3)) * w12, 2)
    v21  = round(max(0, s14 - min(2.5, 5.0 * cm * (sc_p / 6)**3)), 1)
    u24  = round(min(1.2, max(0, (alpha - 0.3) * 3)), 1)
    v24  = round(max(0, v21 - u24), 1)
    w47, w48 = apply_pathology(v24, v12, pathologies, patho_scale, patho_scale)
    return dict(dof_lookup=s13, mtf_lookup=s14,
                dof_raw=v12, mtf_raw=v24,
                mtf_adj=w47, dof_adj=w48,
                dof_grade=dof_grade(w48), vq_grade=vq_grade(w47))

def calc_diffractive(dof_r3, dof_r45, mtf_r3, mtf_r45,
                     dof_pf_min, dof_pf_c_low, dof_pf_c_mid, dof_pf_mid, dof_pf_falloff,
                     mtf_pf_min, mtf_pf_c_low, mtf_pf_c_mid, mtf_pf_mid, mtf_pf_falloff,
                     dof_coma_max, dof_coma_coeff, dof_ang_min, dof_ang_alpha, dof_ang_kappa,
                     mtf_coma_max, mtf_coma_coeff, mtf_ang_min, mtf_ang_alpha, mtf_ang_kappa,
                     patho_scale,
                     ph_p, sc_p, sa, vc, hc, alpha, kappa, pathologies):
    sa_c_ex = sa_col_exact(sa)
    r_dof = dof_r45[sa_c_ex] / dof_r3[sa_c_ex] if dof_r3[sa_c_ex] != 0 else 0
    r_mtf = mtf_r45[sa_c_ex] / mtf_r3[sa_c_ex] if mtf_r3[sa_c_ex] != 0 else 0

    p = ph_p
    if p <= 3:     s9 = max(dof_pf_min, 1 - dof_pf_c_low * (3 - p)**2)
    elif p <= 4:   s9 = 1 - dof_pf_c_mid * (p - 3)**2
    elif p <= 4.5: s9 = dof_pf_mid - (dof_pf_mid - r_dof) * (p - 4) / 0.5
    else:          s9 = max(dof_pf_falloff, r_dof - 0.1 * (p - 4.5))

    q = sc_p
    if q <= 3:     s10 = max(mtf_pf_min, 1 - mtf_pf_c_low * (3 - q)**2)
    elif q <= 4:   s10 = 1 - mtf_pf_c_mid * (q - 3)**2
    elif q <= 4.5: s10 = mtf_pf_mid - (mtf_pf_mid - r_mtf) * (q - 4) / 0.5
    else:          s10 = max(mtf_pf_falloff, r_mtf - 0.07 * (q - 4.5))

    s12 = dof_r3[sa_c_ex] * s9
    s13 = mtf_r3[sa_c_ex] * s10

    cm = coma_magnitude(vc, hc)
    w13 = round(max(0, s12 - min(dof_coma_max, dof_coma_coeff * cm * (ph_p / 6)**2)), 1)
    ang_factor_dof = max(dof_ang_min, 1 - dof_ang_alpha * max(0, alpha - 0.3) - dof_ang_kappa * max(0, kappa - 0.3))
    w14 = round(max(0, w13 * ang_factor_dof), 1)

    w16 = round(max(0, s13 - min(mtf_coma_max, mtf_coma_coeff * cm * (sc_p / 6)**3)), 1)
    ang_factor_mtf = max(mtf_ang_min, 1 - mtf_ang_alpha * max(0, alpha - 0.3) - mtf_ang_kappa * max(0, kappa - 0.3))
    w17 = round(max(0, w16 * ang_factor_mtf), 1)

    w47, w48 = apply_pathology(w17, w14, pathologies, patho_scale, patho_scale)
    return dict(dof_lookup=s12, mtf_lookup=s13,
                dof_raw=w14, mtf_raw=w17,
                mtf_adj=w47, dof_adj=w48,
                dof_grade=dof_grade(w48), vq_grade=vq_grade(w47))

# ─────────────────────────────────────────────────────────────────────────────
# PERSONALITY & NIGHT DRIVING MODIFIERS
# ─────────────────────────────────────────────────────────────────────────────

PERSONALITY_LABELS = {
    1: "Easy Going",
    2: "Somewhat Easy Going",
    3: "Neutral",
    4: "Somewhat Perfectionist",
    5: "Perfectionist",
}

def personality_weights(p, lens_type):
    """
    Interpolate satisfaction formula weights by personality.
    p=1 (Easy Going): DOF weighted heavily, MTF/dysphotopsia lenient.
    p=5 (Perfectionist): MTF and dysphotopsia weighted heavily, DOF less important.
    Returns [w_mtf, w_dof, w_dyspho, w_last]
      where w_last = w_age (monofocal) or w_dof_quality (edof/diffractive)
    """
    t = (p - 1) / 4.0  # 0.0 = easy going, 1.0 = perfectionist
    if lens_type == 'monofocal':
        # [w_mtf, w_dof, w_dyspho, w_age]
        easy = [0.30, 0.35, 0.15, 0.20]
        perf = [0.55, 0.08, 0.30, 0.07]
    else:
        # [w_mtf, w_dof, w_dyspho, w_dof_quality]
        easy = [0.12, 0.60, 0.10, 0.18]
        perf = [0.32, 0.30, 0.28, 0.10]
    return [easy[i] + t * (perf[i] - easy[i]) for i in range(4)]


def mtf_quality_score(mtf):
    """
    Continuous non-linear MTF quality input (replaces old 4-step discrete buckets).
    MTF on 0.0–1.0 scale.
    Below 0.3 falls steeply; below 0.1 near-zero.
    """
    if mtf >= 0.70:  return 1.00
    elif mtf >= 0.50: return 0.80 + (mtf - 0.50) / 0.20 * 0.20   # 0.80–1.00
    elif mtf >= 0.40: return 0.65 + (mtf - 0.40) / 0.10 * 0.15   # 0.65–0.80
    elif mtf >= 0.30: return 0.42 + (mtf - 0.30) / 0.10 * 0.23   # 0.42–0.65
    elif mtf >= 0.20: return 0.20 + (mtf - 0.20) / 0.10 * 0.22   # 0.20–0.42
    elif mtf >= 0.10: return 0.07 + (mtf - 0.10) / 0.10 * 0.13   # 0.07–0.20
    else:             return max(0.0, mtf / 0.10 * 0.07)            # 0.00–0.07


def mtf_satisfaction_gate(mtf):
    """
    Multiplicative gate applied to the FINAL satisfaction score.
    Captures the clinical reality that DOF is meaningless when MTF is
    critically low — no amount of depth of focus helps if the patient
    cannot see clearly at any focal point.
    At MTF = 0.00 → gate ≈ 0.05 (5% of score).
    At MTF = 0.30 → gate = 0.50 (drastic drop begins).
    At MTF ≥ 0.60 → gate = 1.00 (no penalty).
    """
    if mtf >= 0.60:  return 1.00
    elif mtf >= 0.50: return 0.85 + (mtf - 0.50) / 0.10 * 0.15   # 0.85–1.00
    elif mtf >= 0.40: return 0.68 + (mtf - 0.40) / 0.10 * 0.17   # 0.68–0.85
    elif mtf >= 0.30: return 0.48 + (mtf - 0.30) / 0.10 * 0.20   # 0.48–0.68
    elif mtf >= 0.20: return 0.25 + (mtf - 0.20) / 0.10 * 0.23   # 0.25–0.48
    elif mtf >= 0.10: return 0.10 + (mtf - 0.10) / 0.10 * 0.15   # 0.10–0.25
    else:             return max(0.05, mtf / 0.10 * 0.10)           # 0.05–0.10


def night_driving_factor(driving, mtf_val):
    """
    Returns a satisfaction multiplier based on night driving frequency and MTF.
    mtf_val: 0.0–1.0 scale (mtf_adj / 10)
    - None    : always 1.0
    - MTF ≥ 0.6: always 1.0 regardless of driving
    - Below 0.6: graded penalty, more severe for Regular vs Occasional
    """
    if driving == "None" or mtf_val >= 0.6:
        return 1.0
    # Determine tier
    if mtf_val >= 0.5:   tier = 0  # mild
    elif mtf_val >= 0.4: tier = 1  # moderate
    elif mtf_val >= 0.3: tier = 2  # significant
    else:                tier = 3  # severe
    occasional_penalties = [0.03, 0.06, 0.10, 0.15]
    regular_penalties    = [0.07, 0.12, 0.18, 0.25]
    penalty = occasional_penalties[tier] if driving == "Occasional" else regular_penalties[tier]
    return round(1.0 - penalty, 4)

# ─────────────────────────────────────────────────────────────────────────────
# PATIENT SATISFACTION FUNCTIONS (personality + night driving aware)
# ─────────────────────────────────────────────────────────────────────────────

def calc_patient_satisfaction_monofocal(mtf_max, dof, dyspho, age, pathologies,
                                         personality=3, night_driving="None"):
    grade_sum = sum(pathologies.values())
    mtf_q     = mtf_quality_score(mtf_max)   # continuous non-linear quality input
    age_factor = 1.0 if age > 55 else 0.95
    w = personality_weights(personality, 'monofocal')
    raw = (w[0] * mtf_q +
           w[1] * min(dof / 1.75, 1) +
           w[2] * (1 - dyspho / 10) +
           w[3] * age_factor)
    base  = min(0.78, max(0, raw * (1 - 0.2 * min(grade_sum / 9, 1))))
    score = base * night_driving_factor(night_driving, mtf_max)
    score = round(score * mtf_satisfaction_gate(mtf_max), 2)  # MTF gate — final penalty
    return score


def calc_patient_satisfaction_edof(mtf_max, dof, dyspho, pathologies, ref_stab_text,
                                    personality=3, night_driving="None"):
    grade_sum = sum(pathologies.values())
    mtf_q = mtf_quality_score(mtf_max)        # continuous non-linear quality input
    dof_q = 1.0 if dof >= 1.8 else (0.7 if dof >= 1.5 else 0.3)
    w = personality_weights(personality, 'edof')
    raw = (w[0] * mtf_q +
           w[1] * min(dof / 2.5, 1) +
           w[2] * (1 - dyspho / 10) +
           w[3] * dof_q)
    rs_factor = ref_stability_factor(ref_stab_text)
    base  = min(0.95, max(0, raw * (1 - 0.35 * min(grade_sum / 9, 1)) * rs_factor))
    score = base * night_driving_factor(night_driving, mtf_max)
    score = round(score * mtf_satisfaction_gate(mtf_max), 2)  # MTF gate — final penalty
    return score


def calc_patient_satisfaction_diffractive(mtf_max, dof, dyspho, pathologies,
                                           personality=3, night_driving="None"):
    grade_sum = sum(pathologies.values())
    mtf_q = mtf_quality_score(mtf_max)        # continuous non-linear quality input
    dof_q = 1.0 if dof >= 1.8 else (0.7 if dof >= 1.5 else 0.3)
    w = personality_weights(personality, 'edof')
    raw = (w[0] * mtf_q +
           w[1] * min(dof / 2.5, 1) +
           w[2] * (1 - dyspho / 10) +
           w[3] * dof_q)
    base  = min(0.95, max(0, raw * (1 - 0.35 * min(grade_sum / 9, 1))))
    score = base * night_driving_factor(night_driving, mtf_max)
    score = round(score * mtf_satisfaction_gate(mtf_max), 2)  # MTF gate — final penalty
    return score

# ─────────────────────────────────────────────────────────────────────────────
# LOCAL PERSISTENCE
# ─────────────────────────────────────────────────────────────────────────────

RECORDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "patient_records.json")

def load_records_from_disk():
    """Load saved patient records from local JSON file (if it exists)."""
    try:
        if os.path.exists(RECORDS_FILE):
            with open(RECORDS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return []

def save_records_to_disk(patients):
    """Persist patient records to local JSON file."""
    try:
        with open(RECORDS_FILE, "w", encoding="utf-8") as f:
            json.dump(patients, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def delete_records_from_disk():
    """Remove the local records file."""
    try:
        if os.path.exists(RECORDS_FILE):
            os.remove(RECORDS_FILE)
    except Exception:
        pass

# ─────────────────────────────────────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────────────────────────────────────

def generate_excel(patients):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:

        # ── Sheet 1: Patient Summary ─────────────────────────────────────────
        summary_rows = []
        for pt in patients:
            sorted_iols = sorted(
                pt['results'].items(),
                key=lambda x: x[1]['satisfaction'], reverse=True
            )
            pathology_str = (
                ', '.join([f"{k.title()} (Gr{v})" for k, v in pt['pathologies'].items()])
                if pt['pathologies'] else 'None'
            )
            row = {
                'Date / Time':            pt['timestamp'],
                'Patient ID':             pt['patient_id'],
                'Patient Name':           pt['patient_name'],
                'Surgeon':                pt.get('surgeon_name', ''),
                'Eye':                    pt['eye'],
                'Age':                    pt['age'],
                'Photopic Pupil (mm)':    pt['photopic_pupil'],
                'Scotopic Pupil (mm)':    pt['scotopic_pupil'],
                'Corneal SA (µm)':        pt['corneal_sa'],
                'Vertical Coma (µm)':     pt['vc'],
                'Horizontal Coma (µm)':   pt['hc'],
                'Angle Alpha (°)':        pt['alpha'],
                'Angle Kappa (°)':        pt['kappa'],
                'Night Driving':          pt['night_driving'],
                'Personality':            pt['personality_label'],
                'Pathologies':            pathology_str,
                '1st Choice IOL':         sorted_iols[0][0] if len(sorted_iols) > 0 else '',
                '1st Satisfaction':       f"{sorted_iols[0][1]['satisfaction']:.0%}" if len(sorted_iols) > 0 else '',
                '2nd Choice IOL':         sorted_iols[1][0] if len(sorted_iols) > 1 else '',
                '2nd Satisfaction':       f"{sorted_iols[1][1]['satisfaction']:.0%}" if len(sorted_iols) > 1 else '',
                '3rd Choice IOL':         sorted_iols[2][0] if len(sorted_iols) > 2 else '',
                '3rd Satisfaction':       f"{sorted_iols[2][1]['satisfaction']:.0%}" if len(sorted_iols) > 2 else '',
            }
            summary_rows.append(row)
        df_summary = pd.DataFrame(summary_rows)
        df_summary.to_excel(writer, sheet_name='Patient Summary', index=False)

        # ── Sheet 2: IOL Details ─────────────────────────────────────────────
        detail_rows = []
        for pt in patients:
            for iol_name, res in pt['results'].items():
                detail_rows.append({
                    'Date / Time':               pt['timestamp'],
                    'Patient ID':                pt['patient_id'],
                    'Patient Name':              pt['patient_name'],
                    'Surgeon':                   pt.get('surgeon_name', ''),
                    'Eye':                       pt['eye'],
                    'Night Driving':             pt['night_driving'],
                    'Personality':               pt['personality_label'],
                    'IOL':                       iol_name,
                    'Category':                  res['category'],
                    'DOF (D)':                   res['dof'],
                    'MTF max':                   res['mtf'],
                    'Dysphotopsia Score (/10)':  res['dyspho'],
                    'Refractive Stability':      res['ref_stab'],
                    'Patient Satisfaction (%)':  f"{res['satisfaction']:.0%}",
                })
        df_detail = pd.DataFrame(detail_rows)
        df_detail.to_excel(writer, sheet_name='IOL Details', index=False)

    output.seek(0)
    return output.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
# SINGLE-PATIENT REPORT GENERATION
# ─────────────────────────────────────────────────────────────────────────────

def _build_results_list(sph, zasp, nasp, eyh, emv, viv, pan, gem, atl,
                         dys_sph, dys_za, dys_na, dys_eyh, dys_emv, dys_viv,
                         dys_pan, dys_gem, dys_atl,
                         sat_sph, sat_za, sat_na, sat_eyh, sat_emv, sat_viv,
                         sat_pan, sat_gem, sat_atl,
                         ref_mono, ref_eyh, ref_emv, ref_viv, ref_diff):
    """Assemble all 9 IOL results into a standard list of (name, dict) tuples."""
    return [
        ("Spherical Monofocal",
         {"category":"Monofocal", "dof":sph['dof_adj'], "dof_grade":sph['dof_grade'],
          "mtf":round(sph['mtf_adj']/10,2), "vq_grade":sph['vq_grade'],
          "dyspho":dys_sph, "ref_stab":ref_mono, "satisfaction":sat_sph}),
        ("Zero Aspheric Monofocal",
         {"category":"Monofocal", "dof":zasp['dof_adj'], "dof_grade":zasp['dof_grade'],
          "mtf":round(zasp['mtf_adj']/10,2), "vq_grade":zasp['vq_grade'],
          "dyspho":dys_za, "ref_stab":ref_mono, "satisfaction":sat_za}),
        ("Negative Aspheric Monofocal",
         {"category":"Monofocal", "dof":nasp['dof_adj'], "dof_grade":nasp['dof_grade'],
          "mtf":round(nasp['mtf_adj']/10,2), "vq_grade":nasp['vq_grade'],
          "dyspho":dys_na, "ref_stab":ref_mono, "satisfaction":sat_na}),
        ("Eyhance",
         {"category":"EDOF", "dof":eyh['dof_adj'], "dof_grade":eyh['dof_grade'],
          "mtf":round(eyh['mtf_adj']/10,2), "vq_grade":eyh['vq_grade'],
          "dyspho":dys_eyh, "ref_stab":ref_eyh, "satisfaction":sat_eyh}),
        ("Rayner EMV",
         {"category":"EDOF", "dof":emv['dof_adj'], "dof_grade":emv['dof_grade'],
          "mtf":round(emv['mtf_adj']/10,2), "vq_grade":emv['vq_grade'],
          "dyspho":dys_emv, "ref_stab":ref_emv, "satisfaction":sat_emv}),
        ("Vivity",
         {"category":"EDOF", "dof":viv['dof_adj'], "dof_grade":viv['dof_grade'],
          "mtf":round(viv['mtf_adj']/10,2), "vq_grade":viv['vq_grade'],
          "dyspho":dys_viv, "ref_stab":ref_viv, "satisfaction":sat_viv}),
        ("PanOptix",
         {"category":"Diffractive", "dof":pan['dof_adj'], "dof_grade":pan['dof_grade'],
          "mtf":round(pan['mtf_adj']/10,2), "vq_grade":pan['vq_grade'],
          "dyspho":dys_pan, "ref_stab":ref_diff, "satisfaction":sat_pan}),
        ("Gemetric",
         {"category":"Diffractive", "dof":gem['dof_adj'], "dof_grade":gem['dof_grade'],
          "mtf":round(gem['mtf_adj']/10,2), "vq_grade":gem['vq_grade'],
          "dyspho":dys_gem, "ref_stab":ref_diff, "satisfaction":sat_gem}),
        ("Atlisa Tri",
         {"category":"Diffractive", "dof":atl['dof_adj'], "dof_grade":atl['dof_grade'],
          "mtf":round(atl['mtf_adj']/10,2), "vq_grade":atl['vq_grade'],
          "dyspho":dys_atl, "ref_stab":ref_diff, "satisfaction":sat_atl}),
    ]


def generate_patient_pdf(pid, pname, surgeon, eye, age,
                          photopic, scotopic, sa, vc_v, hc_v,
                          alpha_v, kappa_v, night_drv, pers_lbl, pathos,
                          results_list, sorted_res):
    """Generate a formatted single-patient PDF clinical report."""
    try:
        from fpdf import FPDF
    except ImportError:
        return None

    class _Report(FPDF):
        def footer(self):
            self.set_y(-12)
            self.set_font('Helvetica', 'I', 7)
            self.set_text_color(150, 150, 150)
            self.cell(0, 5,
                'Page ' + str(self.page_no()) +
                '  |  Presbyopia IOL Software  |'
                '  Dr Zain Khatib  |  zainkhatib89@gmail.com',
                align='C')

    pdf = _Report('P', 'mm', 'A4')
    pdf.set_margins(15, 15, 15)
    pdf.set_auto_page_break(True, margin=18)
    pdf.add_page()
    EW = 180  # effective width

    def _sec(title):
        pdf.set_fill_color(26, 107, 181)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font('Helvetica', 'B', 10)
        pdf.cell(EW, 7, '  ' + title, fill=True, ln=True)
        pdf.set_text_color(30, 30, 30)

    # Title
    pdf.set_font('Helvetica', 'B', 17)
    pdf.set_text_color(26, 107, 181)
    pdf.cell(EW, 10, 'PRESBYOPIA IOL SELECTION REPORT', align='C', ln=True)
    pdf.set_font('Helvetica', '', 9)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(EW, 5, 'Clinical Decision Support Tool  |  Developed by Dr Zain Khatib', align='C', ln=True)
    pdf.cell(EW, 5, 'Generated: ' + datetime.datetime.now().strftime('%d %B %Y, %H:%M'), align='C', ln=True)
    pdf.ln(5)

    # Patient details
    _sec('PATIENT DETAILS')
    pdf.set_fill_color(245, 248, 252)
    pdf.set_font('Helvetica', '', 9)
    hw = EW / 2
    pdf.cell(hw, 7, '  Patient ID: ' + (pid or '-'), fill=True, border='B')
    pdf.cell(hw, 7, '  Eye: ' + eye, fill=True, border='B', ln=True)
    pdf.cell(hw, 7, '  Name: ' + (pname or '-'), fill=True, border='B')
    pdf.cell(hw, 7, '  Age: ' + str(age) + ' years', fill=True, border='B', ln=True)
    pdf.cell(hw, 7, '  Surgeon: ' + (surgeon or '-'), fill=True, border='B')
    pdf.cell(hw, 7, '', fill=True, border='B', ln=True)
    pdf.ln(4)

    # Clinical parameters
    _sec('CLINICAL PARAMETERS')
    pdf.set_fill_color(245, 248, 252)
    pdf.set_font('Helvetica', '', 9)
    tw = EW / 3
    pdf.cell(tw, 7, f'  Photopic Pupil: {photopic:.1f} mm', fill=True, border='B')
    pdf.cell(tw, 7, f'  Scotopic Pupil: {scotopic:.1f} mm', fill=True, border='B')
    pdf.cell(tw, 7, f'  Corneal SA: {sa:+.2f} um', fill=True, border='B', ln=True)
    pdf.cell(tw, 7, f'  Vertical Coma: {vc_v:.2f} um', fill=True, border='B')
    pdf.cell(tw, 7, f'  Horizontal Coma: {hc_v:.2f} um', fill=True, border='B')
    pdf.cell(tw, 7, f'  Alpha: {alpha_v:.2f} deg  Kappa: {kappa_v:.2f} deg', fill=True, border='B', ln=True)
    pdf.cell(tw, 7, f'  Night Driving: {night_drv}', fill=True, border='B')
    pdf.cell(tw, 7, f'  Personality: {pers_lbl}', fill=True, border='B')
    patho_str = ', '.join([k.title() + ' Gr' + str(v) for k, v in pathos.items()]) or 'None'
    pdf.cell(tw, 7, '  Pathologies: ' + patho_str[:34], fill=True, border='B', ln=True)
    pdf.ln(4)

    # IOL results table
    _sec('IOL RESULTS  (* = top recommendation)')
    cw = [46, 24, 18, 16, 18, 40, 18]
    hdrs = ['IOL', 'Category', 'DOF (D)', 'MTF', 'Dyspho', 'Refractive Stability', 'Satisf.']
    pdf.set_fill_color(41, 128, 185)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font('Helvetica', 'B', 8)
    for h, c in zip(hdrs, cw):
        pdf.cell(c, 7, ' ' + h, border=1, fill=True)
    pdf.ln()

    top3 = [r[0] for r in sorted_res[:3]]
    for idx, (iol_nm, res) in enumerate(results_list):
        is_top = iol_nm in top3
        if is_top:
            pdf.set_fill_color(228, 247, 228)
        elif idx % 2 == 0:
            pdf.set_fill_color(245, 248, 252)
        else:
            pdf.set_fill_color(255, 255, 255)
        pdf.set_text_color(30, 30, 30)
        pdf.set_font('Helvetica', 'B' if is_top else '', 8)
        prefix = '* ' if is_top else '  '
        _ref = res.get('ref_stab', '')
        if 'no clinically significant' in _ref.lower():
            _ref_s = 'Stable'
        elif 'slight' in _ref.lower():
            _ref_s = 'Slight shift'
        elif 'moderate' in _ref.lower():
            _ref_s = 'Moderate shift'
        elif 'significant' in _ref.lower():
            _ref_s = 'Significant shift'
        else:
            _ref_s = (_ref[:20] + '..') if len(_ref) > 20 else _ref
        vals = [prefix + iol_nm, res['category'],
                f" {res['dof']:.2f} D", f" {res['mtf']:.2f}",
                f" {res['dyspho']:.1f}/10", ' ' + _ref_s,
                f" {res['satisfaction']:.0%}"]
        for v, c in zip(vals, cw):
            pdf.cell(c, 7, v, border=1, fill=True)
        pdf.ln()
    pdf.ln(5)

    # Top recommendations
    _sec('TOP RECOMMENDATIONS')
    pdf.ln(2)
    for i, (iol_nm, res) in enumerate(sorted_res[:3], 1):
        pdf.set_fill_color(228, 247, 228)
        pdf.set_text_color(20, 110, 20)
        pdf.set_font('Helvetica', 'B', 10)
        pdf.cell(EW, 7,
                 f'  {i}. {iol_nm} ({res["category"]})  --  {res["satisfaction"]:.0%} estimated satisfaction',
                 fill=True, border=1, ln=True)
        pdf.set_font('Helvetica', '', 9)
        pdf.set_text_color(50, 50, 50)
        pdf.set_fill_color(240, 250, 240)
        _r = res.get('ref_stab', '')
        pdf.cell(EW, 6,
                 f'     DOF: {res["dof"]:.2f} D  |  MTF: {res["mtf"]:.2f}  |  Dysphotopsia: {res["dyspho"]:.1f}/10',
                 fill=True, border='LRB', ln=True)
        short_ref = (_r[:100] + '..') if len(_r) > 100 else _r
        pdf.cell(EW, 6, '     Refractive Stability: ' + (short_ref or 'Stable'),
                 fill=True, border='LRB', ln=True)
        pdf.ln(3)

    # Disclaimer
    pdf.ln(2)
    pdf.set_font('Helvetica', 'I', 7.5)
    pdf.set_text_color(120, 120, 120)
    pdf.set_fill_color(250, 250, 250)
    pdf.multi_cell(EW, 5,
        'DISCLAIMER: This software is a clinical decision-support tool. All recommendations should be '
        'interpreted by a qualified ophthalmologist in the context of a full clinical assessment. '
        'Scores are predictive models and do not substitute clinical judgment.\n'
        'Developed by Dr Zain Khatib  |  zainkhatib89@gmail.com',
        border=1, fill=True)

    return bytes(pdf.output())


def generate_patient_excel_single(pid, pname, surgeon, eye, age,
                                   photopic, scotopic, sa, vc_v, hc_v,
                                   alpha_v, kappa_v, night_drv, pers_lbl, pathos,
                                   results_list, sorted_res):
    """Generate a formatted single-patient Excel clinical report."""
    output = io.BytesIO()
    top3 = [r[0] for r in sorted_res[:3]]

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1 — patient details
        info_rows = [
            ['PRESBYOPIA IOL SELECTION REPORT', ''],
            ['Developed by Dr Zain Khatib | zainkhatib89@gmail.com', ''],
            ['', ''],
            ['Patient ID',          pid or '-'],
            ['Patient Name',        pname or '-'],
            ['Surgeon',             surgeon or '-'],
            ['Eye',                 eye],
            ['Age',                 str(age) + ' years'],
            ['Generated',           datetime.datetime.now().strftime('%Y-%m-%d %H:%M')],
            ['', ''],
            ['--- CLINICAL PARAMETERS ---', ''],
            ['Photopic Pupil (mm)', photopic],
            ['Scotopic Pupil (mm)', scotopic],
            ['Corneal SA (um)',     sa],
            ['Vertical Coma (um)', vc_v],
            ['Horizontal Coma (um)', hc_v],
            ['Angle Alpha (deg)',   alpha_v],
            ['Angle Kappa (deg)',   kappa_v],
            ['Night Driving',       night_drv],
            ['Personality',         pers_lbl],
            ['Pathologies',
             ', '.join([k.title() + ' (Gr' + str(v) + ')' for k, v in pathos.items()]) or 'None'],
        ]
        pd.DataFrame(info_rows, columns=['Parameter', 'Value']).to_excel(
            writer, sheet_name='Patient Details', index=False)

        # Sheet 2 — IOL results, sorted by satisfaction
        result_rows = []
        for iol_nm, res in results_list:
            rank = top3.index(iol_nm) + 1 if iol_nm in top3 else ''
            result_rows.append({
                'Rank':                     rank,
                'IOL':                      iol_nm,
                'Category':                 res['category'],
                'DOF (D)':                  res['dof'],
                'MTF max':                  res['mtf'],
                'Dysphotopsia Score (/10)': res['dyspho'],
                'Refractive Stability':     res['ref_stab'],
                'Patient Satisfaction':     f"{res['satisfaction']:.0%}",
            })
        # Sort: rank 1→3 first, then unranked rows in original order
        ranked   = sorted([r for r in result_rows if isinstance(r['Rank'], int)], key=lambda x: x['Rank'])
        unranked = [r for r in result_rows if not isinstance(r['Rank'], int)]
        pd.DataFrame(ranked + unranked).to_excel(
            writer, sheet_name='IOL Results', index=False)

    output.seek(0)
    return output.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────────────────────────────────────

def main():
    # Session state — load from disk on first run
    if 'patients' not in st.session_state:
        st.session_state.patients = load_records_from_disk()

    st.title("👁️ Presbyopia IOL Selection Software")
    st.markdown(
        "*Clinical decision support tool for IOL selection based on corneal biometry and ocular parameters*"
    )
    st.markdown("---")

    # ── SIDEBAR ─────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("🔬 Patient Parameters")

        # ── Surgeon ────────────────────────────────────────────────────────
        surgeon_name = st.text_input(
            "👨‍⚕️ Surgeon Name (optional)", value="", placeholder="e.g. Dr. Smith"
        )
        st.markdown("---")

        # ── Patient Identification ─────────────────────────────────────────
        st.subheader("👤 Patient Identification")
        patient_id = st.text_input(
            "Patient ID / MRN", value="", placeholder="e.g. PT-001"
        )
        patient_name = st.text_input(
            "Patient Name (optional)", value="", placeholder="e.g. John Doe"
        )
        eye = st.radio("Eye", ["Right Eye (RE)", "Left Eye (LE)"], horizontal=True)

        # ── Demographics ───────────────────────────────────────────────────
        st.subheader("Demographics")
        age = st.number_input(
            "Age (years)", min_value=18, max_value=99, value=65, step=1
        )

        # ── Pupillometry ───────────────────────────────────────────────────
        st.subheader("Pupillometry")
        photopic_pupil = st.number_input(
            "Photopic Pupil (mm)", min_value=1.0, max_value=8.0,
            value=3.0, step=0.1, format="%.1f"
        )
        scotopic_pupil = st.number_input(
            "Scotopic Pupil (mm)", min_value=1.0, max_value=9.0,
            value=5.0, step=0.1, format="%.1f"
        )

        # ── Corneal Aberrometry ────────────────────────────────────────────
        st.subheader("Corneal Aberrometry (6 mm)")
        corneal_sa = st.number_input(
            "Corneal SA Z40 (µm)", min_value=-0.5, max_value=1.0,
            value=0.28, step=0.01, format="%.2f"
        )
        vc = st.number_input(
            "Vertical Coma (µm)", min_value=-3.0, max_value=3.0,
            value=0.0, step=0.01, format="%.2f"
        )
        hc = st.number_input(
            "Horizontal Coma (µm)", min_value=-3.0, max_value=3.0,
            value=0.0, step=0.01, format="%.2f"
        )

        # ── Angles ────────────────────────────────────────────────────────
        st.subheader("Angles of Offset")
        alpha = st.number_input(
            "Angle Alpha (°)", min_value=0.0, max_value=10.0,
            value=0.2, step=0.05, format="%.2f"
        )
        kappa = st.number_input(
            "Angle Kappa (°)", min_value=0.0, max_value=10.0,
            value=0.2, step=0.05, format="%.2f"
        )

        # ── Lifestyle ─────────────────────────────────────────────────────
        st.subheader("🚗 Lifestyle")
        night_driving = st.selectbox(
            "Night Driving Frequency",
            ["None", "Occasional", "Regular"],
            help=(
                "Affects satisfaction when MTF < 0.6.\n"
                "Occasional → mild penalization. Regular → stronger penalization."
            )
        )

        # ── Personality ───────────────────────────────────────────────────
        st.subheader("🧠 Patient Personality")
        personality = st.slider(
            "Personality Spectrum", min_value=1, max_value=5, value=3, step=1,
            help=(
                "1 = Easy Going: prioritises range of vision (DOF) over perfection.\n"
                "5 = Perfectionist: highly sensitive to contrast loss and dysphotopsia."
            )
        )
        pers_label = PERSONALITY_LABELS[personality]
        pers_emoji = {1: "😊", 2: "🙂", 3: "😐", 4: "🤔", 5: "🧐"}[personality]
        st.caption(f"**{pers_emoji} {pers_label}**")
        if personality >= 4:
            st.caption("Higher weight on MTF quality and dysphotopsia tolerance.")
        elif personality <= 2:
            st.caption("Higher weight on depth of focus; more tolerant of visual imperfections.")
        else:
            st.caption("Balanced weighting across all parameters.")

        # ── Ocular Pathologies ────────────────────────────────────────────
        st.subheader("🏥 Ocular Pathologies")
        st.caption("Select any comorbidities and their severity grade")
        PATHO_LIST = [
            'DRY EYE DISEASE', 'MACULAR PATHOLOGY',
            'DIABETIC RETINOPATHY', 'GLAUCOMA',
            'CORNEAL ENDOTHELIAL DISEASE',
        ]
        pathologies = {}
        for p_name in PATHO_LIST:
            col1, col2 = st.columns([3, 2])
            with col1:
                present = st.checkbox(p_name.title(), key=f"p_{p_name}")
            with col2:
                if present:
                    g = st.selectbox("Grade", [1, 2, 3], key=f"g_{p_name}")
                    pathologies[p_name] = g

    # ── CALCULATIONS ────────────────────────────────────────────────────────

    # Group 1 — Monofocal
    sph  = calc_monofocal(SPH_DOF, SPH_MTF, 1.2, 3.0, 0.85,
                           photopic_pupil, scotopic_pupil, corneal_sa, vc, hc, alpha, pathologies)
    zasp = calc_monofocal(ZA_DOF,  ZA_MTF,  0.8, 2.0, 0.82,
                           photopic_pupil, scotopic_pupil, corneal_sa, vc, hc, alpha, pathologies)
    nasp = calc_monofocal(NA_DOF,  NA_MTF,  1.8, 4.5, 0.80,
                           photopic_pupil, scotopic_pupil, corneal_sa, vc, hc, alpha, pathologies)

    dys_sph  = dysphotopsia_score('SPH',      corneal_sa, scotopic_pupil, vc, hc, alpha, kappa, pathologies)
    dys_za   = dysphotopsia_score('ZERO ASP', corneal_sa, scotopic_pupil, vc, hc, alpha, kappa, pathologies)
    dys_na   = dysphotopsia_score('NEG ASP',  corneal_sa, scotopic_pupil, vc, hc, alpha, kappa, pathologies)

    ref_mono = "No clinically significant pupil-dependent refractive shift expected"

    sat_sph  = calc_patient_satisfaction_monofocal(
        sph['mtf_adj']/10,  sph['dof_adj'],  dys_sph,  age, pathologies, personality, night_driving)
    sat_za   = calc_patient_satisfaction_monofocal(
        zasp['mtf_adj']/10, zasp['dof_adj'], dys_za,   age, pathologies, personality, night_driving)
    sat_na   = calc_patient_satisfaction_monofocal(
        nasp['mtf_adj']/10, nasp['dof_adj'], dys_na,   age, pathologies, personality, night_driving)

    # Group 2 — EDOF
    eyh = calc_edof(EYHANCE_DOF, EYHANCE_PUPIL_ROWS, EYHANCE_MTF, EYHANCE_SCOT_ROWS, 0.95,
                    photopic_pupil, scotopic_pupil, corneal_sa, vc, hc, alpha, pathologies)
    emv = calc_edof(RAYNER_DOF,  RAYNER_PUPIL_ROWS,  RAYNER_MTF,  RAYNER_SCOT_ROWS,  0.90,
                    photopic_pupil, scotopic_pupil, corneal_sa, vc, hc, alpha, pathologies)
    viv = calc_edof(VIVITY_DOF,  VIVITY_PUPIL_ROWS,  VIVITY_MTF,  VIVITY_SCOT_ROWS,  1.10,
                    photopic_pupil, scotopic_pupil, corneal_sa, vc, hc, alpha, pathologies)

    dys_eyh = dysphotopsia_score('EYHANCE', corneal_sa, scotopic_pupil, vc, hc, alpha, kappa, pathologies)
    dys_emv = dysphotopsia_score('EMV',     corneal_sa, scotopic_pupil, vc, hc, alpha, kappa, pathologies)
    dys_viv = dysphotopsia_score('VIVITY',  corneal_sa, scotopic_pupil, vc, hc, alpha, kappa, pathologies)

    ref_eyh = refractive_stability_text('EYHANCE',    photopic_pupil, scotopic_pupil, corneal_sa)
    ref_emv = refractive_stability_text('RAYNER EMV', photopic_pupil, scotopic_pupil, corneal_sa)
    ref_viv = refractive_stability_text('VIVITY',     photopic_pupil, scotopic_pupil, corneal_sa)

    sat_eyh = calc_patient_satisfaction_edof(
        eyh['mtf_adj']/10, eyh['dof_adj'], dys_eyh, pathologies, ref_eyh, personality, night_driving)
    sat_emv = calc_patient_satisfaction_edof(
        emv['mtf_adj']/10, emv['dof_adj'], dys_emv, pathologies, ref_emv, personality, night_driving)
    sat_viv = calc_patient_satisfaction_edof(
        viv['mtf_adj']/10, viv['dof_adj'], dys_viv, pathologies, ref_viv, personality, night_driving)

    # Group 3 — Diffractive
    pan = calc_diffractive(
        PAN_DOF_R3, PAN_DOF_R45, PAN_MTF_R3, PAN_MTF_R45,
        0.55, 0.18, 0.12, 0.88, 0.18,
        0.75, 0.08, 0.05, 0.95, 0.60,
        0.5, 1.2, 0.95, 0.06, 0.03,
        3.0, 5.8, 0.85, 0.30, 0.20,
        1.3,
        photopic_pupil, scotopic_pupil, corneal_sa, vc, hc, alpha, kappa, pathologies)

    gem = calc_diffractive(
        GEM_DOF_R3, GEM_DOF_R45, GEM_MTF_R3, GEM_MTF_R45,
        0.60, 0.15, 0.10, 0.90, 0.22,
        0.80, 0.06, 0.035, 0.96, 0.68,
        0.4, 1.0, 0.96, 0.05, 0.025,
        2.7, 5.0, 0.88, 0.25, 0.15,
        1.2,
        photopic_pupil, scotopic_pupil, corneal_sa, vc, hc, alpha, kappa, pathologies)

    atl = calc_diffractive(
        ATL_DOF_R3, ATL_DOF_R45, ATL_MTF_R3, ATL_MTF_R45,
        0.58, 0.16, 0.11, 0.89, 0.20,
        0.78, 0.07, 0.04, 0.955, 0.64,
        0.48, 1.15, 0.95, 0.058, 0.03,
        3.0, 5.6, 0.85, 0.29, 0.19,
        1.25,
        photopic_pupil, scotopic_pupil, corneal_sa, vc, hc, alpha, kappa, pathologies)

    dys_pan = dysphotopsia_score('PANOP',  corneal_sa, scotopic_pupil, vc, hc, alpha, kappa, pathologies)
    dys_gem = dysphotopsia_score('GEME',   corneal_sa, scotopic_pupil, vc, hc, alpha, kappa, pathologies)
    dys_atl = dysphotopsia_score('ATLISA', corneal_sa, scotopic_pupil, vc, hc, alpha, kappa, pathologies)

    ref_diff = "No clinically significant pupil-dependent refractive shift expected"

    sat_pan = calc_patient_satisfaction_diffractive(
        pan['mtf_adj']/10, pan['dof_adj'], dys_pan, pathologies, personality, night_driving)
    sat_gem = calc_patient_satisfaction_diffractive(
        gem['mtf_adj']/10, gem['dof_adj'], dys_gem, pathologies, personality, night_driving)
    sat_atl = calc_patient_satisfaction_diffractive(
        atl['mtf_adj']/10, atl['dof_adj'], dys_atl, pathologies, personality, night_driving)

    # ── DISPLAY ─────────────────────────────────────────────────────────────

    # Patient / Parameter strip
    eye_short = "RE" if "Right" in eye else "LE"
    pid_display = f"{patient_id} — {patient_name}" if patient_name else (patient_id or "—")
    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
    c1.metric("Patient", pid_display[:14] if len(pid_display) > 14 else pid_display)
    c2.metric("Eye", eye_short)
    c3.metric("Age", f"{age} yrs")
    c4.metric("Photopic", f"{photopic_pupil:.1f} mm")
    c5.metric("Scotopic",  f"{scotopic_pupil:.1f} mm")
    c6.metric("Corneal SA", f"{corneal_sa:+.2f} µm")
    c7.metric("Night Drive", night_driving)
    st.markdown("---")

    def sat_color(score):
        if score >= 0.80: return "🟢"
        elif score >= 0.65: return "🟡"
        else: return "🔴"

    def dyspho_color(score):
        if score <= 3: return "🟢"
        elif score <= 6: return "🟡"
        else: return "🔴"

    # ─── Group 1: Monofocal ───────────────────────────────────────────────
    st.subheader("🔵 Monofocal Lenses")
    m1, m2, m3 = st.columns(3)
    for col, name, res, dys, sat in [
        (m1, "Spherical Monofocal",        sph,  dys_sph,  sat_sph),
        (m2, "Zero Aspheric Monofocal",    zasp, dys_za,   sat_za),
        (m3, "Negative Aspheric Monofocal",nasp, dys_na,   sat_na),
    ]:
        with col:
            st.markdown(f"#### {name}")
            st.metric("Predicted DOF (D)",       f"{res['dof_adj']:.2f}")
            st.caption(f"**Grade:** {res['dof_grade']}")
            st.metric("Predicted MTF max",       f"{res['mtf_adj']/10:.2f}")
            st.caption(f"**Visual Quality:** {res['vq_grade']}")
            st.metric("Dysphotopsia Score (/10)", f"{dys:.1f}  {dyspho_color(dys)}")
            st.metric("Refractive Stability",    "✅ Stable")
            st.caption(ref_mono)
            st.metric("Patient Satisfaction",    f"{sat:.0%}  {sat_color(sat)}")
            st.markdown("---")

    # ─── Group 2: EDOF ───────────────────────────────────────────────────
    st.subheader("🟢 Extended Monofocal / EDOF Lenses")
    e1, e2, e3 = st.columns(3)
    for col, name, res, dys, sat, ref in [
        (e1, "Eyhance",    eyh, dys_eyh, sat_eyh, ref_eyh),
        (e2, "Rayner EMV", emv, dys_emv, sat_emv, ref_emv),
        (e3, "Vivity",     viv, dys_viv, sat_viv, ref_viv),
    ]:
        with col:
            st.markdown(f"#### {name}")
            st.metric("Predicted DOF (D)",       f"{res['dof_adj']:.2f}")
            st.caption(f"**Grade:** {res['dof_grade']}")
            st.metric("Predicted MTF max",       f"{res['mtf_adj']/10:.2f}")
            st.caption(f"**Visual Quality:** {res['vq_grade']}")
            st.metric("Dysphotopsia Score (/10)", f"{dys:.1f}  {dyspho_color(dys)}")
            _ref_icon = "✅" if "no clinically significant" in ref.lower() else "⚠️"
            _ref_lbl  = "Stable" if _ref_icon == "✅" else "Shift expected"
            st.metric("Refractive Stability", f"{_ref_icon} {_ref_lbl}")
            st.caption(ref)
            st.metric("Patient Satisfaction",    f"{sat:.0%}  {sat_color(sat)}")
            st.markdown("---")

    # ─── Group 3: Diffractive ─────────────────────────────────────────────
    st.subheader("🟠 Diffractive / Trifocal Lenses")
    d1, d2, d3 = st.columns(3)
    for col, name, res, dys, sat in [
        (d1, "PanOptix",   pan, dys_pan, sat_pan),
        (d2, "Gemetric",   gem, dys_gem, sat_gem),
        (d3, "Atlisa Tri", atl, dys_atl, sat_atl),
    ]:
        with col:
            st.markdown(f"#### {name}")
            st.metric("Predicted DOF (D)",       f"{res['dof_adj']:.2f}")
            st.caption(f"**Grade:** {res['dof_grade']}")
            st.metric("Predicted MTF max",       f"{res['mtf_adj']/10:.2f}")
            st.caption(f"**Visual Quality:** {res['vq_grade']}")
            st.metric("Dysphotopsia Score (/10)", f"{dys:.1f}  {dyspho_color(dys)}")
            st.metric("Refractive Stability",    "✅ Stable")
            st.caption(ref_diff)
            st.metric("Patient Satisfaction",    f"{sat:.0%}  {sat_color(sat)}")
            st.markdown("---")

    # ─── Comparison Table ─────────────────────────────────────────────────
    st.subheader("📊 Comparison Summary")
    data = {
        "IOL": ["Spherical Mono","Zero Aspheric Mono","Neg Aspheric Mono",
                "Eyhance","Rayner EMV","Vivity",
                "PanOptix","Gemetric","Atlisa Tri"],
        "Category": ["Monofocal"]*3 + ["EDOF"]*3 + ["Diffractive"]*3,
        "DOF (D)":  [sph['dof_adj'], zasp['dof_adj'], nasp['dof_adj'],
                     eyh['dof_adj'], emv['dof_adj'],  viv['dof_adj'],
                     pan['dof_adj'], gem['dof_adj'],  atl['dof_adj']],
        "MTF max":  [round(v/10, 2) for v in [
                     sph['mtf_adj'], zasp['mtf_adj'], nasp['mtf_adj'],
                     eyh['mtf_adj'], emv['mtf_adj'],  viv['mtf_adj'],
                     pan['mtf_adj'], gem['mtf_adj'],  atl['mtf_adj']]],
        "Dysphotopsia (/10)": [dys_sph, dys_za, dys_na,
                                dys_eyh, dys_emv, dys_viv,
                                dys_pan, dys_gem, dys_atl],
        "Satisfaction": [f"{s:.0%}" for s in [
                          sat_sph, sat_za, sat_na,
                          sat_eyh, sat_emv, sat_viv,
                          sat_pan, sat_gem, sat_atl]],
        "DOF Grade": [sph['dof_grade'],  zasp['dof_grade'], nasp['dof_grade'],
                      eyh['dof_grade'],  emv['dof_grade'],  viv['dof_grade'],
                      pan['dof_grade'],  gem['dof_grade'],  atl['dof_grade']],
    }
    df = pd.DataFrame(data)
    st.dataframe(df, use_container_width=True, hide_index=True)

    # ─── Clinical MTF warning ────────────────────────────────────────────
    _all_mtf = [mtf_sph/10, mtf_za/10, mtf_na/10,
                eyh['mtf_adj']/10, emv['mtf_adj']/10, viv['mtf_adj']/10,
                mtf_pan/10, mtf_gem/10, mtf_atl/10]
    _best_mtf = max(_all_mtf)
    if _best_mtf < 0.30:
        st.error(
            "\u26a0\ufe0f **Critical optical quality alert:** Predicted MTF is severely compromised "
            f"across all lenses (best = {_best_mtf:.2f}). Typically caused by very high corneal SA, "
            "large scotopic pupil, and/or high coma. Satisfaction scores will be very low "
            "regardless of IOL choice. "
            "**Consider addressing higher-order aberrations or corneal pathology before IOL selection.**"
        )
    elif _best_mtf < 0.45:
        st.warning(
            f"\u26a0\ufe0f **Reduced optical quality:** Best predicted MTF across all lenses = {_best_mtf:.2f}. "
            "Satisfaction scores are limited by the patient\u2019s optical profile. "
            "Corneal SA, coma, and pupil size are the likely limiting factors."
        )

    # ─── Top Recommendations ──────────────────────────────────────────────
    st.subheader("💡 Top Recommendations")
    sat_nums = [sat_sph, sat_za, sat_na, sat_eyh, sat_emv, sat_viv, sat_pan, sat_gem, sat_atl]
    df['sat_num'] = sat_nums
    df_sorted = df.sort_values('sat_num', ascending=False).head(3)
    for _, row in df_sorted.iterrows():
        st.success(
            f"**{row['IOL']}** ({row['Category']}) — "
            f"Satisfaction: {row['Satisfaction']} | "
            f"DOF: {row['DOF (D)']:.2f} D | "
            f"MTF: {row['MTF max']:.2f} | "
            f"Dysphotopsia: {row['Dysphotopsia (/10)']:.1f}/10\n\n"
            f"*{row['DOF Grade']}*"
        )

    # ─── Download Current Patient Report ────────────────────────────────────
    st.markdown("---")
    st.subheader("📄 Download Patient Report")
    st.caption("Download a formatted report for the current patient — no need to save to the database first.")

    _rl = _build_results_list(
        sph, zasp, nasp, eyh, emv, viv, pan, gem, atl,
        dys_sph, dys_za, dys_na, dys_eyh, dys_emv, dys_viv,
        dys_pan, dys_gem, dys_atl,
        sat_sph, sat_za, sat_na, sat_eyh, sat_emv, sat_viv,
        sat_pan, sat_gem, sat_atl,
        ref_mono, ref_eyh, ref_emv, ref_viv, ref_diff,
    )
    _sr    = sorted(_rl, key=lambda x: x[1]['satisfaction'], reverse=True)
    _args  = (patient_id or '-', patient_name, surgeon_name, eye_short, age,
              photopic_pupil, scotopic_pupil, corneal_sa, vc, hc,
              alpha, kappa, night_driving, PERSONALITY_LABELS[personality],
              pathologies, _rl, _sr)
    _pid_s = (patient_id.strip().replace(' ', '_') or 'Patient')
    _date  = datetime.datetime.now().strftime('%Y%m%d')
    _fname = f"IOL_{_pid_s}_{eye_short}_{_date}"

    rpt_c1, rpt_c2 = st.columns(2)
    with rpt_c1:
        try:
            _pdf = generate_patient_pdf(*_args)
            if _pdf:
                st.download_button(
                    label="📄 Download PDF Report",
                    data=_pdf,
                    file_name=_fname + ".pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
        except Exception:
            st.info("PDF requires `fpdf2` — add it to requirements.txt")
    with rpt_c2:
        _xls = generate_patient_excel_single(*_args)
        st.download_button(
            label="📊 Download Excel Report",
            data=_xls,
            file_name=_fname + ".xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # ─── Save Patient Record ──────────────────────────────────────────────
    st.markdown("---")
    st.subheader("💾 Save Patient Record")

    results_payload = {
        "Spherical Monofocal": {
            "category": "Monofocal", "dof": sph['dof_adj'],
            "dof_grade": sph['dof_grade'], "mtf": round(sph['mtf_adj']/10, 2),
            "vq_grade": sph['vq_grade'], "dyspho": dys_sph,
            "ref_stab": ref_mono, "satisfaction": sat_sph,
        },
        "Zero Aspheric Monofocal": {
            "category": "Monofocal", "dof": zasp['dof_adj'],
            "dof_grade": zasp['dof_grade'], "mtf": round(zasp['mtf_adj']/10, 2),
            "vq_grade": zasp['vq_grade'], "dyspho": dys_za,
            "ref_stab": ref_mono, "satisfaction": sat_za,
        },
        "Negative Aspheric Monofocal": {
            "category": "Monofocal", "dof": nasp['dof_adj'],
            "dof_grade": nasp['dof_grade'], "mtf": round(nasp['mtf_adj']/10, 2),
            "vq_grade": nasp['vq_grade'], "dyspho": dys_na,
            "ref_stab": ref_mono, "satisfaction": sat_na,
        },
        "Eyhance": {
            "category": "EDOF", "dof": eyh['dof_adj'],
            "dof_grade": eyh['dof_grade'], "mtf": round(eyh['mtf_adj']/10, 2),
            "vq_grade": eyh['vq_grade'], "dyspho": dys_eyh,
            "ref_stab": ref_eyh, "satisfaction": sat_eyh,
        },
        "Rayner EMV": {
            "category": "EDOF", "dof": emv['dof_adj'],
            "dof_grade": emv['dof_grade'], "mtf": round(emv['mtf_adj']/10, 2),
            "vq_grade": emv['vq_grade'], "dyspho": dys_emv,
            "ref_stab": ref_emv, "satisfaction": sat_emv,
        },
        "Vivity": {
            "category": "EDOF", "dof": viv['dof_adj'],
            "dof_grade": viv['dof_grade'], "mtf": round(viv['mtf_adj']/10, 2),
            "vq_grade": viv['vq_grade'], "dyspho": dys_viv,
            "ref_stab": ref_viv, "satisfaction": sat_viv,
        },
        "PanOptix": {
            "category": "Diffractive", "dof": pan['dof_adj'],
            "dof_grade": pan['dof_grade'], "mtf": round(pan['mtf_adj']/10, 2),
            "vq_grade": pan['vq_grade'], "dyspho": dys_pan,
            "ref_stab": ref_diff, "satisfaction": sat_pan,
        },
        "Gemetric": {
            "category": "Diffractive", "dof": gem['dof_adj'],
            "dof_grade": gem['dof_grade'], "mtf": round(gem['mtf_adj']/10, 2),
            "vq_grade": gem['vq_grade'], "dyspho": dys_gem,
            "ref_stab": ref_diff, "satisfaction": sat_gem,
        },
        "Atlisa Tri": {
            "category": "Diffractive", "dof": atl['dof_adj'],
            "dof_grade": atl['dof_grade'], "mtf": round(atl['mtf_adj']/10, 2),
            "vq_grade": atl['vq_grade'], "dyspho": dys_atl,
            "ref_stab": ref_diff, "satisfaction": sat_atl,
        },
    }

    save_col, dl_col = st.columns([1, 2])
    with save_col:
        if st.button("💾 Save Current Record", type="primary"):
            if not patient_id.strip():
                st.warning("Please enter a Patient ID before saving.")
            else:
                record = {
                    'timestamp':        datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
                    'patient_id':       patient_id.strip(),
                    'patient_name':     patient_name.strip(),
                    'surgeon_name':     surgeon_name.strip(),
                    'eye':              eye_short,
                    'age':              age,
                    'photopic_pupil':   photopic_pupil,
                    'scotopic_pupil':   scotopic_pupil,
                    'corneal_sa':       corneal_sa,
                    'vc':               vc,
                    'hc':               hc,
                    'alpha':            alpha,
                    'kappa':            kappa,
                    'night_driving':    night_driving,
                    'personality':      personality,
                    'personality_label':PERSONALITY_LABELS[personality],
                    'pathologies':      dict(pathologies),
                    'results':          results_payload,
                }
                st.session_state.patients.append(record)
                save_records_to_disk(st.session_state.patients)
                st.success(
                    f"Saved: **{patient_id}**"
                    + (f" ({patient_name})" if patient_name else "")
                    + f" — {eye_short} (auto-saved to disk)"
                )

    if st.session_state.patients:
        # ─── Surgeon Filter ───────────────────────────────────────────────
        all_surgeons = sorted(set(
            pt.get('surgeon_name', '').strip() or '(No surgeon entered)'
            for pt in st.session_state.patients
        ))
        surgeon_options = ['All Surgeons'] + all_surgeons
        selected_surgeon = st.selectbox(
            "🔍 Filter records by Surgeon",
            options=surgeon_options,
            key='surgeon_filter'
        )
        if selected_surgeon == 'All Surgeons':
            filtered_patients = st.session_state.patients
        elif selected_surgeon == '(No surgeon entered)':
            filtered_patients = [pt for pt in st.session_state.patients
                                 if not pt.get('surgeon_name', '').strip()]
        else:
            filtered_patients = [pt for pt in st.session_state.patients
                                 if pt.get('surgeon_name', '').strip() == selected_surgeon]

        with dl_col:
            excel_bytes = generate_excel(filtered_patients)
            surg_tag = '' if selected_surgeon == 'All Surgeons' else '_' + selected_surgeon.replace(' ', '_')
            fname = f"IOL_Records{surg_tag}_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            st.download_button(
                label=f"📥 Download {len(filtered_patients)} Record(s) as Excel",
                data=excel_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        # ─── Saved Records Table ──────────────────────────────────────────
        st.markdown(f"#### 📋 Saved Records — {selected_surgeon} ({len(filtered_patients)} shown)")
        saved_display = []
        for pt in filtered_patients:
            sorted_iols = sorted(
                pt['results'].items(),
                key=lambda x: x[1]['satisfaction'], reverse=True
            )
            saved_display.append({
                'Time':       pt['timestamp'],
                'Patient ID': pt['patient_id'],
                'Name':       pt['patient_name'],
                'Surgeon':    pt.get('surgeon_name', ''),
                'Eye':        pt['eye'],
                'Age':        pt['age'],
                'Night Drive':pt['night_driving'],
                'Personality':pt['personality_label'],
                'Top IOL':    sorted_iols[0][0],
                'Satisfaction':f"{sorted_iols[0][1]['satisfaction']:.0%}",
                '2nd IOL':    sorted_iols[1][0],
                '3rd IOL':    sorted_iols[2][0],
            })
        st.dataframe(pd.DataFrame(saved_display), use_container_width=True, hide_index=True)

        if st.button("🗑️ Clear All Saved Records"):
            st.session_state.patients = []
            delete_records_from_disk()
            st.rerun()

    # ─── Disclaimer ───────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown(
        """
        <div style="background:#f0f4f9;padding:16px 20px;border-radius:8px;border-left:4px solid #1a6bb5;">
        <strong>⚠️ Disclaimer</strong><br>
        This software is a clinical decision-support tool. All recommendations should be interpreted
        by a qualified ophthalmologist in the context of a full clinical assessment.
        Scores are predictive models and do not substitute clinical judgment.<br><br>
        Developed by <strong>Dr Zain Khatib</strong> &nbsp;|&nbsp;
        ✉️ <a href="mailto:zainkhatib89@gmail.com">zainkhatib89@gmail.com</a> &nbsp;—&nbsp;
        For queries or feedback, feel free to reach out.
        </div>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
