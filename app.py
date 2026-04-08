"""
Presbyopia IOL Selection Software
Converted from PRESBYOPIA-IOL-SOFTWARE-v2.xlsx
"""

import streamlit as st
import math
import bisect

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
# DOF rows: pupil ≤2, 2.1-3, 3.1-3.5, 3.6-4.5, >4.5
SPH_DOF = [
    [1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5],
    [1.25,1.25,1.25,1.25,1.25,1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5],
    [1.25,1.25,1.25,1.25,1.25,1.25,1.25,1.25,1.25,1.25,1.25,1.25,1.25],
    [1.25,1.25,1.25,1.0, 1.0, 1.25,1.25,1.0, 0.75,0.75,0.5, 0.25,0.25],
    [0.25,0.25,0.75,0.75,0.75,0.75,0.75,0.5, 0.25,0.25,0.25,0.0, 0.0 ],
]
# MTF rows: scotopic <3, 3-4.5, 4.6-5.5, >5.5
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
# Row 4 (pupil=4) is interpolated ROUND((row3+row5)/2,1) with M5=1.4, N5=1.3 hardcoded
_eyh_r4_dof = []
for i in range(13):
    if i == 9:   _eyh_r4_dof.append(1.4)  # M5 hardcoded
    elif i == 10: _eyh_r4_dof.append(1.3)  # N5 hardcoded
    else:
        raw = (_eyh_r3_dof[i] + _eyh_r5_dof[i]) / 2
        _eyh_r4_dof.append(round(raw, 1))
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
# pupil_bin: 2, 3, 4, 5  |  sa_cat: "0 AND LESS", "0.1 TO 0.3", "0.4 AND ABOVE"
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
    """Round SA to 0.1, clamp to SA_COLS range, find nearest column index."""
    sa_r = round(min(1.0, max(-0.5, sa)), 1)
    sa_r = round(sa_r, 10)
    best = min(range(len(SA_COLS)), key=lambda i: abs(SA_COLS[i] - sa_r))
    return best

def sa_col_le(sa):
    """Less-than-or-equal SA lookup (for EDOF lenses)."""
    sa_c = min(1.0, max(-0.5, sa))
    idx = bisect.bisect_right(SA_COLS, sa_c) - 1
    return max(0, min(idx, len(SA_COLS) - 1))

def row_le(val, sorted_rows):
    """Find 0-based row index using less-than-or-equal MATCH."""
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
    """Return (mtf_product, dof_product) given dict of {name: grade}."""
    MTF_S = {'DRY EYE DISEASE':[1,0.95,0.85,0.7],'MACULAR PATHOLOGY':[1,0.9,0.75,0.55],
             'DIABETIC RETINOPATHY':[1,0.92,0.8,0.6],'GLAUCOMA':[1,0.93,0.82,0.65],
             'CORNEAL ENDOTHELIAL DISEASE':[1,0.9,0.72,0.5]}
    DOF_S = {'DRY EYE DISEASE':[1,0.97,0.9,0.75],'MACULAR PATHOLOGY':[1,0.92,0.8,0.6],
             'DIABETIC RETINOPATHY':[1,0.94,0.85,0.68],'GLAUCOMA':[1,0.95,0.86,0.7],
             'CORNEAL ENDOTHELIAL DISEASE':[1,0.94,0.82,0.6]}
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

def vq_score_from_mtf(mtf_val):
    """V24/W17 → 0-10 quality score."""
    if mtf_val >= 6:   return 10.0
    elif mtf_val >= 4: return round(7 + ((mtf_val-4)/2)*3, 1)
    elif mtf_val >= 3: return round(5 + (mtf_val-3)*2, 1)
    elif mtf_val >= 2: return round(3 + (mtf_val-2)*2, 1)
    else:              return round(max(0, (mtf_val/2)*3), 1)

def dof_grade(w48):
    if w48 < 0.75:  return "Distance only / No functional near"
    if w48 < 1.25:  return "Weak intermediate, poor near"
    if w48 < 1.75:  return "Functional intermediate, limited near"
    if w48 < 2.25:  return "Good intermediate and near"
    return "Excellent near performance"

def vq_grade(w47):
    if w47 >= 6: return "Excellent visual quality across all lighting conditions"
    if w47 >= 5: return "Excellent visual quality in normal lighting; mild reduction in very low-light"
    if w47 >= 4: return "Very good visual quality in bright lighting; noticeable decline in low-light"
    if w47 >= 3: return "Functional visual quality in good lighting; significant deterioration in dim light"
    if w47 >= 2: return "Suboptimal visual quality even in bright conditions; poor performance in low light"
    return "Poor overall visual quality across all lighting conditions"

def coma_magnitude(vc, hc):
    return math.sqrt(vc**2 + (0.7*hc)**2)

def dysphotopsia_score(iol_key, corneal_sa, scotopic_pupil,
                        vc, hc, alpha, kappa, pathologies):
    """Full dysphotopsia calculation (F column from DYSPHOTOPSIA SCORES sheet)."""
    IOL_MULT = {'NEG ASP':0.75,'ZERO ASP':0.78,'SPH':0.8,'EMV':0.85,
                'EYHANCE':0.9,'VIVITY':1.0,'GEME':1.1,'ATLISA':1.15,'PANOP':1.2}
    PATHO_W = {'DRY EYE DISEASE':(0.6,[0,0.5,1.5,3]),'MACULAR PATHOLOGY':(0.5,[0,0.3,0.8,1.5]),
               'DIABETIC RETINOPATHY':(0.5,[0,0.2,0.5,1]),'GLAUCOMA':(0.5,[0,0.2,0.5,1]),
               'CORNEAL ENDOTHELIAL DISEASE':(0.6,[0,0.8,2,4])}

    total_hoa = round(math.sqrt((1.5*corneal_sa)**2 + vc**2 + hc**2), 3)

    # Basic score (C column)
    if iol_key == 'NEG ASP':   diff = abs(corneal_sa - 0.15)
    elif iol_key == 'ZERO ASP': diff = abs(corneal_sa)
    elif iol_key == 'SPH':      diff = abs(corneal_sa + 0.15)
    else:
        fixed = {'EMV':2.0,'EYHANCE':2.2,'VIVITY':3.4,'GEME':4.6,'ATLISA':4.9,'PANOP':5.1}
        diff = None
        basic = fixed.get(iol_key, 5.0)

    if diff is not None:
        basic = 0.0 if diff<=0.1 else 0.4 if diff<=0.25 else 0.8 if diff<=0.5 else 1.5

    # Pupil modifier (D column)
    pupil_mod = round(max(0, ((scotopic_pupil-3)**2)*0.08*basic), 2) + basic

    # Angle/HOA modification (E column)
    e_score = round(min(10, pupil_mod +
                  ((0.6*total_hoa)+(1.2*max(0,alpha-0.2))+(0.6*max(0,kappa-0.2)))
                  *(0.5+0.05*pupil_mod)), 1)

    # Pathology addition (F column)
    dry_g = min(3, pathologies.get('DRY EYE DISEASE', 0))
    corn_g = min(3, pathologies.get('CORNEAL ENDOTHELIAL DISEASE', 0))
    patho_total = 0.0
    for name, (wt, vals) in PATHO_W.items():
        g = min(3, pathologies.get(name, 0))
        patho_total += wt * vals[g]
    if dry_g >= 2 and corn_g >= 2:
        patho_total += 0.5

    mult = IOL_MULT.get(iol_key, 1.0)
    f_score = round(min(10, e_score + patho_total * mult), 1)
    return f_score

def refractive_stability_text(iol_name, photopic, scotopic, corneal_sa):
    """Compute refractive stability string for EDOF lenses."""
    TABLES = {'EYHANCE': EYHANCE_REF, 'RAYNER EMV': RAYNER_REF, 'VIVITY': VIVITY_REF}
    if iol_name not in TABLES:
        return "No clinically significant pupil-dependent refractive shift expected"
    tbl = TABLES[iol_name]

    def sa_cat(sa): 
        return '0 AND LESS' if sa<=0 else ('0.1 TO 0.3' if sa<0.4 else '0.4 AND ABOVE')
    def pupil_bin(p):
        return 2 if p<2.5 else (3 if p<3.5 else (4 if p<4.5 else 5))

    sa_idx = {'0 AND LESS':0, '0.1 TO 0.3':1, '0.4 AND ABOVE':2}[sa_cat(corneal_sa)]
    j = tbl[pupil_bin(photopic)][sa_idx]
    k = tbl[pupil_bin(scotopic)][sa_idx]

    if abs(j)<0.5 and abs(k)<0.5:
        return "No clinically significant pupil-dependent refractive shift expected"
    if j<=-0.5 and abs(k)<0.5:
        return "Beware of myopic shift for distance in bright light / small pupil conditions"
    if j>=0.5 and abs(k)<0.5:
        return "Possible hyperopic shift for distance in bright light / small pupil conditions"
    if abs(j)<0.5 and k<=-0.5:
        return "Chances of low-light myopia / night myopia likely"
    if abs(j)<0.5 and k>=0.5:
        return "Possible hyperopic shift in dim light / larger pupil conditions"
    if j<=-0.5 and k<=-0.5:
        return "Consistent myopic shift tendency across lighting conditions"
    if j>=0.5 and k>=0.5:
        return "Consistent hyperopic shift tendency across lighting conditions"
    if j<=-0.5 and k>=0.5:
        return "Lighting-dependent refractive behavior likely, with myopic tendency in bright light and hyperopic in dim light"
    if j>=0.5 and k<=-0.5:
        return "Lighting-dependent refractive behavior likely, with hyperopic tendency in bright light and night myopia in dim light"
    return "No clinically significant pupil-dependent refractive shift expected"

def ref_stability_factor(text):
    if "No clinically significant" in text:   return 1.0
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
    """Generic monofocal / zero-aspheric / negative-aspheric calculation."""
    sa_c   = sa_col_exact(sa)
    dof_r  = mono_dof_row(ph_p)
    mtf_r  = mono_mtf_row(sc_p)
    s13    = dof_tbl[dof_r][sa_c]
    s14    = mtf_tbl[mtf_r][sa_c]
    cm     = coma_magnitude(vc, hc)
    # DOF with pseudoaccommodative coma boost
    v12    = round(s13 + min(0.45, 1.0*cm*(ph_p/6)**3), 2)
    # MTF with coma penalty
    v21    = round(max(0, s14 - min(2.5, 5.0*cm*(sc_p/6)**3)), 1)
    # Angle penalty (differs per lens type)
    u24    = round(min(ang_max, max(0, (alpha-0.3)*ang_mult)), 1)
    v24    = round(max(0, v21 - u24), 1)
    # Pathology adjustment
    w47, w48 = apply_pathology(v24, v12, pathologies, patho_scale, patho_scale)
    return dict(dof_lookup=s13, mtf_lookup=s14,
                dof_raw=v12, mtf_raw=v24,
                mtf_adj=w47, dof_adj=w48,
                dof_grade=dof_grade(w48), vq_grade=vq_grade(w47))

def calc_edof(dof_tbl, dof_rows, mtf_tbl, mtf_rows, patho_scale,
              ph_p, sc_p, sa, vc, hc, alpha, pathologies):
    """EYHANCE / RAYNER EMV / VIVITY calculation."""
    sa_c   = sa_col_le(sa)
    dof_r  = row_le(max(2, min(5, ph_p)), dof_rows)
    mtf_r  = row_le(max(2, min(5, sc_p)), mtf_rows)
    s13    = dof_tbl[dof_r][sa_c]
    s14    = mtf_tbl[mtf_r][sa_c]
    cm     = coma_magnitude(vc, hc)
    # W12: DOF alpha penalty factor
    if alpha <= 0.25:      w12 = 1.0
    elif alpha <= 0.35:    w12 = 1.0 - ((alpha-0.25)/0.1)*0.08
    elif alpha <= 0.4:     w12 = 0.92 - ((alpha-0.35)/0.05)*0.12
    else:                  w12 = 0.8
    # DOF with coma boost * alpha factor
    v12    = round((s13 + min(0.45, cm*(ph_p/6)**3)) * w12, 2)
    # MTF with coma penalty
    v21    = round(max(0, s14 - min(2.5, 5.0*cm*(sc_p/6)**3)), 1)
    # Angle penalty (same formula as spherical monofocal)
    u24    = round(min(1.2, max(0, (alpha-0.3)*3)), 1)
    v24    = round(max(0, v21 - u24), 1)
    w47, w48 = apply_pathology(v24, v12, pathologies, patho_scale, patho_scale)
    return dict(dof_lookup=s13, mtf_lookup=s14,
                dof_raw=v12, mtf_raw=v24,
                mtf_adj=w47, dof_adj=w48,
                dof_grade=dof_grade(w48), vq_grade=vq_grade(w47))

def calc_diffractive(dof_r3, dof_r45, mtf_r3, mtf_r45,
                     # DOF pupil factor params
                     dof_pf_min, dof_pf_c_low, dof_pf_c_mid, dof_pf_mid, dof_pf_falloff,
                     # MTF pupil factor params
                     mtf_pf_min, mtf_pf_c_low, mtf_pf_c_mid, mtf_pf_mid, mtf_pf_falloff,
                     # Coma/angle params
                     dof_coma_max, dof_coma_coeff, dof_ang_min, dof_ang_alpha, dof_ang_kappa,
                     mtf_coma_max, mtf_coma_coeff, mtf_ang_min, mtf_ang_alpha, mtf_ang_kappa,
                     patho_scale,
                     ph_p, sc_p, sa, vc, hc, alpha, kappa, pathologies):
    """Generic diffractive IOL calculation (PANOPTIX / GEMETRIC / ATLISA TRI)."""
    sa_c_ex  = sa_col_exact(sa)  # exact match for lookup (ROUND(SA,1) in Excel)
    r_dof = dof_r45[sa_c_ex] / dof_r3[sa_c_ex] if dof_r3[sa_c_ex] != 0 else 0
    r_mtf = mtf_r45[sa_c_ex] / mtf_r3[sa_c_ex] if mtf_r3[sa_c_ex] != 0 else 0

    # DOF pupil factor (S9)
    p = ph_p
    if p <= 3:     s9 = max(dof_pf_min, 1 - dof_pf_c_low*(3-p)**2)
    elif p <= 4:   s9 = 1 - dof_pf_c_mid*(p-3)**2
    elif p <= 4.5: s9 = dof_pf_mid - (dof_pf_mid - r_dof)*(p-4)/0.5
    else:          s9 = max(dof_pf_falloff, r_dof - 0.1*(p-4.5))

    # MTF pupil factor (S10)
    q = sc_p
    if q <= 3:     s10 = max(mtf_pf_min, 1 - mtf_pf_c_low*(3-q)**2)
    elif q <= 4:   s10 = 1 - mtf_pf_c_mid*(q-3)**2
    elif q <= 4.5: s10 = mtf_pf_mid - (mtf_pf_mid - r_mtf)*(q-4)/0.5
    else:          s10 = max(mtf_pf_falloff, r_mtf - 0.07*(q-4.5))

    s12 = dof_r3[sa_c_ex] * s9   # DOF base
    s13 = mtf_r3[sa_c_ex] * s10  # MTF base

    cm = coma_magnitude(vc, hc)
    # DOF: subtract coma (uses photopic^2)
    w13 = round(max(0, s12 - min(dof_coma_max, dof_coma_coeff*cm*(ph_p/6)**2)), 1)
    # DOF: angle penalty as multiplier
    ang_factor_dof = max(dof_ang_min, 1 - dof_ang_alpha*max(0,alpha-0.3) - dof_ang_kappa*max(0,kappa-0.3))
    w14 = round(max(0, w13 * ang_factor_dof), 1)

    # MTF: subtract coma (uses scotopic^3)
    w16 = round(max(0, s13 - min(mtf_coma_max, mtf_coma_coeff*cm*(sc_p/6)**3)), 1)
    # MTF: angle penalty as multiplier
    ang_factor_mtf = max(mtf_ang_min, 1 - mtf_ang_alpha*max(0,alpha-0.3) - mtf_ang_kappa*max(0,kappa-0.3))
    w17 = round(max(0, w16 * ang_factor_mtf), 1)

    w47, w48 = apply_pathology(w17, w14, pathologies, patho_scale, patho_scale)
    return dict(dof_lookup=s12, mtf_lookup=s13,
                dof_raw=w14, mtf_raw=w17,
                mtf_adj=w47, dof_adj=w48,
                dof_grade=dof_grade(w48), vq_grade=vq_grade(w47))

def calc_patient_satisfaction_monofocal(mtf_max, dof, dyspho, age, pathologies):
    grade_sum = sum(pathologies.values())
    mtf_q = 1.0 if mtf_max>=0.5 else (0.85 if mtf_max>=0.4 else (0.65 if mtf_max>=0.3 else 0.35))
    raw = (0.45*mtf_q + 0.2*min(dof/1.75,1) + 0.2*(1-dyspho/10) +
           0.15*(1.0 if age>55 else 0.95))
    score = round(min(0.78, max(0, raw*(1-0.2*min(grade_sum/9,1)))), 2)
    return score

def calc_patient_satisfaction_edof(mtf_max, dof, dyspho, pathologies, ref_stab_text):
    grade_sum = sum(pathologies.values())
    mtf_q = 1.0 if mtf_max>=0.5 else (0.85 if mtf_max>=0.4 else (0.65 if mtf_max>=0.3 else 0.3))
    dof_q = 1.0 if dof>=1.8 else (0.7 if dof>=1.5 else 0.3)
    raw = (0.22*mtf_q + 0.48*min(dof/2.5,1) + 0.18*(1-dyspho/10) + 0.12*dof_q)
    rs_factor = ref_stability_factor(ref_stab_text)
    score = round(min(0.95, max(0, raw*(1-0.35*min(grade_sum/9,1))*rs_factor)), 2)
    return score

def calc_patient_satisfaction_diffractive(mtf_max, dof, dyspho, pathologies):
    grade_sum = sum(pathologies.values())
    mtf_q = 1.0 if mtf_max>=0.5 else (0.85 if mtf_max>=0.4 else (0.65 if mtf_max>=0.3 else 0.3))
    dof_q = 1.0 if dof>=1.8 else (0.7 if dof>=1.5 else 0.3)
    raw = (0.22*mtf_q + 0.48*min(dof/2.5,1) + 0.18*(1-dyspho/10) + 0.12*dof_q)
    score = round(min(0.95, max(0, raw*(1-0.35*min(grade_sum/9,1)))), 2)
    return score

# ─────────────────────────────────────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────────────────────────────────────

def main():
    st.title("👁️ Presbyopia IOL Selection Software")
    st.markdown("*Clinical decision support tool for IOL selection based on corneal biometry and ocular parameters*")
    st.markdown("---")

    # ── SIDEBAR INPUTS ──────────────────────────────────────────────────────
    with st.sidebar:
        st.header("🔬 Patient Parameters")

        st.subheader("Demographics")
        age = st.number_input("Age (years)", min_value=18, max_value=99, value=65, step=1)

        st.subheader("Pupillometry")
        photopic_pupil = st.number_input("Photopic Pupil (mm)", min_value=1.0, max_value=8.0,
                                          value=3.0, step=0.1, format="%.1f")
        scotopic_pupil = st.number_input("Scotopic Pupil (mm)", min_value=1.0, max_value=9.0,
                                          value=5.0, step=0.1, format="%.1f")

        st.subheader("Corneal Aberrometry (6mm)")
        corneal_sa = st.number_input("Corneal Spherical Aberration Z40 (µm)",
                                      min_value=-0.5, max_value=1.0,
                                      value=0.28, step=0.01, format="%.2f")
        vc = st.number_input("Vertical Coma (µm)", min_value=-1.0, max_value=1.0,
                              value=0.0, step=0.01, format="%.2f")
        hc = st.number_input("Horizontal Coma (µm)", min_value=-1.0, max_value=1.0,
                              value=0.0, step=0.01, format="%.2f")

        st.subheader("Angles of Offset")
        alpha = st.number_input("Angle Alpha (degrees)", min_value=0.0, max_value=10.0,
                                 value=0.2, step=0.05, format="%.2f")
        kappa = st.number_input("Angle Kappa (degrees)", min_value=0.0, max_value=10.0,
                                 value=0.2, step=0.05, format="%.2f")

        st.subheader("🏥 Ocular Pathologies")
        st.caption("Select any comorbidities and their severity grade")
        PATHO_LIST = ['DRY EYE DISEASE','MACULAR PATHOLOGY',
                      'DIABETIC RETINOPATHY','GLAUCOMA','CORNEAL ENDOTHELIAL DISEASE']
        pathologies = {}
        for p in PATHO_LIST:
            col1, col2 = st.columns([3, 2])
            with col1:
                present = st.checkbox(p.title(), key=f"p_{p}")
            with col2:
                if present:
                    g = st.selectbox("Grade", [1, 2, 3], key=f"g_{p}")
                    pathologies[p] = g

    # ── CALCULATIONS ────────────────────────────────────────────────────────

    # ── Group 1: Monofocal Lenses ─────────────────────────────────────────
    sph  = calc_monofocal(SPH_DOF, SPH_MTF, 1.2, 3.0, 0.85,
                           photopic_pupil, scotopic_pupil, corneal_sa, vc, hc, alpha, pathologies)
    zasp = calc_monofocal(ZA_DOF, ZA_MTF,  0.8, 2.0, 0.82,
                           photopic_pupil, scotopic_pupil, corneal_sa, vc, hc, alpha, pathologies)
    nasp = calc_monofocal(NA_DOF, NA_MTF,  1.8, 4.5, 0.80,
                           photopic_pupil, scotopic_pupil, corneal_sa, vc, hc, alpha, pathologies)

    # Dysphotopsia
    dys_sph  = dysphotopsia_score('SPH',      corneal_sa, scotopic_pupil, vc, hc, alpha, kappa, pathologies)
    dys_za   = dysphotopsia_score('ZERO ASP', corneal_sa, scotopic_pupil, vc, hc, alpha, kappa, pathologies)
    dys_na   = dysphotopsia_score('NEG ASP',  corneal_sa, scotopic_pupil, vc, hc, alpha, kappa, pathologies)

    # Refractive stability (monofocals always stable)
    ref_mono = "No clinically significant pupil-dependent refractive shift expected"

    # Patient satisfaction
    sat_sph  = calc_patient_satisfaction_monofocal(sph['mtf_adj']/10,  sph['dof_adj'],  dys_sph,  age, pathologies)
    sat_za   = calc_patient_satisfaction_monofocal(zasp['mtf_adj']/10, zasp['dof_adj'], dys_za,   age, pathologies)
    sat_na   = calc_patient_satisfaction_monofocal(nasp['mtf_adj']/10, nasp['dof_adj'], dys_na,   age, pathologies)

    # ── Group 2: Extended Monofocal / EDOF ───────────────────────────────
    eyh   = calc_edof(EYHANCE_DOF, EYHANCE_PUPIL_ROWS, EYHANCE_MTF, EYHANCE_SCOT_ROWS, 0.95,
                      photopic_pupil, scotopic_pupil, corneal_sa, vc, hc, alpha, pathologies)
    emv   = calc_edof(RAYNER_DOF,  RAYNER_PUPIL_ROWS,  RAYNER_MTF,  RAYNER_SCOT_ROWS,  0.90,
                      photopic_pupil, scotopic_pupil, corneal_sa, vc, hc, alpha, pathologies)
    viv   = calc_edof(VIVITY_DOF,  VIVITY_PUPIL_ROWS,  VIVITY_MTF,  VIVITY_SCOT_ROWS,  1.10,
                      photopic_pupil, scotopic_pupil, corneal_sa, vc, hc, alpha, pathologies)

    dys_eyh = dysphotopsia_score('EYHANCE', corneal_sa, scotopic_pupil, vc, hc, alpha, kappa, pathologies)
    dys_emv = dysphotopsia_score('EMV',     corneal_sa, scotopic_pupil, vc, hc, alpha, kappa, pathologies)
    dys_viv = dysphotopsia_score('VIVITY',  corneal_sa, scotopic_pupil, vc, hc, alpha, kappa, pathologies)

    ref_eyh = refractive_stability_text('EYHANCE',    photopic_pupil, scotopic_pupil, corneal_sa)
    ref_emv = refractive_stability_text('RAYNER EMV', photopic_pupil, scotopic_pupil, corneal_sa)
    ref_viv = refractive_stability_text('VIVITY',     photopic_pupil, scotopic_pupil, corneal_sa)

    sat_eyh = calc_patient_satisfaction_edof(eyh['mtf_adj']/10, eyh['dof_adj'], dys_eyh, pathologies, ref_eyh)
    sat_emv = calc_patient_satisfaction_edof(emv['mtf_adj']/10, emv['dof_adj'], dys_emv, pathologies, ref_emv)
    sat_viv = calc_patient_satisfaction_edof(viv['mtf_adj']/10, viv['dof_adj'], dys_viv, pathologies, ref_viv)

    # ── Group 3: Diffractive / Trifocal ──────────────────────────────────
    # PANOPTIX
    pan = calc_diffractive(
        PAN_DOF_R3, PAN_DOF_R45, PAN_MTF_R3, PAN_MTF_R45,
        0.55, 0.18, 0.12, 0.88, 0.18,
        0.75, 0.08, 0.05, 0.95, 0.60,
        0.5, 1.2, 0.95, 0.06, 0.03,
        3.0, 5.8, 0.85, 0.30, 0.20,
        1.3,
        photopic_pupil, scotopic_pupil, corneal_sa, vc, hc, alpha, kappa, pathologies)

    # GEMETRIC
    gem = calc_diffractive(
        GEM_DOF_R3, GEM_DOF_R45, GEM_MTF_R3, GEM_MTF_R45,
        0.60, 0.15, 0.10, 0.90, 0.22,
        0.80, 0.06, 0.035, 0.96, 0.68,
        0.4, 1.0, 0.96, 0.05, 0.025,
        2.7, 5.0, 0.88, 0.25, 0.15,
        1.2,
        photopic_pupil, scotopic_pupil, corneal_sa, vc, hc, alpha, kappa, pathologies)

    # ATLISA TRI
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

    sat_pan = calc_patient_satisfaction_diffractive(pan['mtf_adj']/10, pan['dof_adj'], dys_pan, pathologies)
    sat_gem = calc_patient_satisfaction_diffractive(gem['mtf_adj']/10, gem['dof_adj'], dys_gem, pathologies)
    sat_atl = calc_patient_satisfaction_diffractive(atl['mtf_adj']/10, atl['dof_adj'], dys_atl, pathologies)

    # ── DISPLAY ─────────────────────────────────────────────────────────────

    # Summary parameter strip
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("Age",           f"{age} yrs")
    col2.metric("Photopic Pupil", f"{photopic_pupil:.1f} mm")
    col3.metric("Scotopic Pupil", f"{scotopic_pupil:.1f} mm")
    col4.metric("Corneal SA",     f"{corneal_sa:+.2f} µm")
    col5.metric("Angle α",        f"{alpha:.2f}°")
    col6.metric("Angle κ",        f"{kappa:.2f}°")
    st.markdown("---")

    # Colour helper for satisfaction score
    def sat_color(score):
        if score >= 0.85: return "🟢"
        elif score >= 0.70: return "🟡"
        else: return "🔴"

    def dyspho_color(score):
        if score <= 3: return "🟢"
        elif score <= 6: return "🟡"
        else: return "🔴"

    # ─── SECTION 1: MONOFOCAL ───────────────────────────────────────────
    st.subheader("🔵 Monofocal Lenses")
    m1, m2, m3 = st.columns(3)

    for col, name, res, dys, sat in [
        (m1, "Spherical Monofocal",         sph,  dys_sph,  sat_sph),
        (m2, "Zero Aspheric Monofocal",      zasp, dys_za,   sat_za),
        (m3, "Negative Aspheric Monofocal",  nasp, dys_na,   sat_na),
    ]:
        with col:
            st.markdown(f"#### {name}")
            st.metric("Predicted DOF (D)", f"{res['dof_adj']:.2f}")
            st.caption(f"**Grade:** {res['dof_grade']}")
            st.metric("Predicted MTF max", f"{res['mtf_adj']/10:.2f}")
            st.caption(f"**Visual Quality:** {res['vq_grade']}")
            st.metric("Dysphotopsia Score (/10)", f"{dys:.1f}  {dyspho_color(dys)}")
            st.metric("Refractive Stability", "✅ Stable")
            st.caption(ref_mono)
            st.metric("Patient Satisfaction", f"{sat:.0%}  {sat_color(sat)}")
            st.markdown("---")

    # ─── SECTION 2: EDOF / EXTENDED MONOFOCAL ───────────────────────────
    st.subheader("🟢 Extended Monofocal / EDOF Lenses")
    e1, e2, e3 = st.columns(3)

    for col, name, res, dys, sat, ref in [
        (e1, "Eyhance",    eyh, dys_eyh, sat_eyh, ref_eyh),
        (e2, "Rayner EMV", emv, dys_emv, sat_emv, ref_emv),
        (e3, "Vivity",     viv, dys_viv, sat_viv, ref_viv),
    ]:
        with col:
            st.markdown(f"#### {name}")
            st.metric("Predicted DOF (D)", f"{res['dof_adj']:.2f}")
            st.caption(f"**Grade:** {res['dof_grade']}")
            st.metric("Predicted MTF max", f"{res['mtf_adj']/10:.2f}")
            st.caption(f"**Visual Quality:** {res['vq_grade']}")
            st.metric("Dysphotopsia Score (/10)", f"{dys:.1f}  {dyspho_color(dys)}")
            with st.expander("Refractive Stability"):
                st.write(ref)
            st.metric("Patient Satisfaction", f"{sat:.0%}  {sat_color(sat)}")
            st.markdown("---")

    # ─── SECTION 3: DIFFRACTIVE / TRIFOCAL ──────────────────────────────
    st.subheader("🟠 Diffractive / Trifocal Lenses")
    d1, d2, d3 = st.columns(3)

    for col, name, res, dys, sat in [
        (d1, "PanOptix",   pan, dys_pan, sat_pan),
        (d2, "Gemetric",   gem, dys_gem, sat_gem),
        (d3, "Atlisa Tri", atl, dys_atl, sat_atl),
    ]:
        with col:
            st.markdown(f"#### {name}")
            st.metric("Predicted DOF (D)", f"{res['dof_adj']:.2f}")
            st.caption(f"**Grade:** {res['dof_grade']}")
            st.metric("Predicted MTF max", f"{res['mtf_adj']/10:.2f}")
            st.caption(f"**Visual Quality:** {res['vq_grade']}")
            st.metric("Dysphotopsia Score (/10)", f"{dys:.1f}  {dyspho_color(dys)}")
            st.metric("Refractive Stability", "✅ Stable")
            st.caption(ref_diff)
            st.metric("Patient Satisfaction", f"{sat:.0%}  {sat_color(sat)}")
            st.markdown("---")

    # ─── COMPARISON TABLE ────────────────────────────────────────────────
    st.subheader("📊 Comparison Summary")
    import pandas as pd

    data = {
        "IOL": ["Spherical Mono","Zero Aspheric Mono","Neg Aspheric Mono",
                "Eyhance","Rayner EMV","Vivity",
                "PanOptix","Gemetric","Atlisa Tri"],
        "Category": ["Monofocal"]*3 + ["EDOF"]*3 + ["Diffractive"]*3,
        "DOF (D)": [sph['dof_adj'], zasp['dof_adj'], nasp['dof_adj'],
                    eyh['dof_adj'], emv['dof_adj'],  viv['dof_adj'],
                    pan['dof_adj'], gem['dof_adj'],  atl['dof_adj']],
        "MTF max": [round(sph['mtf_adj']/10,2), round(zasp['mtf_adj']/10,2), round(nasp['mtf_adj']/10,2),
                    round(eyh['mtf_adj']/10,2),  round(emv['mtf_adj']/10,2),  round(viv['mtf_adj']/10,2),
                    round(pan['mtf_adj']/10,2),  round(gem['mtf_adj']/10,2),  round(atl['mtf_adj']/10,2)],
        "Dysphotopsia (/10)": [dys_sph, dys_za, dys_na,
                                dys_eyh, dys_emv, dys_viv,
                                dys_pan, dys_gem, dys_atl],
        "Satisfaction": [f"{sat_sph:.0%}", f"{sat_za:.0%}", f"{sat_na:.0%}",
                         f"{sat_eyh:.0%}", f"{sat_emv:.0%}", f"{sat_viv:.0%}",
                         f"{sat_pan:.0%}", f"{sat_gem:.0%}", f"{sat_atl:.0%}"],
        "DOF Grade": [sph['dof_grade'],  zasp['dof_grade'], nasp['dof_grade'],
                      eyh['dof_grade'],  emv['dof_grade'],  viv['dof_grade'],
                      pan['dof_grade'],  gem['dof_grade'],  atl['dof_grade']],
    }
    df = pd.DataFrame(data)
    st.dataframe(df, use_container_width=True, hide_index=True)

    # Best IOL recommendation
    st.subheader("💡 Top Recommendations")
    df['sat_num'] = [sat_sph, sat_za, sat_na, sat_eyh, sat_emv, sat_viv,
                     sat_pan, sat_gem, sat_atl]
    df_sorted = df.sort_values('sat_num', ascending=False).head(3)
    for _, row in df_sorted.iterrows():
        st.success(
            f"**{row['IOL']}** ({row['Category']}) — "
            f"Satisfaction: {row['Satisfaction']} | DOF: {row['DOF (D)']:.2f}D | "
            f"MTF: {row['MTF max']:.2f} | Dysphotopsia: {row['Dysphotopsia (/10)']:.1f}/10\n\n"
            f"*{row['DOF Grade']}*"
        )

    st.markdown("---")
    st.caption(
        "⚠️ This software is a clinical decision-support tool. "
        "All recommendations should be interpreted by a qualified ophthalmologist "
        "in the context of a full clinical assessment. "
        "Scores are predictive models and do not substitute clinical judgment."
    )

if __name__ == "__main__":
    main()
