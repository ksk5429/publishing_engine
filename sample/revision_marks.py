"""
Revision markers for the sample manuscript.

Maps reviewer comment IDs to keyword phrases. Any paragraph containing
one of these phrases will be rendered in RED in the revised manuscript,
making it easy for reviewers to find addressed changes.
"""

REVISION_MARKS = {
    "R1.1": ["three key findings", "sign-consistent binary separation"],
    "R1.2": ["fractional frequency reduction", "calibration is valid for"],
    "R1.3": ["70g", "1:70"],
    "R1.4": ["parked-state", "operational loading", "rotor-induced"],
    "R2.1": ["PCA", "baseline comparison", "cointegration"],
    "R2.2": ["+55.8%", "loss of lateral restraint", "strain distribution"],
    "R2.3": ["two to five", "0.05D", "Jalbi", "Weil"],
    "R2.4": ["reproducibility", "PyPI", "Zenodo", "op3-framework"],
}
