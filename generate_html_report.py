"""
Regenerate TMC_Vision_Accuracy_Report.html from report_template.html and the
tidy CSVs/stats produced by analyze_tmc.py (Manual-vs-Vision: tmc_Baseline.csv,
tmc_1st_Calibration.csv, tmc_stats.json) and consultant_data.py (Consultant-vs-
Vision: consultant_data.csv, consultant_stats.json).

Run analyze_tmc.py / consultant_data.py first if those files are missing or the
source workbooks (NewBaseline.xlsm / TMCValidation.xlsm) have changed.
"""

import json
import pandas as pd

TEMPLATE_PATH = "report_template.html"
STATS_PATH = "tmc_stats.json"
CONSULTANT_STATS_PATH = "consultant_stats.json"
OUT_PATH = "TMC_Vision_Accuracy_Report.html"

MANUAL_VS_VISION_FILES = {"Baseline": "tmc_Baseline.csv", "1st Calibration": "tmc_1st_Calibration.csv"}
CONSULTANT_FILE = "consultant_data.csv"


def main():
    data = {name: pd.read_csv(path).to_dict(orient="records")
            for name, path in MANUAL_VS_VISION_FILES.items()}
    consultant_data = pd.read_csv(CONSULTANT_FILE).to_dict(orient="records")

    with open(STATS_PATH) as f:
        stats = json.load(f)
    with open(CONSULTANT_STATS_PATH) as f:
        consultant_stats = json.load(f)
    with open(TEMPLATE_PATH) as f:
        html = f.read()

    html = html.replace("__DATA_JSON__", json.dumps(data))
    html = html.replace("__STATS_JSON__", json.dumps(stats))
    html = html.replace("__CONSULTANT_DATA_JSON__", json.dumps(consultant_data))
    html = html.replace("__CONSULTANT_STATS_JSON__", json.dumps(consultant_stats))
    html = html.replace("__N_BASELINE__", str(stats["Baseline"]["n_intersections"]))
    html = html.replace("__N_CALIBRATION__", str(stats["1st Calibration"]["n_intersections"]))
    html = html.replace("__N_CONSULTANT__", str(consultant_stats["n_intersections"]))

    with open(OUT_PATH, "w") as f:
        f.write(html)
    print(f"Wrote {OUT_PATH} ({len(html)} bytes)")


if __name__ == "__main__":
    main()
