"""
Plot MAD and MPD bar charts (before calibration vs. Mean Multiplier, 2nd-degree
polynomial, and GPR calibration) from the combined density workbook.

Reads all 12 sheets of *_Combined_Density_30sec_AllMethods.xlsx and produces two
grouped bar charts (one for MAD, one for MPD), each comparing:
    Before Calibration | Mean Multiplier | 2nd Degree Poly | GPR
across the four density proxies: Flow_SMS, Flow_Snapshot, Flow_STOPLine,
Occupancy.

Run this directly in PyCharm. Only FILE_PATH / OUTPUT_DIR below need editing.
"""

import os
import numpy as np
import matplotlib.pyplot as plt
import openpyxl

# ----------------------------------------------------------------------
# CONFIG — edit these if your file/session changes
# ----------------------------------------------------------------------
FILE_PATH = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\120_sec\PCU\05_12_25_1315_1345_Combined_Density_120sec_T1_T2_Pooled_AllMethods.xlsx"

# Images are saved to the PARENT of the 30_sec folder (i.e. the session folder)
OUTPUT_DIR = os.path.dirname(r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\120_sec\PCU")

# Session date/time/aggregation, shown above each chart
SESSION_DATE = "05 December, 2025"
SESSION_TIME = "13:15 to 13:45"
AGGREGATION_LABEL = "120 sec Aggregation (T1 & T2 Combined)"   # e.g. "30 sec Aggregation" / "60 sec Aggregation" / "2 min Aggregation"
OUTPUT_SUFFIX = "120sec"                     # used in the saved PNG filenames

# Sheet-name map: proxy label (as shown on the x-axis) -> {method label: sheet name}
PROXY_SHEETS = {
    "Acc - Flow_SMS": {
        "Mean Multiplier": "M. Mult Acc-Flow_SMS",
        "2nd Degree Poly":  "2nd deg Acc-Flow_SMS",
        "GPR":              "GPR Acc_Flow_SMS",
    },
    "Acc - Flow_SNAPSHOT": {
        "Mean Multiplier": "M. Mult Acc-Flow_Snap",
        "2nd Degree Poly":  "2nd deg Acc-Flow_Snap",
        "GPR":              "GPR Acc_Flow_Snap",
    },
    "Acc - Flow_STOPLINE": {
        "Mean Multiplier": "M. Mult Acc-Flow_STOPLine",
        "2nd Degree Poly":  "2nd deg Acc-Flow_STOPLine",
        "GPR":              "GPR Acc_Flow_STOPLine",
    },
    "Acc - Occupancy": {
        "Mean Multiplier": "M. Mult Acc-Occ",
        "2nd Degree Poly":  "2nd deg Acc-Occ",
        "GPR":              "GPR Acc_Occ",
    },
}

# Chula-branded palette (matches thesis defense deck)
COLORS = {
    "Before Calibration": "#8B1A1A",  # crimson
    "Mean Multiplier":    "#C8A84B",  # gold
    "2nd Degree Poly":    "#0D9488",  # teal
    "GPR":                "#1B3A5C",  # navy
}
METHOD_ORDER = ["Before Calibration", "Mean Multiplier", "2nd Degree Poly", "GPR"]


# ----------------------------------------------------------------------
# EXTRACTION
# ----------------------------------------------------------------------
# Design note: this workbook's MAD/MPD "Averages" row are Excel FORMULAS
# (e.g. =AVERAGE(G2:G62)), not stored values. If the workbook was last saved
# without a full Excel recalculation (e.g. written/edited programmatically),
# openpyxl's data_only=True read returns None for those cells. To avoid
# depending on Excel's cache, MAD_Before/MPD_Before and the Mean-Multiplier /
# 2nd-degree "After" values are recomputed here directly from the two raw
# columns (Density_Accumulation and the raw proxy density) using the exact
# same formulas as the sheet. Only the GPR "After" values (a GP posterior
# mean via array formulas) are too complex to cheaply reproduce, so those
# still read the sheet's cached result — with a clear error if it's missing.

import re


def _normalize(text):
    """Lowercase, strip, and collapse whitespace/underscore differences for matching."""
    return "".join(str(text).strip().lower().split()).replace("-", "_")


def _find_averages_row(ws):
    """Find the row index of the 'Averages ->' label, checked column-by-column
    (label usually sits in column A, but falls back to a full-sheet scan)."""
    for col in ("A", "B"):
        for cell in ws[col]:
            if isinstance(cell.value, str) and "average" in cell.value.lower():
                return cell.row
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and "average" in cell.value.lower():
                return cell.row
    raise ValueError(f"No 'Averages' row found in sheet '{ws.title}'")


def _header_map(ws):
    """Map column index -> normalized header text from row 1."""
    out = {}
    for cell in ws[1]:
        if isinstance(cell.value, str) and cell.value.strip():
            out[cell.column] = _normalize(cell.value)
    return out


def _find_col(headers, must_contain, must_not_contain=()):
    """Find the single column index whose normalized header contains all of
    `must_contain` and none of `must_not_contain`."""
    matches = [
        col for col, h in headers.items()
        if all(s in h for s in must_contain) and not any(s in h for s in must_not_contain)
    ]
    if len(matches) != 1:
        raise ValueError(
            f"Expected exactly one column matching {must_contain} (excluding {must_not_contain}), "
            f"found {len(matches)}: {matches}. Headers were: {list(headers.values())}"
        )
    return matches[0]


def _raw_columns(ws):
    """Return (accum_col, proxy_col, avg_row) for a sheet: the Density_Accumulation
    (reference/ideal) column, the raw proxy density column, and the row index
    of the 'Averages' summary label (data lives in rows 2..avg_row-1)."""
    headers = _header_map(ws)
    accum_col = _find_col(headers, ["density", "accumulation"])
    proxy_col = _find_col(
        headers, ["density"],
        must_not_contain=["accumulation", "modified", "gpr", "posterior"],
    )
    avg_row = _find_averages_row(ws)
    return accum_col, proxy_col, avg_row


def _read_raw_series(ws, accum_col, proxy_col, avg_row):
    """Read (C, D) as numpy arrays over the data rows (2..avg_row-1)."""
    c_vals, d_vals = [], []
    for r in range(2, avg_row):
        c = ws.cell(row=r, column=accum_col).value
        d = ws.cell(row=r, column=proxy_col).value
        if isinstance(c, (int, float)) and isinstance(d, (int, float)):
            c_vals.append(c)
            d_vals.append(d)
    return np.array(c_vals, dtype=float), np.array(d_vals, dtype=float)


def _before_stats(c, d):
    ad = np.abs(c - d)
    pd_ = ad / c * 100
    return float(np.mean(ad)), float(np.mean(pd_))


def _mean_multiplier_after_stats(c, d):
    ratio = d / c
    lam = np.mean(ratio)          # matches Excel's E63 = AVERAGE(E2:E62)
    modified = d / lam            # matches F column
    ad = np.abs(modified - c)
    pd_ = ad / c * 100
    return float(np.mean(ad)), float(np.mean(pd_))


_POLY_FORMULA_RE = re.compile(
    r"=\s*([-\d.eE]+)\s*\+\s*\(([-\d.eE]+)\)\s*\*\s*\w+\d+\^1\s*\+\s*\(([-\d.eE]+)\)\s*\*\s*\w+\d+\^2"
)


def _poly_after_stats(ws_values, ws_formulas, c, d):
    """Parse the 2nd-degree fit coefficients straight from row-2's formula text
    (e.g. '=66.78 + (0.60)*D2^1 + (0.005)*D2^2') and reapply to the full column."""
    headers = _header_map(ws_formulas)
    mod_col = _find_col(headers, ["modified", "density"], must_not_contain=["gpr"])
    formula = ws_formulas.cell(row=2, column=mod_col).value
    if not isinstance(formula, str):
        raise ValueError(
            f"Expected a formula string in row 2 of the modified-density column "
            f"on sheet '{ws_formulas.title}', got: {formula!r}"
        )
    match = _POLY_FORMULA_RE.search(formula.replace(" ", ""))
    if not match:
        raise ValueError(
            f"Could not parse 2nd-degree polynomial coefficients from formula "
            f"on sheet '{ws_formulas.title}': {formula!r}"
        )
    a, b, cf = (float(x) for x in match.groups())
    predicted = a + b * d + cf * d**2
    ad = np.abs(c - predicted)
    pd_ = ad / c * 100
    return float(np.mean(ad)), float(np.mean(pd_))


def _find_hyperparam(ws, *label_variants):
    """Scan column L for a label matching any of `label_variants` (normalized:
    case/whitespace/underscore-insensitive) and return the literal numeric
    value from the same row in column M."""
    targets = {_normalize(v) for v in label_variants}
    for cell in ws["L"]:
        if isinstance(cell.value, str):
            norm = _normalize(cell.value)
            if any(t in norm for t in targets):
                val = ws.cell(row=cell.row, column=13).value  # column M
                if isinstance(val, (int, float)):
                    return float(val)
    raise ValueError(
        f"Sheet '{ws.title}': could not find a numeric hyperparameter for "
        f"any of {label_variants} in column L."
    )


def _gpr_after_stats(ws, accum_col, proxy_col, avg_row):
    """Reproduce the GP posterior mean directly in Python (closed-form RBF +
    white-noise kernel regression) from the sheet's raw data and its fitted
    hyperparameters (Sigma_f^2, Length_scale, Noise_variance -- all stored as
    literal values, not formulas). This avoids depending on Excel's cached
    array-formula results entirely, so it works even if the workbook has
    never been recalculated/saved in Excel."""
    c, d = _read_raw_series(ws, accum_col, proxy_col, avg_row)

    sigma_f2 = _find_hyperparam(ws, "signal variance", "sigma_f^2")
    length_scale = _find_hyperparam(ws, "length_scale", "length scale")
    noise_var = _find_hyperparam(ws, "noise_variance", "noise variance")

    diff = d[:, None] - d[None, :]
    K = sigma_f2 * np.exp(-0.5 * (diff / length_scale) ** 2)
    K_noise = K + noise_var * np.eye(len(d))
    alpha = np.linalg.solve(K_noise, c)
    posterior_mean = K @ alpha

    ad = np.abs(c - posterior_mean)
    pd_ = ad / c * 100
    return float(np.mean(ad)), float(np.mean(pd_))


def load_all_results(path):
    """
    Returns:
        after   : {proxy: {method: {"MAD_After": x, "MPD_After": y}}}
        before  : {proxy: {"MAD_Before": x, "MPD_Before": y}}

    Every value here is computed directly from raw data / stored literal
    hyperparameters -- none of it depends on Excel having cached a formula's
    result, so this works regardless of the workbook's recalculation state.
    """
    # data_only=True: raw numeric columns + literal hyperparameter cells.
    # data_only=False: same file, needed only to read the 2nd-degree fit's
    # formula text (coefficients aren't stored anywhere else in the sheet).
    wb_values = openpyxl.load_workbook(path, data_only=True)
    wb_formulas = openpyxl.load_workbook(path, data_only=False)
    after, before = {}, {}

    for proxy, methods in PROXY_SHEETS.items():
        after[proxy] = {}

        mm_ws = wb_values[methods["Mean Multiplier"]]
        accum_col, proxy_col, avg_row = _raw_columns(mm_ws)
        c, d = _read_raw_series(mm_ws, accum_col, proxy_col, avg_row)

        mad_before, mpd_before = _before_stats(c, d)
        before[proxy] = {"MAD_Before": mad_before, "MPD_Before": mpd_before}

        mad_after, mpd_after = _mean_multiplier_after_stats(c, d)
        after[proxy]["Mean Multiplier"] = {"MAD_After": mad_after, "MPD_After": mpd_after}

        poly_ws_values = wb_values[methods["2nd Degree Poly"]]
        poly_ws_formulas = wb_formulas[methods["2nd Degree Poly"]]
        p_accum_col, p_proxy_col, p_avg_row = _raw_columns(poly_ws_values)
        pc, pd_ = _read_raw_series(poly_ws_values, p_accum_col, p_proxy_col, p_avg_row)
        mad_after, mpd_after = _poly_after_stats(poly_ws_values, poly_ws_formulas, pc, pd_)
        after[proxy]["2nd Degree Poly"] = {"MAD_After": mad_after, "MPD_After": mpd_after}

        gpr_ws = wb_values[methods["GPR"]]
        gpr_accum_col, gpr_proxy_col, gpr_avg_row = _raw_columns(gpr_ws)
        mad_after, mpd_after = _gpr_after_stats(gpr_ws, gpr_accum_col, gpr_proxy_col, gpr_avg_row)
        after[proxy]["GPR"] = {"MAD_After": mad_after, "MPD_After": mpd_after}

    return after, before


# ----------------------------------------------------------------------
# PLOTTING
# ----------------------------------------------------------------------
def plot_metric(after, before, metric_key, metric_label, save_path):
    proxies = list(after.keys())
    n_methods = len(METHOD_ORDER)
    x = np.arange(len(proxies))
    bar_width = 0.8 / n_methods

    # Wider canvas now that there are 4 proxy groups instead of 3
    fig, ax = plt.subplots(figsize=(13, 6.5))

    for i, method in enumerate(METHOD_ORDER):
        if method == "Before Calibration":
            values = [before[p][f"{metric_key}_Before"] for p in proxies]
        else:
            values = [after[p][method][f"{metric_key}_After"] for p in proxies]

        offset = (i - (n_methods - 1) / 2) * bar_width
        bars = ax.bar(
            x + offset, values, bar_width,
            label=method, color=COLORS[method], edgecolor="white", linewidth=0.5,
        )
        ax.bar_label(bars, fmt="%.2f", padding=2, fontsize=8, rotation=0)

    # Date/time line above the chart title
    fig.suptitle(f"Date: {SESSION_DATE}   |   Time: {SESSION_TIME}   |   {AGGREGATION_LABEL}", fontsize=11, y=0.98)
    ax.set_title(f"{metric_key} Comparison: Before Calibration vs. Mean Multiplier, 2nd-Degree, GPR", fontsize=12, pad=12)

    ax.set_xlabel("Calibration Method Comparison")
    ax.set_ylabel(metric_label)
    ax.set_xticks(x)
    ax.set_xticklabels(proxies)
    ax.legend(title="Calibration Stage", loc="upper right", framealpha=0.9)
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    ax.margins(y=0.15)  # headroom so value labels don't get clipped
    fig.tight_layout(rect=[0, 0, 1, 0.94])
    fig.savefig(save_path, dpi=300)
    print(f"Saved: {save_path}")


def main():
    after, before = load_all_results(FILE_PATH)

    plot_metric(
        after, before,
        metric_key="MAD",
        metric_label="MAD",  # no unit
        save_path=os.path.join(OUTPUT_DIR, f"MAD_comparison_Combined{OUTPUT_SUFFIX}.png"),
    )

    plot_metric(
        after, before,
        metric_key="MPD",
        metric_label="MPD (%)",
        save_path=os.path.join(OUTPUT_DIR, f"MPD_comparison_Combined{OUTPUT_SUFFIX}.png"),
    )

    plt.show()


if __name__ == "__main__":
    main()