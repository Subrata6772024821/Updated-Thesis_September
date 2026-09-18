"""
Compare calibration accuracy (MAD, MPD) WITHOUT DTW vs WITH DTW temporal
alignment, across the 4 density proxies (Flow_SMS, Flow_Snapshot,
Flow_STOPLine, Occupancy) and the 3 calibration methods (Mean Multiplier,
2nd-Degree Poly, GPR) -- plus a "Before Calibration" baseline bar (raw
proxy vs. reference, no calibration applied at all) in every group.

Reads two workbooks:
  - the "Combined Density ... AllMethods.xlsx" file (no DTW)
  - the "AllMethods With DTW ...xlsx" file (DTW-aligned)

and produces two grouped bar charts (MAD, MPD), each with 8 x-axis groups
(4 proxies x 2 conditions: "No DTW" / "With DTW"), 4 bars per group
(Before Calibration, Mean Multiplier, 2nd Degree Poly, GPR).

Every value is computed directly from raw data / stored literal
hyperparameters -- nothing depends on Excel having cached a formula's
result, so this works regardless of the workbooks' recalculation state.

Run this directly in PyCharm. Only FILE_PATH_NO_DTW / FILE_PATH_WITH_DTW
below need editing -- everything else (date, time window, aggregation
interval, output folder) is auto-detected.
"""

import os
import re
import numpy as np
import matplotlib.pyplot as plt
import openpyxl

# ----------------------------------------------------------------------
# CONFIG — edit these two paths for a different session/resolution.
# ----------------------------------------------------------------------
FILE_PATH_NO_DTW = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\30_sec\05_12_25_1315_1345_Combined Density_30sec_AllMethods.xlsx"
FILE_PATH_WITH_DTW = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\30_sec\05_12_25_1315_1345_AllMethods_With DTW_30sec.xlsx"

# Optional manual overrides -- leave as None to auto-detect.
SESSION_DATE_OVERRIDE = None       # e.g. "05 December, 2025"
SESSION_TIME_OVERRIDE = None       # e.g. "13:15 to 13:45"
AGGREGATION_OVERRIDE = None        # e.g. "60 sec Aggregation"
OUTPUT_DIR_OVERRIDE = None         # default: parent of the 30_sec/60_sec folder
OUTPUT_SUFFIX_OVERRIDE = None      # default: auto-detected, e.g. "30sec"

# Sheet-name maps: proxy label (x-axis) -> {method label: sheet name}
PROXY_SHEETS_NO_DTW = {
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

PROXY_SHEETS_WITH_DTW = {
    "Acc - Flow_SMS": {
        "Mean Multiplier": "M. Mult Acc-Flow_SMS_DTW",
        "2nd Degree Poly":  "2nd deg Acc-Flow_SMS_DTW",
        "GPR":              "GPR Acc_Flow_SMS_DTW",
    },
    "Acc - Flow_SNAPSHOT": {
        "Mean Multiplier": "M. Mult Acc-Flow_Snap_DTW",
        "2nd Degree Poly":  "2nd deg Acc-Flow_Snap_DTW",
        "GPR":              "GPR Acc_Flow_Snap_DTW",
    },
    "Acc - Flow_STOPLINE": {
        "Mean Multiplier": "M. Mult Acc-STOPLine_DTW",
        "2nd Degree Poly":  "2nd deg Acc-STOPLine_DTW",
        "GPR":              "GPR Acc_STOPLine_DTW",
    },
    "Acc - Occupancy": {
        "Mean Multiplier": "M. Mult Acc-Occ_DTW",
        "2nd Degree Poly":  "2nd deg Acc-Occ_DTW",
        "GPR":              "GPR Acc_Occ_DTW",
    },
}

# "Before Calibration" is always first so it's the leftmost bar in each group.
METHOD_ORDER = ["Before Calibration", "Mean Multiplier", "2nd Degree Poly", "GPR"]
METHOD_COLORS = {
    "Before Calibration": "#8B1A1A",  # crimson -- baseline, not a calibration method
    "Mean Multiplier":    "#C8A84B",  # gold
    "2nd Degree Poly":    "#0D9488",  # teal
    "GPR":                "#1B3A5C",  # navy
}
CONDITION_ORDER = ["No DTW", "With DTW"]


# ----------------------------------------------------------------------
# AUTO-DETECTION: session date/time (from filename) and aggregation (from data)
# ----------------------------------------------------------------------
_SESSION_NAME_RE = re.compile(r"(\d{2})_(\d{2})_(\d{2})_(\d{4})_(\d{4})")
_MONTH_NAMES = [
    "", "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def detect_session_date_time(path):
    match = _SESSION_NAME_RE.search(os.path.basename(path)) or _SESSION_NAME_RE.search(path)
    if not match:
        return None, None
    dd, mm, yy, start, end = match.groups()
    try:
        month_name = _MONTH_NAMES[int(mm)]
        date_str = f"{dd} {month_name}, 20{yy}"
        time_str = f"{start[:2]}:{start[2:]} to {end[:2]}:{end[2:]}"
        return date_str, time_str
    except (IndexError, ValueError):
        return None, None


def detect_aggregation(ws, avg_row):
    """Time-bin midpoints are always in column B (2) on these sheets: 'Time (s)'
    on the plain workbook, 'Ref_Time_s' on the DTW workbook."""
    times = []
    for r in range(2, avg_row):
        v = ws.cell(row=r, column=2).value
        if isinstance(v, (int, float)):
            times.append(v)
    if len(times) < 2:
        return None, None
    diffs = np.diff(sorted(set(times)))
    diffs = diffs[diffs > 0]
    interval = float(np.median(diffs))
    seconds = int(round(interval))
    if seconds % 60 == 0 and seconds >= 60:
        minutes = seconds // 60
        return f"{minutes} min Aggregation", f"{minutes}min"
    return f"{seconds} sec Aggregation", f"{seconds}sec"


# ----------------------------------------------------------------------
# EXTRACTION (generalized to handle both the plain and DTW-aligned column
# naming conventions; everything computed from raw data + literal GPR
# hyperparameters, never from Excel's cached formula results)
# ----------------------------------------------------------------------
def _normalize(text):
    return "".join(str(text).strip().lower().split()).replace("-", "_")


def _find_averages_row(ws):
    for col in ("A", "B", "H"):
        for cell in ws[col]:
            if isinstance(cell.value, str) and "average" in cell.value.lower():
                return cell.row
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and "average" in cell.value.lower():
                return cell.row
    raise ValueError(f"No 'Averages' row found in sheet '{ws.title}'")


def _header_map(ws):
    out = {}
    for cell in ws[1]:
        if isinstance(cell.value, str) and cell.value.strip():
            out[cell.column] = _normalize(cell.value)
    return out


def _find_col(headers, candidate_rules):
    """Try each (must_contain[, must_not_contain]) rule in order; return the
    first single unambiguous column match."""
    for rule in candidate_rules:
        must = rule[0]
        must_not = rule[1] if len(rule) > 1 else ()
        matches = [
            col for col, h in headers.items()
            if all(s in h for s in must) and not any(s in h for s in must_not)
        ]
        if len(matches) == 1:
            return matches[0]
    raise ValueError(
        f"No unambiguous column found for rules {candidate_rules}. "
        f"Headers were: {list(headers.values())}"
    )


def _raw_columns(ws, expect_dtw):
    """Return (reference_col, proxy_col, avg_row): the reference/ideal density
    column, the raw proxy density column, and the 'Averages' row index
    (data lives in rows 2..avg_row-1).

    `expect_dtw` pins which header convention this sheet MUST use -- it does
    NOT fall back to the other convention. Mixing the two silently (e.g.
    matching a leftover 'Density_Accumulation' column inside a DTW sheet
    that also has 'Reference_Value'/'Aligned_Proxy_Value') is exactly what
    caused the DTW and No-DTW 'Before Calibration' bars to come out
    identical: both ended up reading the same un-aligned raw pair.
    """
    headers = _header_map(ws)
    if expect_dtw:
        reference_col = _find_col(headers, [(["reference", "value"],)])
        proxy_col = _find_col(headers, [(["aligned", "proxy", "value"],)])
    else:
        reference_col = _find_col(headers, [(["density", "accumulation"],)])
        proxy_col = _find_col(headers, [
            (["density"], ["accumulation", "modified", "gpr", "posterior"]),
        ])
    avg_row = _find_averages_row(ws)
    return reference_col, proxy_col, avg_row


def _read_raw_series(ws, ref_col, proxy_col, avg_row):
    c_vals, d_vals = [], []
    for r in range(2, avg_row):
        c = ws.cell(row=r, column=ref_col).value
        d = ws.cell(row=r, column=proxy_col).value
        if isinstance(c, (int, float)) and isinstance(d, (int, float)):
            c_vals.append(c)
            d_vals.append(d)
    return np.array(c_vals, dtype=float), np.array(d_vals, dtype=float)


def _before_stats(c, d):
    """Baseline accuracy of the raw proxy against the reference, with no
    calibration (no lambda scaling, no polynomial, no GPR posterior) applied."""
    ad = np.abs(c - d)
    pd_ = ad / c * 100
    return float(np.mean(ad)), float(np.mean(pd_))


def _mean_multiplier_after_stats(c, d):
    ratio = d / c
    lam = np.mean(ratio)
    modified = d / lam
    ad = np.abs(modified - c)
    pd_ = ad / c * 100
    return float(np.mean(ad)), float(np.mean(pd_))


_POLY_FORMULA_RE = re.compile(
    r"=\s*([-\d.eE]+)\s*\+\s*\(([-\d.eE]+)\)\s*\*\s*\w+\d+\^1\s*\+\s*\(([-\d.eE]+)\)\s*\*\s*\w+\d+\^2"
)


def _poly_after_stats(ws_formulas, c, d):
    headers = _header_map(ws_formulas)
    mod_col = _find_col(headers, [(["modified", "density"], ["gpr"])])
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


def _gpr_after_stats(ws, ref_col, proxy_col, avg_row):
    c, d = _read_raw_series(ws, ref_col, proxy_col, avg_row)

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


def load_after_stats(path, proxy_sheets, is_dtw, before_stats=None):
    """Returns {proxy: {method: {"MAD": x, "MPD": y}}} for one workbook.

    `is_dtw` must correctly identify whether `path`/`proxy_sheets` is the
    DTW-aligned workbook or the plain one -- it determines which header
    convention _raw_columns() is allowed to match (see its docstring).

    `before_stats`, if given, is a {proxy: {"MAD": x, "MPD": y}} dict of
    ALREADY-COMPUTED "Before Calibration" baselines (see note below) that
    gets reused as-is instead of being recomputed from this workbook.
    """
    wb_values = openpyxl.load_workbook(path, data_only=True)
    wb_formulas = openpyxl.load_workbook(path, data_only=False)
    results = {}
    computed_before = {}

    for proxy, methods in proxy_sheets.items():
        results[proxy] = {}

        mm_ws = wb_values[methods["Mean Multiplier"]]
        ref_col, proxy_col, avg_row = _raw_columns(mm_ws, is_dtw)
        c, d = _read_raw_series(mm_ws, ref_col, proxy_col, avg_row)

        # "Before Calibration" reflects the raw, unaligned proxy vs. reference
        # pair -- it exists BEFORE DTW is even applied, so it must be the same
        # value whether or not the calibration methods afterward used DTW
        # alignment. We only ever compute it from the plain (non-DTW) workbook
        # and reuse that single value for both conditions -- never recompute
        # it from the DTW-aligned columns, which would (incorrectly) give a
        # different, already-time-aligned baseline.
        if before_stats is not None:
            results[proxy]["Before Calibration"] = before_stats[proxy]
        else:
            mad, mpd = _before_stats(c, d)
            results[proxy]["Before Calibration"] = {"MAD": mad, "MPD": mpd}
            computed_before[proxy] = results[proxy]["Before Calibration"]

        mad, mpd = _mean_multiplier_after_stats(c, d)
        results[proxy]["Mean Multiplier"] = {"MAD": mad, "MPD": mpd}

        poly_ws_values = wb_values[methods["2nd Degree Poly"]]
        poly_ws_formulas = wb_formulas[methods["2nd Degree Poly"]]
        p_ref_col, p_proxy_col, p_avg_row = _raw_columns(poly_ws_values, is_dtw)
        pc, pd_ = _read_raw_series(poly_ws_values, p_ref_col, p_proxy_col, p_avg_row)
        mad, mpd = _poly_after_stats(poly_ws_formulas, pc, pd_)
        results[proxy]["2nd Degree Poly"] = {"MAD": mad, "MPD": mpd}

        gpr_ws = wb_values[methods["GPR"]]
        g_ref_col, g_proxy_col, g_avg_row = _raw_columns(gpr_ws, is_dtw)
        mad, mpd = _gpr_after_stats(gpr_ws, g_ref_col, g_proxy_col, g_avg_row)
        results[proxy]["GPR"] = {"MAD": mad, "MPD": mpd}

    return results, computed_before


# ----------------------------------------------------------------------
# PLOTTING
# ----------------------------------------------------------------------
def plot_metric(no_dtw, with_dtw, metric_key, metric_label, header_line, save_path):
    proxies = list(no_dtw.keys())
    condition_data = {"No DTW": no_dtw, "With DTW": with_dtw}

    # x groups: (proxy, condition) pairs, in order, with extra gap between proxies
    group_labels = []
    group_gap_after = []  # True = insert extra spacing after this group
    for i, proxy in enumerate(proxies):
        for j, cond in enumerate(CONDITION_ORDER):
            group_labels.append(f"{proxy}\n{cond}")
            group_gap_after.append(j == len(CONDITION_ORDER) - 1 and i != len(proxies) - 1)

    n_groups = len(group_labels)
    n_methods = len(METHOD_ORDER)
    bar_width = 0.8 / n_methods

    # compute x positions with extra gap between proxy pairs
    x_positions = []
    pos = 0.0
    for i in range(n_groups):
        x_positions.append(pos)
        pos += 1.0
        if group_gap_after[i]:
            pos += 0.6
    x_positions = np.array(x_positions)

    # Wider canvas now that there are 4 proxies (8 groups) instead of 3 (6 groups)
    fig, ax = plt.subplots(figsize=(18, 6.5))

    for m_idx, method in enumerate(METHOD_ORDER):
        offset = (m_idx - (n_methods - 1) / 2) * bar_width
        values = []
        for proxy in proxies:
            for cond in CONDITION_ORDER:
                values.append(condition_data[cond][proxy][method][metric_key])
        bars = ax.bar(
            x_positions + offset, values, bar_width,
            label=method, color=METHOD_COLORS[method], edgecolor="white", linewidth=0.5,
        )
        ax.bar_label(bars, fmt="%.2f", padding=2, fontsize=7)

    fig.suptitle(header_line, fontsize=11, y=0.98)
    ax.set_title(f"{metric_key} Comparison: No DTW vs. With DTW Temporal Alignment", fontsize=12, pad=12)
    ax.set_xlabel("Density Proxy and DTW Condition")
    ax.set_ylabel(metric_label)
    ax.set_xticks(x_positions)
    ax.set_xticklabels(group_labels)
    ax.legend(title="Calibration Method", loc="upper right", framealpha=0.9)
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    ax.margins(y=0.15)
    fig.tight_layout(rect=[0, 0, 1, 0.94])
    fig.savefig(save_path, dpi=300)
    print(f"Saved: {save_path}")


def main():
    # Compute "Before Calibration" once, from the plain (non-DTW) workbook,
    # then reuse that same baseline for the "With DTW" condition too --
    # see the note in load_after_stats() for why it must not be recomputed
    # from the DTW-aligned columns.
    no_dtw, before_baseline = load_after_stats(FILE_PATH_NO_DTW, PROXY_SHEETS_NO_DTW, is_dtw=False)
    with_dtw, _ = load_after_stats(
        FILE_PATH_WITH_DTW, PROXY_SHEETS_WITH_DTW, is_dtw=True, before_stats=before_baseline,
    )

    session_date, session_time = SESSION_DATE_OVERRIDE, SESSION_TIME_OVERRIDE
    if session_date is None or session_time is None:
        d, t = detect_session_date_time(FILE_PATH_WITH_DTW)
        session_date = session_date or d or "Unknown date"
        session_time = session_time or t or "Unknown time"

    agg_label, agg_suffix = AGGREGATION_OVERRIDE, OUTPUT_SUFFIX_OVERRIDE
    if agg_label is None or agg_suffix is None:
        wb_probe = openpyxl.load_workbook(FILE_PATH_NO_DTW, data_only=True)
        probe_ws = wb_probe[next(iter(PROXY_SHEETS_NO_DTW.values()))["Mean Multiplier"]]
        _, _, probe_avg_row = _raw_columns(probe_ws, expect_dtw=False)
        detected_label, detected_suffix = detect_aggregation(probe_ws, probe_avg_row)
        agg_label = agg_label or detected_label or "Unknown Aggregation"
        agg_suffix = agg_suffix or detected_suffix or "unk"

    header_line = f"Date: {session_date}   |   Time: {session_time}   |   {agg_label}"

    output_dir = OUTPUT_DIR_OVERRIDE or os.path.dirname(os.path.dirname(FILE_PATH_WITH_DTW))
    os.makedirs(output_dir, exist_ok=True)

    plot_metric(
        no_dtw, with_dtw,
        metric_key="MAD", metric_label="MAD",
        header_line=header_line,
        save_path=os.path.join(output_dir, f"MAD_DTW_vs_NoDTW_{agg_suffix}.png"),
    )
    plot_metric(
        no_dtw, with_dtw,
        metric_key="MPD", metric_label="MPD (%)",
        header_line=header_line,
        save_path=os.path.join(output_dir, f"MPD_DTW_vs_NoDTW_{agg_suffix}.png"),
    )

    plt.show()


if __name__ == "__main__":
    main()