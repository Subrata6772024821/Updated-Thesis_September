"""
Time-series comparison of actual (reference) density vs. calibrated density,
for each density proxy (Flow_SMS, Flow_Snapshot, Flow_STOPLine, Occupancy),
in two versions:

    1. No DTW      : actual density vs. Mean Multiplier / 2nd-Degree Poly / GPR
                     calibrated density (calibration fit directly on the raw,
                     unaligned proxy data).
    2. With DTW    : actual density vs. DTW + Mean Multiplier / DTW + 2nd-Degree
                     Poly / DTW + GPR calibrated density (calibration fit on the
                     DTW time-aligned proxy data).

Each of the 8 resulting charts (4 proxies x 2 conditions) is a scatter of the
raw time-series points with a smooth interpolating curve through each series:
Actual, Mean Multiplier, 2nd Degree Poly, GPR.

IMPORTANT -- why the calibrated series are recomputed, not read from the sheet:
the "Modified Density" / "MAD" / "MPD" columns on these sheets are Excel
FORMULAS (including array formulas for GPR). If the workbook was last saved
without Excel recalculating (e.g. written/edited programmatically), openpyxl's
data_only=True read returns None for those cells. To avoid depending on
Excel's cache, every calibrated series here is recomputed directly from the
raw reference/proxy columns plus the literal fitted parameters stored on each
sheet (Mean Multiplier's lambda, the 2nd-degree polynomial's coefficients
parsed from its row-2 formula text, and GPR's kernel hyperparameters) --
so this works regardless of the workbook's recalculation state.

Run this directly in PyCharm. Only FILE_PATH_NO_DTW / FILE_PATH_WITH_DTW
below need editing -- everything else (date, time window, aggregation
interval, output folder) is auto-detected.
"""

import os
import re
import numpy as np
import matplotlib.pyplot as plt
import openpyxl
from scipy.interpolate import make_interp_spline

# ----------------------------------------------------------------------
# CONFIG — edit these two paths for a different session/resolution.
# ----------------------------------------------------------------------
FILE_PATH_NO_DTW = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\120_sec\05_12_25_1315_1345_Combined Density_120sec_AllMethods.xlsx"
FILE_PATH_WITH_DTW = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\120_sec\05_12_25_1315_1345_AllMethods_With DTW_120sec.xlsx"

# Optional manual overrides -- leave as None to auto-detect.
SESSION_DATE_OVERRIDE = None       # e.g. "05 December, 2025"
SESSION_TIME_OVERRIDE = None       # e.g. "13:15 to 13:45"
AGGREGATION_OVERRIDE = None        # e.g. "30 sec Aggregation"
OUTPUT_DIR_OVERRIDE = None         # default: parent of the 30_sec/60_sec folder
OUTPUT_SUFFIX_OVERRIDE = None      # default: auto-detected, e.g. "30sec"

# Sheet-name maps: proxy label -> {method label: sheet name}
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

# High-contrast, colorblind-friendly palette -- chosen so all 4 series stay
# easy to tell apart even with lines overlapping.
BASE_COLORS = {
    "Actual":            "#000000",  # black
    "Mean Multiplier":   "#E67E22",  # orange
    "2nd Degree Poly":   "#27AE60",  # green
    "GPR":               "#8E44AD",  # purple
}


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
    """Time-bin midpoints are always in column B (2) on these sheets."""
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
# EXTRACTION
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
    """Return (reference_col, proxy_col, avg_row).

    `expect_dtw` pins which header convention this sheet MUST use (it does
    NOT fall back to the other convention -- mixing the two silently is what
    causes a DTW sheet's un-aligned leftover columns to get matched instead
    of its actual aligned columns)."""
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


def _read_time_c_d(ws, ref_col, proxy_col, avg_row, time_col=2):
    """Read (time, reference, proxy) arrays over the data rows (2..avg_row-1)."""
    t_vals, c_vals, d_vals = [], [], []
    for r in range(2, avg_row):
        t = ws.cell(row=r, column=time_col).value
        c = ws.cell(row=r, column=ref_col).value
        d = ws.cell(row=r, column=proxy_col).value
        if isinstance(t, (int, float)) and isinstance(c, (int, float)) and isinstance(d, (int, float)):
            t_vals.append(t)
            c_vals.append(c)
            d_vals.append(d)
    return np.array(t_vals, float), np.array(c_vals, float), np.array(d_vals, float)


def _mean_multiplier_series(c, d):
    lam = np.mean(d / c)      # matches Excel's Multiplier-column average
    return d / lam            # matches the "Modified Density" column


_POLY_FORMULA_RE = re.compile(
    r"=\s*([-\d.eE]+)\s*\+\s*\(([-\d.eE]+)\)\s*\*\s*\w+\d+\^1\s*\+\s*\(([-\d.eE]+)\)\s*\*\s*\w+\d+\^2"
)


def _poly_coeffs(ws_formulas):
    """Parse the 2nd-degree fit coefficients from row-2's formula text
    (e.g. '=66.78 + (0.60)*D2^1 + (0.005)*D2^2')."""
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
    return tuple(float(x) for x in match.groups())


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


def _gpr_posterior_series(ws, c, d):
    """Reproduce the GP posterior mean directly (closed-form RBF + white-noise
    kernel regression) from the sheet's literal fitted hyperparameters."""
    sigma_f2 = _find_hyperparam(ws, "signal variance", "sigma_f^2")
    length_scale = _find_hyperparam(ws, "length_scale", "length scale")
    noise_var = _find_hyperparam(ws, "noise_variance", "noise variance")

    diff = d[:, None] - d[None, :]
    K = sigma_f2 * np.exp(-0.5 * (diff / length_scale) ** 2)
    K_noise = K + noise_var * np.eye(len(d))
    alpha = np.linalg.solve(K_noise, c)
    return K @ alpha


def load_time_series(path, proxy_sheets, is_dtw):
    """Returns {proxy: {"time": t, "Actual": c, "Mean Multiplier": .., ..}}."""
    wb_values = openpyxl.load_workbook(path, data_only=True)
    wb_formulas = openpyxl.load_workbook(path, data_only=False)
    out = {}

    for proxy, methods in proxy_sheets.items():
        mm_ws = wb_values[methods["Mean Multiplier"]]
        ref_col, proxy_col, avg_row = _raw_columns(mm_ws, is_dtw)
        t, c, d = _read_time_c_d(mm_ws, ref_col, proxy_col, avg_row)
        order = np.argsort(t)
        t, c, d = t[order], c[order], d[order]

        mm_series = _mean_multiplier_series(c, d)

        poly_ws_formulas = wb_formulas[methods["2nd Degree Poly"]]
        a, b, cf = _poly_coeffs(poly_ws_formulas)
        poly_series = a + b * d + cf * d ** 2

        gpr_ws = wb_values[methods["GPR"]]
        gpr_series = _gpr_posterior_series(gpr_ws, c, d)

        out[proxy] = {
            "time": t,
            "Actual": c,
            "Mean Multiplier": mm_series,
            "2nd Degree Poly": poly_series,
            "GPR": gpr_series,
        }
    return out


# ----------------------------------------------------------------------
# PLOTTING
# ----------------------------------------------------------------------
def _smooth(x, y, n=300):
    """Cubic interpolating spline through the (sorted, de-duplicated) points,
    purely for a smooth visual curve connecting the scatter points."""
    if len(x) < 4:
        return x, y
    x_new = np.linspace(x.min(), x.max(), n)
    spline = make_interp_spline(x, y, k=3)
    return x_new, spline(x_new)


def _plot_on_ax(ax, series, condition_label, method_prefix, proxy):
    t = series["time"]
    labels = {
        "Actual":            "Actual Density (Accumulation)",
        "Mean Multiplier":   f"{method_prefix}Mean Multiplier",
        "2nd Degree Poly":   f"{method_prefix}2nd Degree Polynomial",
        "GPR":               f"{method_prefix}GPR",
    }
    for key in ["Actual", "Mean Multiplier", "2nd Degree Poly", "GPR"]:
        y = series[key]
        color = BASE_COLORS[key]
        ax.scatter(t, y, s=16, color=color, alpha=0.55, zorder=3)
        xs, ys = _smooth(t, y)
        ax.plot(xs, ys, color=color, linewidth=2, label=labels[key], zorder=2)

    ax.set_title(f"{proxy} -- {condition_label}", fontsize=12, pad=10)
    ax.set_ylabel("Density (vehicle/km/lane)")
    ax.legend(loc="upper right", framealpha=0.9, fontsize=9)
    ax.grid(alpha=0.3, linestyle="--")
    ax.margins(y=0.1)


def plot_time_series_combined(no_dtw_data, with_dtw_data, proxy, header_line, save_path):
    """One PNG per proxy: top subplot = No DTW, bottom subplot = With DTW,
    sharing the x-axis so the two conditions line up for direct comparison."""
    fig, (ax_top, ax_bottom) = plt.subplots(
        2, 1, figsize=(13, 11), sharex=True,
        gridspec_kw={"hspace": 0.18},
    )

    _plot_on_ax(ax_top, no_dtw_data, "No DTW", method_prefix="", proxy=proxy)
    _plot_on_ax(ax_bottom, with_dtw_data, "With DTW", method_prefix="DTW + ", proxy=proxy)
    ax_bottom.set_xlabel("Time (s)")

    fig.suptitle(header_line, fontsize=11, y=0.995)
    fig.subplots_adjust(top=0.93, bottom=0.06, left=0.07, right=0.98, hspace=0.18)
    fig.savefig(save_path, dpi=300)
    print(f"Saved: {save_path}")


def main():
    no_dtw_series = load_time_series(FILE_PATH_NO_DTW, PROXY_SHEETS_NO_DTW, is_dtw=False)
    with_dtw_series = load_time_series(FILE_PATH_WITH_DTW, PROXY_SHEETS_WITH_DTW, is_dtw=True)

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

    for proxy in PROXY_SHEETS_NO_DTW:
        safe = proxy.replace(" ", "").replace("-", "_")

        plot_time_series_combined(
            no_dtw_series[proxy], with_dtw_series[proxy], proxy,
            header_line=header_line,
            save_path=os.path.join(output_dir, f"TimeSeries_{safe}_Combined_{agg_suffix}.png"),
        )

    plt.show()


if __name__ == "__main__":
    main()