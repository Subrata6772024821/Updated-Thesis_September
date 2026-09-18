"""
Combined Density Calibration -- POOLED across T1 + T2 time windows
====================================================================

This is a variant of the single-file calibration script. Instead of reading
one 30-second source workbook, it reads BOTH the T1 and T2 source workbooks,
concatenates their rows into one pooled dataset (n_T1 + n_T2 rows), and then
fits M1 (mean multiplier), M2 (2nd-degree polynomial), and M3 (GPR) on the
POOLED data -- exactly the same math as before, just with more rows feeding
each fit. This generally gives more stable calibration curves/hyperparameters
than fitting T1 and T2 separately, since each proxy pair now has roughly
double the sample size to estimate its relationship to the ideal
(Accumulation-based) density from.

Each row's Time-range label (column A) is prefixed with "T1:" or "T2:" so you
can still tell, in the output sheets, which source window a given row came
from -- e.g. "T1: 17-30" vs "T2: 17-45". Nothing else about the sheet layout,
formulas, or column structure changes from the single-file version, so the
existing plotting script's column-name lookups ("Density_Accumulation",
"Density_Flow...", "Modified Density...") still work unmodified against this
combined output workbook.

Run in PyCharm:
    pip install openpyxl pandas numpy scikit-learn scipy
"""

import os
from datetime import datetime

import numpy as np
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.formula import ArrayFormula
from scipy import stats
from sklearn.gaussian_process import GaussianProcessRegressor
from sklearn.gaussian_process.kernels import RBF, WhiteKernel, ConstantKernel

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
FILE_PATH_T1 = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\120_sec\PCU\05_12_25_1315_1345_Combined Density_120sec_T1.xlsx"
FILE_PATH_T2 = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\120_sec\PCU\05_12_25_1315_1345_Combined Density_120sec_T2.xlsx"
SOURCE_SHEET = "Sheet1"

IDEAL_COL = "Density_Accumulation (PCU/km)"
PROXY_COLS = {
    "Flow_SMS": "Density_Flow_SMS (PCU/km)",
    "Flow_Snapshot": "Density_Flow_SNAPSHOT (PCU/km)",
    "Flow_STOPLine": "Density_Flow_STOPLINE (PCU/km)",
    "Occ": "Density_Occupancy (vehicle_varieed length/km)",
}

# Output: saved into the PARENT folder of FILE_PATH_T1's own folder, i.e. one
# level above "...\30_sec\", assumed to be the per-interval-run folder
# "...\05_12_25_1315_1345\". Change OUTPUT_PATH directly if that's wrong
# for your folder layout.
_source_dir = os.path.dirname(FILE_PATH_T1)
_parent_dir = os.path.dirname(_source_dir)
OUTPUT_PATH = os.path.join(
    _parent_dir, "05_12_25_1315_1345_Combined_Density_60sec_T1_T2_Pooled_AllMethods.xlsx"
)

N_RESTARTS = 15
RANDOM_STATE = 42

HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
BLUE_FONT = Font(color="0000FF")
BOLD = Font(bold=True)
NOTE_FONT = Font(italic=True, size=9, color="808080")


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------
def _clean_header(h):
    """Normalize a header cell value for matching: strip stray double-quote
    characters (straight and curly) and leading/trailing whitespace.

    This exists because source workbooks in this pipeline have occasionally
    picked up a trailing literal `"` in a header cell (e.g.
    'Density_Flow_SMS (PCU/km)"'), which breaks an exact-string match
    against PROXY_COLS / IDEAL_COL even though the header looks correct
    at a glance in Excel.
    """
    if not isinstance(h, str):
        return h
    cleaned = h.strip()
    for ch in ('"', "\u201c", "\u201d", "\u2033"):
        cleaned = cleaned.replace(ch, "")
    return cleaned.strip()


def load_source_data(path, sheet_name):
    wb = openpyxl.load_workbook(path, data_only=True)
    src = wb[sheet_name]
    header_row_raw = [c.value for c in src[1]]
    header_row = [_clean_header(h) for h in header_row_raw]
    col_idx = {name: i for i, name in enumerate(header_row)}

    missing = [h for h in [IDEAL_COL, *PROXY_COLS.values()] if h not in col_idx]
    if missing:
        raise KeyError(
            f"Column(s) not found in '{sheet_name}' of {path}: {missing}\n"
            f"Available columns (raw, before cleaning): {header_row_raw}\n"
            f"Available columns (after cleaning): {header_row}"
        )

    time_labels, time_s, ideal = [], [], []
    proxies = {key: [] for key in PROXY_COLS}
    for row in src.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        time_labels.append(row[0])
        time_s.append(row[1])
        ideal.append(row[col_idx[IDEAL_COL]])
        for key, header in PROXY_COLS.items():
            proxies[key].append(row[col_idx[header]])

    return time_labels, time_s, ideal, proxies


def load_and_pool(paths_with_tags, sheet_name):
    """Load several source workbooks and concatenate their rows into one
    pooled dataset. paths_with_tags is a list of (file_path, tag) pairs,
    e.g. [(FILE_PATH_T1, "T1"), (FILE_PATH_T2, "T2")]. The tag is prefixed
    onto each row's Time-range label so the source window stays traceable
    in the output sheets (e.g. "T1: 17-30").

    Rows with a missing (None) value in any proxy or the ideal column are
    dropped, since M1/M2/M3 all require a numeric value in every column --
    the same rule the single-file script effectively required (a blank
    proxy cell there raised a TypeError at write-time)."""
    all_labels, all_time_s, all_ideal = [], [], []
    all_proxies = {key: [] for key in PROXY_COLS}

    for path, tag in paths_with_tags:
        print(f"Reading source data from: {path}  [tag: {tag}]")
        labels, time_s, ideal, proxies = load_source_data(path, sheet_name)
        n_loaded = len(ideal)

        kept = 0
        for i in range(n_loaded):
            row_values = [ideal[i]] + [proxies[key][i] for key in PROXY_COLS]
            if any(v is None for v in row_values):
                print(f"  Skipping row with missing value: {tag} {labels[i]}")
                continue
            all_labels.append(f"{tag}: {labels[i]}")
            all_time_s.append(time_s[i])
            all_ideal.append(ideal[i])
            for key in PROXY_COLS:
                all_proxies[key].append(proxies[key][i])
            kept += 1
        print(f"  {kept}/{n_loaded} rows kept from {tag}.")

    print(f"\nPooled dataset: {len(all_ideal)} total rows across {len(paths_with_tags)} source file(s).")
    return all_labels, all_time_s, all_ideal, all_proxies


# ---------------------------------------------------------------------------
# M1 -- Mean Multiplier
# ---------------------------------------------------------------------------
def write_mean_multiplier_sheet(wb, sheet_name, time_labels, time_s, ideal, proxy,
                                 ideal_header, proxy_header):
    n = len(ideal)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    headers = ["Time (s)", "Time (s)", ideal_header, proxy_header, "Multiplier",
               f"Modified {proxy_header}", "MAD_Before", "MAD_After",
               "MPD_Before", "MPD_After"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = BOLD
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    last_row = n + 1
    for i in range(n):
        r = i + 2
        ws.cell(row=r, column=1, value=time_labels[i])
        ws.cell(row=r, column=2, value=float(time_s[i]))
        ws.cell(row=r, column=3, value=float(ideal[i]))
        ws.cell(row=r, column=4, value=float(proxy[i]))
        ws.cell(row=r, column=5, value=f"=D{r}/C{r}")               # ratio
        ws.cell(row=r, column=6, value=f"=D{r}/$E${last_row + 1}")  # Proxy / mean-multiplier
        ws.cell(row=r, column=7, value=f"=ABS(C{r}-D{r})")          # MAD_Before
        ws.cell(row=r, column=8, value=f"=ABS(F{r}-C{r})")          # MAD_After
        ws.cell(row=r, column=9, value=f"=G{r}/C{r}*100")           # MPD_Before
        ws.cell(row=r, column=10, value=f"=H{r}/C{r}*100")          # MPD_After

    avg_row = last_row + 1
    ws.cell(row=avg_row, column=5, value=f"=AVERAGE(E2:E{last_row})")   # lambda (referenced by col F above)
    for col in ("G", "H", "I", "J"):
        ws.cell(row=avg_row, column=column_index_from_string(col),
                 value=f"=AVERAGE({col}2:{col}{last_row})")
    ws.cell(row=avg_row, column=1, value="Mean multiplier / averages ->").font = BOLD

    ws.column_dimensions["A"].width = 16
    for col in "CDEFGHIJ":
        ws.column_dimensions[col].width = 15
    ws.freeze_panes = "A2"
    return ws


# ---------------------------------------------------------------------------
# M2 -- 2nd-degree polynomial
# ---------------------------------------------------------------------------
def fit_poly2_with_diagnostics(proxy, ideal):
    """Fit the 2nd-degree least-squares polynomial and report whether the
    quadratic term is actually earning its place, statistically -- i.e.
    R^2 for linear-only vs quadratic, and a partial F-test p-value for
    adding the x^2 term."""
    x = np.asarray(proxy, dtype=float)
    y = np.asarray(ideal, dtype=float)
    n = len(y)

    a2, a1, a0 = np.polyfit(x, y, 2)
    y_pred_quad = a0 + a1 * x + a2 * x ** 2
    ss_res_quad = np.sum((y - y_pred_quad) ** 2)
    ss_tot = np.sum((y - np.mean(y)) ** 2)
    r2_quad = 1 - ss_res_quad / ss_tot

    b1, b0 = np.polyfit(x, y, 1)
    y_pred_lin = b0 + b1 * x
    ss_res_lin = np.sum((y - y_pred_lin) ** 2)
    r2_lin = 1 - ss_res_lin / ss_tot

    df1, df2 = 1, n - 3
    f_stat = ((ss_res_lin - ss_res_quad) / df1) / (ss_res_quad / df2)
    p_value = float(stats.f.sf(f_stat, df1, df2))

    return {
        "a0": a0, "a1": a1, "a2": a2,
        "r2_quad": r2_quad, "r2_lin": r2_lin,
        "f_stat": f_stat, "p_value": p_value,
    }


def write_poly2_sheet(wb, sheet_name, time_labels, time_s, ideal, proxy,
                       ideal_header, proxy_header):
    n = len(ideal)
    fit = fit_poly2_with_diagnostics(proxy, ideal)
    a0, a1, a2 = fit["a0"], fit["a1"], fit["a2"]

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    headers = ["Time (s)", "Time (s)", ideal_header, proxy_header,
               f"Modified {proxy_header}", "MAD", "MPD"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = BOLD
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    last_row = n + 1
    for i in range(n):
        r = i + 2
        ws.cell(row=r, column=1, value=time_labels[i])
        ws.cell(row=r, column=2, value=float(time_s[i]))
        ws.cell(row=r, column=3, value=float(ideal[i]))
        ws.cell(row=r, column=4, value=float(proxy[i]))
        ws.cell(row=r, column=5,
                value=f"={a0:.6f} + ({a1:.6f})*D{r}^1 + ({a2:.6f})*D{r}^2")
        ws.cell(row=r, column=6, value=f"=ABS(C{r}-E{r})")   # MAD
        ws.cell(row=r, column=7, value=f"=F{r}/C{r}*100")    # MPD

    avg_row = last_row + 1
    ws.cell(row=avg_row, column=6, value=f"=AVERAGE(F2:F{last_row})")
    ws.cell(row=avg_row, column=7, value=f"=AVERAGE(G2:G{last_row})")
    ws.cell(row=avg_row, column=1, value="Averages ->").font = BOLD

    # ---- Fit-quality diagnostics block: is the quadratic term
    # statistically justified, or just noise? -----------------------------
    diag_row = avg_row + 3
    ws.cell(row=diag_row, column=1, value="Fit diagnostics").font = BOLD
    ws.cell(row=diag_row + 1, column=1, value="R^2 (2nd-degree fit)")
    ws.cell(row=diag_row + 1, column=2, value=round(fit["r2_quad"], 4)).font = BLUE_FONT
    ws.cell(row=diag_row + 2, column=1, value="R^2 (linear-only, for comparison)")
    ws.cell(row=diag_row + 2, column=2, value=round(fit["r2_lin"], 4)).font = BLUE_FONT
    ws.cell(row=diag_row + 3, column=1, value="R^2 gain from adding x^2 term")
    ws.cell(row=diag_row + 3, column=2, value=round(fit["r2_quad"] - fit["r2_lin"], 4)).font = BLUE_FONT
    ws.cell(row=diag_row + 4, column=1, value="p-value: is the x^2 term significant? (F-test)")
    ws.cell(row=diag_row + 4, column=2, value=round(fit["p_value"], 4)).font = BLUE_FONT
    sig_note = ("significant at p<0.05" if fit["p_value"] < 0.05
                else "NOT significant at p<0.05 -- curve is statistically "
                     "indistinguishable from a straight line for this data")
    ws.cell(row=diag_row + 5, column=1, value=f"Interpretation: {sig_note}").font = NOTE_FONT

    note_row = diag_row + 7
    ws.cell(row=note_row, column=1,
            value=(f"Ideal_Density = {a0:.6f} + ({a1:.6f})\u00b7X^1 + "
                    f"({a2:.6f})\u00b7X^2  [numpy.polyfit, fit fresh from pooled T1+T2 data]"))
    ws.cell(row=note_row, column=1).font = NOTE_FONT

    ws.column_dimensions["A"].width = 40
    for col in "CDEFG":
        ws.column_dimensions[col].width = 15
    ws.freeze_panes = "A2"
    return ws


# ---------------------------------------------------------------------------
# M3 -- GPR  (same validated approach as the single-file script)
# ---------------------------------------------------------------------------
def fit_gpr(x, y):
    X = np.asarray(x, dtype=float).reshape(-1, 1)
    Y = np.asarray(y, dtype=float)
    kernel = (
        ConstantKernel(1.0, (1e-3, 1e6)) * RBF(10.0, (1e-2, 1e4))
        + WhiteKernel(1.0, (1e-6, 1e5))
    )
    gp = GaussianProcessRegressor(
        kernel=kernel, n_restarts_optimizer=N_RESTARTS,
        random_state=RANDOM_STATE, normalize_y=False,
    )
    gp.fit(X, Y)
    params = gp.kernel_.get_params()
    return {
        "sigma_f2": float(params["k1__k1__constant_value"]),
        "length_scale": float(params["k1__k2__length_scale"]),
        "noise_var": float(params["k2__noise_level"]),
        "kernel_repr": str(gp.kernel_),
        "log_marginal_likelihood": float(gp.log_marginal_likelihood_value_),
    }


def write_gpr_sheet(wb, sheet_name, time_labels, time_s, ideal, proxy, fit,
                     ideal_header, proxy_header):
    n = len(ideal)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    headers = [
        "Time (s)", "Time (s)", ideal_header, proxy_header,
        "GPR Modified Density (vehicle/km/lane)", "Posterior Std",
        "MAD_Before", "MAD_After", "MPD_Before", "MPD_After",
    ]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = BOLD
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    for i in range(n):
        r = i + 2
        ws.cell(row=r, column=1, value=time_labels[i])
        ws.cell(row=r, column=2, value=float(time_s[i]))
        ws.cell(row=r, column=3, value=float(ideal[i]))
        ws.cell(row=r, column=4, value=float(proxy[i]))
        ws.cell(row=r, column=7, value=f"=ABS(C{r}-D{r})")
        ws.cell(row=r, column=8, value=f"=ABS(C{r}-E{r})")
        ws.cell(row=r, column=9, value=f"=G{r}/C{r}*100")
        ws.cell(row=r, column=10, value=f"=H{r}/C{r}*100")

    avg_row = n + 2
    ws.cell(row=avg_row, column=6, value="Averages ->").font = BOLD
    for col in ("G", "H", "I", "J"):
        ws.cell(row=avg_row, column=column_index_from_string(col),
                 value=f"=AVERAGE({col}2:{col}{n + 1})")

    ws["L1"] = "GPR Hyperparameters (fit via scikit-learn GaussianProcessRegressor, RBF + WhiteKernel, outside Excel, on pooled T1+T2 data)"
    ws["L1"].font = BOLD
    ws["L2"] = "Sigma_f^2 (signal variance)"
    ws["M2"] = fit["sigma_f2"]
    ws["L3"] = "Length_scale"
    ws["M3"] = fit["length_scale"]
    ws["L4"] = "Noise_variance"
    ws["M4"] = fit["noise_var"]
    for addr in ("M2", "M3", "M4"):
        ws[addr].font = BLUE_FONT
    ws["L5"] = "Fitted kernel (sklearn repr.)"
    ws["M5"] = fit["kernel_repr"]
    ws["L6"] = "Log marginal likelihood"
    ws["M6"] = fit["log_marginal_likelihood"]
    ws["L7"] = ("Modified Density = K_row_i . inv(K+noise*I) . y ; "
                "Posterior Std = sqrt((sigma_f^2+noise) - K_row_i . inv(K+noise*I) . K_row_i)")
    ws["L7"].font = NOTE_FONT
    ws.row_dimensions[7].height = 30

    start_col = 15
    kclean_start = start_col
    knoisy_start = kclean_start + n
    kinv_start = knoisy_start + n
    alpha_col = kinv_start + n

    def cl(idx):
        return get_column_letter(idx)

    ws.cell(row=1, column=kclean_start,
            value="K (signal only, RBF) -- used for cross-covariance").font = NOTE_FONT
    ws.cell(row=1, column=knoisy_start,
            value="K + noise*I -- used for inversion").font = NOTE_FONT
    ws.cell(row=1, column=kinv_start,
            value="inv(K + noise*I)  [array formula]").font = NOTE_FONT

    for i in range(n):
        row_i = i + 2
        for j in range(n):
            row_j = j + 2
            base = f"$M$2*EXP(-0.5*(($D{row_i}-$D{row_j})/$M$3)^2)"
            ws.cell(row=row_i, column=kclean_start + j, value=f"={base}")
            noise_term = "+$M$4" if i == j else ""
            ws.cell(row=row_i, column=knoisy_start + j, value=f"={base}{noise_term}")

    kclean_range = f"{cl(kclean_start)}2:{cl(kclean_start + n - 1)}{n + 1}"
    knoisy_range = f"{cl(knoisy_start)}2:{cl(knoisy_start + n - 1)}{n + 1}"
    kinv_range = f"{cl(kinv_start)}2:{cl(kinv_start + n - 1)}{n + 1}"
    alpha_range = f"${cl(alpha_col)}$2:${cl(alpha_col)}${n + 1}"

    ws.cell(row=2, column=kinv_start).value = ArrayFormula(
        ref=kinv_range, text=f"=MINVERSE({knoisy_range})"
    )
    ws.cell(row=1, column=alpha_col, value="alpha = Kinv . y  [array]").font = NOTE_FONT
    ws.cell(row=2, column=alpha_col).value = ArrayFormula(
        ref=f"{cl(alpha_col)}2:{cl(alpha_col)}{n + 1}",
        text=f"=MMULT({kinv_range},$C$2:$C${n + 1})",
    )

    for i in range(n):
        row = i + 2
        row_range = f"${cl(kclean_start)}{row}:${cl(kclean_start + n - 1)}{row}"
        ws.cell(row=row, column=5).value = ArrayFormula(
            ref=f"E{row}:E{row}", text=f"=MMULT({row_range},{alpha_range})",
        )
        ws.cell(row=row, column=6).value = ArrayFormula(
            ref=f"F{row}:F{row}",
            text=(f"=SQRT(MAX(0,($M$2+$M$4)-MMULT(MMULT({row_range},{kinv_range}),"
                  f"TRANSPOSE({row_range}))))"),
        )

    std_col = alpha_col + 2
    for col_idx in range(kclean_start, std_col + 1):
        ws.column_dimensions[cl(col_idx)].hidden = True

    ws.column_dimensions["A"].width = 16
    for col in ("C", "D", "E", "F", "G", "H", "I", "J"):
        ws.column_dimensions[col].width = 14
    ws.column_dimensions["L"].width = 26
    ws.column_dimensions["M"].width = 20
    ws.freeze_panes = "A2"
    return ws


# ---------------------------------------------------------------------------
def main():
    time_labels, time_s, ideal, proxies = load_and_pool(
        [(FILE_PATH_T1, "T1"), (FILE_PATH_T2, "T2")], SOURCE_SHEET
    )
    n = len(ideal)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # drop the default empty sheet

    sheet_names = {
        "Flow_SMS": {"mmult": "M. Mult Acc-Flow_SMS", "poly2": "2nd deg Acc-Flow_SMS", "gpr": "GPR Acc_Flow_SMS"},
        "Flow_Snapshot": {"mmult": "M. Mult Acc-Flow_Snap", "poly2": "2nd deg Acc-Flow_Snap", "gpr": "GPR Acc_Flow_Snap"},
        "Flow_STOPLine": {"mmult": "M. Mult Acc-Flow_STOPLine", "poly2": "2nd deg Acc-Flow_STOPLine", "gpr": "GPR Acc_Flow_STOPLine"},
        "Occ": {"mmult": "M. Mult Acc-Occ", "poly2": "2nd deg Acc-Occ", "gpr": "GPR Acc_Occ"},
    }

    for key, header in PROXY_COLS.items():
        proxy = proxies[key]
        names = sheet_names[key]

        print(f"\n=== {key} ({header}) -- pooled n={n} ===")

        print(f"  M1 Mean Multiplier -> '{names['mmult']}'")
        write_mean_multiplier_sheet(wb, names["mmult"], time_labels, time_s, ideal,
                                     proxy, IDEAL_COL, header)

        print(f"  M2 2nd-degree polynomial -> '{names['poly2']}'")
        write_poly2_sheet(wb, names["poly2"], time_labels, time_s, ideal, proxy,
                           IDEAL_COL, header)

        print(f"  M3 GPR -> '{names['gpr']}' (fitting via scikit-learn...)")
        fit = fit_gpr(proxy, ideal)
        print(f"     sigma_f^2={fit['sigma_f2']:.4f}  length_scale={fit['length_scale']:.4f}  "
              f"noise={fit['noise_var']:.4f}  log-ML={fit['log_marginal_likelihood']:.4f}")
        write_gpr_sheet(wb, names["gpr"], time_labels, time_s, ideal, proxy, fit,
                         IDEAL_COL, header)

    wb.calculation.fullCalcOnLoad = True

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    wb.save(OUTPUT_PATH)
    print(f"\nSaved pooled combined workbook ({len(wb.sheetnames)} sheets, n={n} rows) to:\n  {OUTPUT_PATH}")
    print("Open in Excel and let it recalculate (automatic on open) to populate "
          "the GPR array formulas.")


if __name__ == "__main__":
    main()