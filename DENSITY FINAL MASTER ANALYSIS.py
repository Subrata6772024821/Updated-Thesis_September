"""
Combined Density Calibration -- all 3 methods x 4 proxy pairs
================================================================

Builds ONE output workbook with 12 sheets:

    M. Mult Acc-Flow_SMS       2nd deg Acc-Flow_SMS       GPR Acc_Flow_SMS
    M. Mult Acc-Flow_Snap      2nd deg Acc-Flow_Snap      GPR Acc_Flow_Snap
    M. Mult Acc-Flow_STOPLine  2nd deg Acc-Flow_STOPLine  GPR Acc_Flow_STOPLine
    M. Mult Acc-Occ            2nd deg Acc-Occ            GPR Acc_Occ

NEW IN THIS VERSION: the Flow_STOPLINE proxy (Density_Flow_STOPLINE
(vehicle/km/lane)) is now included as a fourth proxy pair, alongside
Flow_SMS, Flow_Snapshot, and Occupancy. This column is present in your
Combined Density source workbook but was missing from PROXY_COLS in the
previous version of this script, so it never got its own M1/M2/M3
sheets -- that's fixed below.

FIX (this version): load_source_data now normalizes header text before
matching -- stripping stray double-quote characters and leading/trailing
whitespace. This fixes a KeyError caused by the source workbook's
"Density_Flow_SMS (PCU/km)" header actually containing a trailing
literal `"` character (i.e. stored as `Density_Flow_SMS (PCU/km)"`),
which previously failed an exact-string match against PROXY_COLS.

All three methods are LIVE, formula-driven in Excel (not pasted
values) -- consistent with how the M1/M2 sheets were originally built
by hand, and how the GPR sheet was built in the previous script:

  M1  Mean Multiplier:      lambda = AVERAGE(Proxy/Ideal) over all rows
                             Modified = Proxy / lambda
  M2  2nd-degree polynomial: Ideal ~ a0 + a1*Proxy + a2*Proxy^2
                             (least-squares fit done once in Python via
                             numpy.polyfit, then the fitted a0/a1/a2 are
                             written as hardcoded coefficients inside a
                             live Excel formula per row). The sheet also
                             reports R^2 and a partial F-test p-value for
                             the x^2 term, so you can see directly whether
                             the quadratic curvature is statistically
                             justified or just fitting noise.
  M3  GPR:                   RBF + WhiteKernel, hyperparameters fit once
                             via scikit-learn, then kernel matrix /
                             MINVERSE / MMULT posterior mean+std computed
                             natively in-sheet.

NOTE ON M2 COEFFICIENTS: this script always fits fresh from whatever
Sheet1 data is in FILE_PATH, so any hardcoded-coefficient mismatches
from older workbooks self-correct automatically.

Run in PyCharm:
    pip install openpyxl pandas numpy scikit-learn scipy
"""

import os
import shutil
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
FILE_PATH = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\120_sec\PCU\05_12_25_1315_1345_Combined Density_120sec_T2.xlsx"
SOURCE_SHEET = "Sheet1"

IDEAL_COL = "Density_Accumulation (PCU/km)"
PROXY_COLS = {
    "Flow_SMS": "Density_Flow_SMS (PCU/km)",
    "Flow_Snapshot": "Density_Flow_SNAPSHOT (PCU/km)",
    "Flow_STOPLine": "Density_Flow_STOPLINE (PCU/km)",
    "Occ": "Density_Occupancy (vehicle_varieed length/km)",
}

# Output: saved into the PARENT folder of FILE_PATH's own folder, i.e. one
# level above "...\30_sec\", assumed to be the per-interval-run folder
# "...\05_12_25_1315_1345\". Change OUTPUT_PATH directly if that's wrong
# for your folder layout.
_source_dir = os.path.dirname(FILE_PATH)
_parent_dir = os.path.dirname(_source_dir)
_base_name = os.path.splitext(os.path.basename(FILE_PATH))[0]
OUTPUT_PATH = os.path.join(_parent_dir, f"{_base_name}_AllMethods_T2.xlsx")

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

    ws.column_dimensions["A"].width = 12
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
                    f"({a2:.6f})\u00b7X^2  [numpy.polyfit, fit fresh from this sheet's data]"))
    ws.cell(row=note_row, column=1).font = NOTE_FONT

    ws.column_dimensions["A"].width = 38
    for col in "CDEFG":
        ws.column_dimensions[col].width = 15
    ws.freeze_panes = "A2"
    return ws


# ---------------------------------------------------------------------------
# M3 -- GPR  (same validated approach as the previous GPR-only script)
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

    ws["L1"] = "GPR Hyperparameters (fit via scikit-learn GaussianProcessRegressor, RBF + WhiteKernel, outside Excel)"
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

    ws.column_dimensions["A"].width = 12
    for col in ("C", "D", "E", "F", "G", "H", "I", "J"):
        ws.column_dimensions[col].width = 14
    ws.column_dimensions["L"].width = 26
    ws.column_dimensions["M"].width = 20
    ws.freeze_panes = "A2"
    return ws


# ---------------------------------------------------------------------------
def main():
    print(f"Reading source data from: {FILE_PATH}")
    time_labels, time_s, ideal, proxies = load_source_data(FILE_PATH, SOURCE_SHEET)
    n = len(ideal)
    print(f"  {n} rows loaded.")

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

        print(f"\n=== {key} ({header}) ===")

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
    print(f"\nSaved combined workbook ({len(wb.sheetnames)} sheets) to:\n  {OUTPUT_PATH}")
    print("Open in Excel and let it recalculate (automatic on open) to populate "
          "the GPR array formulas.")


if __name__ == "__main__":
    main()