"""
With-DTW Combined Density Calibration -- 3 methods x 4 proxy pairs
=====================================================================

Reads the one-to-one windowed-DTW alignment results (one sheet per
proxy: Acc-Flow_SMS, Acc-Flow_Snapshot, Acc-Flow_STOPLine,
Acc-Occupancy -- each already a clean 61-row, exactly-one-to-one table
with columns Reference_Value / Aligned_Proxy_Value) and builds a
combined output workbook with 12 sheets:

    M. Mult Acc-Flow_SMS_DTW    2nd deg Acc-Flow_SMS_DTW    GPR Acc_Flow_SMS_DTW
    M. Mult Acc-Flow_Snap_DTW   2nd deg Acc-Flow_Snap_DTW   GPR Acc_Flow_Snap_DTW
    M. Mult Acc-STOPLine_DTW    2nd deg Acc-STOPLine_DTW    GPR Acc_STOPLine_DTW
    M. Mult Acc-Occ_DTW         2nd deg Acc-Occ_DTW         GPR Acc_Occ_DTW

NO "Before" MAD/MPD columns here -- those already exist in your
without-DTW workbook and are identical regardless of DTW (the "before"
condition is just the raw, unaligned proxy vs Ideal, which doesn't
change). Only the AFTER-calibration MAD/MPD is computed on each sheet
here, same as before.

Because the source data is now a true one-to-one pairing (61 clean
rows, no repeats on either side -- unlike the earlier
"Ideal_Time_Index appears 22 times" unconstrained-DTW table), MAD/MPD
here are simple, unambiguous 61-row averages -- no "Option A / Option
B" collapsing needed.

All three methods are fit fresh from this sheet's own (Reference,
Aligned_Proxy) pairs and written as LIVE Excel formulas (Mean
Multiplier and 2nd-degree polynomial via numpy, GPR via scikit-learn
with the kernel-matrix/MINVERSE/MMULT array-formula approach used in
every previous script here), so nothing is pasted as static numbers.

Run in PyCharm:
    pip install openpyxl pandas numpy scikit-learn scipy
"""

import os

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
FILE_PATH = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\120_sec\05_12_25_1315_1345_DTW_OneToOne_Results.xlsx"

SOURCE_SHEETS = {
    "Flow_SMS": "Acc-Flow_SMS",
    "Flow_Snapshot": "Acc-Flow_Snapshot",
    "Flow_STOPLine": "Acc-Flow_STOPLine",
    "Occupancy": "Acc-Occupancy",
}

N_RESTARTS = 15
RANDOM_STATE = 42

# Output: saved into the PARENT folder of FILE_PATH's own folder (one
# level above "...\30_sec\"), with "With DTW" in the filename.
_source_dir = os.path.dirname(FILE_PATH)
_parent_dir = os.path.dirname(_source_dir)
_run_id = os.path.basename(_parent_dir)
OUTPUT_PATH = os.path.join(_parent_dir, f"{_run_id}_AllMethods_With DTW_120sec.xlsx")

HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
BLUE_FONT = Font(color="0000FF")
BOLD = Font(bold=True)
NOTE_FONT = Font(italic=True, size=9, color="808080")


# ---------------------------------------------------------------------------
def load_aligned_sheet(path, sheet_name):
    """Reads one proxy's one-to-one alignment sheet, keeping only the
    numeric data rows (the trailing Metric/Value summary block at the
    bottom of each sheet is skipped)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    src = wb[sheet_name]
    header_row = [c.value for c in src[1]]
    col_idx = {name: i for i, name in enumerate(header_row)}

    ref_idx, ref_time, reference, proxy_idx, proxy_time, aligned_proxy = [], [], [], [], [], []
    for row in src.iter_rows(min_row=2, values_only=True):
        if not isinstance(row[col_idx["Ref_Time_Index"]], (int, float)):
            continue  # skip the trailing summary block
        ref_idx.append(row[col_idx["Ref_Time_Index"]])
        ref_time.append(row[col_idx["Ref_Time_s"]])
        reference.append(row[col_idx["Reference_Value"]])
        proxy_idx.append(row[col_idx["Matched_Proxy_Time_Index"]])
        proxy_time.append(row[col_idx["Matched_Proxy_Time_s"]])
        aligned_proxy.append(row[col_idx["Aligned_Proxy_Value"]])

    return ref_idx, ref_time, reference, proxy_idx, proxy_time, aligned_proxy


def write_common_header(ws, extra_headers):
    headers = ["Ref_Time_Index", "Ref_Time_s", "Reference_Value",
               "Matched_Proxy_Time_Index", "Matched_Proxy_Time_s",
               "Aligned_Proxy_Value"] + extra_headers
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = BOLD
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    return headers


def write_base_columns(ws, ref_idx, ref_time, reference, proxy_idx, proxy_time, aligned_proxy):
    n = len(reference)
    for i in range(n):
        r = i + 2
        ws.cell(row=r, column=1, value=ref_idx[i])
        ws.cell(row=r, column=2, value=ref_time[i])
        ws.cell(row=r, column=3, value=float(reference[i]))
        ws.cell(row=r, column=4, value=proxy_idx[i])
        ws.cell(row=r, column=5, value=proxy_time[i])
        ws.cell(row=r, column=6, value=float(aligned_proxy[i]))
    return n


# ---------------------------------------------------------------------------
# M1 -- Mean Multiplier
# ---------------------------------------------------------------------------
def write_mean_multiplier_sheet(wb, sheet_name, ref_idx, ref_time, reference,
                                 proxy_idx, proxy_time, aligned_proxy):
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    write_common_header(ws, ["Multiplier", "Modified Density", "MAD", "MPD"])
    n = write_base_columns(ws, ref_idx, ref_time, reference, proxy_idx, proxy_time, aligned_proxy)

    last_row = n + 1
    for i in range(n):
        r = i + 2
        ws.cell(row=r, column=7, value=f"=F{r}/C{r}")               # ratio
        ws.cell(row=r, column=8, value=f"=F{r}/$G${last_row + 1}")  # Aligned_Proxy / mean-multiplier
        ws.cell(row=r, column=9, value=f"=ABS(H{r}-C{r})")          # MAD
        ws.cell(row=r, column=10, value=f"=I{r}/C{r}*100")          # MPD

    avg_row = last_row + 1
    ws.cell(row=avg_row, column=7, value=f"=AVERAGE(G2:G{last_row})")  # lambda
    for col in ("I", "J"):
        ws.cell(row=avg_row, column=column_index_from_string(col),
                 value=f"=AVERAGE({col}2:{col}{last_row})")
    ws.cell(row=avg_row, column=1, value=f"Averages (N={n}) ->").font = BOLD

    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["E"].width = 12
    for col in "CFGHIJ":
        ws.column_dimensions[col].width = 15
    ws.freeze_panes = "A2"
    return ws


# ---------------------------------------------------------------------------
# M2 -- 2nd-degree polynomial
# ---------------------------------------------------------------------------
def fit_poly2_with_diagnostics(proxy, ideal):
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

    return {"a0": a0, "a1": a1, "a2": a2, "r2_quad": r2_quad, "r2_lin": r2_lin,
            "f_stat": f_stat, "p_value": p_value}


def write_poly2_sheet(wb, sheet_name, ref_idx, ref_time, reference,
                       proxy_idx, proxy_time, aligned_proxy):
    fit = fit_poly2_with_diagnostics(aligned_proxy, reference)
    a0, a1, a2 = fit["a0"], fit["a1"], fit["a2"]

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    write_common_header(ws, ["Modified Density", "MAD", "MPD"])
    n = write_base_columns(ws, ref_idx, ref_time, reference, proxy_idx, proxy_time, aligned_proxy)

    last_row = n + 1
    for i in range(n):
        r = i + 2
        ws.cell(row=r, column=7,
                value=f"={a0:.6f} + ({a1:.6f})*F{r}^1 + ({a2:.6f})*F{r}^2")
        ws.cell(row=r, column=8, value=f"=ABS(G{r}-C{r})")   # MAD
        ws.cell(row=r, column=9, value=f"=H{r}/C{r}*100")    # MPD

    avg_row = last_row + 1
    ws.cell(row=avg_row, column=8, value=f"=AVERAGE(H2:H{last_row})")
    ws.cell(row=avg_row, column=9, value=f"=AVERAGE(I2:I{last_row})")
    ws.cell(row=avg_row, column=1, value=f"Averages (N={n}) ->").font = BOLD

    diag_row = avg_row + 3
    ws.cell(row=diag_row, column=1, value="Fit diagnostics").font = BOLD
    ws.cell(row=diag_row + 1, column=1, value="R^2 (2nd-degree fit)")
    ws.cell(row=diag_row + 1, column=2, value=round(fit["r2_quad"], 4)).font = BLUE_FONT
    ws.cell(row=diag_row + 2, column=1, value="R^2 (linear-only, for comparison)")
    ws.cell(row=diag_row + 2, column=2, value=round(fit["r2_lin"], 4)).font = BLUE_FONT
    ws.cell(row=diag_row + 3, column=1, value="p-value: is the x^2 term significant? (F-test)")
    ws.cell(row=diag_row + 3, column=2, value=round(fit["p_value"], 4)).font = BLUE_FONT
    sig_note = ("significant at p<0.05" if fit["p_value"] < 0.05
                else "NOT significant at p<0.05")
    ws.cell(row=diag_row + 4, column=1, value=f"Interpretation: {sig_note}").font = NOTE_FONT

    eq_row = diag_row + 6
    ws.cell(row=eq_row, column=1,
            value=(f"Reference = {a0:.6f} + ({a1:.6f})\u00b7X^1 + "
                    f"({a2:.6f})\u00b7X^2  [fit on the {n} one-to-one DTW-aligned pairs]")).font = NOTE_FONT

    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["E"].width = 12
    for col in "CFGHI":
        ws.column_dimensions[col].width = 15
    ws.freeze_panes = "A2"
    return ws


# ---------------------------------------------------------------------------
# M3 -- GPR
# ---------------------------------------------------------------------------
def fit_gpr(x, y):
    X = np.asarray(x, dtype=float).reshape(-1, 1)
    Y = np.asarray(y, dtype=float)
    kernel = (
        ConstantKernel(1.0, (1e-3, 1e6)) * RBF(10.0, (1e-2, 1e6))
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


def write_gpr_sheet(wb, sheet_name, ref_idx, ref_time, reference,
                     proxy_idx, proxy_time, aligned_proxy, fit):
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    write_common_header(ws, ["GPR Modified Density", "Posterior Std", "MAD", "MPD"])
    n = write_base_columns(ws, ref_idx, ref_time, reference, proxy_idx, proxy_time, aligned_proxy)

    for i in range(n):
        r = i + 2
        ws.cell(row=r, column=9, value=f"=ABS(G{r}-C{r})")     # MAD
        ws.cell(row=r, column=10, value=f"=I{r}/C{r}*100")     # MPD

    avg_row = n + 2
    ws.cell(row=avg_row, column=8, value=f"Averages (N={n}) ->").font = BOLD
    for col in ("I", "J"):
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

    start_col = 15  # column O
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
            base = f"$M$2*EXP(-0.5*(($F{row_i}-$F{row_j})/$M$3)^2)"
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
        ws.cell(row=row, column=7).value = ArrayFormula(
            ref=f"G{row}:G{row}", text=f"=MMULT({row_range},{alpha_range})",
        )
        ws.cell(row=row, column=8).value = ArrayFormula(
            ref=f"H{row}:H{row}",
            text=(f"=SQRT(MAX(0,($M$2+$M$4)-MMULT(MMULT({row_range},{kinv_range}),"
                  f"TRANSPOSE({row_range}))))"),
        )

    for col_idx in range(kclean_start, alpha_col + 1):
        ws.column_dimensions[cl(col_idx)].hidden = True

    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["E"].width = 12
    for col in ("C", "F", "G", "H", "I", "J"):
        ws.column_dimensions[col].width = 14
    ws.column_dimensions["L"].width = 26
    ws.column_dimensions["M"].width = 20
    ws.freeze_panes = "A2"
    return ws


# ---------------------------------------------------------------------------
def main():
    print(f"Reading with-DTW one-to-one alignment results from: {FILE_PATH}")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    sheet_names = {
        "Flow_SMS": {"mmult": "M. Mult Acc-Flow_SMS_DTW", "poly2": "2nd deg Acc-Flow_SMS_DTW", "gpr": "GPR Acc_Flow_SMS_DTW"},
        "Flow_Snapshot": {"mmult": "M. Mult Acc-Flow_Snap_DTW", "poly2": "2nd deg Acc-Flow_Snap_DTW", "gpr": "GPR Acc_Flow_Snap_DTW"},
        "Flow_STOPLine": {"mmult": "M. Mult Acc-STOPLine_DTW", "poly2": "2nd deg Acc-STOPLine_DTW", "gpr": "GPR Acc_STOPLine_DTW"},
        "Occupancy": {"mmult": "M. Mult Acc-Occ_DTW", "poly2": "2nd deg Acc-Occ_DTW", "gpr": "GPR Acc_Occ_DTW"},
    }

    for key, src_sheet in SOURCE_SHEETS.items():
        print(f"\n=== {key} (source sheet: '{src_sheet}') ===")
        ref_idx, ref_time, reference, proxy_idx, proxy_time, aligned_proxy = load_aligned_sheet(
            FILE_PATH, src_sheet
        )
        n = len(reference)
        print(f"  {n} one-to-one aligned rows loaded.")
        names = sheet_names[key]

        print(f"  M1 Mean Multiplier -> '{names['mmult']}'")
        write_mean_multiplier_sheet(wb, names["mmult"], ref_idx, ref_time, reference,
                                     proxy_idx, proxy_time, aligned_proxy)

        print(f"  M2 2nd-degree polynomial -> '{names['poly2']}'")
        write_poly2_sheet(wb, names["poly2"], ref_idx, ref_time, reference,
                           proxy_idx, proxy_time, aligned_proxy)

        print(f"  M3 GPR -> '{names['gpr']}' (fitting via scikit-learn...)")
        fit = fit_gpr(aligned_proxy, reference)
        print(f"     sigma_f^2={fit['sigma_f2']:.4f}  length_scale={fit['length_scale']:.4f}  "
              f"noise={fit['noise_var']:.4f}  log-ML={fit['log_marginal_likelihood']:.4f}")
        write_gpr_sheet(wb, names["gpr"], ref_idx, ref_time, reference,
                         proxy_idx, proxy_time, aligned_proxy, fit)

    wb.calculation.fullCalcOnLoad = True
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    wb.save(OUTPUT_PATH)
    print(f"\nSaved combined With-DTW workbook ({len(wb.sheetnames)} sheets) to:\n  {OUTPUT_PATH}")
    print("Open in Excel and let it recalculate (automatic on open) to populate "
          "the GPR array formulas.")


if __name__ == "__main__":
    main()