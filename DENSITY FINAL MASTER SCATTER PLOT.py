"""
Plot: Proxy Density vs Reference Density (Accumulation) - Before vs After Calibration
Produces FOUR separate figures - one per proxy method (Flow_SMS, Flow_Snapshot,
Flow_Stopline, Occupancy). Each figure has 3 panels (Multiplicative, 2nd-Degree
Polynomial, GPR) showing before/after calibration points against the reference line.

Run this in PyCharm with your project's Python interpreter.
Requires: pandas, matplotlib, openpyxl  ->  pip install pandas matplotlib openpyxl
"""

import os
import sys
import pandas as pd
import matplotlib.pyplot as plt

# ----------------------------------------------------------------------
# 1. INPUT FILE
# ----------------------------------------------------------------------
file_path = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\120_sec\PCU\05_12_25_1315_1345_Combined Density_120sec_AllMethods_T2.xlsx"

if not os.path.exists(file_path):
    sys.exit(f"File not found:\n{file_path}")

# Diagrams are saved in the same (parent) folder that contains the Excel file
out_dir = os.path.dirname(file_path)

# ----------------------------------------------------------------------
# 1b. FORCE EXCEL TO RECALCULATE THE WORKBOOK FIRST
# ----------------------------------------------------------------------
# The "Modified Density" (after-calibration) columns are live Excel formulas
# (e.g. mean-multiplier, 2nd-degree polynomial, GPR array formulas). If the
# workbook was last saved without being opened/calculated in Excel, those
# cells have NO cached value -> pandas/openpyxl reads them as blank/NaN,
# which is why the "after calibration" points don't show up on the plot.
#
# This step uses Excel itself (via COM automation) to open the file, force a
# full recalculation, and save it - so the values are guaranteed to be there.
# Requires: pip install pywin32   (and Excel installed on this machine)
try:
    import win32com.client as win32

    print("Opening Excel to recalculate formulas (this window may flash briefly)...")
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    wb_com = excel.Workbooks.Open(file_path)
    excel.CalculateFullRebuild()
    wb_com.Save()
    wb_com.Close()
    excel.Quit()
    print("Recalculation complete.\n")
except Exception as e:
    print(f"Could not auto-recalculate via Excel COM ({e}).")
    print("If 'after calibration' points are missing from the plot, open the file")
    print("in Excel yourself, press Ctrl+Alt+F9, save it, and re-run this script.\n")

# ----------------------------------------------------------------------
# 2. SHEET / PROXY / CALIBRATION-METHOD MAPPING
# ----------------------------------------------------------------------
# (proxy suffix used in sheet names, display label)
proxies = [
    ("Flow_SMS",      "Flow_SMS"),
    ("Flow_Snap",     "Flow_Snapshot"),
    ("Flow_STOPLine", "Flow_Stopline"),
    ("Occ",           "Occupancy"),
]

# (sheet-name prefix, display label)
calibration_methods = [
    ("M. Mult", "Mean Multiplier"),
    ("2nd deg", "2nd-Degree Polynomial"),
    ("GPR",     "GPR"),
]

# Distinct, fixed colors so before/after are always easy to tell apart
BEFORE_COLOR = "#1f77b4"   # blue
AFTER_COLOR = "#d62728"    # red

# Unit label used consistently on both axes
DENSITY_UNIT = "PCU/km"

# Time window label appended to each figure's title
TIME_WINDOW_LABEL = "Time Window T2 (120 - 180)"

# Short tag appended to each output PNG's filename (e.g. "T1", "T2", ...)
TIME_WINDOW_TAG = "T2"


def sheet_name(method_prefix, proxy_suffix):
    """Reproduce the exact sheet-naming convention used in the workbook."""
    if proxy_suffix == "Occ":
        joiner = "Acc_Occ" if method_prefix == "GPR" else "Acc-Occ"
    else:
        joiner = f"Acc_{proxy_suffix}" if method_prefix == "GPR" else f"Acc-{proxy_suffix}"
    return f"{method_prefix} {joiner}"


# ----------------------------------------------------------------------
# 3. BUILD ONE FIGURE PER PROXY
# ----------------------------------------------------------------------
any_nan_warning = False

for proxy_suffix, proxy_label in proxies:
    fig, axes = plt.subplots(1, 3, figsize=(18, 6), sharex=True, sharey=True)

    for ax, (method_prefix, method_label) in zip(axes, calibration_methods):
        sn = sheet_name(method_prefix, proxy_suffix)
        df = pd.read_excel(file_path, sheet_name=sn)

        # Locate columns by name pattern (robust to header variations)
        acc_col = [c for c in df.columns if "Density_Accumulation" in c][0]
        raw_col = [c for c in df.columns
                   if c.startswith("Density_Flow") or c.startswith("Density_Occupancy")][0]
        mod_col = [c for c in df.columns if "Modified Density" in c][0]

        x = df[acc_col]
        y_before = df[raw_col]
        y_after = df[mod_col]

        if y_after.isna().all():
            any_nan_warning = True

        # Before calibration: hollow blue circle
        ax.scatter(x, y_before, facecolors='none', edgecolors=BEFORE_COLOR,
                   marker='o', s=40, alpha=0.8, label="Before Calibration")
        # After calibration: filled red x
        ax.scatter(x, y_after, color=AFTER_COLOR,
                   marker='x', s=45, alpha=0.9, label="After Calibration")

        lims = [min(ax.get_xlim()[0], ax.get_ylim()[0]), max(ax.get_xlim()[1], ax.get_ylim()[1])]
        ax.plot(lims, lims, linestyle='--', color='gray', linewidth=1, label='y = x')
        ax.set_xlim(lims)
        ax.set_ylim(lims)
        ax.set_title(method_label, fontsize=13, fontweight='bold')
        ax.set_xlabel(f"Density by Accumulation (Reference) [{DENSITY_UNIT}]")
        ax.grid(True, linestyle=':', alpha=0.5)

    axes[0].set_ylabel(f"Density by {proxy_label} Speed [{DENSITY_UNIT}]")
    handles, labels = axes[0].get_legend_handles_labels()
    fig.legend(handles, labels, loc='lower center', ncol=3, bbox_to_anchor=(0.5, -0.05), fontsize=10)

    fig.suptitle(
        f"Density by {proxy_label} Speed vs Density by Accumulation (Reference)\n"
        f"Before vs After Calibration -- {TIME_WINDOW_LABEL}",
        fontsize=14, fontweight='bold'
    )
    fig.tight_layout(rect=[0, 0.05, 1, 0.93])

    out_path = os.path.join(out_dir, f"Density_Calibration_{proxy_suffix}_{TIME_WINDOW_TAG}.png")
    fig.savefig(out_path, dpi=300, bbox_inches='tight')
    print(f"Saved: {out_path}")
    plt.close(fig)

if any_nan_warning:
    print("\nWARNING: one or more 'Modified Density' (After) columns were entirely blank/NaN.")
    print("This means the workbook's formula cells have no cached calculated values.")
    print("Fix: open the Excel file once, press Ctrl+Alt+F9 to force full recalculation,")
    print("save it, and re-run this script.")

print("\nDone.")