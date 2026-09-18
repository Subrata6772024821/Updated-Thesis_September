"""
Time (continuous, seconds) vs SIGNED Deviation -- one figure per proxy pair
============================================================================
Reference  : Density_Accumulation (the "ideal").
Signed dev : Calibrated (Modified) Density  -  Density_Accumulation
             > 0  -> calibrated value OVERSHOOTS the ideal  ("positive deviation")
             < 0  -> calibrated value UNDERSHOOTS the ideal ("negative deviation")

X axis  : REAL elapsed time in seconds (a continuous numeric axis, not a
          category axis). Each row's own bin edges ("30-60", "45-75", ...)
          are used directly as the bar's [start, end] span via
          ax.bar(..., align="edge"), so a Timestep-2 interval that overlaps
          a Timestep-1 interval (e.g. "T1: 30-60" and "T2: 45-75") is drawn
          overlapping on the plot exactly as it does in real time -- nothing
          is forced into separate side-by-side categorical slots.
Y axis  : Signed deviation (PCU/km). For each interval and each method
          the positive deviations are averaged into an upward bar and the
          negative ones into a downward bar (in practice one row per
          interval per method, so this is just that row's own value).

Colour  now encodes METHOD (Mean Multiplier / 2nd-Degree Polynomial / GPR),
         three distinct colours, as in the very first version of this script.
Line style encodes TIMESTEP:
    Timestep 1 (T1) -> solid outline
    Timestep 2 (T2) -> finely dotted outline
Bars are unfilled (outline only) throughout.

Four PNGs are produced, one per proxy pair.

Run in PyCharm:
    pip install openpyxl numpy matplotlib
"""

import numpy as np
import openpyxl
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.patches import Patch
from pathlib import Path

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
density_file   = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\30_sec\PCU\05_12_25_1315_1345_Combined_Density_30sec_T1_T2_Pooled_AllMethods.xlsx"

OUT_DIR = Path(density_file).parent

PROXY_RUNS = [
    ("Flow_Snapshot", "M. Mult Acc-Flow_Snap",      "2nd deg Acc-Flow_Snap",      "GPR Acc_Flow_Snap"),
    ("Flow_SMS",      "M. Mult Acc-Flow_SMS",       "2nd deg Acc-Flow_SMS",       "GPR Acc_Flow_SMS"),
    ("Flow_STOPLine", "M. Mult Acc-Flow_STOPLine",  "2nd deg Acc-Flow_STOPLine",  "GPR Acc_Flow_STOPLine"),
    ("Occupancy",     "M. Mult Acc-Occ",            "2nd deg Acc-Occ",            "GPR Acc_Occ"),
]

METHOD_ORDER = ["Mean Multiplier", "2nd-Degree Polynomial", "GPR"]
METHOD_COLORS = {
    "Mean Multiplier":       "#1f77b4",   # blue
    "2nd-Degree Polynomial": "#ff7f0e",   # orange
    "GPR":                   "#2ca02c",   # green
}

TIMESTEP_LINESTYLES = {
    "T1": "solid",
    "T2": (0, (1, 1)),   # finely / densely dotted
}

BAR_EDGE_WIDTH = 1.6


# ---------------------------------------------------------------------------
# Parse a row's time-range label -> (regime, t_start, t_end)
# Handles "T1: 30-60", "T2: 45-75", or plain "30-60" (defaults to "T1").
# ---------------------------------------------------------------------------
def parse_time_range(label):
    s = str(label)
    if ":" in s:
        regime_part, range_part = s.split(":", 1)
        regime = regime_part.strip().upper()
        if regime not in ("T1", "T2"):
            regime = "T1"
    else:
        regime, range_part = "T1", s
    a, b = range_part.strip().split("-")
    return regime, float(a), float(b)


def find_col(headers, *substrings):
    for i, h in enumerate(headers):
        if h is None:
            continue
        hl = str(h).lower()
        if all(sub.lower() in hl for sub in substrings):
            return i
    raise KeyError(f"No column containing {substrings}. Headers: {headers[:12]}")


def load_signed_points(path, sheet_name):
    """-> list of (regime, t_start, t_end, signed_deviation)"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet_name]
    rows = ws.iter_rows(values_only=True)
    headers = list(next(rows))
    acc_col = find_col(headers, "Density_Accumulation")
    cal_col = find_col(headers, "Modified Density")

    out = []
    for row in rows:
        if row[1] is None or row[2] is None:
            continue
        acc, cal = row[acc_col], row[cal_col]
        if acc is None or cal is None:
            continue
        regime, t0, t1 = parse_time_range(row[0])
        out.append((regime, t0, t1, float(cal) - float(acc)))
    wb.close()
    return out


def main():
    for proxy_name, m1, m2, m3 in PROXY_RUNS:
        print(f"\n=== {proxy_name} ===")

        # bins[(regime, t0, t1)][method] = [signed values]
        bins = {}
        all_t0, all_t1 = [], []
        for method, sheet in zip(METHOD_ORDER, (m1, m2, m3)):
            pts = load_signed_points(density_file, sheet)
            print(f"  {method:<24} n={len(pts):3d}")
            for regime, t0, t1, dev in pts:
                key = (regime, t0, t1)
                bins.setdefault(key, {}).setdefault(method, []).append(dev)
                all_t0.append(t0)
                all_t1.append(t1)

        fig, ax = plt.subplots(figsize=(16, 7))

        for key, methods_here in bins.items():
            regime, t0, t1 = key
            width = t1 - t0
            ls = TIMESTEP_LINESTYLES[regime]
            for method, vals in methods_here.items():
                vals = np.array(vals)
                pos = vals[vals > 0]
                neg = vals[vals < 0]
                pos_mean = pos.mean() if pos.size else 0.0
                neg_mean = neg.mean() if neg.size else 0.0
                col = METHOD_COLORS[method]
                if pos_mean != 0.0:
                    ax.bar(t0, pos_mean, width=width, align="edge",
                           facecolor="none", edgecolor=col,
                           linewidth=BAR_EDGE_WIDTH, linestyle=ls, zorder=3)
                if neg_mean != 0.0:
                    ax.bar(t0, neg_mean, width=width, align="edge",
                           facecolor="none", edgecolor=col,
                           linewidth=BAR_EDGE_WIDTH, linestyle=ls, zorder=3)

        ax.axhline(0, color="black", linewidth=1.0, zorder=4)
        ax.set_xlim(min(all_t0) - 10, max(all_t1) + 10)
        ax.set_xlabel("Time (s)", fontsize=11)
        ax.set_ylabel("Mean Signed Deviation from Ideal Density (PCU/km)", fontsize=11)
        ax.set_title(f"Time vs Signed Deviation\n"
                     f"Density (Accumulation) vs Density ({proxy_name})",
                     fontsize=13, fontweight="bold")
        ax.grid(True, axis="y", linestyle=":", alpha=0.5, zorder=0)

        method_handles = [Patch(facecolor="none", edgecolor=METHOD_COLORS[m],
                                 linewidth=BAR_EDGE_WIDTH, label=m)
                          for m in METHOD_ORDER]
        timestep_handles = [Line2D([0], [0], color="black",
                                    linestyle=TIMESTEP_LINESTYLES[t],
                                    linewidth=BAR_EDGE_WIDTH, label=f"Timestep {t[1]}")
                            for t in ("T1", "T2")]
        legend1 = ax.legend(handles=method_handles, loc="upper left",
                             fontsize=9, framealpha=0.9, edgecolor="#ccc",
                             title="Method")
        ax.add_artist(legend1)
        ax.legend(handles=timestep_handles, loc="upper right",
                  fontsize=9, framealpha=0.9, edgecolor="#ccc", title="Timestep")

        out = OUT_DIR / f"Time_vs_SignedDeviation_{proxy_name}.png"
        fig.tight_layout()
        fig.savefig(out, dpi=200, bbox_inches="tight")
        plt.close(fig)
        print(f"  Saved -> {out}")

    print("\nDone.")


if __name__ == "__main__":
    main()