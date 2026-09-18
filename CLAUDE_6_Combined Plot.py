import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.cm as cm
import matplotlib.colors as mcolors
import matplotlib.gridspec as gridspec
import matplotlib.patches as mpatches
from pathlib import Path

# ==========================================================
# STEP 6 -- COMBINED FIVE-PANEL FIGURE (T1 + T2 overlaid)
#   Panel 1 (top, large) : Time-Space diagram (speed colormap)
#   Panel 2              : Queue length over time
#   Panel 3              : Multiple-stop events (Gantt bars per vehicle)
#   Panel 4              : Density (PCU/km)         -- T1 and T2 bars
#                           OVERLAID in one panel (different colors,
#                           semi-transparent so both are visible where
#                           they overlap)
#   Panel 5 (bottom)     : Green Time Ratio          -- T1 and T2 bars
#                           OVERLAID in one panel, same styling idea
# All panels share the same X axis (Time [s])
#
# CHANGE FROM THE 7-PANEL VERSION:
#   Panels 4+6 (Density T1, Density T2) are now ONE panel with both
#   sets of bars drawn on the same axes at their own real bin edges
#   (T1: e.g. 30-60s bins; T2: e.g. 45-75s bins) -- distinguished only
#   by color + a legend, with alpha transparency so overlapping
#   regions are still readable. Same idea for panels 5+7 (Green Time
#   Ratio). Total panel count: 5 instead of 7.
# ==========================================================

trajectory_file = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\Trajectory\05_12_25_1315_1345.trajectory_1s_final.xlsx"
queue_file      = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\Trajectory\05_12_25_1315_1345.queue_length.xlsx"
stop_file       = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\Trajectory\05_12_25_1315_1345.stop_analysis.xlsx"
density_file    = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\30_sec\Accumulation\Density_Estimation_ActiveCount_17_1802.xlsx"
greentime_file  = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\GreenTime_GA_Dec05_1315.xlsx"

STOP_LINE_POS = 96.0   # metres
FIGURE_DPI    = 200
FIGURE_SIZE   = (22, 22)   # back down from the 7-panel height, 5 panels now

DENSITY_COL = "Density (PCU/km)"

# One entry per timestep to overlay: (sheet name in the density
# workbook, short tag used in labels/legend, display color, line style)
# Bars are drawn UNFILLED (outline only) so overlapping T1/T2 bars are
# both fully visible rather than one color blending into/covering the
# other -- T1 is a solid outline, T2 a dashed outline, in different
# colors, so the two are distinguishable even where they overlap.
T1_SHEET, T1_TAG, T1_COLOR, T1_LINESTYLE = "30s Buckets",          "T1", "#1f77b4", "-"    # blue, solid
T2_SHEET, T2_TAG, T2_COLOR, T2_LINESTYLE = "Sheet3 Custom Buckets", "T2", "#d62728", "--"   # red, dashed
BAR_LINEWIDTH = 1.6

# Stop event colors per stop number
STOP_COLORS = {1: "#2196F3", 2: "#FF9800", 3: "#4CAF50", 4: "#9C27B0", 5: "#F44336"}

# ----------------------------------------------------------
# LOAD DATA
# ----------------------------------------------------------
print("Loading trajectory data ...")
df_traj = pd.read_excel(trajectory_file, sheet_name="Final_Data")
df_traj.columns = df_traj.columns.str.strip()
if "Reconstructed" in df_traj.columns:
    df_traj = df_traj[df_traj["Reconstructed"] == False].copy()
df_traj = df_traj.sort_values(["Track ID", "Time [s]"]).reset_index(drop=True)
print(f"  {df_traj['Track ID'].nunique():,} vehicles")

print("Loading queue length data ...")
df_queue = pd.read_excel(queue_file, sheet_name="Queue Per Second")
df_queue.columns = df_queue.columns.str.strip()

print("Loading stop analysis data ...")
df_stops_wide = pd.read_excel(stop_file, sheet_name="Multi-Stop Vehicles")
df_stops_wide.columns = df_stops_wide.columns.str.strip()

# Parse wide-format stop columns -> long format
stop_events = []
for _, row in df_stops_wide.iterrows():
    tid   = row["Track ID"]
    vtype = row["Type"] if "Type" in row else "Unknown"
    n_stops = int(row["Total Stops"])
    for n in range(1, n_stops + 1):
        start_col = f"Stop {n} - Start [s]"
        end_col   = f"Stop {n} - End [s]"
        dur_col   = f"Stop {n} - Duration [s]"
        if start_col in row and not pd.isna(row[start_col]):
            stop_events.append({
                "Track ID"    : tid,
                "Type"        : vtype,
                "Stop #"      : n,
                "Start [s]"   : float(row[start_col]),
                "End [s]"     : float(row[end_col]),
                "Duration [s]": float(row[dur_col]),
            })

df_stops_long = pd.DataFrame(stop_events)
multi_ids = sorted(df_stops_wide["Track ID"].tolist())
print(f"  {len(multi_ids)} multi-stop vehicles  |  {len(df_stops_long)} stop events")

# Common time axis
t_min = int(df_traj["Time [s]"].min())
t_max = int(df_traj["Time [s]"].max())


# ----------------------------------------------------------
# DENSITY DATA -- loading helper (per sheet/timestep)
# ----------------------------------------------------------
def load_density(path, sheet_name, density_col=DENSITY_COL):
    df = pd.read_excel(path, sheet_name=sheet_name)
    df.columns = df.columns.str.strip()
    df["t_start"] = df["Time (s)"].str.split("-").str[0].astype(float)
    df["t_end"]   = df["Time (s)"].str.split("-").str[1].astype(float)
    df["t_mid"]   = (df["t_start"] + df["t_end"]) / 2
    df["bin_width"] = df["t_end"] - df["t_start"]
    if density_col not in df.columns:
        raise KeyError(
            f"Column '{density_col}' not found in sheet '{sheet_name}'.\n"
            f"Available columns: {list(df.columns)}"
        )
    return df


# ----------------------------------------------------------
# GREEN TIME DATA -- loading + binning helper
# "Green Phases" sheet: 4 metadata/title rows, then a header row, then
# one row per green phase interval. Bins are computed by overlap with
# whichever density sheet's bin edges are passed in -- T1 and T2 do
# NOT share the same bin boundaries, so this is called once per
# timestep with that timestep's own edges.
# ----------------------------------------------------------
GREENTIME_SHEET = "Green Phases"
GREENTIME_HEADER_ROW = 4   # 0-indexed row containing the real column names


def load_green_time(path, bin_edges_start, bin_edges_end,
                     sheet=GREENTIME_SHEET, header_row=GREENTIME_HEADER_ROW):
    df = pd.read_excel(path, sheet_name=sheet, header=header_row)
    df.columns = df.columns.str.strip()
    df = df.dropna(subset=["Green_Start_s", "Green_End_s"]).copy()
    df["Green_Start_s"] = df["Green_Start_s"].astype(float)
    df["Green_End_s"]   = df["Green_End_s"].astype(float)

    print(f"  {len(df)} green phases loaded "
          f"(total green = {(df['Green_End_s'] - df['Green_Start_s']).sum():.0f}s)")

    green_starts = df["Green_Start_s"].values
    green_ends   = df["Green_End_s"].values

    bin_green_secs = []
    for b_start, b_end in zip(bin_edges_start, bin_edges_end):
        overlap = np.maximum(
            0.0,
            np.minimum(b_end, green_ends) - np.maximum(b_start, green_starts)
        )
        bin_green_secs.append(overlap.sum())

    out = pd.DataFrame({
        "t_start": bin_edges_start,
        "t_end":   bin_edges_end,
        "Green Time [s]": bin_green_secs,
    })
    out["t_mid"]  = (out["t_start"] + out["t_end"]) / 2
    out["bin_width"] = out["t_end"] - out["t_start"]
    out["Green Ratio"] = out["Green Time [s]"] / out["bin_width"]
    return out


print(f"\nLoading density data ({T1_TAG}: '{T1_SHEET}') ...")
df_dens_t1 = load_density(density_file, T1_SHEET)
print(f"Loading green time data ({T1_TAG}, binned to its own windows) ...")
df_green_t1 = load_green_time(greentime_file, df_dens_t1["t_start"].values, df_dens_t1["t_end"].values)

print(f"\nLoading density data ({T2_TAG}: '{T2_SHEET}') ...")
df_dens_t2 = load_density(density_file, T2_SHEET)
print(f"Loading green time data ({T2_TAG}, binned to its own windows) ...")
df_green_t2 = load_green_time(greentime_file, df_dens_t2["t_start"].values, df_dens_t2["t_end"].values)


# ----------------------------------------------------------
# Vertical grid lines helper
# ----------------------------------------------------------
def add_vgrid(ax):
    for xt in np.arange(250, t_max, 250):
        ax.axvline(xt, color="#cccccc", linewidth=0.5, linestyle="--", zorder=0)


# ----------------------------------------------------------
# FIGURE LAYOUT -- 5 rows sharing the X axis
# ----------------------------------------------------------
fig = plt.figure(figsize=FIGURE_SIZE, dpi=FIGURE_DPI)
fig.patch.set_facecolor("white")

gs = gridspec.GridSpec(
    5, 1,
    height_ratios=[3.0, 1.0, 1.2, 1.0, 1.0],
    hspace=0.09,
)
ax1 = fig.add_subplot(gs[0])                      # Time-Space
ax2 = fig.add_subplot(gs[1], sharex=ax1)          # Queue length
ax3 = fig.add_subplot(gs[2], sharex=ax1)          # Multi-stop Gantt
ax4 = fig.add_subplot(gs[3], sharex=ax1)          # Density (T1 + T2 overlaid)
ax5 = fig.add_subplot(gs[4], sharex=ax1)          # Green Time Ratio (T1 + T2 overlaid)

# ----------------------------------------------------------
# PANEL 1 -- TIME-SPACE DIAGRAM
# ----------------------------------------------------------
ax1.set_facecolor("#f8f8f8")

vmin = df_traj["Speed [km/h]"].quantile(0.02)
vmax = df_traj["Speed [km/h]"].quantile(0.98)
norm = mcolors.Normalize(vmin=vmin, vmax=vmax)
cmap = cm.plasma

for track_id, grp in df_traj.groupby("Track ID", sort=False):
    grp    = grp.sort_values("Time [s]")
    t      = grp["Time [s]"].values
    pos    = grp["cumulative_distance"].values
    spd    = grp["Speed [km/h]"].values
    is_multi = track_id in multi_ids
    lw     = 1.4 if is_multi else 0.6
    alpha  = 0.85 if is_multi else 0.50
    for i in range(len(t) - 1):
        ax1.plot([t[i], t[i+1]], [pos[i], pos[i+1]],
                 color=cmap(norm(spd[i])),
                 linewidth=lw, alpha=alpha, solid_capstyle="round")

# Queue zone shading
ax1.fill_between(
    df_queue["Time [s]"],
    STOP_LINE_POS - df_queue["Queue Length [m]"],
    STOP_LINE_POS,
    where=df_queue["Queue Length [m]"] > 0,
    color="#CC000018", zorder=0, label="Queue zone"
)
ax1.axhline(STOP_LINE_POS, color="#CC0000", linewidth=1.0,
            linestyle="--", alpha=0.7, label=f"Stop line ({STOP_LINE_POS:.0f} m)")

# Colorbar
sm = cm.ScalarMappable(cmap=cmap, norm=norm)
sm.set_array([])
cbar = plt.colorbar(sm, ax=ax1, fraction=0.010, pad=0.01)
cbar.set_label("Speed (km/h)", fontsize=8)
cbar.outline.set_linewidth(0.4)
cbar.ax.tick_params(labelsize=7)

ax1.set_ylabel("Space [m]", fontsize=11)
ax1.set_ylim(-2, STOP_LINE_POS * 1.1)
ax1.legend(loc="upper left", fontsize=8, framealpha=0.9, edgecolor="#ccc")
ax1.grid(axis="y", color="#dddddd", linewidth=0.4, linestyle="--", zorder=0)
ax1.set_axisbelow(True)
ax1.tick_params(labelbottom=False, labelsize=9)
ax1.set_title(
    "Space-Time Trajectories  |  Queue Length  |  Multiple Stops  |  "
    f"Density ({T1_TAG} + {T2_TAG} overlaid)  |  Green Time Ratio ({T1_TAG} + {T2_TAG} overlaid)\n"
    "Location: Sathon-Convent, Approach: Sathon Road (All lanes combined)  |  "
    "Dec 05, 2025   13:15 - 13:45",
    fontsize=12, fontweight="bold", pad=8
)
add_vgrid(ax1)

# ----------------------------------------------------------
# PANEL 2 -- QUEUE LENGTH
# ----------------------------------------------------------
ax2.set_facecolor("white")

ax2.fill_between(
    df_queue["Time [s]"],
    df_queue["Queue Length [m]"],
    color="#CC000028", step="post", zorder=1
)
ax2.step(
    df_queue["Time [s]"],
    df_queue["Queue Length [m]"],
    color="#CC0000", linewidth=1.2, where="post",
    label="Queue length [m]", zorder=2
)

q_active = df_queue[df_queue["Queue Length [m]"] > 0]["Queue Length [m]"]
if len(q_active) > 0:
    mean_q = q_active.mean()
    ax2.axhline(mean_q, color="#FF6600", linewidth=1.0, linestyle=":",
                alpha=0.85, label=f"Mean: {mean_q:.1f} m", zorder=3)
    idx_max = df_queue["Queue Length [m]"].idxmax()
    t_peak  = df_queue.loc[idx_max, "Time [s]"]
    q_peak  = df_queue.loc[idx_max, "Queue Length [m]"]
    ax2.annotate(
        f" Max:{q_peak:.1f}m",
        xy=(t_peak, q_peak), xytext=(t_peak + 30, q_peak * 0.85),
        fontsize=7, color="#990000",
        arrowprops=dict(arrowstyle="->", color="#990000", lw=0.7)
    )

ax2.set_ylabel("Queue\n[m]", fontsize=10)
ax2.set_ylim(0, STOP_LINE_POS * 1.1)
ax2.legend(loc="upper right", fontsize=8, framealpha=0.9, edgecolor="#ccc")
ax2.grid(color="#eeeeee", linewidth=0.4, zorder=0)
ax2.set_axisbelow(True)
ax2.tick_params(labelbottom=False, labelsize=9)
add_vgrid(ax2)

# ----------------------------------------------------------
# PANEL 3 -- MULTIPLE STOPS (Gantt chart)
# ----------------------------------------------------------
ax3.set_facecolor("#fafafa")

y_ticks   = list(range(len(multi_ids)))
y_labels  = [str(tid) for tid in multi_ids]
y_map     = {tid: i for i, tid in enumerate(multi_ids)}

for _, ev in df_stops_long.iterrows():
    tid      = ev["Track ID"]
    y_pos    = y_map[tid]
    t_start  = ev["Start [s]"]
    t_end    = ev["End [s]"]
    stop_n   = int(ev["Stop #"])
    color    = STOP_COLORS.get(stop_n, "#888888")

    ax3.barh(
        y_pos,
        width  = max(t_end - t_start, 2.0),
        left   = t_start,
        height = 0.55,
        color  = color,
        edgecolor = "white",
        linewidth = 0.4,
        zorder = 2,
    )
    if (t_end - t_start) > 15:
        ax3.text(
            t_start + (t_end - t_start) / 2, y_pos,
            f"S{stop_n}", ha="center", va="center",
            fontsize=6.5, color="white", fontweight="bold", zorder=3
        )

ax3.set_yticks(y_ticks)
ax3.set_yticklabels(y_labels, fontsize=7.5)
ax3.set_ylabel("Track ID\n(multi-stop)", fontsize=10)
ax3.set_ylim(-0.7, len(multi_ids) - 0.3)
ax3.tick_params(labelbottom=False, labelsize=9)
ax3.grid(axis="x", color="#eeeeee", linewidth=0.4, zorder=0)
ax3.set_axisbelow(True)
for spine in ax3.spines.values():
    spine.set_linewidth(0.5)
add_vgrid(ax3)

legend_handles = [
    mpatches.Patch(color=STOP_COLORS[n], label=f"Stop {n}")
    for n in sorted(STOP_COLORS.keys())
    if n in df_stops_long["Stop #"].values
]
ax3.legend(handles=legend_handles, loc="upper right",
           fontsize=8, framealpha=0.9, edgecolor="#ccc",
           title="Stop event", title_fontsize=8)

# ----------------------------------------------------------
# PANEL 4 -- DENSITY, T1 and T2 OVERLAID (unfilled/outline bars)
# Both timesteps drawn on the SAME axes at their own real bin edges
# (different widths/offsets, e.g. T1: 30-60s bins, T2: 45-75s bins).
# Bars are unfilled (facecolor="none") so both outlines stay fully
# visible wherever they overlap -- T1 solid outline, T2 dashed
# outline, in different colors.
# ----------------------------------------------------------
ax4.set_facecolor("white")

ax4.bar(
    df_dens_t1["t_start"], df_dens_t1[DENSITY_COL],
    width=df_dens_t1["bin_width"], align="edge",
    facecolor="none", edgecolor=T1_COLOR, linewidth=BAR_LINEWIDTH,
    linestyle=T1_LINESTYLE, zorder=2, label=f"Density ({T1_TAG})",
)
ax4.bar(
    df_dens_t2["t_start"], df_dens_t2[DENSITY_COL],
    width=df_dens_t2["bin_width"], align="edge",
    facecolor="none", edgecolor=T2_COLOR, linewidth=BAR_LINEWIDTH,
    linestyle=T2_LINESTYLE, zorder=3, label=f"Density ({T2_TAG})",
)

mean_d1 = df_dens_t1[DENSITY_COL].mean()
mean_d2 = df_dens_t2[DENSITY_COL].mean()
ax4.axhline(mean_d1, color=T1_COLOR, linewidth=1.1, linestyle=":",
            alpha=0.9, zorder=3, label=f"Mean {T1_TAG}: {mean_d1:.1f} PCU/km")
ax4.axhline(mean_d2, color=T2_COLOR, linewidth=1.1, linestyle=":",
            alpha=0.9, zorder=3, label=f"Mean {T2_TAG}: {mean_d2:.1f} PCU/km")

d_max = max(df_dens_t1[DENSITY_COL].max(), df_dens_t2[DENSITY_COL].max())
ax4.set_ylabel("Density\n[PCU/km]", fontsize=10)
ax4.set_ylim(0, d_max * 1.2)
ax4.legend(loc="upper right", fontsize=7.5, framealpha=0.9, edgecolor="#ccc", ncol=2)
ax4.grid(axis="y", color="#eeeeee", linewidth=0.4, zorder=0)
ax4.set_axisbelow(True)
ax4.tick_params(labelbottom=False, labelsize=9)
add_vgrid(ax4)

# ----------------------------------------------------------
# PANEL 5 -- GREEN TIME RATIO, T1 and T2 OVERLAID  (bottom: x-axis shown)
# Same unfilled/outline overlay approach as Panel 4.
# ----------------------------------------------------------
ax5.set_facecolor("white")

ax5.bar(
    df_green_t1["t_start"], df_green_t1["Green Ratio"],
    width=df_green_t1["bin_width"], align="edge",
    facecolor="none", edgecolor=T1_COLOR, linewidth=BAR_LINEWIDTH,
    linestyle=T1_LINESTYLE, zorder=2, label=f"Green Ratio ({T1_TAG})",
)
ax5.bar(
    df_green_t2["t_start"], df_green_t2["Green Ratio"],
    width=df_green_t2["bin_width"], align="edge",
    facecolor="none", edgecolor=T2_COLOR, linewidth=BAR_LINEWIDTH,
    linestyle=T2_LINESTYLE, zorder=3, label=f"Green Ratio ({T2_TAG})",
)

mean_g1 = df_green_t1["Green Ratio"].mean()
mean_g2 = df_green_t2["Green Ratio"].mean()
ax5.axhline(mean_g1, color=T1_COLOR, linewidth=1.1, linestyle=":",
            alpha=0.9, zorder=3, label=f"Mean {T1_TAG}: {mean_g1:.2f}")
ax5.axhline(mean_g2, color=T2_COLOR, linewidth=1.1, linestyle=":",
            alpha=0.9, zorder=3, label=f"Mean {T2_TAG}: {mean_g2:.2f}")

ax5.set_ylabel("Green Time\nRatio [-]", fontsize=10)
ax5.set_ylim(0, 1.05)
ax5.set_xlabel("Time [s]", fontsize=11)
ax5.legend(loc="upper right", fontsize=7.5, framealpha=0.9, edgecolor="#ccc", ncol=2)
ax5.grid(axis="y", color="#eeeeee", linewidth=0.4, zorder=0)
ax5.set_axisbelow(True)
ax5.tick_params(labelsize=9)
add_vgrid(ax5)

# ----------------------------------------------------------
# SHARED X AXIS LIMITS
# ----------------------------------------------------------
ax1.set_xlim(t_min, t_max)

# ----------------------------------------------------------
# SAVE
# ----------------------------------------------------------
output_png = str(
    Path(trajectory_file).parent / "05_12_25_1315_1345.combined_5panel_T1_T2_overlaid.png"
)
plt.savefig(output_png, dpi=FIGURE_DPI, bbox_inches="tight", facecolor="white")
plt.close(fig)
print(f"\nDone. Saved combined 5-panel figure (T1 + T2 overlaid) -> {output_png}")