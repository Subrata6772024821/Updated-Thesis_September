import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.cm as cm
import matplotlib.colors as mcolors
from matplotlib.lines import Line2D
from pathlib import Path

# ==========================================================
# STEP 3 — TIME-SPACE DIAGRAM
# X axis : absolute time (seconds) — full 30-minute session
# Y axis : cumulative distance from entry gate (metres)
# Each vehicle = one diagonal line rising from (entry_t, 0)
#                to (exit_t, traveled_dist)
# Steeper slope = faster vehicle
# Flat / vertical segments = stopped in queue
# ==========================================================

input_file = r"D:\Thesis_Final Data Analysis\05_12_25_1400_1430\Trajectory\1step\05_12_25_1400_1430.trajectory_1s_final.xlsx"
output_png = str(Path(input_file).parent / "05_12_25_1400_1430.timespace_diagram.png")

# ----------------------------------------------------------
# OPTIONS
# ----------------------------------------------------------
COLOR_BY        = "speed"     # "speed" | "type" | "vehicle"
SHOW_TRUNCATED  = True         # True = include truncated vehicles
TRUNCATED_ALPHA = 0.35         # opacity for truncated vehicles
COMPLETE_ALPHA  = 0.80         # opacity for complete vehicles
LINE_WIDTH      = 0.9          # line width (thinner looks cleaner with 1500+ vehicles)
FIGURE_DPI      = 200
FIGURE_SIZE     = (18, 8)      # wide format suits time-space diagrams

# Color per vehicle type
TYPE_COLORS = {
    "Car"           : "#2196F3",
    "Motorcycle"    : "#FF9800",
    "Medium Vehicle": "#4CAF50",
    "Heavy Vehicle" : "#9C27B0",
    "Bus"           : "#F44336",
    "Tuk-Tuk"       : "#00BCD4",
}

# ----------------------------------------------------------
print("Loading final trajectory Excel ...")
df = pd.read_excel(input_file, sheet_name="Final_Data")
df.columns = df.columns.str.strip()
print(f"  {len(df):,} rows  |  {df['Track ID'].nunique():,} vehicles")

# Guard: required columns
required = ["Track ID", "Time [s]", "cumulative_distance", "Speed [km/h]"]
missing  = [c for c in required if c not in df.columns]
if missing:
    raise ValueError(f"Missing columns: {missing}")

# Optional: drop truncated
if not SHOW_TRUNCATED and "Truncated" in df.columns:
    before = df["Track ID"].nunique()
    df = df[df["Truncated"] == False].copy()
    print(f"  Removed {before - df['Track ID'].nunique()} truncated vehicles")

df = df.sort_values(["Track ID", "Time [s]"]).reset_index(drop=True)

# ----------------------------------------------------------
# Colormap setup
if COLOR_BY == "speed":
    cmap = cm.plasma
    vmin = df["Speed [km/h]"].quantile(0.02)
    vmax = df["Speed [km/h]"].quantile(0.98)
    norm = mcolors.Normalize(vmin=vmin, vmax=vmax)

elif COLOR_BY == "type":
    pass  # handled per group

elif COLOR_BY == "vehicle":
    unique_ids = df["Track ID"].unique()
    id_cmap    = cm.tab20
    id_norm    = mcolors.Normalize(vmin=0, vmax=max(len(unique_ids) - 1, 1))
    id_map     = {tid: i for i, tid in enumerate(unique_ids)}

# ----------------------------------------------------------
print(f"Plotting Time-Space diagram (color_by='{COLOR_BY}') ...")

fig, ax = plt.subplots(figsize=FIGURE_SIZE, dpi=FIGURE_DPI)
fig.patch.set_facecolor("white")
ax.set_facecolor("white")

grouped = df.groupby("Track ID", sort=False)
n_drawn = 0

for track_id, grp in grouped:
    grp = grp.sort_values("Time [s]")

    # X = absolute time, Y = position along road (cumulative distance)
    t    = grp["Time [s]"].values
    pos  = grp["cumulative_distance"].values
    spd  = grp["Speed [km/h]"].values

    is_trunc = bool(grp["Truncated"].iloc[0]) if "Truncated" in grp.columns else False
    alpha    = TRUNCATED_ALPHA if is_trunc else COMPLETE_ALPHA
    lw       = LINE_WIDTH * 0.7 if is_trunc else LINE_WIDTH

    if COLOR_BY == "speed":
        # Segment-by-segment color
        for i in range(len(t) - 1):
            c = cmap(norm(spd[i]))
            ax.plot([t[i], t[i+1]], [pos[i], pos[i+1]],
                    color=c, linewidth=lw, alpha=alpha,
                    solid_capstyle="round")

    elif COLOR_BY == "type":
        vtype = str(grp["Type"].iloc[0]).strip() if "Type" in grp.columns else "Car"
        color = TYPE_COLORS.get(vtype, "#888888")
        ax.plot(t, pos, color=color, linewidth=lw, alpha=alpha,
                solid_capstyle="round")

    elif COLOR_BY == "vehicle":
        color = id_cmap(id_norm(id_map[track_id]))
        ax.plot(t, pos, color=color, linewidth=lw, alpha=alpha,
                solid_capstyle="round")

    n_drawn += 1

print(f"  Drew {n_drawn:,} vehicle trajectories")

# ----------------------------------------------------------
# Colorbar / legend
if COLOR_BY == "speed":
    sm = cm.ScalarMappable(cmap=cmap, norm=norm)
    sm.set_array([])
    cbar = plt.colorbar(sm, ax=ax, fraction=0.018, pad=0.01)
    cbar.set_label("Speed (km/h)", fontsize=11)
    cbar.outline.set_linewidth(0.5)

elif COLOR_BY == "type":
    types_present = df["Type"].str.strip().unique() if "Type" in df.columns else []
    handles = [Line2D([0],[0], color=TYPE_COLORS.get(t,"#888"), linewidth=2, label=t)
               for t in TYPE_COLORS if t in types_present]
    ax.legend(handles=handles, loc="upper left", fontsize=9,
              title="Vehicle type", title_fontsize=9,
              framealpha=0.9, edgecolor="#ccc")

# Truncated legend entry
if SHOW_TRUNCATED and "Truncated" in df.columns and df["Truncated"].any():
    h1 = Line2D([0],[0], color="#333", linewidth=LINE_WIDTH,
                alpha=COMPLETE_ALPHA, label="Complete")
    h2 = Line2D([0],[0], color="#333", linewidth=LINE_WIDTH * 0.7,
                alpha=TRUNCATED_ALPHA, linestyle="--",
                label="Truncated (recording cutoff)")
    ax.legend(handles=[h1, h2], loc="upper left", fontsize=9,
              framealpha=0.9, edgecolor="#ccc")

# ----------------------------------------------------------
# Axes
ax.set_xlabel("Time [s]", fontsize=13)
ax.set_ylabel("Space [m]", fontsize=13)
ax.tick_params(labelsize=10)

# X axis: full session in minutes ticks
x_max = df["Time [s]"].max()
x_min = df["Time [s]"].min()
ax.set_xlim(0, np.ceil(x_max / 100) * 100)
ax.set_ylim(-2, df["cumulative_distance"].max() * 1.05)

# Dashed vertical grid every 250s (like the reference image)
for xt in np.arange(250, x_max, 250):
    ax.axvline(xt, color="#cccccc", linewidth=0.6, linestyle="--", zorder=0)

ax.grid(axis="y", color="#cccccc", linewidth=0.4, linestyle="--", zorder=0)
ax.set_axisbelow(True)

for spine in ax.spines.values():
    spine.set_linewidth(0.8)
    spine.set_color("#333")

# ----------------------------------------------------------
# Title
n_complete  = (df[df["Truncated"]==False]["Track ID"].nunique()
               if "Truncated" in df.columns else n_drawn)
n_truncated = (df[df["Truncated"]==True]["Track ID"].nunique()
               if "Truncated" in df.columns else 0)

ax.set_title(
    "Space-Time Trajectories of Vehicle\n"
    "Location: Sathon-Convent, Approach: Sathon Road (All lanes combined)\n"
    f"Dec 05, 2025   1400-1430   |   {n_drawn:,} vehicles   |   color: {COLOR_BY}",
    fontsize=13, fontweight="bold", pad=12
)

# ----------------------------------------------------------
plt.tight_layout()
plt.savefig(output_png, dpi=FIGURE_DPI, bbox_inches="tight",
            facecolor="white")
plt.close()

print(f"\n  Saved -> {output_png}")