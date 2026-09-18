import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import matplotlib.cm as cm
from pathlib import Path

# ==========================================================
# STEP 5 — QUEUE LENGTH DETECTION
#
# Concept:
#   At every second T, a vehicle is "in queue" if:
#       speed < QUEUE_SPEED_KMH   (moving very slowly or stopped)
#   Queue length at time T =
#       STOP_LINE_POS − position of the rearmost queued vehicle
#   If no vehicles are queued → queue length = 0
#
# Outputs (saved to same folder as input):
#   S.queue_length.xlsx  — per-second queue length table + summary stats
#   S.queue_length.png   — two-panel plot: Time-Space + queue length over time
# ==========================================================

input_file   = r"D:\Thesis_Final Data Analysis\04_12_25_1715_1745\Trajectory\04_12_25_1715_1745.trajectory_1s_final.xlsx"
output_excel = str(Path(input_file).parent / "04_12_25_1715_1745.queue_length.xlsx")
output_png   = str(Path(input_file).parent / "04_12_25_1715_1745.queue_length.png")

# ----------------------------------------------------------
# PARAMETERS
# ----------------------------------------------------------
QUEUE_SPEED_KMH = 4.5    # km/h — pedestrian walking speed threshold (Ashqer et al., 2022)
STOP_LINE_POS   = 96.0   # metres — position of exit gate / stop bar
                          # (= gate-to-gate distance, ~96m from your data)
FIGURE_DPI      = 200
FIGURE_SIZE     = (18, 10)

# ----------------------------------------------------------
print("Loading final trajectory Excel ...")
df = pd.read_excel(input_file, sheet_name="Final_Data")
df.columns = df.columns.str.strip()

# Exclude reconstructed vehicles (no real speed/position data)
if "Reconstructed" in df.columns:
    df = df[df["Reconstructed"] == False].copy()

df = df.sort_values(["Track ID", "Time [s]"]).reset_index(drop=True)
print(f"  {df['Track ID'].nunique():,} vehicles  |  {len(df):,} rows")

# ----------------------------------------------------------
# QUEUE LENGTH CALCULATION — per second
# ----------------------------------------------------------
print(f"\nCalculating queue length  (queue speed threshold: < {QUEUE_SPEED_KMH} km/h) ...")

t_min = int(np.ceil(df["Time [s]"].min()))
t_max = int(np.floor(df["Time [s]"].max()))
time_grid = np.arange(t_min, t_max + 1, 1)

queue_records = []

for T in time_grid:
    # All vehicles present at this second
    snapshot = df[df["Time [s]"] == T]

    if snapshot.empty:
        queue_records.append({
            "Time [s]"             : T,
            "Queue Length [m]"     : 0.0,
            "Queued Vehicles"      : 0,
            "Total Vehicles"       : 0,
            "Rearmost Position [m]": np.nan,
            "Queued Track IDs"     : "",
        })
        continue

    total_vehicles = len(snapshot)

    # Vehicles in queue: speed below threshold
    in_queue = snapshot[snapshot["Speed [km/h]"] < QUEUE_SPEED_KMH]
    n_queued = len(in_queue)

    if n_queued == 0:
        queue_records.append({
            "Time [s]"             : T,
            "Queue Length [m]"     : 0.0,
            "Queued Vehicles"      : 0,
            "Total Vehicles"       : total_vehicles,
            "Rearmost Position [m]": np.nan,
            "Queued Track IDs"     : "",
        })
        continue

    # Rearmost queued vehicle = lowest cumulative_distance (furthest from stop line)
    rearmost_pos   = in_queue["cumulative_distance"].min()
    queue_length   = max(0.0, STOP_LINE_POS - rearmost_pos)
    queued_ids     = ",".join(str(int(x)) for x in sorted(in_queue["Track ID"].tolist()))

    queue_records.append({
        "Time [s]"             : T,
        "Queue Length [m]"     : round(queue_length, 2),
        "Queued Vehicles"      : n_queued,
        "Total Vehicles"       : total_vehicles,
        "Rearmost Position [m]": round(rearmost_pos, 2),
        "Queued Track IDs"     : queued_ids,
    })

df_queue = pd.DataFrame(queue_records)

# ----------------------------------------------------------
# SUMMARY STATISTICS
# ----------------------------------------------------------
q_nonzero = df_queue[df_queue["Queue Length [m]"] > 0]["Queue Length [m]"]

summary = {
    "Metric": [
        "Observation period",
        "Queue speed threshold",
        "Total seconds observed",
        "Seconds with active queue",
        "Queue occupancy (%)",
        "Mean queue length (when active)",
        "Max queue length",
        "Time of max queue length",
        "Mean queued vehicles (when active)",
        "Max queued vehicles at one instant",
    ],
    "Value": [
        f"{t_min}s – {t_max}s  ({(t_max-t_min)/60:.1f} min)",
        f"< {QUEUE_SPEED_KMH} km/h",
        f"{len(df_queue):,}",
        f"{len(q_nonzero):,}",
        f"{len(q_nonzero)/len(df_queue)*100:.1f}%",
        f"{q_nonzero.mean():.1f} m"     if len(q_nonzero) > 0 else "n/a",
        f"{q_nonzero.max():.1f} m"      if len(q_nonzero) > 0 else "n/a",
        f"{df_queue.loc[df_queue['Queue Length [m]'].idxmax(), 'Time [s]']:.0f}s" if len(q_nonzero) > 0 else "n/a",
        f"{df_queue[df_queue['Queued Vehicles']>0]['Queued Vehicles'].mean():.1f}" if len(q_nonzero) > 0 else "n/a",
        f"{df_queue['Queued Vehicles'].max()}",
    ]
}
df_summary = pd.DataFrame(summary)

print(f"\n  Queue statistics:")
for _, row in df_summary.iterrows():
    print(f"    {row['Metric']:45s}: {row['Value']}")

# ----------------------------------------------------------
# SAVE EXCEL
# ----------------------------------------------------------
print(f"\nSaving Excel -> {output_excel}")
with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
    df_queue.to_excel(writer,   sheet_name="Queue Per Second", index=False)
    df_summary.to_excel(writer, sheet_name="Summary",          index=False)

    wb = writer.book
    hdr_fmt = wb.add_format({
        "bold": True, "bg_color": "#D9E1F2",
        "border": 1,  "align": "center"
    })
    num_fmt = wb.add_format({"num_format": "0.00", "align": "center"})
    int_fmt = wb.add_format({"num_format": "0",    "align": "center"})

    # Queue Per Second sheet formatting
    ws1 = writer.sheets["Queue Per Second"]
    col_widths = [10, 18, 16, 16, 22, 40]
    for i, (col, w) in enumerate(zip(df_queue.columns, col_widths)):
        ws1.set_column(i, i, w)
        ws1.write(0, i, col, hdr_fmt)
    ws1.freeze_panes(1, 0)

    # Conditional color: highlight rows where queue > 0
    highlight_fmt = wb.add_format({"bg_color": "#FFF2CC"})
    ws1.conditional_format(
        1, 1, len(df_queue), 1,
        {"type": "cell", "criteria": ">", "value": 0, "format": highlight_fmt}
    )

    # Summary sheet
    ws2 = writer.sheets["Summary"]
    ws2.set_column(0, 0, 45)
    ws2.set_column(1, 1, 30)
    for i, col in enumerate(df_summary.columns):
        ws2.write(0, i, col, hdr_fmt)

print("  Saved.")

# ----------------------------------------------------------
# PLOT: two-panel figure
# ----------------------------------------------------------
print(f"\nPlotting -> {output_png}")

fig, (ax1, ax2) = plt.subplots(
    2, 1, figsize=FIGURE_SIZE, dpi=FIGURE_DPI,
    gridspec_kw={"height_ratios": [2, 1]}, sharex=True
)
fig.patch.set_facecolor("white")

# ── TOP PANEL: Time-Space diagram with queue shading ──────
ax1.set_facecolor("#f8f8f8")

# Speed colormap
vmin = df["Speed [km/h]"].quantile(0.02)
vmax = df["Speed [km/h]"].quantile(0.98)
norm = mcolors.Normalize(vmin=vmin, vmax=vmax)
cmap = cm.plasma

# Draw all vehicle trajectories
for track_id, grp in df.groupby("Track ID", sort=False):
    grp = grp.sort_values("Time [s]")
    t   = grp["Time [s]"].values
    pos = grp["cumulative_distance"].values
    spd = grp["Speed [km/h]"].values
    for i in range(len(t) - 1):
        ax1.plot([t[i], t[i+1]], [pos[i], pos[i+1]],
                 color=cmap(norm(spd[i])),
                 linewidth=0.7, alpha=0.6, solid_capstyle="round")

# Shade queue region at each second
ax1.fill_between(
    df_queue["Time [s]"],
    STOP_LINE_POS - df_queue["Queue Length [m]"],
    STOP_LINE_POS,
    where=df_queue["Queue Length [m]"] > 0,
    color="#CC000022", label="Queue zone"
)

# Stop line
ax1.axhline(STOP_LINE_POS, color="#CC0000", linewidth=1.2,
            linestyle="--", alpha=0.8, label=f"Stop line ({STOP_LINE_POS}m)")

# Colorbar
sm = cm.ScalarMappable(cmap=cmap, norm=norm)
sm.set_array([])
cbar = plt.colorbar(sm, ax=ax1, fraction=0.015, pad=0.01)
cbar.set_label("Speed (km/h)", fontsize=10)
cbar.outline.set_linewidth(0.5)

ax1.set_ylabel("Space [m]", fontsize=12)
ax1.set_ylim(-2, STOP_LINE_POS * 1.08)
ax1.legend(loc="upper left", fontsize=9, framealpha=0.9)
ax1.grid(axis="y", color="#cccccc", linewidth=0.4, linestyle="--", zorder=0)
ax1.set_axisbelow(True)
for xt in np.arange(250, t_max, 250):
    ax1.axvline(xt, color="#cccccc", linewidth=0.5, linestyle="--", zorder=0)
ax1.set_title(
    "Time-Space Diagram with Queue Zone\n"
    "Location: Sathon-Convent, Approach: Sathon Road  |  Dec 04, 2025  1715–1745",
    fontsize=12, fontweight="bold"
)

# ── BOTTOM PANEL: Queue length over time ──────────────────
ax2.set_facecolor("white")

ax2.fill_between(
    df_queue["Time [s]"],
    df_queue["Queue Length [m]"],
    color="#CC000033", step="post"
)
ax2.step(
    df_queue["Time [s]"],
    df_queue["Queue Length [m]"],
    color="#CC0000", linewidth=1.2, where="post",
    label="Queue length [m]"
)

# Mark maximum
idx_max = df_queue["Queue Length [m]"].idxmax()
t_peak  = df_queue.loc[idx_max, "Time [s]"]
q_peak  = df_queue.loc[idx_max, "Queue Length [m]"]
ax2.annotate(
    f"  Max: {q_peak:.1f}m\n  t={t_peak:.0f}s",
    xy=(t_peak, q_peak),
    xytext=(t_peak + 30, q_peak * 0.85),
    fontsize=8, color="#990000",
    arrowprops=dict(arrowstyle="->", color="#990000", lw=0.8)
)

# Mean line
mean_q = q_nonzero.mean() if len(q_nonzero) > 0 else 0
ax2.axhline(mean_q, color="#FF6600", linewidth=1.0,
            linestyle=":", alpha=0.8,
            label=f"Mean (active periods): {mean_q:.1f}m")

ax2.set_xlabel("Time [s]", fontsize=12)
ax2.set_ylabel("Queue length [m]", fontsize=12)
ax2.set_xlim(t_min, t_max)
ax2.set_ylim(0, STOP_LINE_POS * 1.1)
ax2.legend(loc="upper right", fontsize=9, framealpha=0.9)
ax2.grid(color="#eeeeee", linewidth=0.5, zorder=0)
ax2.set_axisbelow(True)
for xt in np.arange(250, t_max, 250):
    ax2.axvline(xt, color="#cccccc", linewidth=0.5, linestyle="--", zorder=0)
ax2.set_title(
    f"Queue Length over Time  "
    f"(threshold: speed < {QUEUE_SPEED_KMH} km/h  |  "
    f"max: {q_peak:.1f}m  |  "
    f"mean active: {mean_q:.1f}m)",
    fontsize=11, fontweight="bold"
)

plt.tight_layout()
plt.savefig(output_png, dpi=FIGURE_DPI, bbox_inches="tight", facecolor="white")
plt.close()
print(f"  Saved.")
print(f"\n  Done.")
print(f"   Excel : {output_excel}")
print(f"   Plot  : {output_png}")