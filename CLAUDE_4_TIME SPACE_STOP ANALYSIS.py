import pandas as pd
import numpy as np
from pathlib import Path

# ==========================================================
# STEP 4 — STOP DETECTION
#
# Definition of a STOP (parallel to X-axis on Time-Space diagram):
#   Speed < STOP_SPEED_KMH  for at least  MIN_STOP_DURATION  consecutive seconds
#
# Output Excel (same folder as input):
#   Sheet "Multi-Stop Vehicles" — only vehicles that stop 2 or more times
#   Columns:
#     Track ID | Type | Total Stops | Stop 1 Start [s] | Stop 1 End [s] | Stop 1 Duration [s]
#              | Stop 2 Start [s] | Stop 2 End [s] | Stop 2 Duration [s] | ...
# ==========================================================

input_file  = r"D:\Thesis_Final Data Analysis\04_12_25_1715_1745\Trajectory\04_12_25_1715_1745.trajectory_1s_final.xlsx"
output_file = str(Path(input_file).parent / "04_12_25_1715_1745.stop_analysis.xlsx")

# ----------------------------------------------------------
# PARAMETERS
# ----------------------------------------------------------
STOP_SPEED_KMH    = 2.0   # km/h  — complete cessation of movement (two stops require genuine re-acceleration between them)
MIN_STOP_DURATION = 3     # seconds — must stay stopped for at least this long
                           # (avoids flagging brief deceleration as a real stop)
MIN_STOP_COUNT    = 2     # only keep vehicles with this many stops or more

# ----------------------------------------------------------
print("Loading final trajectory Excel ...")
df = pd.read_excel(input_file, sheet_name="Final_Data")
df.columns = df.columns.str.strip()

# Exclude reconstructed vehicles (no real speed data)
if "Reconstructed" in df.columns:
    df = df[df["Reconstructed"] == False].copy()

df = df.sort_values(["Track ID", "Time [s]"]).reset_index(drop=True)
print(f"  {df['Track ID'].nunique():,} vehicles loaded")

# ----------------------------------------------------------
# STOP DETECTION — per vehicle
# ----------------------------------------------------------
print(f"\nDetecting stops  "
      f"(speed < {STOP_SPEED_KMH} km/h  for >= {MIN_STOP_DURATION} consecutive seconds) ...")

results        = []   # vehicles with >= MIN_STOP_COUNT stops
single_results = []   # vehicles with exactly 1 stop
stop_records   = []   # every individual stop event (for plotting)

for track_id, grp in df.groupby("Track ID"):
    grp    = grp.sort_values("Time [s]").reset_index(drop=True)
    times  = grp["Time [s]"].values
    speeds = grp["Speed [km/h]"].values
    vtype  = grp["Type"].iloc[0] if "Type" in grp.columns else "Unknown"

    is_stopped = speeds < STOP_SPEED_KMH

    # ── Find consecutive stop blocks ──
    stop_events = []
    in_stop     = False
    start_idx   = 0

    for i in range(len(is_stopped)):
        if is_stopped[i] and not in_stop:
            # vehicle just stopped
            in_stop   = True
            start_idx = i

        elif not is_stopped[i] and in_stop:
            # vehicle just moved again
            in_stop  = False
            duration = times[i - 1] - times[start_idx]
            if duration >= MIN_STOP_DURATION:
                stop_events.append((
                    round(times[start_idx], 1),   # start time
                    round(times[i - 1],    1),   # end time
                    round(duration,        1),   # duration
                ))
                stop_records.append({
                    "Track ID"      : track_id,
                    "Type"          : vtype,
                    "Stop #"        : len(stop_events),
                    "Start time [s]": round(times[start_idx], 1),
                    "End time [s]"  : round(times[i - 1],    1),
                    "Duration [s]"  : round(duration,        1),
                })

    # Handle stop that persists until end of trajectory
    if in_stop:
        duration = times[-1] - times[start_idx]
        if duration >= MIN_STOP_DURATION:
            stop_events.append((
                round(times[start_idx], 1),
                round(times[-1],        1),
                round(duration,         1),
            ))
            stop_records.append({
                "Track ID"      : track_id,
                "Type"          : vtype,
                "Stop #"        : len(stop_events),
                "Start time [s]": round(times[start_idx], 1),
                "End time [s]"  : round(times[-1],        1),
                "Duration [s]"  : round(duration,         1),
            })

    # Build row dict for this vehicle
    row = {
        "Track ID"   : track_id,
        "Type"       : vtype,
        "Total Stops": len(stop_events),
    }
    for n, (t_start, t_end, t_dur) in enumerate(stop_events, start=1):
        row[f"Stop {n} - Start [s]"   ] = t_start
        row[f"Stop {n} - End [s]"     ] = t_end
        row[f"Stop {n} - Duration [s]"] = t_dur

    # ── Route to multi-stop or single-stop list ──
    if len(stop_events) >= MIN_STOP_COUNT:
        results.append(row)
    elif len(stop_events) == 1:
        single_results.append(row)

# ----------------------------------------------------------
# BUILD OUTPUT DATAFRAME
# ----------------------------------------------------------
if not results:
    print(f"\n  No vehicles found with >= {MIN_STOP_COUNT} stops under these parameters.")
    print("  Try increasing STOP_SPEED_KMH or decreasing MIN_STOP_DURATION.")
else:
    df_out = pd.DataFrame(results)
    df_out = df_out.sort_values("Total Stops", ascending=False).reset_index(drop=True)

    print(f"\n  Vehicles with >= {MIN_STOP_COUNT} stops: {len(df_out):,}")
    print(f"  Max stops by a single vehicle: {df_out['Total Stops'].max()}")
    print(f"  Stop count distribution:")
    for n, cnt in df_out["Total Stops"].value_counts().sort_index().items():
        print(f"    {n} stops: {cnt} vehicles")

    # ----------------------------------------------------------
    # SAVE EXCEL
    # ----------------------------------------------------------
    print(f"\nSaving -> {output_file}")
    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        df_out.to_excel(writer, sheet_name="Multi-Stop Vehicles", index=False)

        # Formatting
        wb  = writer.book
        ws  = writer.sheets["Multi-Stop Vehicles"]

        # Header format
        hdr_fmt = wb.add_format({
            "bold"    : True,
            "bg_color": "#D9E1F2",
            "border"  : 1,
            "align"   : "center",
            "valign"  : "vcenter",
            "text_wrap": True,
        })
        # Stop-column group colors (cycle through 3 shades)
        stop_colors = ["#FFF2CC", "#E2EFDA", "#FCE4D6"]
        num_fmt  = wb.add_format({"num_format": "0.0", "align": "center"})
        int_fmt  = wb.add_format({"num_format": "0",   "align": "center", "bold": True})
        text_fmt = wb.add_format({"align": "center"})

        for col_idx, col_name in enumerate(df_out.columns):
            # Column width
            max_len = max(len(col_name),
                          df_out[col_name].astype(str).str.len().max())
            ws.set_column(col_idx, col_idx, min(max_len + 2, 22))

            # Header cell
            ws.write(0, col_idx, col_name, hdr_fmt)

            # Color stop columns by stop number
            if "Stop" in col_name:
                try:
                    stop_n = int(col_name.split("Stop")[1].split("-")[0].strip()) - 1
                    color  = stop_colors[stop_n % len(stop_colors)]
                    col_fmt = wb.add_format({
                        "bg_color"  : color,
                        "num_format": "0.0",
                        "align"     : "center",
                        "border"    : 1,
                    })
                    ws.set_column(col_idx, col_idx, min(max_len + 2, 22), col_fmt)
                except Exception:
                    pass

        # Bold the Total Stops column
        tot_col = list(df_out.columns).index("Total Stops")
        ws.set_column(tot_col, tot_col, 12, int_fmt)

        # Freeze top row
        ws.freeze_panes(1, 0)

        # ── Single-Stop Vehicles sheet ──
        if single_results:
            df_single = pd.DataFrame(single_results)
            df_single = df_single.sort_values("Track ID").reset_index(drop=True)
            df_single.to_excel(writer, sheet_name="Single-Stop Vehicles", index=False)

            ws_s = writer.sheets["Single-Stop Vehicles"]
            for col_idx, col_name in enumerate(df_single.columns):
                max_len = max(len(str(col_name)),
                              df_single[col_name].astype(str).str.len().max())
                ws_s.set_column(col_idx, col_idx, min(max_len + 2, 22))
                ws_s.write(0, col_idx, col_name, hdr_fmt)
            # Color the stop columns
            for col_idx, col_name in enumerate(df_single.columns):
                if "Stop" in col_name:
                    col_fmt_s = wb.add_format({
                        "bg_color"  : stop_colors[0],
                        "num_format": "0.0",
                        "align"     : "center",
                        "border"    : 1,
                    })
                    max_len = max(len(str(col_name)),
                                  df_single[col_name].astype(str).str.len().max())
                    ws_s.set_column(col_idx, col_idx, min(max_len + 2, 22), col_fmt_s)
            ws_s.freeze_panes(1, 0)
            print(f"  Single-stop vehicles written: {len(df_single):,}")

        # Also add a compact summary sheet
        summary_data = {
            "Parameter"           : ["Stop speed threshold", "Min stop duration",
                                     "Min stops to include",
                                     "Single-stop vehicles",
                                     "Multi-stop vehicles (≥2 stops)",
                                     "Max stops (single vehicle)"],
            "Value"               : [f"< {STOP_SPEED_KMH} km/h",
                                     f"≥ {MIN_STOP_DURATION} s",
                                     f"≥ {MIN_STOP_COUNT}",
                                     len(single_results),
                                     len(df_out),
                                     df_out["Total Stops"].max()],
        }
        pd.DataFrame(summary_data).to_excel(
            writer, sheet_name="Parameters", index=False)

    print(f"  Saved successfully.")
    print(f"  Single-stop vehicles : {len(single_results):,}")
    print(f"\n  Preview (top 5 vehicles):")
    preview_cols = ["Track ID", "Type", "Total Stops",
                    "Stop 1 - Start [s]", "Stop 1 - End [s]", "Stop 1 - Duration [s]"]
    preview_cols = [c for c in preview_cols if c in df_out.columns]
    print(df_out[preview_cols].head().to_string(index=False))

    # Build long-format stop events DataFrame for plotting
    df_stops = pd.DataFrame(stop_records)
    # Keep only stops belonging to multi-stop vehicles
    multi_ids = set(df_out["Track ID"].tolist())
    df_stops = df_stops[df_stops["Track ID"].isin(multi_ids)].reset_index(drop=True)

    # ----------------------------------------------------------
    # TIME-SPACE DIAGRAM — MULTI-STOP VEHICLES ONLY
    # ----------------------------------------------------------
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.cm as cm
    import matplotlib.colors as mcolors
    import matplotlib.patches as mpatches
    from matplotlib.lines import Line2D

    output_png = str(Path(input_file).parent / "04_12_25_1715_1745.timespace_multistop.png")
    print(f"\nPlotting Time-Space diagram for multi-stop vehicles ...")

    multi_ids = set(df_out["Track ID"].tolist())
    df_multi  = df[df["Track ID"].isin(multi_ids)].copy()

    # Build stop event lookup: track_id -> list of stop dicts
    stop_lookup = {}
    for _, s in df_stops.iterrows():
        tid = s["Track ID"]
        if tid not in stop_lookup:
            stop_lookup[tid] = []
        grp_t    = df_multi[df_multi["Track ID"] == tid]
        mask     = (grp_t["Time [s]"] >= s["Start time [s]"]) & \
                   (grp_t["Time [s]"] <= s["End time [s]"])
        pos_vals = grp_t.loc[mask, "cumulative_distance"].values
        stop_lookup[tid].append({
            "t_start": s["Start time [s]"],
            "t_end"  : s["End time [s]"],
            "p_min"  : pos_vals.min()  if len(pos_vals) > 0 else 0,
            "p_max"  : pos_vals.max()  if len(pos_vals) > 0 else 0,
            "stop_n" : s["Stop #"],
        })

    # Per-vehicle color (tab20 so each vehicle is distinct)
    unique_ids = df_multi["Track ID"].unique()
    id_cmap    = cm.tab20
    id_norm    = mcolors.Normalize(vmin=0, vmax=max(len(unique_ids) - 1, 1))
    id_map     = {tid: i for i, tid in enumerate(unique_ids)}

    fig, axes = plt.subplots(1, 2, figsize=(22, 9), dpi=200,
                              gridspec_kw={"width_ratios": [1.7, 1]})
    fig.patch.set_facecolor("white")

    # ── LEFT PANEL: Time-Space diagram ──────────────────────────
    ax = axes[0]
    ax.set_facecolor("#f8f8f8")

    # All other vehicles as grey background context
    for track_id, grp in df.groupby("Track ID", sort=False):
        if track_id in multi_ids:
            continue
        ax.plot(grp["Time [s]"].values, grp["cumulative_distance"].values,
                color="#cccccc", linewidth=0.5, alpha=0.35, zorder=1)

    # Multi-stop vehicles: colored lines + stop boxes
    for track_id, grp in df_multi.groupby("Track ID", sort=False):
        grp   = grp.sort_values("Time [s]")
        color = id_cmap(id_norm(id_map[track_id]))

        ax.plot(grp["Time [s]"].values, grp["cumulative_distance"].values,
                color=color, linewidth=1.8, alpha=0.92, zorder=2)

        # Track ID label at start of trajectory
        ax.text(grp["Time [s]"].iloc[0] - 2,
                grp["cumulative_distance"].iloc[0],
                str(track_id),
                fontsize=6.5, color=color,
                ha="right", va="center", zorder=5)

        # Red box for each stop event
        if track_id in stop_lookup:
            for st in stop_lookup[track_id]:
                p_spread = max(st["p_max"] - st["p_min"], 1.0)
                rect = mpatches.FancyBboxPatch(
                    (st["t_start"], st["p_min"] - 0.5),
                    width    = max(st["t_end"] - st["t_start"], 1.0),
                    height   = p_spread + 1.0,
                    boxstyle = "round,pad=0.3",
                    linewidth= 1.0,
                    edgecolor= "#CC0000",
                    facecolor= "#FF000025",
                    zorder   = 3
                )
                ax.add_patch(rect)
                # Label stop number (S1, S2, ...)
                ax.text(st["t_start"] + (st["t_end"] - st["t_start"]) / 2,
                        st["p_min"] - 1.8,
                        f'S{st["stop_n"]}',
                        fontsize=5.5, color="#CC0000",
                        ha="center", va="top", zorder=5)

    ax.set_xlabel("Time [s]", fontsize=12)
    ax.set_ylabel("Space [m]", fontsize=12)
    ax.set_xlim(0, np.ceil(df["Time [s]"].max() / 100) * 100)
    ax.set_ylim(-5, df["cumulative_distance"].max() * 1.08)
    for xt in np.arange(250, df["Time [s]"].max(), 250):
        ax.axvline(xt, color="#cccccc", linewidth=0.5, linestyle="--", zorder=0)
    ax.grid(axis="y", color="#cccccc", linewidth=0.4, linestyle="--", zorder=0)
    ax.set_axisbelow(True)
    for spine in ax.spines.values():
        spine.set_linewidth(0.8)

    h_multi = Line2D([0], [0], color="#555", lw=1.8,
                     label=f"Multi-stop vehicle  (n={len(multi_ids)})")
    h_other = Line2D([0], [0], color="#cccccc", lw=0.8,
                     label="Other vehicles (context)")
    h_stop  = mpatches.Patch(facecolor="#FF000025", edgecolor="#CC0000",
                              lw=1.0, label="Stop event  (S1, S2, ...)")
    ax.legend(handles=[h_multi, h_other, h_stop],
              loc="upper left", fontsize=9, framealpha=0.9, edgecolor="#ccc")
    ax.set_title(
        "Time-Space Diagram — Multi-Stop Vehicles Highlighted\n"
        "Location: Sathon-Convent, Approach: Sathon Road  |  Dec 04, 2025  1715–1745",
        fontsize=11, fontweight="bold"
    )

    # ── RIGHT PANEL: horizontal bar chart — stop count per vehicle ──
    ax2 = axes[1]
    ax2.set_facecolor("white")

    stop_counts = df_out.set_index("Track ID")["Total Stops"]
    sorted_ids  = stop_counts.sort_values(ascending=True).index.tolist()
    bar_colors  = [id_cmap(id_norm(id_map[tid])) for tid in sorted_ids]

    bars = ax2.barh(
        [str(tid) for tid in sorted_ids],
        [stop_counts[tid] for tid in sorted_ids],
        color=bar_colors, edgecolor="white", linewidth=0.5, height=0.7
    )
    # Count label on each bar
    for bar, tid in zip(bars, sorted_ids):
        ax2.text(bar.get_width() + 0.05,
                 bar.get_y() + bar.get_height() / 2,
                 str(stop_counts[tid]),
                 va="center", ha="left", fontsize=9, fontweight="bold")

    ax2.set_xlabel("Number of Stops", fontsize=11)
    ax2.set_ylabel("Track ID", fontsize=11)
    ax2.set_title(
        f"Stop Count per Vehicle\n"
        f"(speed < {STOP_SPEED_KMH} km/h  for ≥ {MIN_STOP_DURATION}s)",
        fontsize=11, fontweight="bold"
    )
    ax2.set_xlim(0, stop_counts.max() + 1.2)
    ax2.xaxis.set_major_locator(plt.MaxNLocator(integer=True))
    ax2.grid(axis="x", color="#eeeeee", linewidth=0.6, zorder=0)
    ax2.set_axisbelow(True)
    for spine in ax2.spines.values():
        spine.set_linewidth(0.6)

    plt.tight_layout()
    plt.savefig(output_png, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close()
    print(f"  Plot saved -> {output_png}")