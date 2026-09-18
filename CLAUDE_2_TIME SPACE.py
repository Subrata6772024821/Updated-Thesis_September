import pandas as pd
import numpy as np

# ==========================================================
# STEP 2 — Kalman smooth  ->  1-second interpolation  ->  Cumulative distance
#
# Handles 3 vehicle categories:
#   A. Complete     : all coordinate frames captured, distance accurate
#   B. Truncated    : recording ends before exit gate, distance scaled to reference
#   C. Reconstructed: no coordinate data at all, distance computed from avg_speed * time
# ==========================================================

input_file  = r"D:\Thesis_Final Data Analysis\05_12_25_1400_1430\Trajectory\1step\05_12_25_1400_1430.cross_test2_pivoted.xlsx"
output_file = r"D:\Thesis_Final Data Analysis\05_12_25_1400_1430\Trajectory\1step\05_12_25_1400_1430.trajectory_1s_final.xlsx"

# ----------------------------------------------------------
def kalman_cv_zvd(group):
    group = group.sort_values("Time [s]").reset_index(drop=True)
    if len(group) == 0:
        return group

    x     = group["x [m]"].to_numpy()
    y     = group["y [m]"].to_numpy()
    speed = group["Speed [km/h]"].to_numpy()
    n     = len(group)

    X      = np.zeros((4, n))
    P      = np.eye(4)
    X[0,0] = x[0]
    X[1,0] = y[0]

    Q_base = np.diag([0.1, 0.1, 1.0, 1.0])
    R      = np.diag([5.0, 5.0])
    H      = np.array([[1,0,0,0],[0,1,0,0]])

    for i in range(1, n):
        dt = group["Time [s]"].iloc[i] - group["Time [s]"].iloc[i-1]
        if pd.isna(dt) or dt <= 0:
            dt = 0.033

        F            = np.array([[1,0,dt,0],[0,1,0,dt],[0,0,1,0],[0,0,0,1]])
        displacement = np.sqrt((x[i]-x[i-1])**2 + (y[i]-y[i-1])**2)
        is_stop      = (speed[i] < 0.5) and (displacement < 0.1)
        Q            = Q_base * (0.01 if is_stop else 1.0)

        X_pred = F @ X[:,i-1]
        P_pred = F @ P @ F.T + Q
        Z      = np.array([x[i], y[i]])
        Y      = Z - H @ X_pred
        S      = H @ P_pred @ H.T + R
        K      = P_pred @ H.T @ np.linalg.inv(S)
        X[:,i] = X_pred + K @ Y
        P      = (np.eye(4) - K @ H) @ P_pred

        if is_stop:
            X[2,i] = 0.0
            X[3,i] = 0.0

    group["x_smooth"] = X[0]
    group["y_smooth"] = X[1]
    return group

# ----------------------------------------------------------
print("Loading workbook ...")
all_sheets = pd.read_excel(input_file, sheet_name=None)
print(f"  Sheets: {list(all_sheets.keys())}")

final_list = []

for sheet_name, df in all_sheets.items():
    print(f"\nProcessing sheet: {sheet_name}")

    df.columns = (df.columns.astype(str).str.strip()
                  .str.replace('"', '', regex=False)
                  .str.replace('\u00A0', ' ', regex=False))

    if "Track ID" not in df.columns:
        possible = [c for c in df.columns if "track" in c.lower() and "id" in c.lower()]
        if possible:
            df = df.rename(columns={possible[0]: "Track ID"})
        else:
            print(f"  Skipping: 'Track ID' not found")
            continue

    required = ["Track ID", "x [m]", "y [m]", "Speed [km/h]", "Time [s]"]
    missing  = [c for c in required if c not in df.columns]
    if missing:
        print(f"  Skipping: missing {missing}")
        continue

    for c in ["x [m]", "y [m]", "Speed [km/h]", "Time [s]"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    if "Truncated"     not in df.columns: df["Truncated"]     = False
    if "Reconstructed" not in df.columns: df["Reconstructed"] = False

    # Per-vehicle reference lookup
    ref_lookup = {}
    if "Traveled Dist. [m]" in df.columns:
        ref_lookup = df.groupby("Track ID")["Traveled Dist. [m]"].first().to_dict()

    df = (df.sort_values(["Track ID", "Time [s]"])
            .drop_duplicates(subset=["Track ID", "Time [s]"], keep="first")
            .reset_index(drop=True))

    interp_list = []

    # ----------------------------------------------------------
    # Split into normal (has coordinates) vs reconstructed (no coordinates)
    recon_ids  = set(df.loc[df["Reconstructed"] == True, "Track ID"].unique())
    normal_ids = set(df["Track ID"].unique()) - recon_ids

    # ── NORMAL VEHICLES: Kalman + interpolation ──
    if normal_ids:
        df_normal = df[df["Track ID"].isin(normal_ids)].copy()
        df_normal = df_normal.dropna(subset=["x [m]", "y [m]"])

        print(f"  Kalman filtering {len(normal_ids):,} normal vehicles ...")
        kalman_results = []
        for track_id, grp in df_normal.groupby("Track ID"):
            result = kalman_cv_zvd(grp)
            result["Track ID"] = track_id
            kalman_results.append(result)

        if kalman_results:
            df_normal = pd.concat(kalman_results, ignore_index=True)
            df_normal = df_normal.sort_values(["Track ID", "Time [s]"]).reset_index(drop=True)

            print(f"  Interpolating to 1-second intervals ...")
            for track_id, group in df_normal.groupby("Track ID"):
                group = group.sort_values("Time [s]").set_index("Time [s]")
                t_min = int(np.ceil(group.index.min()))
                t_max = int(np.floor(group.index.max()))
                if t_min > t_max:
                    continue

                new_time     = np.arange(t_min, t_max + 1, 1)
                numeric_cols = group.select_dtypes(include=[np.number]).columns
                num_part     = group[numeric_cols]
                num_part     = num_part.reindex(num_part.index.union(new_time))
                num_part     = num_part.interpolate(method="linear")
                num_part     = num_part.loc[new_time]
                num_part     = num_part.dropna(subset=["x_smooth", "y_smooth"])
                if len(num_part) == 0:
                    continue

                # Re-stamp non-numeric columns lost during select_dtypes interpolation
                veh_row = df_normal.loc[df_normal["Track ID"] == track_id].iloc[0]
                num_part["Track ID"]      = track_id
                num_part["Sheet"]         = sheet_name
                num_part["Type"]          = str(veh_row["Type"]).strip() if "Type" in df_normal.columns else "Unknown"
                num_part["Truncated"]     = bool(veh_row["Truncated"])    if "Truncated"      in df_normal.columns else False
                num_part["Reconstructed"] = False
                interp_list.append(
                    num_part.reset_index().rename(columns={"index": "Time [s]"})
                )

    # ── RECONSTRUCTED VEHICLES: synthetic trajectory ──
    if recon_ids:
        print(f"  Rebuilding {len(recon_ids):,} reconstructed vehicles from metadata ...")
        df_recon = df[df["Track ID"].isin(recon_ids)].copy()

        for track_id, grp in df_recon.groupby("Track ID"):
            ref_dist = ref_lookup.get(track_id, grp["Traveled Dist. [m]"].iloc[0]
                                      if "Traveled Dist. [m]" in grp.columns else 96.0)
            avg_spd  = grp["Speed [km/h]"].iloc[0]
            entry_t  = grp["Entry Time [s]"].iloc[0] if "Entry Time [s]" in grp.columns else grp["Time [s]"].min()
            exit_t   = grp["Exit Time [s]"].iloc[0]  if "Exit Time [s]"  in grp.columns else grp["Time [s]"].max()

            t_steps  = grp["Time [s]"].values
            n        = len(t_steps)
            if n == 0:
                continue

            # Linear position: distance = avg_speed * elapsed_time
            elapsed  = t_steps - t_steps[0]
            duration = exit_t - entry_t
            pos      = ref_dist * elapsed / duration if duration > 0 else np.zeros(n)

            vtype  = str(grp["Type"].iloc[0]).strip() if "Type" in grp.columns else "Unknown"
            result = pd.DataFrame({
                "Time [s]"           : t_steps,
                "x_smooth"           : np.nan,
                "y_smooth"           : np.nan,
                "Speed [km/h]"       : avg_spd,
                "step_distance"      : np.concatenate([[0], np.diff(pos)]),
                "cumulative_distance": pos,
                "Track ID"           : track_id,
                "Type"               : vtype,
                "Sheet"              : sheet_name,
                "Truncated"          : True,
                "Reconstructed"      : True,
            })
            interp_list.append(result)

    if not interp_list:
        continue

    df_interp = pd.concat(interp_list, ignore_index=True)
    df_interp = df_interp.sort_values(["Track ID", "Time [s]"]).reset_index(drop=True)

    # ----------------------------------------------------------
    # Distance for normal vehicles
    print("  Computing cumulative distances ...")
    normal_mask = df_interp["Reconstructed"] == False

    step_dist = pd.Series(0.0, index=df_interp.index)

    for track_id, grp in df_interp[normal_mask].groupby("Track ID"):
        grp  = grp.sort_values("Time [s]")
        idx  = grp.index
        dx   = grp["x_smooth"].values
        dy   = grp["y_smooth"].values
        dist = np.sqrt(np.diff(dx)**2 + np.diff(dy)**2)
        dist[dist > 15]   = 0.0
        dist[dist < 0.05] = 0.0
        step_dist.loc[idx] = np.concatenate([[0.0], dist])

    df_interp.loc[normal_mask, "step_distance"] = step_dist[normal_mask]
    df_interp.loc[normal_mask, "cumulative_distance"] = (
        df_interp[normal_mask].groupby("Track ID")["step_distance"].cumsum().values
    )

    # Scale truncated (non-reconstructed) vehicles
    print("  Scaling truncated vehicles ...")
    n_scaled = 0
    for track_id, grp_idx in df_interp[normal_mask & (df_interp["Truncated"]==True)].groupby("Track ID").groups.items():
        grp       = df_interp.loc[grp_idx]
        ref_dist  = ref_lookup.get(track_id)
        coord_dist= grp["cumulative_distance"].iloc[-1]
        if ref_dist is None or coord_dist <= 0:
            continue
        scale = ref_dist / coord_dist
        df_interp.loc[grp_idx, "step_distance"]       *= scale
        df_interp.loc[grp_idx, "cumulative_distance"] *= scale
        n_scaled += 1

    print(f"  Scaled {n_scaled} truncated vehicles")

    # Drop raw columns
    drop_cols = ["Entry Gate", "Entry Time [s]", "Exit Gate", "Exit Time [s]",
                 "Traveled Dist. [m]", "Avg. Speed [km/h]",
                 "Tan. Acc. [m/s2]", "Lat. Acc. [m/s2]"]
    df_interp = df_interp.drop(
        columns=[c for c in drop_cols if c in df_interp.columns], errors="ignore")

    final_list.append(df_interp)
    print(f"  Done: {len(df_interp):,} rows | {df_interp['Track ID'].nunique():,} vehicles")

# ----------------------------------------------------------
if not final_list:
    print("\nNo output produced.")
else:
    print("\nCombining all sheets ...")
    final_df = pd.concat(final_list, ignore_index=True)
    final_df = final_df.sort_values(["Sheet", "Track ID", "Time [s]"]).reset_index(drop=True)

    print("Saving ...")
    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        final_df.to_excel(writer, sheet_name="Final_Data", index=False)

    print(f"\n  Saved -> {output_file}")
    print(f"   Total rows    : {len(final_df):,}")
    print(f"   Total vehicles: {final_df['Track ID'].nunique():,}  (all 1515)")
    print(f"   Complete      : {final_df[final_df['Truncated']==False]['Track ID'].nunique():,}")
    print(f"   Truncated     : {final_df[(final_df['Truncated']==True) & (final_df['Reconstructed']==False)]['Track ID'].nunique():,}")
    print(f"   Reconstructed : {final_df[final_df['Reconstructed']==True]['Track ID'].nunique():,}")