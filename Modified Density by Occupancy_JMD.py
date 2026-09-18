import pandas as pd
import numpy as np
import os
import csv
import sys

def log(msg):
    """Print with an immediate flush so progress is visible right away."""
    print(msg, flush=True)

csv.field_size_limit(10_000_000)

# ======================================
# FILE PATHS
# ======================================
entry_file_path = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\120_sec\Occupancy\Raw_Occupancy_Density_Entry_5step.csv"
exit_file_path  = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\120_sec\Occupancy\Raw_Occupancy_Density_Exit_5step.csv"

# ======================================
# ==========  EDIT HERE: VEHICLE LENGTHS (VARIED METHOD)  ==========
# ======================================
VEHICLE_LENGTH_M = {
    "Motorcycle": 2.0,
    "Car": 4.5,
    "Tuk-Tuk": 2.5,
    "Medium Vehicle": 6.0,
    "Bus": 10.0,
    "Heavy Vehicle": 12.0,
}
CATEGORIES = list(VEHICLE_LENGTH_M.keys())

# ======================================
# ==========  EDIT HERE: RAW "Type" LABEL -> CATEGORY MAPPING  ==========
# Any raw label not listed here (e.g. "Pedestrian", "Bicycle") is reported
# and excluded, since no length was given for it.
# ======================================
TYPE_MAP = {
    "Bus": "Bus",
    "Car": "Car",
    "Heavy Vehicle": "Heavy Vehicle",
    "Medium Vehicle": "Medium Vehicle",
    "Motorcycle": "Motorcycle",
    "Tuk-Tuk": "Tuk-Tuk",
}

# ======================================
# ==========  EDIT HERE: FIXED VEHICLE LENGTH (OLD/FIXED METHOD)  ==========
# Your original code never actually divided occupied-time by a length --
# it only computed the merged occupied-time itself. Since the fixed
# constant you previously used for the density conversion wasn't given,
# this defaults to 5.0 m. CHANGE THIS to whatever value you actually used
# before, if it wasn't 5.0 m.
# ======================================
FIXED_VEHICLE_LENGTH_M = 5.0

# ======================================
# ==========  EDIT HERE: BUCKET SIZES FOR THE TWO SHEETS  ==========
# Sheet 1: fixed-size buckets (0-30, 30-60, 60-90 ...)
# Sheet 2: first bucket is SHEET2_FIRST_BUCKET wide, every bucket after is
#          SHEET2_REST_BUCKET wide (default: 0-15, then 15-45, 45-75 ...)
# ======================================
SHEET1_BUCKET = 120
SHEET2_FIRST_BUCKET = 60
SHEET2_REST_BUCKET = 120

# ======================================
# ROBUST CSV READER (same parser used throughout this project -- handles
# the ragged trajectory tail and any stray/unbalanced quotes safely)
# ======================================
def read_occupancy_csv(path, label):
    dist_col = "Traveled Dist. [m]"
    log(f"\nOpening {label} file: {path}")
    log(f"File size: {os.path.getsize(path) / 1_000_000:.2f} MB")

    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f, delimiter=",", quotechar='"')
        rows = [r for r in reader if r]

    header = [h.strip() for h in rows[0]]
    data_rows = rows[1:]
    max_len = max(len(r) for r in data_rows)
    if len(header) < max_len:
        header = header + [f"traj_col_{i}" for i in range(len(header), max_len)]
    data_rows = [r + [""] * (max_len - len(r)) for r in data_rows]

    df = pd.DataFrame(data_rows, columns=header[:max_len])
    df.columns = df.columns.str.strip()
    df = df.replace("", np.nan)

    entry_time_col = "Entry Time [s]"
    exit_time_col = "Exit Time [s]"
    type_col = "Type"

    df = df[df[type_col].notna()].copy()
    df[entry_time_col] = pd.to_numeric(df[entry_time_col], errors="coerce")
    df[exit_time_col] = pd.to_numeric(df[exit_time_col], errors="coerce")
    df[type_col] = df[type_col].astype(str).str.strip()

    before = len(df)
    df = df.dropna(subset=[entry_time_col, exit_time_col])
    if len(df) < before:
        log(f"NOTE: dropped {before - len(df)} {label} rows missing Entry or Exit Time.")

    # ---- Detector zone length (L_d) ----
    # "Traveled Dist. [m]" is how far each vehicle moved between the
    # Entry-line and Exit-line of THIS narrow detection zone -- i.e. it
    # IS the physical zone length for this file. Averaged across every
    # vehicle (regardless of category, since zone length is a property of
    # the detector, not of any one vehicle type), this gives L_d for the
    # k = O / (L_v + L_d) x 1000 formula.
    df[dist_col] = pd.to_numeric(df[dist_col], errors="coerce")
    L_d = float(df[dist_col].mean())
    log(f"{label} detector zone length (L_d, average Traveled Dist.): {L_d:.4f} m "
        f"(from {df[dist_col].notna().sum()} vehicles)")

    df["Type_raw"] = df[type_col]
    df[type_col] = df["Type_raw"].map(TYPE_MAP)
    unmapped = df.loc[df[type_col].isna(), "Type_raw"].unique()
    unmapped = [x for x in unmapped if x not in ("", "nan")]
    if len(unmapped):
        log(f"WARNING: {label} file has {len(unmapped)} unmapped Type label(s), "
            f"excluded (no vehicle length given for them): {list(unmapped)}")

    df = df[df[type_col].notna()].reset_index(drop=True)
    df["Length_m"] = df[type_col].map(VEHICLE_LENGTH_M)

    log(f"Loaded {len(df)} usable {label} vehicles.")
    return df.reset_index(drop=True), L_d

entry_df, ENTRY_L_D = read_occupancy_csv(entry_file_path, "ENTRY")
exit_df, EXIT_L_D = read_occupancy_csv(exit_file_path, "EXIT")

entry_time_col = "Entry Time [s]"
exit_time_col = "Exit Time [s]"

data_max_time = int(max(
    entry_df[entry_time_col].max(), entry_df[exit_time_col].max(),
    exit_df[entry_time_col].max(), exit_df[exit_time_col].max(),
))
log(f"\nMax time available across both files: {data_max_time}")
sys.stdout.flush()

# ======================================
# MANUAL INPUT: ANALYSIS WINDOW (START / END)
# ======================================
while True:
    try:
        sys.stdout.flush()
        start_time = int(input("Enter the START time (seconds) to begin the analysis from (e.g. 20): ").strip())
        sys.stdout.flush()
        end_time = int(input("Enter the END time (seconds) to stop the analysis at (e.g. 1802): ").strip())
        if start_time < 0 or end_time < 0:
            print("Times cannot be negative. Please re-enter both.")
            continue
        if end_time <= start_time:
            print("END time must be greater than START time. Please re-enter both.")
            continue
        if start_time > data_max_time:
            print(f"START time ({start_time}) is beyond the data's max time ({data_max_time}).")
            continue
        if end_time > data_max_time:
            print(f"NOTE: END time ({end_time}) is beyond the data's max time ({data_max_time}). "
                  f"The analysis will still run up to {end_time} as requested.")
        break
    except ValueError:
        print("Please enter whole numbers for both START and END.")

log(f"Analyzing only the window [{start_time}, {end_time}) seconds.\n")

# ======================================
# GRID-BUCKET HELPERS
# ======================================
def make_fixed_grid(bucket_size, upto_time):
    boundaries = []
    t = 0
    while t < upto_time:
        boundaries.append((t, t + bucket_size))
        t += bucket_size
    return boundaries

def make_first_then_fixed_grid(first_size, rest_size, upto_time):
    boundaries = []
    t = 0
    size = first_size
    while t < upto_time:
        boundaries.append((t, t + size))
        t += size
        size = rest_size
    return boundaries

# ======================================
# PER-WINDOW OCCUPANCY -> DENSITY (one gate/file at a time)
#
# FIXED method: replicates the original given code's merged-interval
# occupied time EXACTLY (union of overlapping vehicle presence -- no
# double counting), then applies the standard occupancy-to-density
# formula k = O / (L_v + L_d) x 1000, using ONE constant vehicle length
# for everyone plus this file's own detector zone length L_d:
#     Density_fixed = (merged_occupied_time / duration) x 1000 / (FIXED_LENGTH_M + L_d)
#
# VARIED method: sums each vehicle's OWN occupied-time contribution
# separately (NOT merged -- if two vehicles genuinely overlap, that's
# more vehicles/density, not double-counted time), each divided by that
# vehicle's OWN length plus the same L_d, then added together:
#     Density_varied = sum_i [ (occupied_time_i / duration) x 1000 / (length_i_m + L_d) ]
# ======================================
def compute_window_density(df, eff_start, eff_end, L_d):
    duration = eff_end - eff_start
    mask = (df[exit_time_col] > eff_start) & (df[entry_time_col] < eff_end)
    sub = df[mask]
    n_vehicles = len(sub)

    if n_vehicles == 0:
        return 0.0, 0.0, 0

    clip_start = sub[entry_time_col].clip(lower=eff_start)
    clip_end = sub[exit_time_col].clip(upper=eff_end)

    # ---- FIXED method: merge overlapping intervals ----
    intervals = sorted(zip(clip_start, clip_end))
    merged = []
    for s, e in intervals:
        if merged and s <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], e))
        else:
            merged.append((s, e))
    merged_occupied_time = sum(e - s for s, e in merged)
    density_fixed = (merged_occupied_time / duration) * 1000 / (FIXED_VEHICLE_LENGTH_M + L_d)

    # ---- VARIED method: per-vehicle, own length + L_d, not merged ----
    individual_times = (clip_end - clip_start).to_numpy()
    lengths_plus_ld_m = sub["Length_m"].to_numpy() + L_d
    density_varied = float(np.sum((individual_times / duration) * 1000 / lengths_plus_ld_m))

    return density_fixed, density_varied, n_vehicles

def build_density_sheet(boundaries, sheet_name=""):
    out_rows = []
    n_buckets = len(boundaries)
    for bi, (g_start, g_end) in enumerate(boundaries):
        if bi % 20 == 0:
            log(f"  [{sheet_name}] bucket {bi + 1}/{n_buckets} ...")

        eff_start = max(g_start, start_time)
        eff_end = min(g_end, end_time)
        if eff_start >= eff_end:
            continue

        entry_fixed, entry_varied, n_entry = compute_window_density(entry_df, eff_start, eff_end, ENTRY_L_D)
        exit_fixed, exit_varied, n_exit = compute_window_density(exit_df, eff_start, eff_end, EXIT_L_D)

        total_n = n_entry + n_exit
        if total_n > 0:
            final_fixed = (entry_fixed * n_entry + exit_fixed * n_exit) / total_n
            final_varied = (entry_varied * n_entry + exit_varied * n_exit) / total_n
        else:
            final_fixed = None
            final_varied = None

        out_rows.append({
            "Time (s)": f"{eff_start}-{eff_end}",
            "Entry: Vehicle Count": n_entry,
            "Exit: Vehicle Count": n_exit,
            "Entry Density - Fixed Length (vehicle/km)": round(entry_fixed, 3),
            "Exit Density - Fixed Length (vehicle/km)": round(exit_fixed, 3),
            "Final Density by Occupancy - Fixed Length (vehicle/km)":
                round(final_fixed, 3) if final_fixed is not None else None,
            "Entry Density - Varied Length (vehicle/km)": round(entry_varied, 3),
            "Exit Density - Varied Length (vehicle/km)": round(exit_varied, 3),
            "Final Density by Occupancy - Varied Length (vehicle/km)":
                round(final_varied, 3) if final_varied is not None else None,
        })

    log(f"  [{sheet_name}] done -- {len(out_rows)} rows built.")
    return pd.DataFrame(out_rows)

# ======================================
# BUILD BOTH SHEETS
# ======================================
sheet1_boundaries = make_fixed_grid(SHEET1_BUCKET, end_time)
log(f"\nBuilding Sheet 1 ({SHEET1_BUCKET}s fixed buckets, {len(sheet1_boundaries)} buckets total) ...")
sheet1_df = build_density_sheet(sheet1_boundaries, sheet_name="Sheet1")

sheet2_boundaries = make_first_then_fixed_grid(SHEET2_FIRST_BUCKET, SHEET2_REST_BUCKET, end_time)
log(f"\nBuilding Sheet 2 (alternate {SHEET2_FIRST_BUCKET}/{SHEET2_REST_BUCKET}s buckets, "
    f"{len(sheet2_boundaries)} buckets total) ...")
sheet2_df = build_density_sheet(sheet2_boundaries, sheet_name="Sheet2")

# ======================================
# CONFIG REFERENCE SHEET (for audit/transparency)
# ======================================
config_rows = [
    {"Setting": "Fixed vehicle length (m) -- OLD method", "Value": FIXED_VEHICLE_LENGTH_M},
    {"Setting": "Entry detector zone length L_d (m) -- avg Traveled Dist.", "Value": round(ENTRY_L_D, 4)},
    {"Setting": "Exit detector zone length L_d (m) -- avg Traveled Dist.", "Value": round(EXIT_L_D, 4)},
]
for cat, length in VEHICLE_LENGTH_M.items():
    config_rows.append({"Setting": f"{cat} length (m) -- VARIED method", "Value": length})
config_df = pd.DataFrame(config_rows)

# ======================================
# SAVE
# ======================================
output_dir = os.path.dirname(os.path.abspath(entry_file_path))
output_file = os.path.join(output_dir, f"Occupancy_Density_FixedVsVaried_{start_time}_{end_time}.xlsx")

counter = 1
while os.path.exists(output_file):
    output_file = os.path.join(output_dir, f"Occupancy_Density_FixedVsVaried_{start_time}_{end_time}_{counter}.xlsx")
    counter += 1

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    sheet1_df.to_excel(writer, sheet_name=f"{SHEET1_BUCKET}s Occupancy Density", index=False)
    sheet2_df.to_excel(writer, sheet_name="Alt Timeframe Occupancy Density", index=False)
    config_df.to_excel(writer, sheet_name="Config_Reference", index=False)

log("\n✅ Excel file created!")
log(f"Saved at: {output_file}")