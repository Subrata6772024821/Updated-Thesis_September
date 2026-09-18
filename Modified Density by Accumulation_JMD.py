import pandas as pd
import numpy as np
import os
import csv

# ======================================
# FILE PATH
# ======================================
file_path = r"D:\Thesis_Final Data Analysis\05_12_25_1400_1430\30_sec\Accumulation\Raw_Accumulation_Density_1step.csv"

# ======================================
# ==========  EDIT HERE: PCE (PASSENGER CAR EQUIVALENT) FACTORS  ==========
# Shared by BOTH output workbooks.
# ======================================
PCE_FACTORS = {
    "Bus": 3,
    "Car": 1,
    "Heavy Vehicle": 3,
    "Medium Vehicle": 2,
    "Motorcycle": 0.2,
    "Tuk-Tuk": 1,
}
CATEGORIES = list(PCE_FACTORS.keys())

# ======================================
# ==========  EDIT HERE: RAW "Type" LABEL -> CATEGORY MAPPING  ==========
# If the CSV's "Type" column already spells categories exactly as above,
# leave this as-is. If your raw file uses different labels, map them to
# one of the 6 CATEGORIES below. Any raw label NOT listed here is reported
# and excluded from the analysis.
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
# ==========  EDIT HERE: SHEET 3 TIME BUCKET PATTERN  ==========
# Sheet 2 always uses fixed 30-sec buckets (0-30, 30-60, 60-90...).
# Sheet 3 uses a DIFFERENT pattern: the first bucket is SHEET3_FIRST_BUCKET
# seconds wide, every bucket after that is SHEET3_REST_BUCKET seconds wide
# (default: 0-15, then 15-45, 45-75, 75-105...). Edit these two numbers to
# change the pattern -- nothing else needs to change. Shared by BOTH
# output workbooks.
# ======================================
SHEET3_FIRST_BUCKET = 15
SHEET3_REST_BUCKET = 30
SHEET2_BUCKET = 30  # fixed bucket size used by Sheet 2 (both workbooks)

# ======================================
# MANUAL ROW-BY-ROW PARSE (ragged CSV)
# ======================================
FIXED_COLUMNS = [
    "Track ID", "Type", "Entry Gate", "Entry Time [s]",
    "Exit Gate", "Exit Time [s]", "Traveled Dist. [m]", "Avg. Speed [km/h]",
]

rows = []
with open(file_path, encoding="utf-8-sig", newline="") as f:
    reader = csv.reader(f, skipinitialspace=True)
    header = next(reader)  # skip header line
    for raw_row in reader:
        if not raw_row:
            continue
        rows.append(raw_row[:len(FIXED_COLUMNS)])

df = pd.DataFrame(rows, columns=FIXED_COLUMNS)

for col in ("Type", "Entry Gate", "Exit Gate"):
    df[col] = df[col].astype(str).str.strip()

print("Columns after parsing:", df.columns.tolist())
print(f"Shape: {df.shape}")

# ======================================
# MAP RAW "Type" -> STANDARD CATEGORY
# ======================================
df["Type_raw"] = df["Type"]
df["Type"] = df["Type_raw"].map(TYPE_MAP)

unmapped_labels = df.loc[df["Type"].isna(), "Type_raw"].unique()
unmapped_labels = [x for x in unmapped_labels if x not in ("", "nan")]
if len(unmapped_labels):
    print(f"\nWARNING: {len(unmapped_labels)} raw Type label(s) not found in "
          f"TYPE_MAP and will be EXCLUDED from the analysis: {list(unmapped_labels)}")
    print("Add them to TYPE_MAP above if they should be included.\n")

df = df[df["Type"].notna()]

# ======================================
# NUMERIC CONVERSION
# ======================================
entry_col = "Entry Time [s]"
exit_col = "Exit Time [s]"
dist_col = "Traveled Dist. [m]"

df[entry_col] = pd.to_numeric(df[entry_col], errors="coerce")
df[exit_col] = pd.to_numeric(df[exit_col], errors="coerce")
df[dist_col] = pd.to_numeric(df[dist_col], errors="coerce")

# The active-count method needs BOTH entry and exit times to know when a
# vehicle was on the segment. Rows missing either are dropped (reported).
before = len(df)
dropped = df[df[entry_col].isna() | df[exit_col].isna()]
df = df.dropna(subset=[entry_col, exit_col], how="any")
if len(dropped):
    print(f"NOTE: dropped {len(dropped)} of {before} rows missing Entry or Exit "
          f"Time (no way to determine when they were on the segment).")

data_max_time = int(max(df[entry_col].max(), df[exit_col].max()))
print("Max time available in the raw data:", data_max_time)

# ======================================
# MANUAL INPUT: ANALYSIS WINDOW (START / END)
# Shared by BOTH output workbooks -- same window, same underlying data.
# ======================================
while True:
    try:
        start_time = int(input("Enter the START time (seconds) to begin the analysis from (e.g. 17): ").strip())
        end_time = int(input("Enter the END time (seconds) to stop the analysis at (e.g. 1745): ").strip())
        if start_time < 0 or end_time < 0:
            print("Times cannot be negative. Please re-enter both.")
            continue
        if end_time <= start_time:
            print("END time must be greater than START time. Please re-enter both.")
            continue
        if start_time > data_max_time:
            print(f"START time ({start_time}) is beyond the raw data's max time ({data_max_time}).")
            continue
        if end_time > data_max_time:
            print(f"NOTE: END time ({end_time}) is beyond the raw data's max time ({data_max_time}). "
                  f"The analysis will still run up to {end_time} as requested.")
        break
    except ValueError:
        print("Please enter whole numbers for both START and END.")

print(f"Analyzing only the window [{start_time}, {end_time}) seconds -- "
      f"{end_time - start_time} one-second intervals.\n")

seconds = list(range(start_time, end_time))

# ======================================
# MANUAL INPUT: VEHICLES ALREADY ON THE SEGMENT AT t = START, BY CATEGORY
# Only used by the FLOW-BASED workbook (Workbook 2). The ACTIVE-COUNT
# workbook (Workbook 1) does not need this -- it works out presence
# directly from each vehicle's own Entry/Exit Time.
# ======================================
print(f"Enter the number of vehicles physically present in the segment at "
      f"t = {start_time}, BY CATEGORY (counted directly, e.g. from drone footage).")
print("(This is only used for the flow-based workbook -- Workbook 2.)")
N0 = {}
for cat in CATEGORIES:
    while True:
        try:
            val = int(input(f"  {cat}: ").strip())
            if val < 0:
                print("  Cannot be negative."); continue
            N0[cat] = val
            break
        except ValueError:
            print("  Please enter a whole number.")
N0_total = sum(N0.values())
print(f"Total N0 (all categories) = {N0_total}\n")

# ======================================
# AVERAGE TRAVELED DISTANCE (vehicles entering within the analysis window)
# Shared by BOTH workbooks.
# ======================================
window_vehicles = df[(df[entry_col] >= start_time) & (df[entry_col] < end_time)]
avg_dist_m = window_vehicles[dist_col].mean()
avg_dist_km = avg_dist_m / 1000.0 if pd.notna(avg_dist_m) and avg_dist_m > 0 else np.nan

print(f"Average traveled distance of vehicles in the analysis window: "
      f"{avg_dist_m:.2f} m ({avg_dist_km:.5f} km)")
if pd.isna(avg_dist_km) or avg_dist_km == 0:
    print("WARNING: average traveled distance is missing or zero -- "
          "density (veh/km, PCU/km) columns will be blank.\n")

df_by_cat = {cat: df[df["Type"] == cat] for cat in CATEGORIES}

# ============================================================
# ============================================================
#   METHOD 1: ACTIVE-COUNT ACCUMULATION (-> Workbook 1)
#   A vehicle counts as present at second t if
#   Entry Time <= t < Exit Time. Computed directly from the full
#   trajectory data -- no manual headcount, no drift.
# ============================================================
# ============================================================
seconds_arr = np.array(seconds)
category_accum_active = {}

for cat in CATEGORIES:
    sub = df_by_cat[cat]
    entries = sub[entry_col].to_numpy()
    exits = sub[exit_col].to_numpy()
    active = (entries[None, :] <= seconds_arr[:, None]) & (exits[None, :] > seconds_arr[:, None])
    category_accum_active[cat] = active.sum(axis=1).astype(int).tolist()

# ============================================================
# ============================================================
#   METHOD 2: FLOW-BASED ACCUMULATION (-> Workbook 2)
#   Accumulation(cat, t) = N0[cat] + running total of
#   (Flow In - Flow Out) up to and including second t.
# ============================================================
# ============================================================
flow_in = {cat: [] for cat in CATEGORIES}
flow_out = {cat: [] for cat in CATEGORIES}
category_accum_flow = {cat: [] for cat in CATEGORIES}
cum_net = {cat: 0 for cat in CATEGORIES}

for t in seconds:
    for cat in CATEGORIES:
        sub = df_by_cat[cat]
        f_in = int(((sub[entry_col] >= t) & (sub[entry_col] < t + 1)).sum())
        f_out = int(((sub[exit_col] >= t) & (sub[exit_col] < t + 1)).sum())
        flow_in[cat].append(f_in)
        flow_out[cat].append(f_out)
        cum_net[cat] += (f_in - f_out)
        category_accum_flow[cat].append(N0[cat] + cum_net[cat])

# ======================================
# GRID-BUCKET HELPERS (shared by both workbooks' Sheet 2 / Sheet 3)
# ======================================
def make_fixed_grid(bucket_size, upto_time):
    """Boundaries 0-bucket_size, bucket_size-2*bucket_size, ... up to >= upto_time."""
    boundaries = []
    t = 0
    while t < upto_time:
        boundaries.append((t, t + bucket_size))
        t += bucket_size
    return boundaries

def make_first_then_fixed_grid(first_size, rest_size, upto_time):
    """First bucket is first_size wide, every bucket after is rest_size wide."""
    boundaries = []
    t = 0
    size = first_size
    while t < upto_time:
        boundaries.append((t, t + size))
        t += size
        size = rest_size
    return boundaries

def build_summary_sheet(boundaries, seconds, category_accum, avg_dist_km, start_time, end_time):
    """
    Clips each grid bucket to [start_time, end_time), skips empty buckets,
    and computes per-category + total average accumulation, vehicle/hour,
    PCU/hour, density (veh/km) and density (PCU/km) for each bucket.
    Every second belongs to exactly one bucket -> no duplication, no gaps.
    """
    second_index = {int(t): i for i, t in enumerate(seconds)}
    out_rows = []

    for g_start, g_end in boundaries:
        eff_start = max(g_start, start_time)
        eff_end = min(g_end, end_time)
        if eff_start >= eff_end:
            continue  # bucket falls entirely outside the analysis window

        idxs = [second_index[t] for t in range(eff_start, eff_end)]
        duration = eff_end - eff_start

        row = {"Time (s)": f"{eff_start}-{eff_end}"}
        total_vehicle_avg = 0.0
        total_pcu_avg = 0.0
        for cat in CATEGORIES:
            cat_vals = [category_accum[cat][i] for i in idxs]
            cat_avg = sum(cat_vals) / len(cat_vals)
            row[f"Avg Accumulation - {cat} (vehicle)"] = round(cat_avg, 3)
            total_vehicle_avg += cat_avg
            total_pcu_avg += cat_avg * PCE_FACTORS[cat]

        row["Avg Total Accumulation (vehicle)"] = round(total_vehicle_avg, 3)
        row["Avg Total Accumulation (PCU)"] = round(total_pcu_avg, 3)
        row["Vehicle/hour"] = round(total_vehicle_avg * (3600 / duration), 2)
        row["PCU/hour"] = round(total_pcu_avg * (3600 / duration), 2)

        if pd.notna(avg_dist_km) and avg_dist_km > 0:
            row["Density (vehicle/km)"] = round(total_vehicle_avg / avg_dist_km, 3)
            row["Density (PCU/km)"] = round(total_pcu_avg / avg_dist_km, 3)
        else:
            row["Density (vehicle/km)"] = None
            row["Density (PCU/km)"] = None

        row["Avg Traveled Distance (km)"] = round(avg_dist_km, 5) if pd.notna(avg_dist_km) else None
        out_rows.append(row)

    return pd.DataFrame(out_rows)


def build_sheet1_active(seconds, category_accum):
    """Sheet 1 for the active-count workbook: category columns + totals."""
    out_rows = []
    for i, t in enumerate(seconds):
        row = {"Time (s)": f"{t}-{t+1}"}
        total_vehicle = 0
        total_pcu = 0.0
        for cat in CATEGORIES:
            val = category_accum[cat][i]
            row[cat] = val
            total_vehicle += val
            total_pcu += val * PCE_FACTORS[cat]
        row["Total accumulated vehicles (vehicle)"] = total_vehicle
        row["Total accumulated vehicles (PCU)"] = round(total_pcu, 3)
        out_rows.append(row)
    return pd.DataFrame(out_rows)


def build_sheet1_flow(seconds, flow_in, flow_out, category_accum, N0):
    """Sheet 1 for the flow-based workbook: In/Out/Accumulation per category + totals."""
    out_rows = []
    for i, t in enumerate(seconds):
        row = {"Time (s)": f"{t}-{t+1}"}
        total_vehicle = 0
        total_pcu = 0.0
        for cat in CATEGORIES:
            row[f"{cat} In"] = flow_in[cat][i]
            row[f"{cat} Out"] = flow_out[cat][i]
            row[f"{cat} Accumulation"] = category_accum[cat][i]
            total_vehicle += category_accum[cat][i]
            total_pcu += category_accum[cat][i] * PCE_FACTORS[cat]
        row["Total accumulated vehicles (vehicle)"] = total_vehicle
        row["Total accumulated vehicles (PCU)"] = round(total_pcu, 3)
        out_rows.append(row)
    df_out = pd.DataFrame(out_rows)
    n_rows = len(df_out)
    for cat in CATEGORIES:
        df_out[f"N0_{cat}"] = [N0[cat]] + [None] * (n_rows - 1)
    df_out["Analysis_Start_Time_s"] = [start_time] + [None] * (n_rows - 1)
    return df_out


# ======================================
# BUILD ALL SHEETS FOR BOTH WORKBOOKS
# ======================================
sheet2_boundaries = make_fixed_grid(SHEET2_BUCKET, end_time)
sheet3_boundaries = make_first_then_fixed_grid(SHEET3_FIRST_BUCKET, SHEET3_REST_BUCKET, end_time)

# ---- Workbook 1: Active-count method ----
wb1_sheet1 = build_sheet1_active(seconds, category_accum_active)
wb1_sheet2 = build_summary_sheet(sheet2_boundaries, seconds, category_accum_active,
                                  avg_dist_km, start_time, end_time)
wb1_sheet3 = build_summary_sheet(sheet3_boundaries, seconds, category_accum_active,
                                  avg_dist_km, start_time, end_time)

# ---- Workbook 2: Flow-based method (manual N0) ----
wb2_sheet1 = build_sheet1_flow(seconds, flow_in, flow_out, category_accum_flow, N0)
wb2_sheet2 = build_summary_sheet(sheet2_boundaries, seconds, category_accum_flow,
                                  avg_dist_km, start_time, end_time)
wb2_sheet3 = build_summary_sheet(sheet3_boundaries, seconds, category_accum_flow,
                                  avg_dist_km, start_time, end_time)

# ======================================
# SAVE TWO SEPARATE WORKBOOKS
# ======================================
output_dir = os.path.dirname(file_path)

wb1_path = os.path.join(output_dir, f"Density_Estimation_ActiveCount_{start_time}_{end_time}.xlsx")
with pd.ExcelWriter(wb1_path, engine="openpyxl") as writer:
    wb1_sheet1.to_excel(writer, sheet_name="1sec Category Accumulation", index=False)
    wb1_sheet2.to_excel(writer, sheet_name=f"{SHEET2_BUCKET}s Buckets", index=False)
    wb1_sheet3.to_excel(writer, sheet_name="Sheet3 Custom Buckets", index=False)

wb2_path = os.path.join(output_dir, f"Density_Estimation_FlowBased_{start_time}_{end_time}.xlsx")
with pd.ExcelWriter(wb2_path, engine="openpyxl") as writer:
    wb2_sheet1.to_excel(writer, sheet_name="1sec Category Accumulation", index=False)
    wb2_sheet2.to_excel(writer, sheet_name=f"{SHEET2_BUCKET}s Buckets", index=False)
    wb2_sheet3.to_excel(writer, sheet_name="Sheet3 Custom Buckets", index=False)

print("\n✅ Two Excel workbooks created!")
print("Workbook 1 (Active-count method):", wb1_path)
print("Workbook 2 (Flow-based method,  N0 =", N0, "):", wb2_path)