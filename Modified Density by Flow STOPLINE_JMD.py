import pandas as pd
import numpy as np
import os
import csv
import sys

def log(msg):
    """Print with an immediate flush so progress is visible right away,
    even when stdout is being redirected/buffered (a common reason a
    running script LOOKS frozen when it's actually still working)."""
    print(msg, flush=True)

# Raise the csv module's field size limit -- a stray, unbalanced " character
# in the export makes the parser treat everything after it as one giant
# unterminated field, which trips the default 131072-byte limit.
csv.field_size_limit(10_000_000)

# ======================================
# FILE PATH
# No separate trajectory file needed this time -- STOPLINE speed reads
# each vehicle's own flattened trajectory directly out of this same CSV
# (the (x, y, Speed, Tan.Acc, Lat.Acc, Time) columns after the first 8
# fixed metadata columns).
# ======================================
file_path = r"D:\Thesis_Final Data Analysis\05_12_25_1315_1345\120_sec\Flow\Raw_Flow_Density_5step.csv"

# ======================================
# ==========  EDIT HERE: PCE (PASSENGER CAR EQUIVALENT) FACTORS  ==========
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
# ==========  EDIT HERE: GATE DETECTION  ==========
# The Entry/Exit Gate labels (e.g. "Gate 14", "Gate 15") differ between
# datasets, so by default the script AUTO-DETECTS them instead of relying
# on hardcoded numbers:
#   - The EXIT gate is taken as whichever Exit Gate label is the most
#     common in the whole file -- i.e. the segment's actual downstream
#     gate that most trips complete at.
#   - The ENTRY gate(s) are every distinct Entry Gate value that shows up
#     among trips ending at that detected exit gate -- this naturally
#     picks up a single entry (14->15) or multiple merging entries
#     (14/16->15) without you specifying anything.
# These same auto-detected gates are also used below to decide which
# Entry/Exit crossings count as "stopline" events for the speed calc.
# Set AUTO_DETECT_GATES = False and fill in MANUAL_ENTRY_GATES /
# MANUAL_EXIT_GATE below only if you need to override the detection.
# ======================================
AUTO_DETECT_GATES = True
MANUAL_ENTRY_GATES = ["Gate 14"]   # only used if AUTO_DETECT_GATES = False
MANUAL_EXIT_GATE = "Gate 15"       # only used if AUTO_DETECT_GATES = False

# ======================================
# ==========  EDIT HERE: BUCKET SIZES FOR THE TWO SUMMARY SHEETS  ==========
# Sheet 1: fixed-size buckets (0-30, 30-60, 60-90 ...)
# Sheet 2: first bucket is SHEET2_FIRST_BUCKET wide, every bucket after is
#          SHEET2_REST_BUCKET wide (default: 0-15, then 15-45, 45-75 ...)
# ======================================
SHEET1_BUCKET = 120
SHEET2_FIRST_BUCKET = 60
SHEET2_REST_BUCKET = 120

# ======================================
# READ CSV
# The file is genuinely comma-delimited, but the trajectory portion of
# each row is one large quoted field (quotechar='"') protecting internal
# commas -- and one stray/unbalanced quote somewhere in the export was
# desyncing pandas' own quote handling, which is what caused the earlier
# "Expected N fields, saw ~2N" and "field larger than field limit" errors.
# Parsing manually with Python's csv module sidesteps both problems.
# ======================================
log(f"Opening file: {file_path}")
log(f"File size: {os.path.getsize(file_path) / 1_000_000:.2f} MB")
log("Reading and parsing CSV row-by-row (this is the slow step on large "
    "files with long trajectory fields) ...")

with open(file_path, "r", encoding="utf-8-sig", newline="") as f:
    reader = csv.reader(f, delimiter=",", quotechar='"')
    rows = []
    for i, r in enumerate(reader):
        if r:
            rows.append(r)
        if (i + 1) % 5000 == 0:
            log(f"  ... read {i + 1} raw lines so far")

log(f"Finished reading CSV: {len(rows)} raw lines (including header).")

header = [h.strip() for h in rows[0]]
data_rows_raw = rows[1:]  # kept UNPADDED -- needed later to extract each
                          # vehicle's own actual (not padded) trajectory length

max_len = max(len(r) for r in data_rows_raw)
if len(header) < max_len:
    header = header + [f"traj_col_{i}" for i in range(len(header), max_len)]

log("Padding ragged rows to a rectangular table (for the Flow/category table only) ...")
# Pad short rows (fewer trajectory points) so the table is rectangular.
# This padded copy is only used to build the Flow/category DataFrame
# below -- the STOPLINE speed calculation further down uses
# data_rows_raw (unpadded) instead, so padding never corrupts a vehicle's
# trajectory sample count.
data_rows_padded = [r + [""] * (max_len - len(r)) for r in data_rows_raw]

log("Building DataFrame ...")
df = pd.DataFrame(data_rows_padded, columns=header[:max_len])
df.columns = df.columns.str.strip()
df = df.replace("", np.nan)  # csv.reader gives "" for blanks/padding, not NaN

log(f"Parsed {len(df)} data rows, {len(df.columns)} columns.")
log("Detected Columns (first 10):")
log(str(df.columns[:10].tolist()))

# ======================================
# DEFINE COLUMNS
# ======================================
entry_gate_col = "Entry Gate"
exit_gate_col  = "Exit Gate"
entry_time_col = "Entry Time [s]"
exit_time_col  = "Exit Time [s]"
type_col       = "Type"
dist_col       = "Traveled Dist. [m]"

# ======================================
# CLEAN DATA
# ======================================
df = df[df[type_col].notna()].copy()

df[entry_time_col] = pd.to_numeric(df[entry_time_col], errors="coerce")
df[exit_time_col]  = pd.to_numeric(df[exit_time_col], errors="coerce")
df[dist_col]       = pd.to_numeric(df[dist_col], errors="coerce")

df[entry_gate_col] = df[entry_gate_col].astype(str).str.replace('"', '').str.strip()
df[exit_gate_col]  = df[exit_gate_col].astype(str).str.replace('"', '').str.strip()
df[type_col]       = df[type_col].astype(str).str.strip()

before = len(df)
dropped = df[df[entry_time_col].isna() | df[exit_time_col].isna()]
df = df.dropna(subset=[entry_time_col, exit_time_col])
if len(dropped):
    log(f"NOTE: dropped {len(dropped)} of {before} rows missing Entry or Exit "
        f"Time (no way to determine when they were on the segment).")

df = df.reset_index(drop=True)

# ======================================
# AUTO-DETECT (OR APPLY MANUAL) ENTRY/EXIT GATES
# ======================================
if AUTO_DETECT_GATES:
    exit_gate_counts = df[exit_gate_col].value_counts()
    exit_gate_counts = exit_gate_counts[
        exit_gate_counts.index.notna() & (~exit_gate_counts.index.isin(["", "nan"]))
    ]
    if exit_gate_counts.empty:
        raise ValueError(
            "AUTO_DETECT_GATES is on, but no usable Exit Gate values were found "
            "in the CSV. Set AUTO_DETECT_GATES = False and fill in "
            "MANUAL_ENTRY_GATES / MANUAL_EXIT_GATE instead."
        )

    VALID_EXIT_GATE = exit_gate_counts.idxmax()
    detected_entries = (
        df.loc[df[exit_gate_col] == VALID_EXIT_GATE, entry_gate_col]
        .dropna()
        .unique()
        .tolist()
    )
    VALID_ENTRY_GATES = sorted(g for g in detected_entries if g not in ("", "nan"))

    log(f"\n🔎 Auto-detected EXIT gate: '{VALID_EXIT_GATE}' "
        f"({exit_gate_counts.max()} of {exit_gate_counts.sum()} trips)")
    log(f"🔎 Auto-detected ENTRY gate(s) feeding it: {VALID_ENTRY_GATES}")
    log("   (If this looks wrong for your dataset, set AUTO_DETECT_GATES = False "
        "above and specify MANUAL_ENTRY_GATES / MANUAL_EXIT_GATE instead.)\n")
else:
    VALID_ENTRY_GATES = MANUAL_ENTRY_GATES
    VALID_EXIT_GATE = MANUAL_EXIT_GATE
    log(f"\nUsing manually specified gates -- Entry: {VALID_ENTRY_GATES}, "
        f"Exit: {VALID_EXIT_GATE}\n")

# ======================================
# VALID TRIPS (auto-detected or manual gate pair above) -- kept for the
# Gate_Distance_Summary reference sheet only. Built from the FULL,
# unfiltered dataset -- BEFORE the category mapping below -- so this
# reference distance is never accidentally shrunk by a Type label that
# doesn't map cleanly.
# ======================================
traj = df[
    (df[exit_gate_col] == VALID_EXIT_GATE) &
    (df[entry_gate_col].isin(VALID_ENTRY_GATES))
].copy()

dist_summary = traj.groupby(entry_gate_col)[dist_col].agg(["mean", "std", "count"])
log(f"\n📏 Average measured distance per section ({'/'.join(VALID_ENTRY_GATES)} -> {VALID_EXIT_GATE} trips):")
log(str(dist_summary))

# ======================================
# MAP RAW "Type" -> STANDARD CATEGORY
# This filtering is used ONLY for the per-category Flow/PCU counts below
# -- it must not affect the distance reference sheet above, and must not
# affect the STOPLINE speed events below either (those come from ALL
# vehicles regardless of category, matching the original STOPLINE script).
# ======================================
df["Type_raw"] = df[type_col]
df[type_col] = df["Type_raw"].map(TYPE_MAP)

unmapped_labels = df.loc[df[type_col].isna(), "Type_raw"].unique()
unmapped_labels = [x for x in unmapped_labels if x not in ("", "nan")]
if len(unmapped_labels):
    log(f"\nWARNING: {len(unmapped_labels)} raw Type label(s) not found in "
        f"TYPE_MAP and will be EXCLUDED from the Flow/PCU counts: {list(unmapped_labels)}")
    log("Add them to TYPE_MAP above if they should be categorized.\n")

df = df[df[type_col].notna()].reset_index(drop=True)

data_max_time = int(max(df[entry_time_col].max(), df[exit_time_col].max()))
log(f"\nMax time available in the raw data: {data_max_time}")
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
            print(f"START time ({start_time}) is beyond the raw data's max time ({data_max_time}).")
            continue
        if end_time > data_max_time:
            print(f"NOTE: END time ({end_time}) is beyond the raw data's max time ({data_max_time}). "
                  f"The analysis will still run up to {end_time} as requested.")
        break
    except ValueError:
        print("Please enter whole numbers for both START and END.")

log(f"Analyzing only the window [{start_time}, {end_time}) seconds.\n")

# ======================================
# STOPLINE SPEED -- build entry/exit spot-speed events directly from the
# raw CSV's own flattened trajectory columns (no separate file needed).
#
# For every vehicle:
#   - if it entered via one of VALID_ENTRY_GATES, interpolate its own
#     trajectory to get its instantaneous (spot) speed AT the exact entry
#     time, and record (entry_time, spot_speed) as an "entry event".
#   - if it exited via VALID_EXIT_GATE, do the same at its exit time as
#     an "exit event".
#
# For a given interval [t0, t1): pool every entry-event speed AND every
# exit-event speed whose crossing time falls in [t0, t1) into ONE list,
# and take the HARMONIC MEAN of that pooled list -- this is the
# interval's "stopline" speed. Harmonic mean is the standard traffic
# space-mean-speed estimator: slower vehicles (which spend longer on the
# road) get more influence, exactly as physically appropriate. Pooling
# entry+exit together (rather than averaging two separate harmonic
# means) keeps every vehicle's statistical weight equal and needs no
# special-casing when one side has zero crossings.
# ======================================
log("Extracting per-vehicle trajectories for STOPLINE speed (from this same CSV) ...")

def interp_speed_at(times, speeds, t):
    """Instantaneous speed at time t, linearly interpolated between the
    two bracketing trajectory samples (clamped to the endpoints if t
    falls outside the recorded trajectory range)."""
    if len(times) == 0:
        return np.nan
    if t <= times[0]:
        return float(speeds[0])
    if t >= times[-1]:
        return float(speeds[-1])
    return float(np.interp(t, times, speeds))

entry_events = []  # list of (time, spot_speed_kmh)
exit_events = []
n_skipped = 0

for raw_row in data_rows_raw:
    if len(raw_row) < 8:
        n_skipped += 1
        continue
    try:
        r_entry_gate = raw_row[2].strip().strip('"')
        r_entry_time = float(raw_row[3])
        r_exit_gate = raw_row[4].strip().strip('"')
        r_exit_time = float(raw_row[5])
    except (ValueError, IndexError):
        n_skipped += 1
        continue

    is_valid_entry = r_entry_gate in VALID_ENTRY_GATES
    is_valid_exit = r_exit_gate == VALID_EXIT_GATE
    if not (is_valid_entry or is_valid_exit):
        continue  # this vehicle contributes no stopline event at all

    traj_fields = raw_row[8:]  # this vehicle's OWN actual length -- not padded
    n_samples = len(traj_fields) // 6
    if n_samples == 0:
        continue

    times = np.empty(n_samples)
    speeds = np.empty(n_samples)
    ok = True
    for k in range(n_samples):
        base = k * 6
        try:
            speeds[k] = float(traj_fields[base + 2])
            times[k] = float(traj_fields[base + 5])
        except (ValueError, IndexError):
            ok = False
            break
    if not ok:
        n_skipped += 1
        continue

    if is_valid_entry:
        entry_events.append((r_entry_time, interp_speed_at(times, speeds, r_entry_time)))
    if is_valid_exit:
        exit_events.append((r_exit_time, interp_speed_at(times, speeds, r_exit_time)))

log(f"Built {len(entry_events)} entry events and {len(exit_events)} exit events "
    f"for STOPLINE speed ({n_skipped} vehicle rows skipped -- malformed/short trajectory).")

def events_in_window(events, lo, hi):
    return [speed for t, speed in events if lo <= t < hi]

def harmonic_mean(speeds):
    """Space-mean-speed estimator: n / sum(1/v_i). Non-positive or
    missing speeds are excluded before the calculation."""
    arr = np.array([s for s in speeds if s and s > 0 and not np.isnan(s)], dtype=float)
    if len(arr) == 0:
        return None
    return float(len(arr) / np.sum(1.0 / arr))

def compute_stopline_speed(t0, t1):
    ev_entry = events_in_window(entry_events, t0, t1)
    ev_exit = events_in_window(exit_events, t0, t1)
    entry_hm = harmonic_mean(ev_entry)   # reference only, for Speed_Detail
    exit_hm = harmonic_mean(ev_exit)     # reference only, for Speed_Detail
    pooled = ev_entry + ev_exit
    stopline_kmh = harmonic_mean(pooled)
    return stopline_kmh, entry_hm, exit_hm, len(ev_entry), len(ev_exit), len(pooled)

# ======================================
# PRECOMPUTE PER-CATEGORY SUBSETS ONCE
# ======================================
log("Pre-splitting data by category ...")
df_by_cat = {cat: df[df[type_col] == cat] for cat in CATEGORIES}
for cat in CATEGORIES:
    log(f"  {cat}: {len(df_by_cat[cat])} rows")

# ======================================
# GRID-BUCKET HELPERS
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

# ======================================
# BUILD ONE FLOW / DENSITY SUMMARY SHEET FOR A GIVEN SET OF BUCKET
# BOUNDARIES. Each bucket is clipped to [start_time, end_time), so with
# start_time=20 the first bucket becomes "20-30" (not "0-30"), and with
# end_time=1802 the last bucket becomes a short "1800-1802" instead of
# running past the requested end. Every second belongs to exactly ONE
# bucket -- no gaps, no double-counting.
# ======================================
speed_detail_rows = []  # collected across BOTH sheets, tagged by sheet name

def build_flow_density_sheet(boundaries, sheet_name=""):
    out_rows = []
    n_buckets = len(boundaries)
    for bi, (g_start, g_end) in enumerate(boundaries):
        if bi % 20 == 0:
            log(f"  [{sheet_name}] bucket {bi + 1}/{n_buckets} ...")

        eff_start = max(g_start, start_time)
        eff_end = min(g_end, end_time)
        if eff_start >= eff_end:
            continue

        duration = eff_end - eff_start
        row = {"Time (s)": f"{eff_start}-{eff_end}"}

        total_in = 0
        total_out = 0
        total_avg_vehicle = 0.0
        total_avg_pcu = 0.0

        for cat in CATEGORIES:
            sub = df_by_cat[cat]  # precomputed once, not re-filtered here

            # Flow In / Out = simply based on Entry Time / Exit Time falling
            # inside this bucket -- no gate restriction (unchanged from the
            # snapshot version -- this is the per-category Flow/PCU part,
            # separate from the gate-restricted STOPLINE speed events).
            f_in = sub[
                (sub[entry_time_col] >= eff_start) &
                (sub[entry_time_col] < eff_end)
            ].shape[0]

            f_out = sub[
                (sub[exit_time_col] >= eff_start) &
                (sub[exit_time_col] < eff_end)
            ].shape[0]

            avg_flow_cat = (f_in + f_out) / 2
            avg_flow_cat_pcu = avg_flow_cat * PCE_FACTORS[cat]

            row[f"{cat} Flow In"] = f_in
            row[f"{cat} Flow Out"] = f_out
            row[f"{cat} Avg Flow (vehicle)"] = round(avg_flow_cat, 3)

            total_in += f_in
            total_out += f_out
            total_avg_vehicle += avg_flow_cat
            total_avg_pcu += avg_flow_cat_pcu

        row["Total Flow In (vehicle)"] = total_in
        row["Total Flow Out (vehicle)"] = total_out
        row["Average Flow (vehicle)"] = round(total_avg_vehicle, 3)
        row["Average Flow (vehicle/hour)"] = round(total_avg_vehicle * (3600 / duration), 2)
        row["Average Flow (PCU)"] = round(total_avg_pcu, 3)
        row["Average Flow (PCU/hour)"] = round(total_avg_pcu * (3600 / duration), 2)

        # ---- STOPLINE speed (pooled harmonic mean) for this bucket ----
        stopline_kmh, entry_hm, exit_hm, n_entry, n_exit, n_pooled = compute_stopline_speed(eff_start, eff_end)
        stopline_ms = stopline_kmh / 3.6 if stopline_kmh else None

        row["Stopline Speed (m/s)"] = round(stopline_ms, 4) if stopline_ms else None
        row["Stopline Speed (km/h)"] = round(stopline_kmh, 4) if stopline_kmh else None

        # ---- Density (vehicle/km, PCU/km) = Flow (per hour) / Speed (km/h) ----
        # NOTE: no lanes adjustment here, to match the structure of the
        # script this one is based on. The original STOPLINE script
        # divides by NUM_LANES as well (Flow / (Speed * Lanes)) -- add
        # that multiplication back in below if you want that instead.
        if stopline_kmh and stopline_kmh > 0:
            row["Density (vehicle/km)"] = round(row["Average Flow (vehicle/hour)"] / stopline_kmh, 3)
            row["Density (PCU/km)"] = round(row["Average Flow (PCU/hour)"] / stopline_kmh, 3)
        else:
            row["Density (vehicle/km)"] = None
            row["Density (PCU/km)"] = None

        out_rows.append(row)

        speed_detail_rows.append([
            sheet_name, f"{eff_start}-{eff_end}",
            n_entry, entry_hm, n_exit, exit_hm, n_pooled, stopline_kmh
        ])

    log(f"  [{sheet_name}] done -- {len(out_rows)} rows built.")
    return pd.DataFrame(out_rows)

# ======================================
# BUILD BOTH SHEETS
# ======================================
sheet1_boundaries = make_fixed_grid(SHEET1_BUCKET, end_time)
log(f"\nBuilding Sheet 1 ({SHEET1_BUCKET}s fixed buckets, {len(sheet1_boundaries)} buckets total) ...")
sheet1_df = build_flow_density_sheet(sheet1_boundaries, sheet_name="Sheet1")

sheet2_boundaries = make_first_then_fixed_grid(SHEET2_FIRST_BUCKET, SHEET2_REST_BUCKET, end_time)
log(f"\nBuilding Sheet 2 (alternate {SHEET2_FIRST_BUCKET}/{SHEET2_REST_BUCKET}s buckets, "
    f"{len(sheet2_boundaries)} buckets total) ...")
sheet2_df = build_flow_density_sheet(sheet2_boundaries, sheet_name="Sheet2")

speed_detail_df = pd.DataFrame(speed_detail_rows, columns=[
    "Sheet", "Interval (s)",
    "Entry: Vehicle Count", "Entry: Harmonic-Mean Speed [km/h] (ref. only)",
    "Exit: Vehicle Count", "Exit: Harmonic-Mean Speed [km/h] (ref. only)",
    "Pooled Vehicle Count", "Stopline Speed [km/h] (harmonic mean of pooled entry+exit)"
])

# ======================================
# SAVE
# ======================================
output_dir = os.path.dirname(os.path.abspath(file_path))
output_file = os.path.join(output_dir, f"Category_Flow_Density_STOPLINE_{start_time}_{end_time}.xlsx")

counter = 1
while os.path.exists(output_file):
    output_file = os.path.join(output_dir, f"Category_Flow_Density_STOPLINE_{start_time}_{end_time}_{counter}.xlsx")
    counter += 1

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    sheet1_df.to_excel(writer, sheet_name=f"{SHEET1_BUCKET}s Flow-Density", index=False)
    sheet2_df.to_excel(writer, sheet_name="Alt Timeframe Flow-Density", index=False)
    speed_detail_df.to_excel(writer, sheet_name="Speed_Detail", index=False)
    dist_summary.to_excel(writer, sheet_name="Gate_Distance_Summary")

log("\n✅ Excel file created!")
log(f"Saved at: {output_file}")