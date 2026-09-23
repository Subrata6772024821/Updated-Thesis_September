"""
Per-Signal-Cycle Traffic Analysis: GTR, Density, Vehicle Counts, and Degree of Saturation
================================================================================
For every complete signal cycle (from one green start to the next), this
script computes and exports:

  Number of Cycle, Cycle Time Start, Cycle Time End, Green Period, Green Time
  Ratio, Density (vehicle/km), Density (PCU/km), Total Number of Vehicles,
  Entry Flow Rate (vehicle/hour), Number and Percentage of No-Stop /
  Single-Stop / Multi-Stop / (Single+Multi) Vehicles, the same four counts
  in PCU, and Degree of Saturation (DoS).

DATA SOURCES (all four of your files are used together)
------------------------------------------------------
  - GREEN_TIME_FILE      -> signal cycle boundaries and Green Time Ratio
  - RAW_TRAJECTORY_CSV   -> every vehicle's Track ID, Type, and Entry Time
                             (used for Total Vehicles, Entry Flow Rate, and
                             to split vehicles into No-Stop/Single/Multi)
  - STOP_ANALYSIS_FILE   -> which Track IDs are Single-Stop vs Multi-Stop
                             (everything else is No-Stop)
  - DENSITY_FILE         -> 1-second vehicle/PCU accumulation (density) and
                             the PCU-equivalence weights per vehicle type,
                             which this script derives directly from your
                             own data rather than assuming standard textbook
                             values

METHODOLOGY NOTES - read before using the numbers
------------------------------------------------------
1. CYCLE DEFINITION: a cycle runs from one green-phase start to the next.
   The final phase has no known next start within the recorded session, so
   its cycle is undefined and excluded (see console output).

2. TWO DIFFERENT "COUNTING LOGICS" ARE USED ON PURPOSE:
     - Density (vehicle/km, PCU/km) is an INSTANTANEOUS OCCUPANCY measure:
       for every second, it counts how many vehicles are physically present
       on the road at that instant (Entry Time <= t and Exit Time > t), then
       averages those per-second snapshots across the cycle. A vehicle that
       lingers longer (e.g. a stopped vehicle) is counted in more snapshots.
     - Total Vehicles / Entry Flow / No-Stop / Single-Stop / Multi-Stop are
       ARRIVAL EVENT counts: every vehicle is counted exactly ONCE, based on
       its Entry Time falling inside the cycle window - regardless of how
       long it stayed.
   These are intentionally different, standard traffic-engineering measures
   (occupancy vs. flow). One consequence: a vehicle that enters near the end
   of one cycle but doesn't exit until the next cycle will count toward THIS
   cycle's vehicle totals (by Entry Time) while still contributing to the
   NEXT cycle's density (by physical presence). This is a normal edge effect
   of chopping continuous traffic into discrete cycles, not an error.

3. PCU WEIGHTS are solved directly from your Density file (least squares
   against its own category counts vs. its Total PCU column), so this stays
   correct even if a different session uses different PCU factors.

4. PERCENTAGE COLUMNS (No-Stop / Single-Stop / Multi-Stop / Stopped) are
   calculated from VEHICLE COUNTS, not PCU counts. A "stop" is a per-vehicle
   event (one motorcycle stopping is one stopping event, exactly as much as
   one bus stopping is one), so weighting it by physical road-space (PCU)
   would distort the percentage toward whatever vehicle mix happened to
   stop, rather than reflecting what fraction of actual vehicles/drivers
   experienced a stop. This also matches standard traffic-engineering
   practice - "percentage of vehicles stopped" is a well-established MOE
   (Webster's method, HCM) and is always defined per-vehicle, not
   PCU-weighted.

5. DEGREE OF SATURATION (DoS = v / c, where c = s x N_LANES x GTR):
     - v (Entry Flow Rate in PCU/h) uses the SAME Entry-Time-based counting
       logic as the vehicle counts above, converted to an hourly rate.
     - s (saturation flow rate per lane) is a STANDARD DEFAULT of 1,900
       PCU/h/lane - not measured on-site. Change S_PER_LANE below if you
       have a field-measured value instead.
     - N_LANES is now entered MANUALLY at runtime (the script will prompt
       for it each time you run it), rather than being hardcoded in CONFIG.
       This was previously a visual estimate from a single drone frame (3
       lanes looked certain, a 4th was ambiguous) - confirm against
       satellite imagery if possible, since DoS scales directly (inversely)
       with this number.
     - DoS > 1.0 means the cycle is oversaturated (demand exceeds what the
       green time can discharge); values much above ~1.2-1.3 should be read
       as "clearly oversaturated" rather than trusted to the third decimal,
       since standard delay models become unstable well before that point.

Output: one Excel workbook with three sheets - "Per-Cycle Summary",
"Assumptions and Notes", and "PCU Weights Used" - saved to the SAME FOLDER
as the raw trajectory CSV (its parent directory). The output filename is
auto-derived from the Green Time file's name, so it can't go stale when you
reuse this script for a different session (see OUTPUT_EXCEL_NAME_OVERRIDE
below).
"""

import os
import csv
import re
import openpyxl
import numpy as np
import pandas as pd

# =====================================================================
# CONFIG - edit these
# =====================================================================
GREEN_TIME_FILE = r"D:\Thesis_Final Data Analysis\Customs Intersection\Green Time Data\GreenTime_South_12-15-12-33.xlsx"
DENSITY_FILE = r"D:\Thesis_Final Data Analysis\Customs Intersection\South\09_09_26_1215_1234\Density_Estimation_ActiveCount_88_1010_1.xlsx"
STOP_ANALYSIS_FILE = r"D:\Thesis_Final Data Analysis\Customs Intersection\South\09_09_26_1215_1234\Trajectory\09_09_26_1215_1234.stop_analysis.xlsx"
RAW_TRAJECTORY_CSV = r"D:\Thesis_Final Data Analysis\Customs Intersection\South\09_09_26_1215_1234\Raw_Density_Accumulation_1step.csv"

# Output Excel is saved in the SAME FOLDER as RAW_TRAJECTORY_CSV (its parent
# directory). The filename is auto-derived from GREEN_TIME_FILE's own name
# (e.g. "GreenTime_FINAL_West_14-28-15-03.xlsx" -> "West_14-28-15-03"), so it
# always matches whichever session's files you point this script at. Set
# OUTPUT_EXCEL_NAME_OVERRIDE below if you want a fixed name instead.
OUTPUT_EXCEL_NAME_OVERRIDE = None  # e.g. "My_Custom_Name.xlsx", or leave as None to auto-derive

# ---- Degree of Saturation assumptions ----
# N_LANES is no longer set here - the script asks for it on every run (see
# get_n_lanes_from_user() below), since it's a visual/manual estimate that
# can change per intersection/session. DEFAULT_N_LANES is only used as the
# pre-filled suggestion shown in the prompt, and as a fallback if the script
# is run non-interactively (e.g. piped input with no value / EOF).
DEFAULT_N_LANES = 4
S_PER_LANE = 1900       # PCU/hour/lane, standard textbook default

# ---- Density data coverage requirement ----
# A cycle's density is only computed if at least this fraction of its seconds
# actually have density data available (some sessions' density files don't
# cover the full recording window). Below this threshold, both density
# columns are left BLANK for that cycle rather than averaged over an
# incomplete, systematically-biased subset of seconds (see chat discussion -
# missing seconds are usually right after a green start, which is not a
# random/representative slice of the cycle). The exact coverage percentage
# is always reported in its own column regardless of whether it passed.
DENSITY_COVERAGE_MIN_FRACTION = 0.80  # 80%

# Default sheet/column names (auto-detected by content if not found, so this
# tolerates workbooks whose tabs are named slightly differently)
SHEET_GREEN_PHASES = "Green Phases"
SHEET_SECOND_ACCUMULATION = "1sec Category Accumulation"
SHEET_BUCKETS_FOR_DISTANCE = "30s Buckets"
SHEET_MULTISTOP = "Multi-Stop Vehicles"
SHEET_SINGLESTOP = "Single-Stop Vehicles"
VEHICLE_CATEGORIES = ["Bus", "Car", "Heavy Vehicle", "Medium Vehicle", "Motorcycle", "Tuk-Tuk"]


# =====================================================================
# USER INPUT
# =====================================================================
def get_n_lanes_from_user(default=DEFAULT_N_LANES):
    """Prompts the user to type in the number of lanes for this intersection/
    approach, used in the Degree of Saturation capacity calculation
    (c = S_PER_LANE x N_LANES x GTR). Keeps asking until a valid positive
    integer is entered. Pressing Enter with no input accepts `default`.
    If input is not available at all (e.g. script run non-interactively
    with no stdin), falls back to `default` automatically so the script
    doesn't hang."""
    while True:
        try:
            raw = input(f"Enter number of lanes for this approach [default={default}]: ").strip()
        except EOFError:
            print(f"No interactive input available - using default N_LANES={default}.")
            return default
        if raw == "":
            print(f"Using default N_LANES={default}.")
            return default
        try:
            n = int(raw)
            if n <= 0:
                print("Number of lanes must be a positive whole number. Try again.")
                continue
            return n
        except ValueError:
            print(f"'{raw}' is not a valid whole number. Try again.")


# =====================================================================
# LOADERS
# =====================================================================
def _find_sheet_with_column(wb, column_name, scan_rows=15):
    for sn in wb.sheetnames:
        ws = wb[sn]
        for r in range(1, min(scan_rows, ws.max_row) + 1):
            if column_name in [c.value for c in ws[r]]:
                return sn
    available = {sn: [c.value for c in wb[sn][1]] for sn in wb.sheetnames}
    raise ValueError(f"Could not find a sheet containing '{column_name}'. "
                      f"Sheets/headers found: {available}")


def _find_header_row(ws, marker_value, max_scan=15):
    for r in range(1, min(max_scan, ws.max_row) + 1):
        if ws.cell(r, 1).value == marker_value:
            return r
    return None


def derive_output_name(green_time_file_path):
    """Builds the output Excel filename from the Green Time file's own name,
    e.g. 'GreenTime_FINAL_West_14-28-15-03.xlsx' -> 'Final_Per_Cycle_Analysis_West_14-28-15-03.xlsx'.
    Falls back to a generic name if the expected 'GreenTime_FINAL_' prefix isn't found.
    Splits on both '/' and '\\' explicitly (rather than relying on os.path) so this
    behaves the same regardless of which OS the script happens to run on."""
    filename_only = re.split(r"[\\/]", green_time_file_path)[-1]
    base = os.path.splitext(filename_only)[0]
    m = re.match(r"^GreenTime_FINAL_(.+)$", base)
    suffix = m.group(1) if m else base
    return f"Final_Per_Cycle_Analysis_{suffix}.xlsx"


def load_green_phases(path, sheet_name=SHEET_GREEN_PHASES):
    """Returns a list of (green_start_s, green_end_s, is_partial) tuples, in order.
    is_partial is True whenever the sheet's 'Partial_Phase' column says 'Yes' for
    that row (case-insensitive) - meaning that phase's cycle cannot be closed out
    (no subsequent green observed in this session). If the sheet has no
    'Partial_Phase' column at all, every phase is treated as not-partial (False),
    so this stays backward-compatible with older Green Time files."""
    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet_name not in wb.sheetnames:
        sheet_name = _find_sheet_with_column(wb, "Dataset_ID")
    ws = wb[sheet_name]
    header_row = _find_header_row(ws, "Dataset_ID")
    if header_row is None:
        raise ValueError(f"Could not find header row ('Dataset_ID') in '{sheet_name}'.")
    col = {ws.cell(header_row, c).value: c for c in range(1, ws.max_column + 1)}
    partial_col = col.get("Partial_Phase")  # None if this file doesn't have the column

    phases = []
    for r in range(header_row + 1, ws.max_row + 1):
        gs, ge = ws.cell(r, col["Green_Start_s"]).value, ws.cell(r, col["Green_End_s"]).value
        if gs is None or ge is None:
            break
        is_partial = False
        if partial_col is not None:
            flag = ws.cell(r, partial_col).value
            is_partial = str(flag).strip().lower() == "yes" if flag is not None else False
        phases.append((float(gs), float(ge), is_partial))
    return phases


def load_density_data(path, sheet_name=SHEET_SECOND_ACCUMULATION,
                       distance_sheet=SHEET_BUCKETS_FOR_DISTANCE,
                       categories=VEHICLE_CATEGORIES):
    """
    Returns (sec_df, dist_km, pcu_weight):
      sec_df     - DataFrame[t, veh, pcu] at 1-second resolution
      dist_km    - constant average traveled distance (km)
      pcu_weight - dict {category: PCU-equivalence weight}, solved by least
                   squares from this file's own category counts vs. its
                   Total PCU column (so it's always internally consistent,
                   never a guessed/standard value).
    """
    wb = openpyxl.load_workbook(path, data_only=True)

    if sheet_name not in wb.sheetnames:
        sheet_name = _find_sheet_with_column(wb, "Time (s)")
    ws = wb[sheet_name]
    header = [c.value for c in ws[1]]
    time_col = header.index("Time (s)") + 1
    veh_col = header.index("Total accumulated vehicles (vehicle)") + 1
    pcu_col = header.index("Total accumulated vehicles (PCU)") + 1
    cat_cols = [header.index(c) + 1 for c in categories]

    rows, X, y = [], [], []
    for r in range(2, ws.max_row + 1):
        t_str = ws.cell(r, time_col).value
        if t_str is None:
            continue
        t0 = int(str(t_str).split("-")[0])
        veh, pcu = ws.cell(r, veh_col).value, ws.cell(r, pcu_col).value
        rows.append((t0, veh, pcu))
        X.append([ws.cell(r, c).value for c in cat_cols])
        y.append(pcu)
    sec_df = pd.DataFrame(rows, columns=["t", "veh", "pcu"]).sort_values("t").reset_index(drop=True)

    X, y = np.array(X, dtype=float), np.array(y, dtype=float)
    weights, *_ = np.linalg.lstsq(X, y, rcond=None)
    pcu_weight = dict(zip(categories, weights))

    if distance_sheet not in wb.sheetnames:
        distance_sheet = _find_sheet_with_column(wb, "Avg Traveled Distance (km)")
    ws2 = wb[distance_sheet]
    header2 = [c.value for c in ws2[1]]
    dist_col = header2.index("Avg Traveled Distance (km)") + 1
    dist_km = None
    for r in range(2, ws2.max_row + 1):
        v = ws2.cell(r, dist_col).value
        if v is not None:
            dist_km = float(v)
            break
    if dist_km is None:
        raise ValueError(f"Could not find 'Avg Traveled Distance (km)' in '{distance_sheet}'.")

    return sec_df, dist_km, pcu_weight


def load_raw_trajectories(path):
    """Returns a list of (Track ID, Type, Entry Time [s]) for every tracked vehicle."""
    with open(path, encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        header = next(reader)
        header = [h.strip() for h in header]
        tid_i, type_i, entry_i = header.index("Track ID"), header.index("Type"), header.index("Entry Time [s]")
        vehicles = [(int(row[tid_i]), row[type_i].strip(), float(row[entry_i])) for row in reader]
    return vehicles


def load_stop_track_ids(path, sheet_name):
    """Returns the set of Track IDs listed in the given sheet (Multi-Stop or Single-Stop)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet_name not in wb.sheetnames:
        sheet_name = _find_sheet_with_column(wb, "Track ID")
    ws = wb[sheet_name]
    header = [c.value for c in ws[1]]
    id_col = header.index("Track ID") + 1
    return set(ws.cell(r, id_col).value for r in range(2, ws.max_row + 1) if ws.cell(r, id_col).value is not None)


# =====================================================================
# CORE CALCULATIONS
# =====================================================================
def density_coverage_and_means(t0, t1, sec_df, dist_km, min_fraction=DENSITY_COVERAGE_MIN_FRACTION):
    """
    Checks how much of the cycle window [t0, t1) actually has density data
    available, and returns (coverage_pct, dens_veh, dens_pcu):
      coverage_pct - the ACTUAL percentage of the cycle's seconds that have
                      density data, reported regardless of whether it passed
                      the threshold (never hidden).
      dens_veh, dens_pcu - the mean density over whatever seconds ARE
                      available, UNLESS coverage_pct is below min_fraction,
                      in which case both are NaN (left blank in the output)
                      rather than silently averaged over a partial,
                      systematically-biased subset of the cycle.
    """
    sub = sec_df[(sec_df.t >= t0) & (sec_df.t < t1)]
    expected_seconds = t1 - t0
    actual_seconds = len(sub)
    coverage = min(1.0, actual_seconds / expected_seconds) if expected_seconds > 0 else 0.0
    coverage_pct = round(100 * coverage, 1)

    if coverage < min_fraction or actual_seconds == 0:
        return coverage_pct, np.nan, np.nan
    return coverage_pct, sub["veh"].mean() / dist_km, sub["pcu"].mean() / dist_km


def safe_pct(numerator, denominator):
    """Returns numerator/denominator as a percentage, or NaN if denominator is 0
    (an empty cycle shouldn't silently show 0% - it should show 'not applicable')."""
    return round(100 * numerator / denominator, 2) if denominator else np.nan


def build_cycle_table(green_phases, sec_df, dist_km, pcu_weight, vehicles,
                       multi_ids, single_ids, n_lanes, s_per_lane):
    starts = [p[0] for p in green_phases]
    durations = [p[1] - p[0] for p in green_phases]
    is_partial_flags = [p[2] for p in green_phases]
    s_group = n_lanes * s_per_lane

    records = []
    for i in range(len(starts) - 1):
        if is_partial_flags[i]:
            print(f"[excluded] Cycle {i + 1} (Phase {i + 1}, green start={starts[i]}s) skipped: "
                  f"marked Partial_Phase = Yes in the Green Time file (cycle cannot be closed out).")
            continue
        t0, t1 = starts[i], starts[i + 1]
        cyc_len = t1 - t0
        g = durations[i]
        gtr = g / cyc_len

        dens_coverage_pct, dens_veh, dens_pcu = density_coverage_and_means(t0, t1, sec_df, dist_km)
        if np.isnan(dens_veh):
            print(f"[density blanked] Cycle {i + 1} ({t0}-{t1}s): only {dens_coverage_pct}% density data "
                  f"coverage (< {DENSITY_COVERAGE_MIN_FRACTION*100:.0f}% required) - "
                  f"Density columns left blank for this cycle.")

        in_cycle = [(tid, typ) for tid, typ, en in vehicles if t0 <= en < t1]
        total_veh = len(in_cycle)
        n_multi = sum(1 for tid, _ in in_cycle if tid in multi_ids)
        n_single = sum(1 for tid, _ in in_cycle if tid in single_ids)
        n_nostop = total_veh - n_multi - n_single
        n_stopped = n_single + n_multi

        total_pcu = sum(pcu_weight[typ] for _, typ in in_cycle)
        multi_pcu = sum(pcu_weight[typ] for tid, typ in in_cycle if tid in multi_ids)
        single_pcu = sum(pcu_weight[typ] for tid, typ in in_cycle if tid in single_ids)
        nostop_pcu = total_pcu - multi_pcu - single_pcu

        v_veh_h = total_veh * (3600 / cyc_len)
        v_pcu_h = total_pcu * (3600 / cyc_len)
        c = s_group * gtr
        dos = v_pcu_h / c

        records.append({
            "Number of Cycle": i + 1,
            "Cycle Time Start (s)": t0,
            "Cycle Time End (s)": t1,
            "Green Period (s)": g,
            "Green Time Ratio": round(gtr, 4),
            "Density Data Coverage (%)": dens_coverage_pct,
            "Density (vehicle/km)": round(dens_veh, 2),
            "Density (PCU/km)": round(dens_pcu, 2),
            "Total Number of Vehicles": total_veh,
            "Entry Flow Rate (vehicle/hour)": round(v_veh_h, 1),
            "Number of NO STOP Vehicles": n_nostop,
            "Number of Single Stop Vehicles": n_single,
            "Number of Multiple Stop Vehicles": n_multi,
            "Percentage of NO STOP Vehicles (%)": safe_pct(n_nostop, total_veh),
            "Percentage of Single Stop Vehicles (%)": safe_pct(n_single, total_veh),
            "Percentage of Multiple Stop Vehicles (%)": safe_pct(n_multi, total_veh),
            "Percentage of Stop Vehicles (Single+Multiple) (%)": safe_pct(n_stopped, total_veh),
            "Total Number of Vehicles (PCU)": round(total_pcu, 2),
            "Entry Flow Rate (PCU/hour)": round(v_pcu_h, 1),
            "Number of NO STOP Vehicles (PCU)": round(nostop_pcu, 2),
            "Number of Single Stop Vehicles (PCU)": round(single_pcu, 2),
            "Number of Multiple Stop Vehicles (PCU)": round(multi_pcu, 2),
            "Degree of Saturation (DoS)": round(dos, 3),
        })
    return pd.DataFrame(records)


# =====================================================================
# MAIN
# =====================================================================
def main():
    output_dir = os.path.dirname(RAW_TRAJECTORY_CSV)
    os.makedirs(output_dir, exist_ok=True)
    output_excel_name = OUTPUT_EXCEL_NAME_OVERRIDE or derive_output_name(GREEN_TIME_FILE)

    n_lanes = get_n_lanes_from_user()

    green_phases = load_green_phases(GREEN_TIME_FILE)
    sec_df, dist_km, pcu_weight = load_density_data(DENSITY_FILE)
    vehicles = load_raw_trajectories(RAW_TRAJECTORY_CSV)
    multi_ids = load_stop_track_ids(STOP_ANALYSIS_FILE, SHEET_MULTISTOP)
    single_ids = load_stop_track_ids(STOP_ANALYSIS_FILE, SHEET_SINGLESTOP)

    print(f"Loaded {len(green_phases)} green phases.")
    print(f"Loaded {len(sec_df)} one-second density rows. Distance constant = {dist_km} km.")
    print(f"Auto-derived PCU weights: { {k: round(v, 3) for k, v in pcu_weight.items()} }")
    print(f"Loaded {len(vehicles)} tracked vehicles from raw trajectory CSV.")
    print(f"Multi-stop IDs: {len(multi_ids)}, Single-stop IDs: {len(single_ids)}")
    print(f"DoS assumption: N_LANES={n_lanes} (entered manually), S_PER_LANE={S_PER_LANE} PCU/h "
          f"-> lane-group saturation flow = {n_lanes * S_PER_LANE} PCU/h")
    print(f"Output will be saved as: {output_excel_name}\n")

    df = build_cycle_table(green_phases, sec_df, dist_km, pcu_weight, vehicles,
                            multi_ids, single_ids, n_lanes, S_PER_LANE)
    print(df.to_string(index=False))

    notes = [
        "Cycle definition: interval from one green-phase start to the next. A phase is "
        "excluded from becoming a cycle if the Green Time file marks it Partial_Phase = "
        "Yes (meaning its cycle cannot be closed out - no subsequent green was observed "
        "in that session). This is read directly from the Partial_Phase column, not "
        "inferred structurally, so it also catches a partial phase that isn't simply the "
        "last one in the file.",
        "Green Time Ratio (GTR) = this phase's green duration / this cycle's length.",
        "Density (vehicle/km) and Density (PCU/km): the mean of per-second INSTANTANEOUS "
        "occupancy snapshots (vehicles physically present at each second-mark) across the "
        "cycle, divided by the constant average traveled distance (km).",
        f"Density Data Coverage (%): the actual percentage of a cycle's seconds that have "
        f"density data available (some sessions' density files don't cover the full "
        f"recording window). This is always reported, even when it's 100%. If coverage is "
        f"below {DENSITY_COVERAGE_MIN_FRACTION*100:.0f}%, BOTH density columns are left "
        f"BLANK for that cycle rather than averaged over an incomplete, systematically-"
        f"biased subset (the missing seconds are typically right after a green start, "
        f"which is not a representative slice of the whole cycle) - see console output "
        f"for which cycles, if any, were blanked this run.",
        "Total Number of Vehicles / Entry Flow Rate / No-Stop / Single-Stop / Multi-Stop "
        "counts: an ARRIVAL EVENT count - every vehicle counted exactly once, assigned to "
        "the cycle its Entry Time [s] falls into (from the raw trajectory CSV). Single-Stop "
        "and Multi-Stop status comes from Track ID membership in the stop-analysis "
        "workbook's sheets; No-Stop = Total - Single-Stop - Multi-Stop.",
        "PERCENTAGE columns (No-Stop / Single-Stop / Multi-Stop / Stopped) are calculated "
        "from VEHICLE COUNTS, not PCU counts. A stop is a per-vehicle event, so weighting "
        "it by physical road-space (PCU) would distort the percentage toward whichever "
        "vehicle mix happened to stop, rather than reflecting the true share of vehicles "
        "that experienced a stop. This also matches standard practice: 'percentage of "
        "vehicles stopped' is a well-established measure of effectiveness (Webster's "
        "method, HCM) and is always defined per-vehicle. 'Percentage of Stop Vehicles "
        "(Single+Multiple)' = (Single-Stop + Multi-Stop) / Total, i.e. any vehicle that "
        "stopped at least once, regardless of how many times.",
        "PCU-suffixed count columns apply the PCU-equivalence weight (solved from this "
        "session's own density data, printed above / in the 'PCU Weights Used' sheet) to "
        "each vehicle instead of counting it as 1 - used for Density and DoS, where "
        "physical road-space matters, not for the percentage columns.",
        "Entry Flow Rate = the cycle's vehicle/PCU count scaled to an hourly rate: "
        "count x (3600 / cycle length in seconds).",
        f"Degree of Saturation (DoS) = Entry Flow Rate (PCU/h) / capacity, where capacity = "
        f"{S_PER_LANE} PCU/h/lane x N_LANES x GTR. S_PER_LANE is a standard textbook "
        f"default, not field-measured. N_LANES is now entered manually at the start of each "
        f"run (used value: {n_lanes}) rather than hardcoded - confirm your lane count "
        f"against satellite imagery if possible, since DoS scales directly with this input.",
        "DoS > 1.0 indicates an oversaturated cycle (demand exceeds what the green time can "
        "discharge). Values well above ~1.2-1.3 should be read qualitatively rather than "
        "trusted to the third decimal, since standard delay models become unstable near and "
        "beyond v/c = 1.",
        "Note: Density uses an instantaneous-presence counting logic, while the vehicle/PCU "
        "count columns use an entry-time arrival logic. These are intentionally different, "
        "standard traffic-engineering measures (occupancy vs. flow) and are not expected to "
        "count the same underlying 'instances'.",
    ]

    out_path = os.path.join(output_dir, output_excel_name)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Per-Cycle Summary", index=False)
        pd.DataFrame({"Notes": notes}).to_excel(writer, sheet_name="Assumptions and Notes", index=False)
        pd.DataFrame({"Vehicle Category": list(pcu_weight.keys()),
                      "PCU Weight (auto-derived)": [round(v, 3) for v in pcu_weight.values()]}
                     ).to_excel(writer, sheet_name="PCU Weights Used", index=False)

    print(f"\nSaved -> {out_path}")


if __name__ == "__main__":
    main()