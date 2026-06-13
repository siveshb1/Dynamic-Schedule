import streamlit as st
import pandas as pd
import plotly.express as px
from ortools.sat.python import cp_model
from datetime import datetime, timedelta
import math

# =============================================================================
# Page Config
# =============================================================================
st.set_page_config(layout="wide", page_title="Live Batch Scheduler Pro")
st.title("🏭 Shop Floor Live Batch Scheduler")
st.markdown("""
**Operator Instructions:**
1. Upload your `BatchSchedule_Template.xlsx` via the sidebar.
2. View the optimised timeline below. The red dashed line is **Current Time (Now)**.
3. To pin a task, edit its start/end time, tick **Lock?**, then click **🚀 Re-Optimise**.
""")

# =============================================================================
# Fallback hardcoded recipe
# =============================================================================
DEFAULT_RECIPE = [
    [1,  "LC_PPM",                 1.0,  [],               [],                       0,   []],
    [2,  "PPM_Disp",               2.0,  [1],              [],                       0,   []],
    [3,  "LC_RM",                  2.0,  [],               [],                       0,   []],
    [4,  "RM_Disp",                5.0,  [3],              [],                       0,   []],
    [5,  "FV01_CIP",               6.0,  [],               ["CIP01","FV03"],         0,   []],
    [6,  "TV51_CIP",               6.0,  [],               ["CIP01"],                0,   []],
    [7,  "FV03_CIP",               6.0,  [],               ["CIP01","FV03"],         0,   []],
    [8,  "TV51_PreFIT",            0.75, [2],              ["ITM55"],                0,   []],
    [9,  "TV51_PHTC",              1.0,  [6,8],            ["ITMSKID"],              24,  [6,8]],
    [10, "TV51_SIP",               7.0,  [9],              ["Facility_temp"],        0,   []],
    [11, "TV51_PHTS",              1.0,  [10],             ["ITMSKID"],              0,   []],
    [12, "FV01_PreFIT",            1.5,  [2],              ["ITM55"],                0,   []],
    [13, "FV01_PHTC",              1.0,  [5,12,29,30],     ["ITMSKID","FV03"],       24,  [5,12,29,30]],
    [14, "FV01_SIP",               7.0,  [13],             ["Facility_temp","FV03"], 0,   []],
    [15, "FV01_PHTS",              1.0,  [14],             ["ITMSKID","FV03"],       0,   []],
    [16, "FV03_PreFIT",            0.75, [2],              ["ITM55"],                0,   []],
    [17, "FV03_PHTC",              1.0,  [7,16],           ["FV03"],                 24,  [7,16]],
    [18, "FV03_SIP",               6.0,  [17],             ["Facility_temp","FV03"], 0,   []],
    [19, "FV03_PHTS",              1.0,  [18],             ["FV03"],                 0,   []],
    [20, "MV60_CIP",               1.5,  [],               ["CIP01","CIP09"],        0,   []],
    [21, "MV60_PreFIT",            0.5,  [2],              ["ITM55"],                0,   []],
    [22, "MV60_SIP",               3.0,  [20,21],          ["CIP09"],                24,  [20,21]],
    [23, "MV61_CIP",               1.5,  [],               ["CIP01","CIP09"],        0,   []],
    [24, "MV61_PreFIT",            0.5,  [2],              ["ITM55"],                0,   []],
    [25, "MV61_SIP",               3.0,  [23,24],          ["CIP09"],                24,  [25]],
    [26, "MV56_CIP",               1.5,  [],               ["CIP01","CIP09"],        0,   []],
    [27, "MV56_PreFIT",            0.5,  [2],              ["ITM55"],                0,   []],
    [28, "MV56_SIP",               3.0,  [26,27],          ["CIP09"],                24,  [28]],
    [29, "ITM51_PreFIT_EMFLON5",   1.0,  [2],              ["ITM55"],                0,   []],
    [30, "ITM51_PreFIT_Milidisk",  1.0,  [2],              ["ITM55"],                0,   []],
    [31, "PlateExp",               0.5,  [37],             [],                       0,   []],
    [32, "WaterColl_FV01",         1.0,  [15],             [],                       72,  [15]],
    [33, "WaterColl_FV03",         1.0,  [19],             [],                       72,  [19]],
    [34, "WaterColl_FV60",         1.0,  [22],             ["UP52"],                 24,  [22]],
    [35, "WaterColl_FV61",         1.0,  [25],             ["UP52"],                 24,  [25]],
    [36, "WaterColl_FV56",         1.0,  [28],             ["UP52"],                 24,  [28]],
    [37, "Durapore PreFIT",        1.0,  [2,15,19],        ["ITM55"],                0,   []],
    [38, "DocVerify",              3.0,  [11,15,19,22,25,28,37], [],                 0,   []],
    [39, "Formulation",           10.0,  [4,31,32,33,34,35,36,38], [],              0,   []],
    [40, "Filtration",             3.0,  [39],             [],                       72,  [11]],
    [41, "Postfit Durapore",       2.0,  [40],             ["ITM55"],                0,   []],
]

# =============================================================================
# Sidebar
# =============================================================================
st.sidebar.header("📁 Upload Production Template")
uploaded_file = st.sidebar.file_uploader(
    "BatchSchedule_Template.xlsx", type=["xlsx"],
    help="Upload the template Excel file. All fields are read from the 'BatchSchedule' sheet."
)

st.sidebar.header("🗓️ Batch Calendar Start")
batch_date = st.sidebar.date_input("Start Date", datetime(2026, 6, 13).date())
batch_time = st.sidebar.time_input("Start Time", datetime(2026, 6, 13, 7, 22).time())
batch_start_dt = datetime.combine(batch_date, batch_time)

enforce_now = st.sidebar.checkbox(
    "Enforce Current-Time Floor",
    value=True,
    help="When ON, unlocked tasks cannot be scheduled before NOW."
)

# =============================================================================
# Excel Parser helpers
# =============================================================================
def _parse_int_list(val) -> list:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return []
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return []
    result = []
    for tok in s.split(","):
        tok = tok.strip()
        if tok:
            try:
                result.append(int(float(tok)))
            except ValueError:
                pass
    return result

def _parse_str_list(val) -> list:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return []
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return []
    return [x.strip() for x in s.split(",") if x.strip()]

def _to_datetime(val):
    """Coerce Excel datetime / Timestamp / string to naive Python datetime, or None."""
    if val is None:
        return None
    try:
        if isinstance(val, pd.Timestamp):
            if pd.isna(val):
                return None
            dt = val.to_pydatetime()
        elif isinstance(val, datetime):
            dt = val
        else:
            ts = pd.to_datetime(val)
            if pd.isna(ts):
                return None
            dt = ts.to_pydatetime()
        return dt.replace(tzinfo=None)
    except Exception:
        return None

# =============================================================================
# Excel Parser
# =============================================================================
def parse_excel_template(df: pd.DataFrame, batch_start_dt: datetime):
    """
    Returns
    -------
    recipe    : list[tag, name, duration, deps, resources, lag_w, lag_deps]
    overrides : dict  tag -> {Start(Hr), End(Hr), Lock?, Started}
    errors    : list[str]  non-fatal warnings

    Override states
    ───────────────
    Lock=TRUE + Start + End  →  completed & pinned   (Lock?=True,  Started=True)
    Start only, no Lock      →  in-progress           (Lock?=False, Started=True)
    Nothing                  →  pending               (not in overrides)
    """
    recipe, overrides, errors = [], {}, []
    df.columns = [str(c).strip() for c in df.columns]

    required = {"Tag", "Task Name", "Duration (Hr)"}
    missing  = required - set(df.columns)
    if missing:
        raise ValueError(f"Excel is missing required columns: {missing}")

    for row_num, row in df.iterrows():
        excel_row = row_num + 2
        try:
            tag      = int(float(row["Tag"]))
            name     = str(row["Task Name"]).strip()
            duration = float(row["Duration (Hr)"])
        except (ValueError, TypeError) as exc:
            errors.append(f"Row {excel_row}: skipped — {exc}")
            continue

        deps     = _parse_int_list(row.get("Dependencies"))
        resources= _parse_str_list(row.get("Resources"))
        lag_w    = 0.0
        try:
            raw_lag = row.get("Lag Window (Hr)")
            if raw_lag is not None and not (isinstance(raw_lag, float) and pd.isna(raw_lag)):
                lag_w = float(raw_lag)
        except (ValueError, TypeError):
            errors.append(f"Row {excel_row} (tag {tag}): bad Lag Window — defaulting to 0")
        lag_deps = _parse_int_list(row.get("Lag Dependencies"))

        recipe.append([tag, name, duration, deps, resources, lag_w, lag_deps])

        lock_raw  = row.get("Lock Status")
        is_locked = bool(lock_raw) if not (isinstance(lock_raw, float) and pd.isna(lock_raw)) else False
        s_dt = _to_datetime(row.get("Start Override"))
        e_dt = _to_datetime(row.get("End Override"))

        if is_locked and s_dt and e_dt:
            # Fully completed and pinned
            try:
                s_hr = (s_dt - batch_start_dt).total_seconds() / 3600.0
                e_hr = (e_dt - batch_start_dt).total_seconds() / 3600.0
                if not (math.isfinite(s_hr) and math.isfinite(e_hr)):
                    raise ValueError("non-finite hours")
                overrides[tag] = {
                    "Start (Hr)": s_hr, "End (Hr)": e_hr,
                    "Lock?": True, "Started": True,
                }
            except Exception as exc:
                errors.append(
                    f"Row {excel_row} (tag {tag}): could not compute override hours ({exc}) "
                    f"— treated as unlocked."
                )
        elif s_dt and not is_locked:
            # Task started but not yet completed — lag constraint is consumed
            try:
                s_hr = (s_dt - batch_start_dt).total_seconds() / 3600.0
                if math.isfinite(s_hr):
                    overrides[tag] = {
                        "Start (Hr)": s_hr, "End (Hr)": None,
                        "Lock?": False, "Started": True,
                    }
            except Exception:
                pass
        elif is_locked and (not s_dt or not e_dt):
            errors.append(
                f"Row {excel_row} (tag {tag}): Lock=TRUE but Start/End Override missing "
                f"— treated as unlocked."
            )

    return recipe, overrides, errors

# =============================================================================
# Lag Window Analyser
# =============================================================================
def analyse_lag(recipe, overrides, batch_start_dt):
    """
    For every task with a lag window that has NOT yet started:
      - Identify which of its lag-deps are completed (locked with end time)
      - Check if dep_end is still within lag_w hours BEFORE now
        i.e.  now - lag_w  <=  dep_end  <=  now
        If dep_end < now - lag_w  →  dep is EXPIRED (too old), must be repeated
        If dep_end is within window →  still valid, normal constraint applies

    Returns
    -------
    repeat_tags : set of original tag IDs whose completed run is expired
                  → these must be treated as unlocked (repeated) in the solver
    warnings    : list of (task_tag, task_name, dep_tag, dep_name, dep_end_abs, window_close_abs)
                  for UI display
    """
    # Build lookups
    tag_info    = {t[0]: t for t in recipe}   # tag -> full recipe row
    completed   = {                            # tag -> end_hr  (locked tasks with end time)
        t[0]: overrides[t[0]]["End (Hr)"]
        for t in recipe
        if t[0] in overrides
        and overrides[t[0]].get("Lock?")
        and overrides[t[0]].get("End (Hr)") is not None
    }
    started     = {                            # tags whose lag is consumed
        t[0] for t in recipe
        if t[0] in overrides and overrides[t[0]].get("Started")
    }

    now_hr = (datetime.now() - batch_start_dt).total_seconds() / 3600.0

    repeat_tags = set()   # dep tags that need to be repeated
    warnings    = []

    for t in recipe:
        tag, name, _, _, _, lag_w, lag_deps = t
        if lag_w <= 0 or not lag_deps:
            continue
        # Rule 1: task already started → lag consumed, skip
        if tag in started:
            continue

        for ld in lag_deps:
            if ld not in completed:
                continue   # dep not done → solver handles the window normally

            dep_end_hr   = completed[ld]
            # Valid window: dep must have ended no earlier than (now - lag_w)
            # i.e. dep_end >= now - lag_w  →  dep is still fresh
            # i.e. dep_end <  now - lag_w  →  dep is EXPIRED, needs repeat
            earliest_valid_dep_end = now_hr - lag_w

            if dep_end_hr < earliest_valid_dep_end:
                # Expired — dep must be repeated
                repeat_tags.add(ld)
                dep_name      = tag_info[ld][1] if ld in tag_info else str(ld)
                dep_end_abs   = batch_start_dt + timedelta(hours=dep_end_hr)
                window_close  = dep_end_abs + timedelta(hours=lag_w)
                warnings.append({
                    "task_tag":       tag,
                    "task_name":      name,
                    "dep_tag":        ld,
                    "dep_name":       dep_name,
                    "dep_end_abs":    dep_end_abs,
                    "window_close":   window_close,
                    "lag_w":          lag_w,
                })

    return repeat_tags, warnings

# =============================================================================
# OR-Tools Solver
# =============================================================================
def solve_schedule(recipe, batch_start_dt, overrides=None,
                   enforce_now_flag=True, repeat_tags=None):
    """
    repeat_tags : set of tag IDs whose locked override should be IGNORED
                  (treated as unlocked / to be repeated).
                  The solver will re-schedule them freely so the lag window
                  of their dependent task is satisfied.
    """
    if overrides    is None: overrides    = {}
    if repeat_tags  is None: repeat_tags  = set()

    model   = cp_model.CpModel()
    SCALE   = 100
    horizon = int(400 * SCALE)

    task_starts, task_ends, task_intervals, resource_ivs = {}, {}, {}, {}

    current_hr     = (datetime.now() - batch_start_dt).total_seconds() / 3600.0
    current_scaled = max(0, int(current_hr * SCALE))

    for t in recipe:
        tag, _, duration, _, resources, _, _ = t

        # Pin task only if locked AND not flagged for repeat
        if tag in overrides and overrides[tag].get("Lock?") and tag not in repeat_tags:
            raw_s = overrides[tag]["Start (Hr)"]
            raw_e = overrides[tag]["End (Hr)"]
            if not (math.isfinite(raw_s) and math.isfinite(raw_e)):
                raise ValueError(
                    f"Tag {tag} locked but has invalid override hours "
                    f"(Start={raw_s}, End={raw_e})."
                )
            s_val   = int(raw_s * SCALE)
            e_val   = int(raw_e * SCALE)
            dur_val = e_val - s_val
            start_v = model.NewConstant(s_val)
            end_v   = model.NewConstant(e_val)
        else:
            # Free variable — solver places it optimally
            dur_val = int(duration * SCALE)
            start_v = model.NewIntVar(0, horizon, f"s_{tag}")
            end_v   = model.NewIntVar(0, horizon, f"e_{tag}")
            if enforce_now_flag and current_scaled > 0:
                model.Add(start_v >= current_scaled)

        interval = model.NewIntervalVar(start_v, dur_val, end_v, f"iv_{tag}")
        task_starts[tag]    = start_v
        task_ends[tag]      = end_v
        task_intervals[tag] = interval

        for res in resources:
            resource_ivs.setdefault(res, []).append(interval)

    # ── Precedence & lag constraints ─────────────────────────────────────────
    for t in recipe:
        tag, _, _, deps, _, lag_w, lag_deps = t

        for d in deps:
            if d in task_ends:
                model.Add(task_starts[tag] >= task_ends[d])

        if lag_w > 0:
            for ld in lag_deps:
                if ld in task_ends:
                    # Upper bound: task must start within lag_w of dep ending
                    model.Add(task_starts[tag] <= task_ends[ld] + int(lag_w * SCALE))

    # ── No-overlap on shared resources ───────────────────────────────────────
    for res, ivs in resource_ivs.items():
        if len(ivs) > 1:
            model.AddNoOverlap(ivs)

    # ── Minimise makespan ────────────────────────────────────────────────────
    makespan = model.NewIntVar(0, horizon, "makespan")
    model.AddMaxEquality(makespan, list(task_ends.values()))
    model.Minimize(makespan)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 30.0
    status = solver.Solve(model)

    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return {
            t[0]: (
                solver.Value(task_starts[t[0]]) / SCALE,
                solver.Value(task_ends[t[0]])   / SCALE,
            )
            for t in recipe
        }
    return None

# =============================================================================
# Load data from Excel or defaults
# =============================================================================
file_key  = (uploaded_file.name if uploaded_file else "default") + str(batch_start_dt)
state_key = f"sched_{file_key}"

parse_errors = []

if uploaded_file is not None:
    try:
        raw_df = pd.read_excel(uploaded_file, sheet_name="BatchSchedule")
        base_recipe, sheet_overrides, parse_errors = parse_excel_template(raw_df, batch_start_dt)
        if parse_errors:
            with st.expander("⚠️ Parser warnings (non-fatal)", expanded=False):
                for e in parse_errors:
                    st.warning(e)
    except Exception as exc:
        st.error(f"❌ Could not parse Excel file: {exc}. Falling back to built-in defaults.")
        base_recipe, sheet_overrides = DEFAULT_RECIPE, {}
else:
    base_recipe, sheet_overrides = DEFAULT_RECIPE, {}

# =============================================================================
# Session-state initialisation / global reset
# =============================================================================
reset_clicked = st.sidebar.button("♻️ Apply Start Time / Reset")

if state_key not in st.session_state or reset_clicked:

    # Analyse lag violations before solving
    repeat_tags, lag_warnings = analyse_lag(base_recipe, sheet_overrides, batch_start_dt)

    # Show warnings
    if lag_warnings:
        tag_lookup = {t[0]: t[1] for t in base_recipe}
        for w in lag_warnings:
            st.warning(
                f"⏰ **[{w['task_tag']}] {w['task_name']}** — "
                f"Tag **{w['dep_tag']} ({w['dep_name']})** completed at "
                f"`{w['dep_end_abs'].strftime('%Y-%m-%d %H:%M')}`, "
                f"{w['lag_w']:.0f}-hr window closed at "
                f"`{w['window_close'].strftime('%Y-%m-%d %H:%M')}`. "
                f"**Tag {w['dep_tag']} must be repeated.** "
                f"Solver has auto-scheduled a new run. 🔄"
            )

    with st.spinner("Running initial optimisation…"):
        init_results = solve_schedule(
            base_recipe, batch_start_dt,
            overrides=sheet_overrides,
            enforce_now_flag=enforce_now,
            repeat_tags=repeat_tags,
        )

    if init_results is None:
        st.error("❌ Infeasible schedule. Check dependencies / lock overrides.")
        st.stop()

    rows = []
    for t in base_recipe:
        tag, name, duration, _, _, _, _ = t
        s_hr, e_hr = init_results[tag]
        is_locked  = (
            sheet_overrides.get(tag, {}).get("Lock?", False)
            and tag not in repeat_tags
        )
        is_repeat  = tag in repeat_tags
        rows.append({
            "Tag":                   tag,
            "Task Name":             name,
            "Planned Duration (Hr)": duration,
            "Start Time":            batch_start_dt + timedelta(hours=s_hr),
            "End Time":              batch_start_dt + timedelta(hours=e_hr),
            "Lock?":                 is_locked,
            "Repeat":                is_repeat,
        })
    st.session_state[state_key] = pd.DataFrame(rows)

# =============================================================================
# Interactive table
# =============================================================================
st.subheader("📝 Real-Time Calendar Log Matrix")
col1, col2 = st.columns([1, 6])
with col1:
    reoptimise = st.button("🚀 Re-Optimise Schedule", type="primary")
with col2:
    if st.button("🔄 Reset All Locks"):
        if state_key in st.session_state:
            del st.session_state[state_key]
        st.rerun()

edited_df = st.data_editor(
    st.session_state[state_key],
    disabled=["Tag", "Task Name", "Planned Duration (Hr)", "Repeat"],
    hide_index=True,
    use_container_width=True,
    column_config={
        "Start Time": st.column_config.DatetimeColumn(
            "Start Time", format="YYYY-MM-DD HH:mm", step=60),
        "End Time":   st.column_config.DatetimeColumn(
            "End Time",   format="YYYY-MM-DD HH:mm", step=60),
        "Repeat":     st.column_config.CheckboxColumn(
            "Repeat 🔄", help="Auto-flagged tasks whose lag dep was expired and rescheduled"),
    }
)

# =============================================================================
# Re-optimisation logic
# =============================================================================
if reoptimise:
    overrides_from_ui = {}
    for _, row in edited_df.iterrows():
        s_hr = (pd.to_datetime(row["Start Time"]) - batch_start_dt).total_seconds() / 3600.0
        e_hr = (pd.to_datetime(row["End Time"])   - batch_start_dt).total_seconds() / 3600.0
        overrides_from_ui[int(row["Tag"])] = {
            "Start (Hr)": s_hr,
            "End (Hr)":   e_hr,
            "Lock?":      bool(row["Lock?"]),
            "Started":    bool(row["Lock?"]),  # locked = completed = started
        }

    # Re-analyse lag with current UI overrides
    repeat_tags, lag_warnings = analyse_lag(base_recipe, overrides_from_ui, batch_start_dt)

    if lag_warnings:
        for w in lag_warnings:
            st.warning(
                f"⏰ **[{w['task_tag']}] {w['task_name']}** — "
                f"Tag **{w['dep_tag']} ({w['dep_name']})** completed at "
                f"`{w['dep_end_abs'].strftime('%Y-%m-%d %H:%M')}`, "
                f"{w['lag_w']:.0f}-hr window closed at "
                f"`{w['window_close'].strftime('%Y-%m-%d %H:%M')}`. "
                f"**Tag {w['dep_tag']} must be repeated.** "
                f"Solver has auto-scheduled a new run. 🔄"
            )

    with st.spinner("Solving…"):
        new_results = solve_schedule(
            base_recipe, batch_start_dt,
            overrides=overrides_from_ui,
            enforce_now_flag=enforce_now,
            repeat_tags=repeat_tags,
        )

    if new_results is None:
        st.error("❌ Infeasible! Check locked tasks don't violate dependency order.")
    else:
        st.success("✅ Schedule updated successfully.")
        for idx, row in edited_df.iterrows():
            tag       = int(row["Tag"])
            is_repeat = tag in repeat_tags
            if not row["Lock?"] or is_repeat:
                edited_df.at[idx, "Start Time"] = batch_start_dt + timedelta(hours=new_results[tag][0])
                edited_df.at[idx, "End Time"]   = batch_start_dt + timedelta(hours=new_results[tag][1])
            edited_df.at[idx, "Repeat"] = is_repeat
        st.session_state[state_key] = edited_df
        st.rerun()

# =============================================================================
# Gantt chart
# =============================================================================
st.subheader("📊 Dynamic Live Gantt Visual")

tag_to_resources = {t[0]: t[4] for t in base_recipe}
chart_rows = []
current_df = st.session_state[state_key]

for _, row in current_df.iterrows():
    tag       = int(row["Tag"])
    is_repeat = bool(row.get("Repeat", False))
    is_locked = bool(row["Lock?"])
    res       = tag_to_resources.get(tag, [])


    if is_repeat:
        status = "Repeat 🔄"
    elif is_locked:
        status = "Locked ✅"
    else:
        status = "Dynamic"

    chart_rows.append({
        "Task Label":        f"[{tag}] {row['Task Name']}" + (" (Repeat)" if is_repeat else ""),
        "Start Window":      pd.to_datetime(row["Start Time"]),
        "End Window":        pd.to_datetime(row["End Time"]),
        "Assigned Resource": ", ".join(res) if res else "Core Process",
        "Status":            status,
    })

df_chart = pd.DataFrame(chart_rows)

fig = px.timeline(
    df_chart,
    x_start="Start Window",
    x_end="End Window",
    y="Task Label",
    color="Assigned Resource",
    pattern_shape="Status",
    title="<b>Live Operational Production Timeline</b>",
)
fig.update_yaxes(autorange="reversed")

now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
fig.add_shape(
    type="line",
    x0=now_str, x1=now_str, y0=0, y1=1,
    xref="x", yref="paper",
    line=dict(color="red", width=2, dash="dash"),
)
fig.add_annotation(
    x=now_str, y=1.02,
    xref="x", yref="paper",
    text="▼ Now",
    showarrow=False,
    xanchor="center",
    font=dict(color="red", size=11),
)
fig.update_layout(
    height=900,
    xaxis=dict(tickformat="%m/%d %H:%M"),
    legend=dict(orientation="h", y=-0.12),
)
st.plotly_chart(fig, use_container_width=True)
