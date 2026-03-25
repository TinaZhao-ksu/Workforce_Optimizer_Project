
import hashlib
import streamlit as st
import pandas as pd
import pulp

# ─── PAGE CONFIG ──────────────────────────────────────────────────────────────
# Set the browser tab title, icon, and use full-width layout
st.set_page_config(page_title="Workforce Optimizer", page_icon="⚡", layout="wide")

# ─── SIDEBAR WIDTH ────────────────────────────────────────────────────────────
# Streamlit's default sidebar max-width is ~350px which clips longer labels.
# This CSS override bumps it to 420px and lets it resize freely up to that point.
st.markdown(
    """
    <style>
        [data-testid="stSidebar"] { min-width: 320px; max-width: 420px; }
        [data-testid="stSidebar"] .stFileUploader { width: 100%; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ─── CONSTANTS ────────────────────────────────────────────────────────────────
# Emoji color indicators for each skill type used in display throughout the UI
SKILL_COLORS = {"A": "🟠", "B": "🔵", "C": "🟢", "D": "🟣", "E": "🔴"}

# Expected column names for each sheet in the uploaded Excel template.
# These are validated on upload to give the user clear feedback if something is missing.
REQUIRED_EMPLOYEE_COLS = {"employee", "capacity", "skills", "hourly_rate ($)"}
REQUIRED_PROJECT_COLS  = {"project", "reimbursable", "max_total", "budget ($)"}
REQUIRED_TASK_COLS     = {"project", "task_id", "task_type", "min_hours"}


def parse_excel_bool(value) -> bool:
    """Parse booleans from Excel safely (handles TRUE/FALSE strings and 0/1)."""
    if pd.isna(value):
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    s = str(value).strip().lower()
    if s in {"true", "t", "yes", "y", "1"}:
        return True
    if s in {"false", "f", "no", "n", "0", ""}:
        return False
    raise ValueError(f"Invalid reimbursable value: {value!r}. Use TRUE/FALSE, YES/NO, or 1/0.")


# ─── BLANK TEMPLATE BUILDER ───────────────────────────────────────────────────
# Generates an in-memory Excel file with the three required sheets pre-populated
# with headers and one example row each, plus a Data Dictionary sheet.
# Returns raw bytes so Streamlit can serve it as a download without touching disk.
def build_blank_template() -> bytes:
    import io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()

    # ── color palette ──
    HEADER_FILL  = PatternFill("solid", start_color="1F4E79")   # dark navy
    EXAMPLE_FILL = PatternFill("solid", start_color="D9E1F2")   # light blue tint
    DICT_FILL    = PatternFill("solid", start_color="375623")    # dark green
    THIN = Side(style="thin", color="BFBFBF")
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    def style_header(cell, green=False):
        cell.font      = Font(bold=True, color="FFFFFF", size=11)
        cell.fill      = DICT_FILL if green else HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = BORDER

    def style_example(cell):
        cell.fill      = EXAMPLE_FILL
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border    = BORDER

    def style_note(cell):
        cell.font      = Font(italic=True, color="595959", size=9)
        cell.alignment = Alignment(horizontal="left")

    def set_col_widths(ws, widths):
        for col, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = w

    # ── Employees sheet ──
    ws_emp = wb.active
    ws_emp.title = "Employees"
    emp_headers  = ["employee", "capacity", "skills", "hourly_rate ($)", "employee_type"]
    emp_example  = ["E001", 40, "A,B", 67.50, "Senior"]
    emp_example2 = ["E002", 30, "B,C,D", 55.00, "Mid-Level"]
    for c, h in enumerate(emp_headers, 1):
        style_header(ws_emp.cell(1, c, h))
    for c, v in enumerate(emp_example, 1):
        style_example(ws_emp.cell(2, c, v))
    for c, v in enumerate(emp_example2, 1):
        style_example(ws_emp.cell(3, c, v))
    # note row
    notes = ["Unique ID (e.g. E001)", "Max hrs/week (integer)",
             "Comma-separated A-E", "Hourly billing rate ($)", "Junior / Mid-Level / Senior"]
    for c, n in enumerate(notes, 1):
        style_note(ws_emp.cell(4, c, n))
    set_col_widths(ws_emp, [16, 14, 18, 18, 16])
    ws_emp.row_dimensions[1].height = 20

    # ── Projects sheet ──
    ws_proj = wb.create_sheet("Projects")
    proj_headers  = ["project", "reimbursable", "max_total", "budget ($)"]
    proj_example  = ["P001", True, 90, 12000]
    proj_example2 = ["P002", False, None, 8500]
    for c, h in enumerate(proj_headers, 1):
        style_header(ws_proj.cell(1, c, h))
    for c, v in enumerate(proj_example, 1):
        style_example(ws_proj.cell(2, c, v))
    for c, v in enumerate(proj_example2, 1):
        style_example(ws_proj.cell(3, c, v))
    notes_p = ["Unique ID (e.g. P001)", "TRUE = client-billable, FALSE = internal",
               "Max hours cap (leave blank if not reimbursable)", "Total dollar budget ($)"]
    for c, n in enumerate(notes_p, 1):
        style_note(ws_proj.cell(4, c, n))
    set_col_widths(ws_proj, [16, 16, 14, 14])
    ws_proj.row_dimensions[1].height = 20

    # ── Tasks sheet ──
    ws_task = wb.create_sheet("Tasks")
    task_headers  = ["project", "task_id", "task_type", "min_hours", "Urgency"]
    task_example  = ["P001", "T001", "A", 10, None]
    task_example2 = ["P001", "T002", "B", 14, None]
    for c, h in enumerate(task_headers, 1):
        style_header(ws_task.cell(1, c, h))
    for c, v in enumerate(task_example, 1):
        style_example(ws_task.cell(2, c, v))
    for c, v in enumerate(task_example2, 1):
        style_example(ws_task.cell(3, c, v))
    notes_t = ["Must match a project ID", "Unique task ID (e.g. T001)",
               "Skill type required (A / B / C / D / E)", "Minimum hours to complete", "Optional — leave blank"]
    for c, n in enumerate(notes_t, 1):
        style_note(ws_task.cell(4, c, n))
    set_col_widths(ws_task, [14, 14, 14, 14, 12])
    ws_task.row_dimensions[1].height = 20

    # ── Data Dictionary sheet ──
    ws_dict = wb.create_sheet("Data Dictionary")
    dict_headers = ["Sheet", "Column", "Type", "Description"]
    dict_rows = [
        ("Employees", "employee",       "ID",        "Unique employee identifier (e.g. E001)"),
        ("Employees", "capacity",       "Integer",   "Max weekly hours available"),
        ("Employees", "skills",         "String",    "Comma-separated skill types (A-E)"),
        ("Employees", "hourly_rate ($)","Float ($)",  "Hourly billing/wage rate"),
        ("Employees", "employee_type",  "Category",  "Junior (<25 h), Mid-Level (25-39 h), Senior (40+ h)"),
        ("Projects",  "project",        "ID",        "Unique project identifier (e.g. P001)"),
        ("Projects",  "reimbursable",   "Boolean",   "TRUE = client-billable; FALSE = internal/overhead"),
        ("Projects",  "max_total",      "Integer",   "Max hours budget for reimbursable projects (blank = no cap)"),
        ("Projects",  "budget ($)",     "Float ($)",  "Total dollar budget. LP uses this as a wage cost ceiling."),
        ("Tasks",     "project",        "ID",        "Project this task belongs to"),
        ("Tasks",     "task_id",        "ID",        "Unique task identifier (e.g. T001)"),
        ("Tasks",     "task_type",      "Category",  "Required skill type (A / B / C / D / E)"),
        ("Tasks",     "min_hours",      "Integer",   "Minimum hours required to complete task"),
        ("Tasks",     "Urgency",        "Optional",  "Not used by optimizer — reserved for future use"),
    ]
    for c, h in enumerate(dict_headers, 1):
        style_header(ws_dict.cell(1, c, h), green=True)
    for r, row in enumerate(dict_rows, 2):
        for c, v in enumerate(row, 1):
            cell        = ws_dict.cell(r, c, v)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="left", vertical="center")
    set_col_widths(ws_dict, [14, 18, 12, 52])
    ws_dict.row_dimensions[1].height = 20

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─── TEMPLATE VALIDATION ──────────────────────────────────────────────────────
# Checks that the uploaded file has the expected sheets and required columns.
# Returns (True, None) if valid or (False, error_message) if something is wrong.
def validate_template(xl: pd.ExcelFile) -> tuple[bool, str | None]:
    required_sheets = {"Employees", "Projects", "Tasks"}
    missing_sheets  = required_sheets - set(xl.sheet_names)
    if missing_sheets:
        return False, f"Missing sheet(s): {', '.join(sorted(missing_sheets))}. Expected: Employees, Projects, Tasks."

    emp_df  = pd.read_excel(xl, sheet_name="Employees", nrows=0)
    proj_df = pd.read_excel(xl, sheet_name="Projects",  nrows=0)
    task_df = pd.read_excel(xl, sheet_name="Tasks",     nrows=0)

    emp_cols  = set(emp_df.columns.str.strip())
    proj_cols = set(proj_df.columns.str.strip())
    task_cols = set(task_df.columns.str.strip())

    missing_emp  = REQUIRED_EMPLOYEE_COLS - emp_cols
    missing_proj = REQUIRED_PROJECT_COLS  - proj_cols
    missing_task = REQUIRED_TASK_COLS     - task_cols

    errors = []
    if missing_emp:
        errors.append(f"Employees sheet missing column(s): {', '.join(missing_emp)}")
    if missing_proj:
        errors.append(f"Projects sheet missing column(s): {', '.join(missing_proj)}")
    if missing_task:
        errors.append(f"Tasks sheet missing column(s): {', '.join(missing_task)}")

    if errors:
        return False, "\n".join(errors)
    return True, None


# ─── DATA LOADER ──────────────────────────────────────────────────────────────
# Reads the three core sheets from the uploaded Excel file and converts them
# into the list-of-dicts format the optimizer and UI expect.
# Strips whitespace from string fields and handles NaN gracefully.
def load_data(xl: pd.ExcelFile) -> tuple[list, list, list]:
    # --- Employees ---
    emp_df = pd.read_excel(xl, sheet_name="Employees")
    emp_df.columns = emp_df.columns.str.strip()
    employees = []
    for _, row in emp_df.iterrows():
        raw_skills = str(row["skills"]).strip()
        # Skills stored as comma-separated string (e.g. "A,B,C") → parse to list
        skills = [s.strip() for s in raw_skills.split(",") if s.strip()]
        employees.append({
            "id":       str(row["employee"]).strip(),
            "capacity": int(row["capacity"]),
            "skills":   skills,
            "rate":     float(row["hourly_rate ($)"]),   # hourly wage used in budget LP constraint
            "type":     str(row.get("employee_type", "")).strip(),
        })

    # --- Projects ---
    proj_df = pd.read_excel(xl, sheet_name="Projects")
    proj_df.columns = proj_df.columns.str.strip()
    projects = []
    for _, row in proj_df.iterrows():
        max_total = None if pd.isna(row["max_total"]) else int(row["max_total"])
        budget    = None if pd.isna(row["budget ($)"]) else float(row["budget ($)"])
        projects.append({
            "id":          str(row["project"]).strip(),
            "reimbursable": parse_excel_bool(row["reimbursable"]),
            "maxTotal":    max_total,   # max hours cap (reimbursable projects only)
            "budget":      budget,      # dollar budget cap used in LP wage constraint
        })

    # --- Tasks ---
    task_df = pd.read_excel(xl, sheet_name="Tasks")
    task_df.columns = task_df.columns.str.strip()
    tasks = []
    for _, row in task_df.iterrows():
        tasks.append({
            "project":  str(row["project"]).strip(),
            "id":       str(row["task_id"]).strip(),
            "type":     str(row["task_type"]).strip(),
            "minHours": int(row["min_hours"]),
        })

    return employees, projects, tasks


# ─── LINEAR PROGRAM OPTIMIZER ─────────────────────────────────────────────────
# Uses PuLP to formulate and solve a linear program that maximizes task coverage.
#
# Decision variables:
#   x[i][j]  = continuous hours employee i spends on task j  (>= 0)
#   s[j]     = unmet/slack hours for task j                  (>= 0)
#
# Objective: minimize sum(weight_j * s[j])
#   where weight_j is 2 for reimbursable-project tasks, 1 for non-reimbursable.
#   This ensures reimbursable tasks are prioritized by the optimizer.
#
# Constraints:
#   1. Employee capacity:  sum_j(x[i][j]) <= capacity[i]          for each employee i
#   2. Task demand:        sum_i(x[i][j]) + s[j] >= minHours[j]   for each task j
#   3. Skill match:        x[i][j] = 0  if employee i lacks the required skill for task j
#   4. Project hour cap:   sum_{j in p} sum_i(x[i][j]) <= maxTotal[p]  (reimbursable)
#   5. Project $ budget:   sum_{j in p} sum_i(x[i][j] * rate[i]) <= budget[p]  for each project p
#
# Returns a dict with:
#   asgn     - {task_id: {employee, hours, partial}} assignment map
#   load     - {employee_id: total hours assigned}
#   p_hours  - {project_id: total hours assigned}
#   p_cost   - {project_id: total wage cost assigned ($)}
#   status   - solver status string
def run_optimizer(employees, tasks, projects):
    proj_map = {p["id"]: p for p in projects}

    # Build index lookups for clean variable naming
    emp_ids  = [e["id"] for e in employees]
    task_ids = [t["id"] for t in tasks]
    emp_map  = {e["id"]: e for e in employees}
    task_map = {t["id"]: t for t in tasks}

    # Determine which (employee, task) pairs are valid based on skill match
    valid_pairs = [
        (eid, tid)
        for eid in emp_ids
        for tid in task_ids
        if task_map[tid]["type"] in emp_map[eid]["skills"]
    ]

    # Initialize the LP problem (minimization)
    prob = pulp.LpProblem("workforce_optimizer", pulp.LpMinimize)

    # --- Decision variables ---
    # x[(eid, tid)] = hours employee eid works on task tid
    x = {
        (eid, tid): pulp.LpVariable(f"x_{eid}_{tid}", lowBound=0)
        for (eid, tid) in valid_pairs
    }
    # s[tid] = unmet hours for task tid (slack variable)
    s = {
        tid: pulp.LpVariable(f"s_{tid}", lowBound=0)
        for tid in task_ids
    }

    # --- Objective: prioritize coverage, then minimize labor cost ---
    unmet_penalty = 10000
    prob += pulp.lpSum(
        unmet_penalty * (2 if proj_map.get(task_map[tid]["project"], {}).get("reimbursable") else 1) * s[tid]
        for tid in task_ids
    ) + 0.01 * pulp.lpSum(
        x[(eid, tid)] * emp_map[eid]["rate"] for (eid, tid) in valid_pairs
    )

    # --- Constraint 1: Employee capacity ---
    # Each employee cannot work more hours than their weekly capacity
    for eid in emp_ids:
        emp_tasks = [tid for (e, tid) in valid_pairs if e == eid]
        if emp_tasks:
            prob += pulp.lpSum(x[(eid, tid)] for tid in emp_tasks) <= emp_map[eid]["capacity"]

    # --- Constraint 2: Task demand (met hours + slack >= minHours) ---
    # Ensures every task is either covered by assigned employees or accounted for as unmet
    for tid in task_ids:
        assigned_vars = [x[(eid, tid)] for (eid, t) in valid_pairs if t == tid]
        if assigned_vars:
            prob += pulp.lpSum(assigned_vars) + s[tid] >= task_map[tid]["minHours"]
        else:
            # No eligible employee exists — slack must absorb the full demand
            prob += s[tid] >= task_map[tid]["minHours"]

    # --- Constraint 3: Project hour cap (reimbursable projects only) ---
    # Prevents total hours billed to a reimbursable project from exceeding maxTotal
    for p in projects:
        if p["reimbursable"] and p["maxTotal"] is not None:
            project_pairs = [(eid, tid) for (eid, tid) in valid_pairs if task_map[tid]["project"] == p["id"]]
            if project_pairs:
                prob += pulp.lpSum(x[(eid, tid)] for (eid, tid) in project_pairs) <= p["maxTotal"]

    # --- Constraint 4: Project dollar budget cap ---
    # Prevents the total wage cost (hours × rate) assigned to any project from exceeding its budget
    for p in projects:
        if p["budget"] is not None:
            project_pairs = [(eid, tid) for (eid, tid) in valid_pairs if task_map[tid]["project"] == p["id"]]
            if project_pairs:
                prob += pulp.lpSum(
                    x[(eid, tid)] * emp_map[eid]["rate"]
                    for (eid, tid) in project_pairs
                ) <= p["budget"]

    # --- Solve ---
    # Use CBC solver silently (msg=0 suppresses console output)
    prob.solve(pulp.PULP_CBC_CMD(msg=0, threads=1))

    # --- Extract results ---
    # Build assignment map: for each task, find which employee has the most hours assigned
    asgn    = {}
    load    = {eid: 0.0 for eid in emp_ids}
    p_hours = {p["id"]: 0.0 for p in projects}
    p_cost  = {p["id"]: 0.0 for p in projects}

    # Aggregate hours assigned per task across all eligible employees
    task_assigned = {tid: {} for tid in task_ids}
    for (eid, tid) in valid_pairs:
        val = pulp.value(x[(eid, tid)])
        if val and val > 0.01:  # ignore near-zero floating point noise
            task_assigned[tid][eid] = val
            load[eid]                           += val
            p_hours[task_map[tid]["project"]]   += val
            p_cost[task_map[tid]["project"]]    += val * emp_map[eid]["rate"]

    # Determine the primary employee for each task (highest hours assigned)
    for tid in task_ids:
        assigned = task_assigned[tid]
        if not assigned:
            # No employee was assigned any hours to this task
            asgn[tid] = {"employee": None, "hours": 0, "reason": "no_skill_or_budget"}
        else:
            # Primary owner = employee with the most hours on this task
            primary_eid = max(assigned, key=assigned.get)
            hours       = sum(assigned.values())
            partial     = hours < task_map[tid]["minHours"] - 0.5  # flagged if short by >0.5 h
            asgn[tid]   = {
                "employee": primary_eid,
                "hours":    round(hours, 1),
                "partial":  partial,
                "all_assigned": assigned,   # full breakdown for multi-employee tasks
            }

    return {
        "asgn":    asgn,
        "load":    {k: round(v, 1) for k, v in load.items()},
        "p_hours": {k: round(v, 1) for k, v in p_hours.items()},
        "p_cost":  {k: round(v, 2) for k, v in p_cost.items()},
        "status":  pulp.LpStatus[prob.status],
    }


# ─── DEPARTURE ANALYSIS ───────────────────────────────────────────────────────
# Simulates the impact of a specific employee leaving by:
#  1. Running the optimizer without that employee
#  2. Checking which of their formerly-owned tasks can be re-covered
#  3. Ranking remaining employees as replacement candidates
def analyze_departure(leaving_id, employees, tasks, projects):
    # Run baseline to find what the departing employee currently owns
    baseline    = run_optimizer(employees, tasks, projects)
    base_asgn   = baseline["asgn"]
    leaving_emp = next((e for e in employees if e["id"] == leaving_id), None)
    if not leaving_emp:
        return None

    # Tasks the departing employee is the primary owner of
    owned      = [t for t in tasks if base_asgn.get(t["id"], {}).get("employee") == leaving_id]
    need_skills = list(set(t["type"] for t in owned))
    total_h    = sum(t["minHours"] for t in owned)

    # Re-run optimizer excluding the departing employee
    rest       = [e for e in employees if e["id"] != leaving_id]
    new_result = run_optimizer(rest, tasks, projects)
    n_asgn     = new_result["asgn"]
    n_load     = new_result["load"]

    # Check coverage status for each previously-owned task
    plan = []
    for t in owned:
        a      = n_asgn.get(t["id"], {})
        status = (
            "ok"         if a.get("employee") and not a.get("partial")
            else "overloaded" if a.get("partial")
            else "failed"
        )
        plan.append({"task": t, "newOwner": a.get("employee"), "status": status})

    # Score remaining employees as replacement candidates
    cands = []
    for e in rest:
        overlap  = [s for s in need_skills if s in e["skills"]]
        sc       = len(overlap) / max(len(need_skills), 1)
        free     = e["capacity"] - (n_load.get(e["id"]) or 0)
        can_all  = free >= total_h
        cr       = min(free / total_h, 1) if total_h > 0 else 1
        score    = sc * 0.6 + cr * 0.3 + (0.1 if can_all else 0)
        if sc > 0:
            cands.append({"e": e, "overlap": overlap, "sc": sc,
                          "free": free, "canAll": can_all, "score": score})
    cands = sorted(cands, key=lambda x: -x["score"])[:5]

    return {
        "leaving":    leaving_emp,
        "owned":      owned,
        "needSkills": need_skills,
        "totalH":     total_h,
        "cands":      cands,
        "plan":       plan,
    }


# ─── EMERGENCY COVERAGE ANALYSIS ─────────────────────────────────────────────
# Identifies tasks disrupted when one or more employees are marked absent.
# Re-runs the optimizer without those employees and reports auto-coverage and gaps.
def analyze_emergency(absent_ids, employees, tasks, projects):
    if not absent_ids:
        return None

    proj_map   = {p["id"]: p for p in projects}

    # Baseline assignment to find which tasks the absent employees own
    baseline   = run_optimizer(employees, tasks, projects)
    base_asgn  = baseline["asgn"]

    avail      = [e for e in employees if e["id"] not in absent_ids]
    disrupted  = [t for t in tasks if base_asgn.get(t["id"], {}).get("employee") in absent_ids]

    # Re-run without absent employees
    new_result = run_optimizer(avail, tasks, projects)
    n_asgn     = new_result["asgn"]
    n_load     = new_result["load"]

    # Build substitution map: auto-covered vs. needs manual assignment
    sub_map = {}
    for t in disrupted:
        a = n_asgn.get(t["id"], {})
        if a.get("employee") and not a.get("partial"):
            sub_map[t["id"]] = {"autoSub": a["employee"]}
        else:
            # Rank available employees as manual sub candidates
            cands = []
            for e in avail:
                if t["type"] in e["skills"]:
                    free       = e["capacity"] - (n_load.get(e["id"]) or 0)
                    util_after = (
                        ((n_load.get(e["id"]) or 0) + t["minHours"]) / e["capacity"] * 100
                        if e["capacity"] > 0 else 999
                    )
                    cands.append({"e": e, "free": free,
                                  "canTake": free >= t["minHours"],
                                  "utilAfter": util_after})
            cands = sorted(cands, key=lambda x: (-(1 if x["canTake"] else 0), -x["free"]))[:3]
            sub_map[t["id"]] = {"autoSub": None, "cands": cands}

    auto_cov = sum(1 for s in sub_map.values() if s.get("autoSub"))

    # Summarize skill gaps created by the absences
    gaps = {}
    for e in employees:
        if e["id"] in absent_ids:
            for sk in e["skills"]:
                if sk not in gaps:
                    gaps[sk] = {"lost": 0, "lostH": 0, "remain": 0, "demH": 0}
                gaps[sk]["lost"]  += 1
                gaps[sk]["lostH"] += e["capacity"]
    for sk in gaps:
        gaps[sk]["remain"] = len([e for e in avail if sk in e["skills"]])
        gaps[sk]["demH"]   = sum(t["minHours"] for t in disrupted if t["type"] == sk)

    # Group disrupted tasks by project for priority display
    urg_proj = {}
    for t in disrupted:
        pid = t["project"]
        if pid not in urg_proj:
            urg_proj[pid] = {"proj": proj_map[pid], "tasks": [], "total": 0, "risk": 0, "cov": 0}
        urg_proj[pid]["tasks"].append(t)
        urg_proj[pid]["total"] += t["minHours"]
        if sub_map.get(t["id"], {}).get("autoSub"):
            urg_proj[pid]["cov"] += 1
        else:
            urg_proj[pid]["risk"] += 1

    return {
        "absentEmps": [e for e in employees if e["id"] in absent_ids],
        "disrupted":  disrupted,
        "urgProj":    urg_proj,
        "subMap":     sub_map,
        "gaps":       gaps,
        "autoCov":    auto_cov,
        "total":      len(disrupted),
    }


# ─── SKILL GAP ANALYSIS ───────────────────────────────────────────────────────
# Compares the total hours demanded across all tasks of each skill type
# against the available capacity of employees who hold that skill.
# Flags skill types as surplus / healthy / tight / critical.
def analyze_skill_gap(employees, tasks, projects):
    skill_types  = sorted(set(t["type"] for t in tasks))
    total_demand = {s: 0 for s in skill_types}
    total_supply = {s: 0 for s in skill_types}
    emp_count    = {s: 0 for s in skill_types}

    for t in tasks:
        total_demand[t["type"]] += t["minHours"]
    for e in employees:
        for s in e["skills"]:
            if s in total_supply:
                total_supply[s] += e["capacity"]
                emp_count[s]    += 1

    opt = run_optimizer(employees, tasks, projects)

    results = []
    for skill in skill_types:
        dem      = total_demand[skill]
        sup      = total_supply.get(skill, 0)
        gap      = dem - sup
        ratio    = dem / sup if sup > 0 else 99
        coverage = min(sup / dem, 1) * 100 if dem > 0 else 100
        status   = (
            "surplus"  if ratio <= 0.6
            else "healthy"  if ratio <= 1.0
            else "tight"    if ratio <= 1.3
            else "critical"
        )
        hours_short  = max(0, gap)
        hires_needed = max(0, -(-hours_short // 30)) if hours_short > 0 else 0

        # Find employees who could be trained in this skill (have bandwidth and <3 skills)
        can_learn = []
        for e in employees:
            if skill not in e["skills"] and len(e["skills"]) < 3:
                free = e["capacity"] - (opt["load"].get(e["id"]) or 0)
                if free >= 5:
                    can_learn.append({"e": e, "free": free})
        can_learn = sorted(can_learn, key=lambda x: -x["free"])[:3]

        results.append({
            "skill":       skill,
            "dem":         dem,
            "sup":         sup,
            "gap":         gap,
            "ratio":       ratio,
            "coverage":    coverage,
            "status":      status,
            "hoursShort":  hours_short,
            "hiresNeeded": hires_needed,
            "canLearn":    can_learn,
            "empCount":    emp_count.get(skill, 0),
        })

    overall_ok = all(r["status"] in ("healthy", "surplus") for r in results)
    return {"results": results, "overallOk": overall_ok}


# ─── STYLE HELPER ─────────────────────────────────────────────────────────────
# Maps status strings to emoji-colored labels for consistent UI display
def status_badge(status):
    colors = {
        "OK":          "🟢", "Overloaded": "🟡", "Failed":     "🔴",
        "Fulfilled":   "🟢", "Partial":    "🟡", "Over Budget":"🔴",
        "critical":    "🔴", "tight":      "🟡", "healthy":    "🟢",
        "surplus":     "⚪", "Reassigned": "🟢",
    }
    return f"{colors.get(status, '')} {status}"


# ─── HEADER ───────────────────────────────────────────────────────────────────
st.markdown("## ⚡ Workforce Optimizer")
st.caption("Upload your budget data template to run the optimizer.")


# ─── FILE UPLOAD (SIDEBAR) ────────────────────────────────────────────────────
# The sidebar houses the upload widget so it stays accessible across all tabs.
# Users upload an Excel file matching the provided template structure.
with st.sidebar:
    st.header("📂 Data Import")
    st.markdown(
        "Upload an Excel file with three sheets: **Employees**, **Projects**, and **Tasks**."
    )

    # ── Downloadable blank template ──
    # Build the template fresh each session so the download is always current.
    # Wrapped in a try so a packaging issue never blocks the upload flow.
    try:
        template_bytes = build_blank_template()
        st.download_button(
            label="⬇ Download blank template",
            data=template_bytes,
            file_name="workforce_optimizer_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception:
        st.caption("Template download unavailable — openpyxl may not be installed.")

    st.divider()

    uploaded_file = st.file_uploader(
        label="Choose Excel file (.xlsx)",
        type=["xlsx"],
        help="File must contain sheets: Employees, Projects, Tasks",
    )



# ─── GUARD: require upload before anything else runs ─────────────────────────
# If no file is uploaded, show instructions and stop rendering the rest of the app.
if uploaded_file is None:
    st.info("👈 Upload your Excel template in the sidebar to get started.")
    st.stop()


# ─── PARSE AND VALIDATE UPLOAD ────────────────────────────────────────────────
# Parse the uploaded bytes into a pandas ExcelFile, validate the structure,
# then load into the data structures the optimizer expects.
try:
    xl = pd.ExcelFile(uploaded_file)
except Exception as e:
    st.error(f"Could not read the file: {e}")
    st.stop()

valid, err_msg = validate_template(xl)
if not valid:
    st.error(f"**Invalid template:**\n\n{err_msg}")
    st.stop()

# Load data — wrapped in try/except to surface any parsing errors clearly
try:
    EMPLOYEES, PROJECTS, TASKS = load_data(xl)
except Exception as e:
    st.error(f"Error reading data from file: {e}")
    st.stop()

# Confirm successful load to the user
st.sidebar.success(
    f"Loaded: {len(EMPLOYEES)} employees · {len(PROJECTS)} projects · {len(TASKS)} tasks"
)


# ─── RUN OPTIMIZER ────────────────────────────────────────────────────────────
# Run the LP optimizer once on page load and cache results in session_state
# so switching tabs doesn't re-solve unnecessarily.
# The cache key is the file name + size so re-uploading a new file re-solves.
file_bytes = uploaded_file.getvalue()
cache_key = hashlib.md5(file_bytes).hexdigest()
if st.session_state.get("_cache_key") != cache_key:
    with st.spinner("Running LP optimizer..."):
        opt_result = run_optimizer(EMPLOYEES, TASKS, PROJECTS)
    st.session_state["opt_result"]  = opt_result
    st.session_state["_cache_key"] = cache_key

opt_result = st.session_state["opt_result"]
asgn       = opt_result["asgn"]
load       = opt_result["load"]
p_hours    = opt_result["p_hours"]
p_cost     = opt_result["p_cost"]

# Convenience lookups used across multiple tabs
proj_map     = {p["id"]: p for p in PROJECTS}
task_map     = {t["id"]: t for t in TASKS}

# Summary counts for the top-level metrics
total_cap  = sum(e["capacity"] for e in EMPLOYEES)
total_load = sum(load.values())
util       = round(total_load / total_cap * 100) if total_cap > 0 else 0
n_ok       = sum(1 for a in asgn.values() if a.get("employee") and not a.get("partial"))
n_fail     = sum(1 for a in asgn.values() if not a.get("employee"))
n_part     = sum(1 for a in asgn.values() if a.get("partial"))

# Build per-employee task list for the Employees tab
tasks_by_emp = {e["id"]: [] for e in EMPLOYEES}
for t in TASKS:
    a = asgn.get(t["id"], {})
    if a.get("employee"):
        tasks_by_emp[a["employee"]].append({**t, **a})

# Run skill gap analysis (uses optimizer internally)
skill_gap_data = analyze_skill_gap(EMPLOYEES, TASKS, PROJECTS)
critical_gaps  = [r for r in skill_gap_data["results"] if r["status"] in ("critical", "tight")]


# ─── TOP METRICS ──────────────────────────────────────────────────────────────
# Three headline KPIs shown above the tabs for a quick status read
col1, col2, col3 = st.columns(3)
col1.metric("Utilization",   f"{util}%")
col2.metric("Tasks Covered", f"{n_ok} / {len(TASKS)}")
col3.metric("Skill Alerts",  len(critical_gaps))

# Surface the LP solver status in case the model is infeasible or unbounded
solver_status = opt_result.get("status", "Unknown")
if solver_status not in ("Optimal", "Not Solved"):
    st.warning(f"LP solver status: **{solver_status}** — results may be incomplete.")

st.divider()


# ─── TABS ─────────────────────────────────────────────────────────────────────
tab_labels = [
    "📊 Dashboard", "👥 Employees", "📁 Projects",
    "📋 Assignments", "🎯 Skill Gap Advisor",
    "👤 Departure Planner", "🚨 Emergency Coverage",
]
tabs = st.tabs(tab_labels)


# ══════════════════════════════════════════════════════════════════ DASHBOARD
with tabs[0]:
    # Summary KPI row
    c1, c2, c3, c4 = st.columns(4)
    reimb_count = sum(1 for p in PROJECTS if p["reimbursable"])
    c1.metric("Total Employees",     len(EMPLOYEES), f"{total_cap} total h/week")
    c2.metric("Active Projects",     len(PROJECTS),
              f"{reimb_count} reimbursable · {len(PROJECTS)-reimb_count} standard")
    c3.metric("Tasks Assigned",      f"{n_ok} / {len(TASKS)}",
              f"{n_fail + n_part} need attention" if n_fail + n_part > 0 else "All tasks covered")
    c4.metric("Workforce Utilization", f"{util}%",
              f"{round(total_load, 1)} of {total_cap} h used")

    # Employee utilization table with progress bar
    st.subheader("Employee Utilization")
    emp_rows = []
    for e in EMPLOYEES:
        l = load.get(e["id"], 0)
        p = round(l / e["capacity"] * 100) if e["capacity"] > 0 else 0
        emp_rows.append({
            "Employee":      e["id"],
            "Type":          e.get("type", ""),
            "Rate ($/h)":    f"${e['rate']:.2f}",
            "Skills":        " ".join(SKILL_COLORS.get(s, s) for s in e["skills"]),
            "Assigned (h)":  round(l, 1),
            "Capacity (h)":  e["capacity"],
            "Utilization %": p,
        })
    df_emp = pd.DataFrame(emp_rows)
    st.dataframe(
        df_emp, use_container_width=True, hide_index=True,
        column_config={"Utilization %": st.column_config.ProgressColumn(
            "Utilization %", min_value=0, max_value=100, format="%d%%"
        )},
    )

    # Skill demand vs supply table
    st.subheader("Skill Demand vs Supply")
    skill_types = sorted(set(t["type"] for t in TASKS))
    skill_rows = []
    for sk in skill_types:
        dem    = sum(t["minHours"] for t in TASKS if t["type"] == sk)
        sup    = sum(e["capacity"] for e in EMPLOYEES if sk in e["skills"])
        ratio  = dem / sup if sup > 0 else 99
        status = ("surplus" if ratio <= 0.6 else "healthy" if ratio <= 1.0
                  else "tight" if ratio <= 1.3 else "critical")
        n_emp  = len([e for e in EMPLOYEES if sk in e["skills"]])
        n_task = len([t for t in TASKS if t["type"] == sk])
        skill_rows.append({
            "Skill":       f"{SKILL_COLORS.get(sk, sk)} {sk}",
            "Tasks":       n_task, "Employees": n_emp,
            "Demand (h)":  dem, "Supply (h)": sup,
            "Coverage %":  round(min(sup / dem, 1) * 100) if dem > 0 else 100,
            "Status":      status.upper(),
        })
    df_skills = pd.DataFrame(skill_rows)
    st.dataframe(
        df_skills, use_container_width=True, hide_index=True,
        column_config={"Coverage %": st.column_config.ProgressColumn(
            "Coverage %", min_value=0, max_value=100, format="%d%%"
        )},
    )

    if critical_gaps:
        st.warning(f"⚠ {len(critical_gaps)} skill(s) under pressure — see **Skill Gap Advisor** for recommendations.")

    # Project summary table including dollar budget tracking
    st.subheader("Project Summary")
    proj_rows = []
    for p in PROJECTS:
        dem     = sum(t["minHours"] for t in TASKS if t["project"] == p["id"])
        asn     = p_hours.get(p["id"], 0)
        cost    = p_cost.get(p["id"], 0)
        cov_pct = round(asn / dem * 100) if dem > 0 else 100
        ov_hr   = p["maxTotal"] is not None and asn > p["maxTotal"]
        ov_bud  = p["budget"]   is not None and cost > p["budget"]
        ok      = asn >= dem
        status  = "Over Budget" if (ov_hr or ov_bud) else ("Fulfilled" if ok else "Partial")
        proj_rows.append({
            "Project":       p["id"],
            "Type":          "Reimbursable" if p["reimbursable"] else "Non-Reimb.",
            "Demand (h)":    dem,
            "Assigned (h)":  round(asn, 1),
            "Hour Cap":      f"{p['maxTotal']} h" if p["maxTotal"] is not None else "—",
            "Wage Cost ($)": f"${cost:,.0f}",
            "Budget ($)":    f"${p['budget']:,.0f}" if p["budget"] is not None else "—",
            "Coverage %":    cov_pct,
            "Status":        status_badge(status),
        })
    st.dataframe(
        pd.DataFrame(proj_rows), use_container_width=True, hide_index=True,
        column_config={"Coverage %": st.column_config.ProgressColumn(
            "Coverage %", min_value=0, max_value=100, format="%d%%"
        )},
    )


# ══════════════════════════════════════════════════════════════════ EMPLOYEES
with tabs[1]:
    st.subheader("Employee Roster")
    for e in EMPLOYEES:
        l          = load.get(e["id"], 0)
        p          = round(l / e["capacity"] * 100) if e["capacity"] > 0 else 0
        skill_icons = " ".join(SKILL_COLORS.get(s, s) + s for s in e["skills"])
        label      = (f"**{e['id']}** [{e.get('type','')}] — {skill_icons} — "
                      f"${e['rate']:.2f}/h — {p}% utilized ({round(l,1)}/{e['capacity']} h)")
        with st.expander(label):
            tasks = tasks_by_emp.get(e["id"], [])
            if not tasks:
                st.write("No tasks assigned.")
            else:
                rows = [{
                    "Task":    t["id"],
                    "Project": t["project"],
                    "Skill":   f"{SKILL_COLORS.get(t['type'], t['type'])} {t['type']}",
                    "Min Hrs": t["minHours"],
                    "Assigned Hrs": round(t.get("hours", t["minHours"]), 1),
                    "Status":  "Partial" if t.get("partial") else "OK",
                } for t in tasks]
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════ PROJECTS
with tabs[2]:
    st.subheader("Project Overview")
    for p in PROJECTS:
        pt    = [t for t in TASKS if t["project"] == p["id"]]
        dem   = sum(t["minHours"] for t in pt)
        asn   = p_hours.get(p["id"], 0)
        cost  = p_cost.get(p["id"], 0)
        ov_hr = p["maxTotal"] is not None and asn > p["maxTotal"]
        ov_bd = p["budget"]   is not None and cost > p["budget"]
        ok    = asn >= dem
        label = (
            "🟢 Fulfilled"   if ok and not ov_hr and not ov_bd
            else "🔴 Over Budget" if (ov_hr or ov_bd)
            else "🟡 Partial"
        )
        ptype = "Reimbursable" if p["reimbursable"] else "Non-Reimbursable"
        header = (f"**{p['id']}** — {ptype} — {label} "
                  f"({round(asn,1)}/{dem} h · ${cost:,.0f} cost)")
        with st.expander(header):
            # Show hour budget progress bar for reimbursable projects
            if p["maxTotal"] is not None:
                st.progress(
                    min(asn / p["maxTotal"], 1.0),
                    text=f"Hour budget: {round(asn,1)} / {p['maxTotal']} h",
                )
            # Show dollar budget progress bar if a budget is set
            if p["budget"] is not None:
                st.progress(
                    min(cost / p["budget"], 1.0),
                    text=f"Wage budget: ${cost:,.0f} / ${p['budget']:,.0f}",
                )
            rows = []
            for t in pt:
                a = asgn.get(t["id"], {})
                rows.append({
                    "Task":        t["id"],
                    "Skill":       f"{SKILL_COLORS.get(t['type'], t['type'])} {t['type']}",
                    "Min Hrs":     t["minHours"],
                    "Assigned Hrs":round(a.get("hours", 0), 1),
                    "Assigned To": a.get("employee") or "—",
                    "Status": (
                        "🟡 Partial" if a.get("partial")
                        else "🟢 OK"   if a.get("employee")
                        else "🔴 Unassigned"
                    ),
                })
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════ ASSIGNMENTS
with tabs[3]:
    st.subheader("Full Task Assignment Matrix")
    rows = []
    for t in TASKS:
        a     = asgn.get(t["id"], {})
        s_str = "Failed" if not a.get("employee") else ("Overloaded" if a.get("partial") else "OK")
        reimb = "R" if proj_map.get(t["project"], {}).get("reimbursable") else "N"
        rows.append({
            "Task":        t["id"],
            "Project":     f"{t['project']} ({reimb})",
            "Skill":       f"{SKILL_COLORS.get(t['type'], t['type'])} {t['type']}",
            "Min Hours":   t["minHours"],
            "Assigned Hrs":round(a.get("hours", 0), 1),
            "Assigned To": a.get("employee") or "—",
            "Status":      status_badge(s_str),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════ SKILL GAP
with tabs[4]:
    st.subheader("Skill Gap Advisor")
    data = skill_gap_data

    if data["overallOk"]:
        st.success("Your current team's skill coverage is sufficient for all active projects.")
    else:
        st.error("Current skill supply cannot meet project demand. Immediate investment recommended.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Skills Healthy",   f"{sum(1 for r in data['results'] if r['status'] in ('healthy','surplus'))}/{len(data['results'])}")
    c2.metric("Tight Supply",     sum(1 for r in data['results'] if r['status'] == 'tight'))
    c3.metric("Critical Gaps",    sum(1 for r in data['results'] if r['status'] == 'critical'))
    c4.metric("Est. Hires Needed",sum(r["hiresNeeded"] for r in data["results"]))

    order = {"critical": 0, "tight": 1, "healthy": 2, "surplus": 3}
    for r in sorted(data["results"], key=lambda x: order[x["status"]]):
        icon = ("🔴" if r["status"] == "critical" else "🟡" if r["status"] == "tight"
                else "🟢" if r["status"] == "healthy" else "⚪")
        with st.expander(
            f"{SKILL_COLORS.get(r['skill'], r['skill'])} **Skill {r['skill']}** — "
            f"{icon} {r['status'].upper()} — {r['coverage']:.0f}% covered"
        ):
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Demand",    f"{r['dem']} h")
            c2.metric("Supply",    f"{r['sup']} h")
            c3.metric("Gap",       f"{'+' if r['gap'] > 0 else '-'}{abs(r['gap'])} h")
            c4.metric("Employees", r["empCount"])
            c5.metric("Coverage",  f"{r['coverage']:.0f}%")
            st.progress(r["coverage"] / 100, text=f"Supply vs Demand: {r['coverage']:.0f}%")

            if r["status"] in ("critical", "tight"):
                col_a, col_b = st.columns(2)
                with col_a:
                    st.markdown("**Option A — Upskill Existing Staff**")
                    if not r["canLearn"]:
                        st.caption("No employees have bandwidth to learn this skill right now.")
                    else:
                        for c in r["canLearn"]:
                            skills_str = " ".join(SKILL_COLORS.get(s, s) + s for s in c["e"]["skills"])
                            st.caption(f"**{c['e']['id']}** ({skills_str}) — {c['free']} h free")
                        total_free = sum(c["free"] for c in r["canLearn"])
                        st.info(f"Training would add ~{total_free} h of Skill {r['skill']} capacity without new hires.")
                with col_b:
                    st.markdown("**Option B — Hire New Employee**")
                    st.metric("Hours Short",       f"{r['hoursShort']} h")
                    st.metric("Est. Hires Needed", r["hiresNeeded"], "@ avg 30 h/week")
            elif r["status"] == "surplus":
                st.caption(f"Surplus supply — consider reassigning some Skill {r['skill']} capacity to support tighter areas.")


# ══════════════════════════════════════════════════════════════════ DEPARTURE
with tabs[5]:
    st.subheader("Departure Impact Planner")
    st.info(
        "Select an employee who is leaving. The engine analyzes their task portfolio, "
        "identifies coverage gaps, and ranks the best internal replacements."
    )

    emp_options = ["— Select departing employee —"] + [
        f"{e['id']} · {e.get('type','')} · Skills: {', '.join(e['skills'])} · {e['capacity']} h/week"
        for e in EMPLOYEES
    ]
    selection = st.selectbox("Departing Employee", emp_options)

    if selection != emp_options[0]:
        dept_id = selection.split(" · ")[0]
        if st.button("Analyze Departure"):
            with st.spinner("Analyzing departure impact..."):
                st.session_state["dept_result"] = analyze_departure(
                    dept_id, EMPLOYEES, TASKS, PROJECTS
                )

    if "dept_result" in st.session_state and st.session_state["dept_result"]:
        dr      = st.session_state["dept_result"]
        leaving = dr["leaving"]
        auto_ok = sum(1 for p in dr["plan"] if p["status"] == "ok")
        at_risk = sum(1 for p in dr["plan"] if p["status"] != "ok")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Tasks Owned",       len(dr["owned"]))
        c2.metric("Hours to Cover",    f"{dr['totalH']} h")
        c3.metric("Auto-Redistributed",auto_ok)
        c4.metric("Needs Manual Cover",at_risk)

        skills_str = " ".join(SKILL_COLORS.get(s, s) + s for s in leaving["skills"])
        need_str   = " ".join(SKILL_COLORS.get(s, s) + s for s in dr["needSkills"])
        st.markdown(f"**Departing:** {leaving['id']} ({leaving.get('type','')}) — "
                    f"{skills_str} — ${leaving['rate']:.2f}/h — {leaving['capacity']} h/week capacity")
        st.markdown(f"**Skills needed for coverage:** {need_str}")

        if at_risk > 0:
            st.error(f"{at_risk} task(s) cannot be auto-covered without a replacement.")

        st.subheader("Top Replacement Candidates")
        for i, c in enumerate(dr["cands"]):
            label   = "Best Match" if i == 0 else f"#{i+1}"
            c_skills = " ".join(SKILL_COLORS.get(s, s) + s for s in c["e"]["skills"])
            with st.expander(f"{label} — **{c['e']['id']}** — {c_skills} — Score: {round(c['score']*100)}"):
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Skill Coverage", f"{round(c['sc']*100)}%")
                col2.metric("Free Hours",     f"{c['free']} h")
                col3.metric("Can Absorb All", "Yes" if c["canAll"] else "No")
                col4.metric("Fit Score",       round(c["score"] * 100))
                missing = [s for s in dr["needSkills"] if s not in c["e"]["skills"]]
                if missing:
                    st.warning(f"Missing skills: {', '.join(missing)} — additional resource needed.")

        st.subheader("Task Redistribution Plan")
        rows = [{
            "Task":     p["task"]["id"],
            "Project":  p["task"]["project"],
            "Skill":    f"{SKILL_COLORS.get(p['task']['type'], p['task']['type'])} {p['task']['type']}",
            "Hours":    p["task"]["minHours"],
            "New Owner":p["newOwner"] or "—",
            "Status":   status_badge("Reassigned" if p["status"] == "ok"
                                     else p["status"].capitalize()),
        } for p in dr["plan"]]
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        if st.button("Clear Analysis"):
            del st.session_state["dept_result"]
            st.rerun()


# ══════════════════════════════════════════════════════════════════ EMERGENCY
with tabs[6]:
    st.subheader("Emergency Coverage Planner")
    st.error(
        "Mark employees as absent (sick, unavailable). The engine instantly identifies "
        "disrupted tasks and recommends the best available subs — reimbursable projects "
        "are treated as highest priority."
    )

    absent_ids = st.multiselect(
        "Mark Absent Employees",
        options=[e["id"] for e in EMPLOYEES],
        format_func=lambda eid: (
            f"{eid} [{next(e for e in EMPLOYEES if e['id']==eid).get('type','')}] "
            f"{next(e for e in EMPLOYEES if e['id']==eid)['skills']}"
        ),
    )

    if absent_ids:
        if st.button("Run Coverage Analysis"):
            with st.spinner("Analyzing emergency coverage..."):
                em_result = analyze_emergency(absent_ids, EMPLOYEES, TASKS, PROJECTS)
            st.session_state["em_result"] = em_result

    if "em_result" in st.session_state and st.session_state.get("em_result"):
        em_result = st.session_state["em_result"]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Employees Absent", len(em_result["absentEmps"]))
        c2.metric("Tasks Disrupted",  em_result["total"])
        c3.metric("Auto-Covered",     em_result["autoCov"])
        at_risk_n = em_result["total"] - em_result["autoCov"]
        c4.metric("Needs Manual Sub", at_risk_n)

        if at_risk_n == 0:
            st.success(f"All {em_result['total']} disrupted tasks can be auto-redistributed.")
        else:
            st.error(f"{at_risk_n} task(s) require manual reassignment. See recommendations below.")

        # Skill gap impact from the absences
        if em_result["gaps"]:
            st.subheader("Skill Gap Impact")
            gap_rows = [{
                "Skill":                f"{SKILL_COLORS.get(sk, sk)} {sk}",
                "Lost Employees":       g["lost"],
                "Lost Hours":           g["lostH"],
                "Remaining Employees":  g["remain"],
                "At-Risk Demand (h)":   g["demH"],
            } for sk, g in em_result["gaps"].items()]
            st.dataframe(pd.DataFrame(gap_rows), use_container_width=True, hide_index=True)

        # Affected projects sorted by priority (reimbursable first, then most at-risk)
        if em_result["urgProj"]:
            st.subheader("Affected Projects — Priority Order")
            proj_rows = sorted(
                [{
                    "Project":     pid,
                    "Type":        "🔵 Reimbursable · High Priority"
                                   if info["proj"].get("reimbursable") else "🟣 Non-Reimbursable",
                    "Tasks":       len(info["tasks"]),
                    "Total (h)":   info["total"],
                    "Auto-Covered":info["cov"],
                    "At Risk":     info["risk"],
                } for pid, info in em_result["urgProj"].items()],
                key=lambda x: (0 if "Reimbursable" in x["Type"] else 1, -x["At Risk"]),
            )
            st.dataframe(pd.DataFrame(proj_rows), use_container_width=True, hide_index=True)

        # Per-task substitution recommendations
        st.subheader("Immediate Sub Recommendations")
        disrupted_sorted = sorted(
            em_result["disrupted"],
            key=lambda t: (
                -(1 if proj_map.get(t["project"], {}).get("reimbursable") else 0),
                -t["minHours"],
            ),
        )
        for t in disrupted_sorted:
            sub      = em_result["subMap"].get(t["id"], {})
            p        = proj_map.get(t["project"], {})
            priority = "🔵 High Priority" if p.get("reimbursable") else "🟣 Standard"
            auto     = sub.get("autoSub")
            header   = (
                f"**{t['id']}** — {SKILL_COLORS.get(t['type'], t['type'])} Skill {t['type']} "
                f"— {t['project']} — {t['minHours']} h — {priority}"
            )
            header += f" — 🟢 Auto-assigned → {auto}" if auto else " — 🔴 Needs Manual Assignment"
            with st.expander(header):
                if not auto and sub.get("cands"):
                    st.markdown("**Best Available Subs:**")
                    for i, cand in enumerate(sub["cands"]):
                        label     = "Top Pick" if i == 0 else f"#{i+1}"
                        ok_label  = "No overload" if cand["canTake"] else "Slight overload"
                        st.caption(
                            f"{label} — **{cand['e']['id']}** — {cand['free']} h free "
                            f"— {round(cand['utilAfter'])}% util after — {ok_label}"
                        )
                elif not auto:
                    st.error(
                        f"No available employees with Skill {t['type']} — "
                        "consider external hire or deferral."
                    )
    elif not absent_ids:
        st.caption("Select one or more absent employees above, then click Run Coverage Analysis.")


# ─── BOTTOM OF PAGE: DATA REQUIREMENTS + METHODOLOGY ─────────────────────────
# Shown below all tabs so it doesn't clutter the sidebar or the working area.
# Both sections render as collapsed expanders by default.
st.divider()

# ── Data Requirements ──────────────────────────────────────────────────────────
with st.expander("📋 Data requirements — click to expand"):
    st.markdown("##### Employees sheet")
    st.code("employee | capacity | skills | hourly_rate ($) | employee_type", language="text")
    st.markdown("""
| Column | Type | Notes |
|--------|------|-------|
| `employee` | Text | Unique ID, e.g. `E001` |
| `capacity` | Integer | Max hours available per week |
| `skills` | Text | Comma-separated skill types: `A`, `B`, `C`, `D`, or `E` — e.g. `A,B,C` |
| `hourly_rate ($)` | Number | Billing/wage rate used in the budget constraint |
| `employee_type` | Text | `Junior` / `Mid-Level` / `Senior` (display only) |
""")
    st.markdown("##### Projects sheet")
    st.code("project | reimbursable | max_total | budget ($)", language="text")
    st.markdown("""
| Column | Type | Notes |
|--------|------|-------|
| `project` | Text | Unique ID, e.g. `P001` |
| `reimbursable` | Boolean | `TRUE` = client-billable (gets higher optimization priority); `FALSE` = internal |
| `max_total` | Integer | Hard cap on total hours billed. Leave blank for non-reimbursable projects. |
| `budget ($)` | Number | Dollar ceiling. The optimizer will not assign more wage cost than this. |
""")
    st.markdown("##### Tasks sheet")
    st.code("project | task_id | task_type | min_hours", language="text")
    st.markdown("""
| Column | Type | Notes |
|--------|------|-------|
| `project` | Text | Must exactly match a project ID in the Projects sheet |
| `task_id` | Text | Unique task ID, e.g. `T001` |
| `task_type` | Text | The skill required to perform this task (`A` / `B` / `C` / `D` / `E`) |
| `min_hours` | Integer | Minimum hours the optimizer must try to cover |
""")


# ── Model Methodology ──────────────────────────────────────────────────────────
with st.expander("📐 Model methodology — click to expand"):
    st.markdown("#### What this optimizer is doing, in plain terms")
    st.markdown("""
The optimizer is solving a **Linear Programming (LP)** problem — a mathematical technique that
finds the best possible answer to a resource allocation problem given a set of hard rules (constraints).

**In plain terms:** it's trying to assign every task to the right employee in a way that covers as
many hours as possible, without exceeding any employee's weekly availability, any project's hour cap,
or any project's dollar budget. Reimbursable (client-billable) tasks are treated as higher priority
than internal ones.

Think of it like a scheduling puzzle: the optimizer is testing millions of possible combinations
simultaneously and picking the one that leaves the fewest uncovered hours, while respecting every
constraint at once.
""")

    st.markdown("#### The four rules (constraints)")
    st.markdown("""
**1. Employee capacity** — No employee can be assigned more hours than their stated weekly capacity.
If someone has 35 hours available, the model will never try to give them 36.

**2. Task demand** — Every task has a minimum number of hours it needs. The model tries to meet that
minimum. If it can't (due to budget or skill gaps), it records the shortfall rather than pretending
the task is covered.

**3. Skill matching** — An employee can only be assigned to a task if they hold the required skill
type. A Skill B task cannot be assigned to someone who only has Skills C and D.

**4. Project budget ceiling** — For each project, the total wage cost of all hours assigned
(hours × each employee's hourly rate) cannot exceed the project's dollar budget. This is the key
financial constraint that links headcount decisions to real dollars.

For reimbursable projects there is also a **project hour cap** (`max_total`) — a separate ceiling
on the raw number of billable hours regardless of cost.
""")

    st.markdown("#### The objective — what the model is maximizing")
    st.markdown("""
The optimizer minimizes **uncovered (slack) hours** across all tasks, with reimbursable tasks
weighted twice as heavily as internal tasks. This means when capacity is tight and trade-offs must
be made, client-facing work is always covered first.
""")

    st.markdown("#### Equations")
    st.latex(r"""
\text{Minimize} \quad \sum_{j} w_j \cdot s_j
""")
    st.markdown("""where:
- $j$ = a task
- $s_j$ = uncovered (slack) hours for task $j$ — how many hours the model couldn't fill
- $w_j = 2$ if the task belongs to a reimbursable project, $w_j = 1$ otherwise
""")

    st.latex(r"""
\text{Subject to:}
""")
    st.markdown("**Employee capacity constraint** — for each employee $i$:")
    st.latex(r"""
\sum_{j} x_{ij} \leq \text{capacity}_i
""")
    st.markdown("**Task demand constraint** — for each task $j$:")
    st.latex(r"""
\sum_{i} x_{ij} + s_j \geq \text{minHours}_j
""")
    st.markdown("**Project hour cap** — for each reimbursable project $p$:")
    st.latex(r"""
\sum_{j \in p} \sum_{i} x_{ij} \leq \text{maxTotal}_p
""")
    st.markdown("**Project dollar budget** — for each project $p$:")
    st.latex(r"""
\sum_{j \in p} \sum_{i} x_{ij} \cdot \text{rate}_i \leq \text{budget}_p
""")
    st.markdown("**Skill matching and non-negativity:**")
    st.latex(r"""
x_{ij} = 0 \quad \text{if employee } i \text{ lacks the skill for task } j, \qquad x_{ij} \geq 0, \quad s_j \geq 0
""")
    st.markdown("""
where $x_{ij}$ is the number of hours employee $i$ is assigned to task $j$.
The solver finds the values of every $x_{ij}$ and $s_j$ that minimize total uncovered hours
while satisfying all constraints simultaneously.
""")
