import streamlit as st
import pandas as pd
import time
import os
from datetime import datetime, timedelta
from database import ExcelDB

# 1. Page Configuration
st.set_page_config(page_title="WorkManager – Daily Productivity Hub", page_icon="💼", layout="wide")

# 2. Modern and Fancy CSS UI Styling Injection (using st.markdown)
custom_styles = '''
<style>
    /* Premium Font and smooth animations */
    html, body, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }
    
    /* Hover effects for custom metric cards and task cards */
    .metric-card-adaptive {
        transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease;
    }
    .metric-card-adaptive:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0,0,0,0.04), 0 4px 6px -4px rgba(0,0,0,0.04);
        border-color: rgba(37, 99, 235, 0.3) !important;
    }
    
    /* Clean, modern Streamlit Buttons */
    .stButton > button {
        border-radius: 8px !important;
        font-weight: 600 !important;
        transition: all 0.2s ease !important;
        border: 1px solid rgba(0,0,0,0.08) !important;
    }
    .stButton > button:hover {
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05), 0 2px 4px -1px rgba(0,0,0,0.05) !important;
        transform: scale(1.01) !important;
    }
    .stButton > button:active {
        transform: scale(0.99) !important;
    }
    
    /* Modern Form Inputs and Selectboxes */
    div[data-baseweb="select"], div[data-baseweb="input"] {
        border-radius: 8px !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.02) !important;
    }
    
    /* Custom style for headings */
    h2, h3, h4 {
        font-weight: 800 !important;
        letter-spacing: -0.025em !important;
    }
    
    /* Kanban columns container styling */
    div[data-testid="stHorizontalBlock"] > div[data-testid="stVerticalBlock"] {
        background-color: rgba(128,128,128,0.02);
        padding: 5px;
        border-radius: 12px;
        border: 1px solid rgba(128,128,128,0.08);
    }
</style>
'''
st.markdown(custom_styles, unsafe_allow_html=True)

# 3. Configurations & Metadata
CATEGORIES = {
    "work": {"label": "Work", "icon": "💼", "color": "#1d4ed8", "bg": "#eff6ff"},
    "meetings": {"label": "Meetings", "icon": "🗓", "color": "#b45309", "bg": "#fffbeb"},
    "learning": {"label": "Learning", "icon": "📚", "color": "#0e7490", "bg": "#ecfeff"},
    "health": {"label": "Health", "icon": "💪", "color": "#15803d", "bg": "#f0fdf4"},
    "finance": {"label": "Finance", "icon": "💰", "color": "#047857", "bg": "#ecfdf5"},
    "project": {"label": "Project", "icon": "🚀", "color": "#be123c", "bg": "#fff1f2"},
    "personal": {"label": "Personal", "icon": "👤", "color": "#7e34cd", "bg": "#f5f3ff"}
}

PRIORITIES = {
    "critical": {"label": "Critical", "color": "#ef4444", "bg": "#fef2f2", "border": "#fecaca", "dot": "🔴"},
    "high": {"label": "High", "color": "#f97316", "bg": "#fff7ed", "border": "#fed7aa", "dot": "🟠"},
    "medium": {"label": "Medium", "color": "#eab308", "bg": "#fefce8", "border": "#fef08a", "dot": "🟡"},
    "low": {"label": "Low", "color": "#22c55e", "bg": "#f0fdf4", "border": "#bbf7d0", "dot": "🟢"}
}

STATUSES = {
    "todo": {"label": "To Do", "color": "#475569", "bg": "#f1f5f9"},
    "in-progress": {"label": "In Progress", "color": "#1d4ed8", "bg": "#eff6ff"},
    "review": {"label": "Review", "color": "#7e34cd", "bg": "#f5f3ff"},
    "done": {"label": "Done", "color": "#047857", "bg": "#ecfdf5"},
    "blocked": {"label": "Blocked", "color": "#b91c1c", "bg": "#fef2f2"}
}

# 4. Helper Formatters
def format_minutes(m):
    if m < 60:
        return f"{m}m"
    h = m // 60
    rem = m % 60
    return f"{h}h {rem}m" if rem > 0 else f"{h}h"

def format_due_date(due_date_str):
    today_str = datetime.now().strftime("%Y-%m-%d")
    tomorrow_str = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")
    if due_date_str == today_str:
        return "Today"
    elif due_date_str == tomorrow_str:
        return "Tomorrow"
    elif due_date_str < today_str:
        try:
            diff = (datetime.strptime(today_str, "%Y-%m-%d") - datetime.strptime(due_date_str, "%Y-%m-%d")).days
            return f"{diff}d overdue"
        except:
            return due_date_str
    else:
        try:
            diff = (datetime.strptime(due_date_str, "%Y-%m-%d") - datetime.strptime(today_str, "%Y-%m-%d")).days
            return f"In {diff}d"
        except:
            return due_date_str

def is_overdue(due_date_str, status):
    return due_date_str < datetime.now().strftime("%Y-%m-%d") and status != "done"

# 5. Adaptive Card Renderers
def adaptive_metric_card(icon, value, title, subtext, accent_color="#3b82f6"):
    return f'''
    <div class="metric-card-adaptive" style="border-left: 5px solid {accent_color} !important;">
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <span style="font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; opacity: 0.7;">{title}</span>
            <span style="font-size: 18px;">{icon}</span>
        </div>
        <div style="font-size: 28px; font-weight: 800; margin-top: 4px; margin-bottom: 2px; letter-spacing: -0.025em;">{value}</div>
        <div style="font-size: 11px; opacity: 0.6; font-weight: 600;">{subtext}</div>
    </div>
    '''

def progress_bar(percentage, color="#3b82f6"):
    return f'''
    <div style="width: 100%; background-color: rgba(128, 128, 128, 0.2); border-radius: 9999px; height: 8px; margin: 6px 0;">
        <div style="background-color: {color}; width: {min(percentage, 100)}%; height: 8px; border-radius: 9999px; transition: width 0.3s ease;"></div>
    </div>
    '''

def category_list_item(icon, label, count_text, color="#3b82f6"):
    return f'''
    <div class="category-list-item-class" style="display: flex; justify-content: space-between; align-items: center; padding: 8px 0; border-bottom: 1px solid rgba(128, 128, 128, 0.1); font-size: 13px;">
        <div style="display: flex; align-items: center; gap: 8px;">
            <span style="font-size: 14px;">{icon}</span>
            <span style="font-weight: 600;">{label}</span>
        </div>
        <span style="font-weight: 700; color: {color} !important; font-size: 11px; background-color: rgba(128, 128, 128, 0.1); padding: 2px 8px; border-radius: 9999px;">{count_text}</span>
    </div>
    '''

def schedule_item_html(start, end, title, category, icon, color, bg):
    return f'''
    <div class="metric-card-adaptive" style="
        border-left: 4px solid {color} !important;
        background-color: rgba(128, 128, 128, 0.05) !important;
    ">
        <div style="font-weight: 600; font-size: 11px; color: gray; margin-bottom: 2px;">⏱ {start} – {end}</div>
        <div style="font-weight: 700; margin-bottom: 4px;">{title}</div>
        <div style="display: inline-block; font-size: 10px; font-weight: 700; padding: 2px 8px; border-radius: 6px; background-color: {bg}; color: {color};">
            {icon} {category}
        </div>
    </div>
    '''

# Fetch local Excel DB instances
ExcelDB.initialize_db()
tasks = ExcelDB.get_tasks()
time_blocks = ExcelDB.get_time_blocks()
activity_logs = ExcelDB.get_activity_logs()
notes_list = ExcelDB.get_notes()

today_str = datetime.now().strftime("%Y-%m-%d")

# 5. Session State Keys
if "selected_task_id" not in st.session_state:
    st.session_state.selected_task_id = None
if "edit_mode" not in st.session_state:
    st.session_state.edit_mode = False
if "pomo_running" not in st.session_state:
    st.session_state.pomo_running = False
if "pomo_time_left" not in st.session_state:
    st.session_state.pomo_time_left = 25 * 60
if "pomo_mode" not in st.session_state:
    st.session_state.pomo_mode = "Pomodoro"
if "pending_block" not in st.session_state:
    st.session_state.pending_block = None
if "show_overlap_warning" not in st.session_state:
    st.session_state.show_overlap_warning = False
if "task_select_mode" not in st.session_state:
    st.session_state.task_select_mode = False
if "selected_task_ids" not in st.session_state:
    st.session_state.selected_task_ids = []

# 6. Sidebar Navigation
with st.sidebar:
    st.markdown("<h2 style='margin-bottom:0;'>💼 WorkManager</h2>", unsafe_allow_html=True)
    st.caption("Daily Productivity Hub")
    st.divider()
    
    # Track task overdue count for badging
    overdue_count = len([t for t in tasks if is_overdue(t["dueDate"], t["status"])])
    overdue_badge = f" 🔴 {overdue_count}" if overdue_count > 0 else ""
    
    nav_options = [
        "📊 Dashboard",
        f"✅ Tasks{overdue_badge}",
        "🗂 Kanban Board",
        "⏱ Time Manager",
        "📝 Daily Notes",
        "📜 Activity Log"
    ]
    
    selected_nav = st.sidebar.radio("Navigation", nav_options, index=0, label_visibility="collapsed")
    
    # Bulletproof navigation router mapping
    nav_mapping = {
        "📊 Dashboard": "dashboard",
        f"✅ Tasks{overdue_badge}": "tasks",
        "🗂 Kanban Board": "kanban",
        "⏱ Time Manager": "time",
        "📝 Daily Notes": "notes",
        "📜 Activity Log": "activity"
    }
    nav_route = nav_mapping.get(selected_nav, "dashboard")
    
    st.divider()
    
    # POMODORO TIMER SIDEBAR INTEGRATION (High Performance Client-Side JS)
    st.markdown("### ⏱ Focus Pomodoro")
    
    # Mode selector buttons
    p_col1, p_col2, p_col3 = st.columns(3)
    with p_col1:
        if st.button("🍅 Pomo", key="p_m1", use_container_width=True):
            st.session_state.pomo_mode = "Pomodoro"
            st.session_state.pomo_time_left = 25 * 60
            st.session_state.pomo_running = False
    with p_col2:
        if st.button("☕ Short", key="p_m2", use_container_width=True):
            st.session_state.pomo_mode = "Short Break"
            st.session_state.pomo_time_left = 5 * 60
            st.session_state.pomo_running = False
    with p_col3:
        if st.button("🌴 Long", key="p_m3", use_container_width=True):
            st.session_state.pomo_mode = "Long Break"
            st.session_state.pomo_time_left = 15 * 60
            st.session_state.pomo_running = False
            
    # Target task picker for time logging
    active_tasks = [t for t in tasks if t["status"] != "done"]
    task_pomo_opts = {t["id"]: f"[{CATEGORIES[t['category']]['icon']}] {t['title']}" for t in active_tasks}
    task_pomo_opts[None] = "-- General Focus Session --"
    pomo_task_id = st.selectbox("Focus Target Task", options=list(task_pomo_opts.keys()), format_func=lambda x: task_pomo_opts[x])
    
    # Timer Display HTML / JS Ticking clock (Runs on client-side, allows seamless navigation!)
    target_title = task_pomo_opts.get(pomo_task_id, "General Focus Session")
    
    iframe_timer_html = f'''
    <html>
    <head>
    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            margin: 0;
            padding: 0;
            color: inherit;
        }}
    </style>
    </head>
    <body>
    <div style="text-align: center; padding: 12px; border-radius: 10px; background-color: rgba(128,128,128,0.06); border: 1px solid rgba(128,128,128,0.12); margin-bottom: 5px;">
        <div style="font-size: 10px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; opacity: 0.6; margin-bottom: 3px;">{st.session_state.pomo_mode}</div>
        <div id="pomo-clock" style="font-size: 38px; font-weight: 800; font-family: monospace; letter-spacing: -0.02em; line-height: 1; margin-bottom: 3px;">00:00</div>
        <div style="font-size: 11px; font-weight: 500; opacity: 0.7; margin-bottom: 8px; max-width: 100%; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="{target_title}">🎯 {target_title}</div>
        
        <div style="display: flex; gap: 8px; justify-content: center;">
            <button id="pomo-toggle-btn" style="background-color: #2563eb; color: white; border: none; padding: 5px 12px; border-radius: 4px; font-size: 11px; font-weight: 600; cursor: pointer; transition: background-color 0.2s;">Start</button>
            <button id="pomo-reset-btn" style="background-color: transparent; border: 1px solid rgba(128,128,128,0.3); padding: 5px 12px; border-radius: 4px; font-size: 11px; font-weight: 600; cursor: pointer; color: inherit;">Reset</button>
        </div>
    </div>
    
    <script>
        let duration = {st.session_state.pomo_time_left};
        let isRunning = false;
        let timerId = null;
        
        const clockDiv = document.getElementById("pomo-clock");
        const toggleBtn = document.getElementById("pomo-toggle-btn");
        const resetBtn = document.getElementById("pomo-reset-btn");
        
        function formatTime(s) {{
            const mins = Math.floor(s / 60);
            const secs = s % 60;
            return (mins < 10 ? "0" : "") + mins + ":" + (secs < 10 ? "0" : "") + secs;
        }}
        
        function updateDisplay() {{
            clockDiv.textContent = formatTime(duration);
            toggleBtn.textContent = isRunning ? "Pause" : "Start";
            toggleBtn.style.backgroundColor = isRunning ? "#dc2626" : "#2563eb";
        }}
        
        function tick() {{
            if (isRunning && duration > 0) {{
                duration--;
                updateDisplay();
                if (duration === 0) {{
                    isRunning = false;
                    clearInterval(timerId);
                    updateDisplay();
                    alert("Focus session complete! Please click 'Log Completed Session' in the sidebar to save progress.");
                }}
            }}
        }}
        
        toggleBtn.addEventListener("click", () => {{
            isRunning = !isRunning;
            updateDisplay();
            if (isRunning && !timerId) {{
                timerId = setInterval(tick, 1000);
            }} else if (!isRunning) {{
                clearInterval(timerId);
                timerId = null;
            }}
        }});
        
        resetBtn.addEventListener("click", () => {{
            isRunning = false;
            clearInterval(timerId);
            timerId = null;
            duration = {st.session_state.pomo_time_left};
            updateDisplay();
        }});
        
        updateDisplay();
    </script>
    </body>
    </html>
    '''
    # Using st.iframe to replace the deprecated st.components.v1.html cleanly!
    st.iframe(iframe_timer_html, height=140)
    
    # Save button for Streamlit to persist Excel logs when timer finishes
    pomo_minutes = 25 if st.session_state.pomo_mode == "Pomodoro" else (5 if st.session_state.pomo_mode == "Short Break" else 15)
    if st.button(f"📥 Log Completed {pomo_minutes}m Block", key="complete_pomo_log_btn", use_container_width=True):
        if pomo_task_id:
            ExcelDB.log_time(pomo_task_id, pomo_minutes)
            st.toast(f"Success! Logged {pomo_minutes}m to task!")
        else:
            ExcelDB.add_activity_log("None", "General Focus Session", f"Completed a {pomo_minutes}m focus Pomodoro block.")
            st.toast(f"Success! General session logged.")
        st.balloons()
        st.rerun()

    st.divider()
    
    total_tasks = len(tasks)
    completed_tasks = len([t for t in tasks if t["status"] == "done"])
    st.markdown(f"**{total_tasks} Total Tasks**")
    st.caption(f"{completed_tasks} Completed Tasks")

# 7. AN INSPIRING MOTIVATIONAL HEADLINE
st.markdown("<h4 style='text-align: center; color: var(--text-color); opacity: 0.8; font-weight: 800; background-color: rgba(128, 128, 128, 0.05); padding: 12px; border-radius: 12px; border: 1px solid rgba(128, 128, 128, 0.1); margin-bottom: 20px;'>🚀 FORGE YOUR LEGACY TODAY: STAY FOCUSED, CONQUER YOUR LIMITS, AND SLAY YOUR MILESTONES!</h4>", unsafe_allow_html=True)

# 8. ROUTE SWITCHING
if nav_route == "dashboard":
    st.markdown("## 📊 Dashboard")
    st.caption(f"Welcome back! {datetime.now().strftime('%A, %B %d, %Y')}")
    
    # A. KPI Cards Section
    today_tasks = [t for t in tasks if t["dueDate"] == today_str]
    today_done = len([t for t in today_tasks if t["status"] == "done"])
    in_progress_count = len([t for t in tasks if t["status"] == "in-progress"])
    
    completion_rate = round(completed_tasks / total_tasks * 100) if total_tasks > 0 else 0
    
    kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
    with kpi_col1:
        st.markdown(adaptive_metric_card("📋", len(today_tasks), "Today's Tasks", f"{today_done} done", "#2563eb"), unsafe_allow_html=True)
    with kpi_col2:
        st.markdown(adaptive_metric_card("⚡", in_progress_count, "In Progress", "active now", "#d97706"), unsafe_allow_html=True)
    with kpi_col3:
        st.markdown(adaptive_metric_card("⚠️", overdue_count, "Overdue Tasks", "needs attention" if overdue_count > 0 else "all caught up", "#dc2626"), unsafe_allow_html=True)
    with kpi_col4:
        st.markdown(adaptive_metric_card("✅", f"{completion_rate}%", "Completion Rate", f"{completed_tasks} of {total_tasks} done", "#10b981"), unsafe_allow_html=True)
        
    st.markdown("<br>", unsafe_allow_html=True)
    
    # B. Two-Pane Content Grid (With "Recent Activities" removed as requested!)
    body_col1, body_col2 = st.columns([1, 1.3])
    
    with body_col1:
        # Time logging metrics
        with st.container(border=True):
            st.markdown("#### ⏱ Time Summary")
            total_logged = sum(t["loggedMinutes"] for t in tasks)
            total_estimated = sum(t["estimatedMinutes"] for t in tasks)
            
            st.markdown(f"**Total Logged:** {format_minutes(total_logged)} &nbsp;&nbsp;|&nbsp;&nbsp; **Estimated Total:** {format_minutes(total_estimated)}")
            perc = (total_logged / total_estimated * 100) if total_estimated > 0 else 0
            st.markdown(progress_bar(perc, "#2563eb"), unsafe_allow_html=True)
            st.caption(f"{round(perc)}% of estimated time utilized")
            
            st.divider()
            
            # Category summaries
            st.markdown("<div style='font-size:12px; font-weight:600; text-transform:uppercase; color:gray;'>By Category</div>", unsafe_allow_html=True)
            for c_id, c_info in CATEGORIES.items():
                cat_tasks = [t for t in tasks if t["category"] == c_id]
                if cat_tasks:
                    done_c = len([t for t in cat_tasks if t["status"] == "done"])
                    st.markdown(category_list_item(c_info["icon"], c_info["label"], f"{done_c}/{len(cat_tasks)} tasks", c_info["color"]), unsafe_allow_html=True)
                    
        # Priority items count
        with st.container(border=True):
            st.markdown("#### 🔥 Open by Priority")
            for p_id, p_info in PRIORITIES.items():
                open_p_tasks = len([t for t in tasks if t["priority"] == p_id and t["status"] != "done"])
                if open_p_tasks > 0:
                    st.markdown(f'''
                    <div style="display:flex; justify-content:space-between; align-items:center; padding:6px 0; font-size:13px;">
                        <span>{p_info["dot"]} <b>{p_info["label"]}</b></span>
                        <span style="font-weight:600; background-color:{p_info["bg"]}; color:{p_info["color"]}; padding:2px 8px; border-radius:12px; font-size:11px; border: 1px solid {p_info["border"]}">{open_p_tasks} open</span>
                    </div>
                    ''', unsafe_allow_html=True)
                    
    with body_col2:
        # Timeline Schedule summary
        with st.container(border=True):
            st.markdown("#### 🗓 Today's Schedule")
            today_blocks = [b for b in time_blocks if b["date"] == today_str]
            if not today_blocks:
                st.caption("No schedule blocks allocated for today.")
            else:
                for b in sorted(today_blocks, key=lambda x: x["startTime"]):
                    c_info = CATEGORIES.get(b["category"], {"icon": "📌", "color": "gray", "bg": "#f3f4f6", "label": b["category"]})
                    st.markdown(schedule_item_html(b["startTime"], b["endTime"], b["taskTitle"], c_info["label"], c_info["icon"], c_info["color"], c_info["bg"]), unsafe_allow_html=True)

    # C. Status Summary row
    st.markdown("#### 🗂 Status Overview")
    stat_cols = st.columns(5)
    for idx, (s_id, s_info) in enumerate(STATUSES.items()):
        s_tasks = len([t for t in tasks if t["status"] == s_id])
        with stat_cols[idx]:
            st.markdown(f'''
            <div style="background-color: {s_info["bg"]}; color: {s_info["color"]}; border-radius: 8px; padding: 12px; text-align: center;">
                <div style="font-size: 24px; font-weight: 700;">{s_tasks}</div>
                <div style="font-size: 11px; font-weight: 600; text-transform: uppercase;">{s_info["label"]}</div>
            </div>
            ''', unsafe_allow_html=True)

elif nav_route == "tasks":
    st.markdown("## ✅ Task Manager")
    
    # Collapsible Creator Form
    with st.expander("➕ Create New Task", expanded=False):
        with st.form("new_task_form", clear_on_submit=True):
            f_title = st.text_input("Task Title *", placeholder="What are you working on?")
            f_desc = st.text_area("Description", placeholder="Task description...")
            
            col_a, col_b = st.columns(2)
            with col_a:
                f_cat = st.selectbox("Category", list(CATEGORIES.keys()), format_func=lambda x: f"{CATEGORIES[x]['icon']} {CATEGORIES[x]['label']}")
                f_prio = st.selectbox("Priority", list(PRIORITIES.keys()), index=2, format_func=lambda x: f"{PRIORITIES[x]['dot']} {PRIORITIES[x]['label']}")
                f_status = st.selectbox("Status", list(STATUSES.keys()), index=0, format_func=lambda x: STATUSES[x]["label"])
            with col_b:
                f_date = st.date_input("Due Date", value=datetime.now())
                f_est = st.number_input("Estimated Minutes", min_value=0, value=60, step=15)
                f_tags = st.text_input("Tags (comma separated)", placeholder="finance, report, roadmap")
                
            submitted = st.form_submit_button("Create Task", use_container_width=True)
            if submitted:
                if not f_title.strip():
                    st.error("Task title cannot be empty.")
                else:
                    new_id = ExcelDB.add_task(
                        title=f_title,
                        description=f_desc,
                        category=f_cat,
                        priority=f_prio,
                        status=f_status,
                        dueDate=f_date.strftime("%Y-%m-%d"),
                        estimatedMinutes=f_est,
                        tags=f_tags
                    )
                    st.success(f"Task '{f_title}' added successfully!")
                    st.rerun()

    # Search and Filter Bars
    st.markdown("#### 🔍 Filter & Sort Tasks")
    filt_col1, filt_col2, filt_col3, filt_col4, filt_col5 = st.columns([1.5, 1, 1, 1, 1])
    with filt_col1:
        search_q = st.text_input("Search", placeholder="Search by title or tags...", label_visibility="collapsed")
    with filt_col2:
        cat_filter = st.selectbox("Category", ["All"] + list(CATEGORIES.keys()), format_func=lambda x: "All Categories" if x == "All" else f"{CATEGORIES[x]['icon']} {CATEGORIES[x]['label']}", label_visibility="collapsed")
    with filt_col3:
        prio_filter = st.selectbox("Priority", ["All"] + list(PRIORITIES.keys()), format_func=lambda x: "All Priorities" if x == "All" else f"{PRIORITIES[x]['dot']} {PRIORITIES[x]['label']}", label_visibility="collapsed")
    with filt_col4:
        status_filter = st.selectbox("Status", ["All"] + list(STATUSES.keys()), format_func=lambda x: "All Statuses" if x == "All" else STATUSES[x]["label"], label_visibility="collapsed")
    with filt_col5:
        sort_by = st.selectbox("Sort", ["Due Date", "Priority", "Created At"], label_visibility="collapsed")

    # Apply filters mathematically
    prio_ranks = {"critical": 0, "high": 1, "medium": 2, "low": 3}
    filtered_tasks = [
        t for t in tasks
        if (not search_q or search_q.lower() in t["title"].lower() or any(search_q.lower() in tag.lower() for tag in t["tags"]))
        and (cat_filter == "All" or t["category"] == cat_filter)
        and (prio_filter == "All" or t["priority"] == prio_filter)
        and (status_filter == "All" or t["status"] == status_filter)
    ]
    
    # Apply Sorting
    if sort_by == "Due Date":
        filtered_tasks.sort(key=lambda x: x["dueDate"])
    elif sort_by == "Priority":
        filtered_tasks.sort(key=lambda x: prio_ranks.get(x["priority"], 99))
    else:
        filtered_tasks.sort(key=lambda x: x["createdAt"], reverse=True)

    # Multi-Column details alignment
    list_col, details_col = st.columns([1.6, 1]) if st.session_state.selected_task_id else (st.columns([1, 0.001])[0], None)
    
    with list_col:
        # BULK ACTIONS CONTROLS
        st.session_state.task_select_mode = st.checkbox("Toggle Selection Mode (Bulk Edit) 🗂", value=st.session_state.task_select_mode)
        
        if st.session_state.task_select_mode:
            c1, c2 = st.columns([0.2, 2])
            with c1:
                if st.checkbox("Select All", key="select_all_tasks_box"):
                    st.session_state.selected_task_ids = [t["id"] for t in filtered_tasks]
                else:
                    if len(st.session_state.selected_task_ids) == len(filtered_tasks):
                        st.session_state.selected_task_ids = []
            
        st.markdown(f"**Showing {len(filtered_tasks)} tasks**")
        
        if not filtered_tasks:
            st.info("No tasks matched your filter criteria.")
        else:
            for t in filtered_tasks:
                c_info = CATEGORIES.get(t["category"], {})
                p_info = PRIORITIES.get(t["priority"], {})
                s_info = STATUSES.get(t["status"], {})
                is_od = is_overdue(t["dueDate"], t["status"])
                perc = (t["loggedMinutes"] / t["estimatedMinutes"] * 100) if t["estimatedMinutes"] > 0 else 0
                
                # Inline action checkmarks
                card_cols = st.columns([0.08, 0.92])
                with card_cols[0]:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.session_state.task_select_mode:
                        is_selected = st.checkbox("Select", value=(t["id"] in st.session_state.selected_task_ids), key=f"sel_check_{t['id']}", label_visibility="collapsed")
                        if is_selected and t["id"] not in st.session_state.selected_task_ids:
                            st.session_state.selected_task_ids.append(t["id"])
                        elif not is_selected and t["id"] in st.session_state.selected_task_ids:
                            st.session_state.selected_task_ids.remove(t["id"])
                    else:
                        checked = st.checkbox("Complete Task", value=(t["status"] == "done"), key=f"check_{t['id']}", label_visibility="collapsed")
                        if checked != (t["status"] == "done"):
                            ExcelDB.update_task(t["id"], {"status": "done" if checked else "todo"})
                            st.rerun()
                        
                with card_cols[1]:
                    # Adaptive task container
                    with st.container(border=True):
                        strike = "text-decoration: line-through; opacity: 0.6;" if t["status"] == "done" else ""
                        st.markdown(f"<div style='font-size:16px; font-weight:700; {strike}'>{t['title']}</div>", unsafe_allow_html=True)
                        if t["description"]:
                            st.markdown(f"<div style='font-size:13px; opacity:0.8; margin-top:2px;'>{t['description']}</div>", unsafe_allow_html=True)
                            
                        # Metadata rows
                        od_badge = "🔴 Overdue" if is_od else ""
                        tag_pills = " ".join([f"<span style='background-color:rgba(128,128,128,0.1); padding:2px 8px; border-radius:4px; font-size:11px; margin-right:4px;'>#{tag}</span>" for tag in t["tags"]])
                        st.markdown(f'''
                        <div style="display:flex; flex-wrap:wrap; gap:10px; align-items:center; margin-top:8px; font-size:12px;">
                            <span style="background-color:{c_info.get("bg")}; color:{c_info.get("color")}; padding:2px 8px; border-radius:12px; font-weight:600;">{c_info.get("icon")} {c_info.get("label")}</span>
                            <span style="background-color:{p_info.get("bg")}; color:{p_info.get("color")}; padding:2px 8px; border-radius:12px; font-weight:600; border: 1px solid {p_info.get("border")}">{p_info.get("label")}</span>
                            <span style="color:{'#dc2626' if is_od else 'gray'}; font-weight:{'700' if is_od else 'normal'}">📅 {format_due_date(t["dueDate"])} ({t["dueDate"]}) {od_badge}</span>
                            <span>⏱ {format_minutes(t["loggedMinutes"])} / {format_minutes(t["estimatedMinutes"])}</span>
                        </div>
                        ''', unsafe_allow_html=True)
                        if tag_pills:
                            st.markdown(f"<div style='margin-top:6px;'>{tag_pills}</div>", unsafe_allow_html=True)
                        
                        if t["estimatedMinutes"] > 0:
                            st.markdown(progress_bar(perc, "#2563eb"), unsafe_allow_html=True)
                            
                        btn_cols = st.columns([1, 4])
                        with btn_cols[0]:
                            if st.button("View Details", key=f"det_btn_{t['id']}", use_container_width=True):
                                st.session_state.selected_task_id = t["id"]
                                st.session_state.edit_mode = False
                                st.rerun()

    # FLOATING ACTION BAR FOR BULK DELETE
    if st.session_state.task_select_mode and st.session_state.selected_task_ids:
        st.markdown(
            f'''
            <div style="
                position: fixed;
                bottom: 1.5rem;
                left: 55%;
                transform: translateX(-50%);
                background-color: #fca5a5;
                color: #7f1d1d;
                border: 1px solid #f87171;
                padding: 12px 24px;
                border-radius: 12px;
                box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1);
                display: flex;
                align-items: center;
                gap: 20px;
                z-index: 999;
            ">
                <span style="font-weight: 700; font-size:14px;">🚨 {len(st.session_state.selected_task_ids)} Tasks Selected</span>
            </div>
            ''',
            unsafe_allow_html=True
        )
        if st.button("🗑️ Delete Selected Tasks", key="bulk_delete_confirm_btn", use_container_width=True):
            ExcelDB.delete_multiple_tasks(st.session_state.selected_task_ids)
            st.success(f"Successfully deleted {len(st.session_state.selected_task_ids)} tasks.")
            st.session_state.selected_task_ids = []
            st.rerun()

    # C. Detail Overlays Column
    if details_col and st.session_state.selected_task_id:
        t_id = st.session_state.selected_task_id
        task_detail = next((t for t in tasks if str(t["id"]) == str(t_id)), None)
        
        if task_detail:
            c_info = CATEGORIES.get(task_detail["category"], {})
            p_info = PRIORITIES.get(task_detail["priority"], {})
            s_info = STATUSES.get(task_detail["status"], {})
            
            with details_col:
                with st.container(border=True):
                    det_header_col = st.columns([8, 1])
                    with det_header_col[0]:
                        st.markdown(f"### {c_info.get('icon')} {task_detail['title']}")
                    with det_header_col[1]:
                        if st.button("✕", key="close_details_btn"):
                            st.session_state.selected_task_id = None
                            st.rerun()
                            
                    st.divider()
                    
                    # Form toggle view
                    edit_mode = st.toggle("Edit Mode ⚙️", value=st.session_state.edit_mode, key="toggle_edit_mode")
                    st.session_state.edit_mode = edit_mode
                    
                    if edit_mode:
                        with st.form("edit_task_details_form"):
                            e_title = st.text_input("Title", value=task_detail["title"])
                            e_desc = st.text_area("Description", value=task_detail["description"])
                            
                            col_e1, col_e2 = st.columns(2)
                            with col_e1:
                                e_cat = st.selectbox("Category", list(CATEGORIES.keys()), index=list(CATEGORIES.keys()).index(task_detail["category"]), format_func=lambda x: f"{CATEGORIES[x]['icon']} {CATEGORIES[x]['label']}")
                                e_prio = st.selectbox("Priority", list(PRIORITIES.keys()), index=list(PRIORITIES.keys()).index(task_detail["priority"]), format_func=lambda x: f"{PRIORITIES[x]['dot']} {PRIORITIES[x]['label']}")
                            with col_e2:
                                e_status = st.selectbox("Status", list(STATUSES.keys()), index=list(STATUSES.keys()).index(task_detail["status"]), format_func=lambda x: STATUSES[x]["label"])
                                e_date = st.date_input("Due Date", value=datetime.strptime(task_detail["dueDate"], "%Y-%m-%d"))
                                
                            e_est = st.number_input("Estimated Minutes", min_value=0, value=int(task_detail["estimatedMinutes"]), step=15)
                            e_tags = st.text_input("Tags (comma separated)", value=", ".join(task_detail["tags"]))
                            
                            saved = st.form_submit_button("Save Changes", use_container_width=True)
                            if saved:
                                ExcelDB.update_task(t_id, {
                                    "title": e_title,
                                    "description": e_desc,
                                    "category": e_cat,
                                    "priority": e_prio,
                                    "status": e_status,
                                    "dueDate": e_date.strftime("%Y-%m-%d"),
                                    "estimatedMinutes": e_est,
                                    "tags": [x.strip() for x in e_tags.split(",") if x.strip()]
                                })
                                st.success("Task updated successfully!")
                                st.session_state.edit_mode = False
                                st.rerun()
                    else:
                        if task_detail["description"]:
                            st.markdown(f"*{task_detail['description']}*")
                        else:
                            st.caption("No description provided.")
                            
                        st.divider()
                        
                        grid_col1, grid_col2 = st.columns(2)
                        with grid_col1:
                            st.markdown(f"**Category:** {c_info.get('icon')} {c_info.get('label')}")
                            st.markdown(f"**Priority:** {p_info.get('dot')} {p_info.get('label')}")
                        with grid_col2:
                            st.markdown(f"**Status:** {STATUSES[task_detail['status']]['label']}")
                            st.markdown(f"**Due Date:** {format_due_date(task_detail['dueDate'])} ({task_detail['dueDate']})")
                            
                        st.divider()
                        
                        # Interactive Time Logging
                        st.markdown("##### ⏱ Log Focus Time")
                        time_log_cols = st.columns([2, 1])
                        with time_log_cols[0]:
                            log_min = st.number_input("Log Minutes", min_value=1, value=30, step=5, key="det_log_minutes")
                        with time_log_cols[1]:
                            st.markdown("<br>", unsafe_allow_html=True)
                            if st.button("Log Time", key="det_log_time_btn", use_container_width=True):
                                ExcelDB.log_time(t_id, log_min)
                                st.success(f"Logged {log_min} minutes for this task!")
                                st.rerun()
                                
                        st.divider()

                        # ADD COMMENT WITH TIMESTAMP after creation
                        st.markdown("##### 💬 Comments & Notes")
                        with st.form("task_comment_form", clear_on_submit=True):
                            new_comment_text = st.text_input("Add comment...", placeholder="Write a note, update, or reminder...")
                            add_comment_submitted = st.form_submit_button("Add Comment", use_container_width=True)
                            if add_comment_submitted and new_comment_text.strip():
                                ExcelDB.add_comment(t_id, new_comment_text)
                                st.success("Comment added!")
                                st.rerun()

                        # List task comments
                        task_comments = ExcelDB.get_comments(t_id)
                        if task_comments:
                            for comm in task_comments:
                                dt = datetime.strptime(comm["timestamp"], "%Y-%m-%d %H:%M:%S")
                                time_str = dt.strftime("%b %d, %H:%M")
                                st.markdown(f'''
                                <div style="font-size:12px; border-bottom:1px solid rgba(128,128,128,0.05); padding:6px 0; margin-bottom: 4px;">
                                    <div style="font-weight:600; opacity:0.8;">{comm["commentText"]}</div>
                                    <div style="color:gray; font-size:10px; text-align:right; margin-top:2px;">⏱ {time_str}</div>
                                </div>
                                ''', unsafe_allow_html=True)
                        else:
                            st.caption("No comments posted yet.")

                        st.divider()
                        
                        # OUTLOOK ICS INTEGRATION BUTTONS
                        st.markdown("##### 📅 Save to Calendar (Outlook)")
                        raw_ics_content = ExcelDB.generate_ics(
                            task_id=task_detail["id"],
                            title=task_detail["title"],
                            description=task_detail["description"],
                            due_date_str=task_detail["dueDate"]
                        )
                        st.download_button(
                            label="📥 Download .ics",
                            data=raw_ics_content,
                            file_name=f"task_{task_detail['id']}.ics",
                            mime="text/calendar",
                            use_container_width=True,
                            help="Save this event to local Outlook or Apple Calendar"
                        )
                        
                        st.divider()
                        
                        if st.button("🗑 Delete Task", key="delete_task_btn", use_container_width=True):
                            ExcelDB.delete_task(t_id)
                            st.warning("Task deleted successfully!")
                            st.session_state.selected_task_id = None
                            st.rerun()

elif nav_route == "kanban":
    st.markdown("## 🗂 Kanban Board")
    st.caption("Lay out and interact with columns chronologically")
    
    k_cols = st.columns(5)
    for s_idx, (s_id, s_info) in enumerate(STATUSES.items()):
        column_tasks = [t for t in tasks if t["status"] == s_id]
        
        with k_cols[s_idx]:
            st.markdown(f'''
            <div style="
                background-color: {s_info["bg"]};
                color: {s_info["color"]};
                border-radius: 8px 8px 0 0;
                padding: 10px 14px;
                display: flex;
                justify-content: space-between;
                align-items: center;
                border-bottom: 2px solid {s_info["color"]}40;
                font-weight: 700;
                font-size: 14px;
            ">
                <span>{s_info["label"]}</span>
                <span style="background-color:rgba(255,255,255,0.6); padding:2px 8px; border-radius:9999px; font-size:11px;">{len(column_tasks)}</span>
            </div>
            ''', unsafe_allow_html=True)
            
            with st.container(border=True):
                if not column_tasks:
                    st.caption("No tasks in this status.")
                else:
                    for t in column_tasks:
                        c_info = CATEGORIES.get(t["category"], {})
                        p_info = PRIORITIES.get(t["priority"], {})
                        is_od = is_overdue(t["dueDate"], t["status"])
                        
                        st.markdown(f'''
                        <div class="metric-card-adaptive" style="
                            background-color: rgba(128,128,128,0.03);
                            border: 1px solid rgba(128,128,128,0.12);
                            border-left: 3px solid {p_info.get("color")};
                            border-radius: 8px;
                            padding: 12px;
                            margin-bottom: 10px;
                            font-size: 13px;
                            box-shadow: 0 1px 2px 0 rgba(0,0,0,0.02);
                        ">
                            <div style="display:flex; align-items:center; gap:6px; font-weight:600; margin-bottom:4px;">
                                <span>{p_info.get("dot")}</span>
                                <span style="font-weight:700;">{t["title"]}</span>
                            </div>
                            <div style="display:flex; flex-wrap:wrap; gap:4px; margin-bottom:6px;">
                                <span style="background-color:{c_info.get("bg")}; color:{c_info.get("color")}; padding:1px 4px; border-radius:4px; font-size:10px; font-weight:600;">{c_info.get("icon")} {c_info.get("label")}</span>
                                <span style="font-size:10px; color:{'#dc2626' if is_od else 'gray'}; font-weight:{'700' if is_od else 'normal'}">📅 {format_due_date(t["dueDate"])}</span>
                            </div>
                        </div>
                        ''', unsafe_allow_html=True)
                        
                        # Direct state adjustments
                        m_cols = st.columns(2)
                        with m_cols[0]:
                            new_s = st.selectbox("Move To", list(STATUSES.keys()), index=list(STATUSES.keys()).index(t["status"]), format_func=lambda x: STATUSES[x]["label"], key=f"kanban_s_{t['id']}", label_visibility="collapsed")
                            if new_s != t["status"]:
                                ExcelDB.update_task(t["id"], {"status": new_s})
                                st.rerun()
                        with m_cols[1]:
                            new_p = st.selectbox("Priority", list(PRIORITIES.keys()), index=list(PRIORITIES.keys()).index(t["priority"]), format_func=lambda x: PRIORITIES[x]["label"], key=f"kanban_p_{t['id']}", label_visibility="collapsed")
                            if new_p != t["priority"]:
                                ExcelDB.update_task(t["id"], {"priority": new_p})
                                st.rerun()

elif nav_route == "time":
    st.markdown("## ⏱ Time Manager")
    st.caption("Allocate daily hours and review category allocations")
    
    time_date = st.date_input("Schedule Date", value=datetime.now())
    selected_date_str = time_date.strftime("%Y-%m-%d")
    
    # DUP CHECK AND FORCE-SCHEDULE WARNING PROMPT
    if st.session_state.show_overlap_warning and st.session_state.pending_block:
        pending = st.session_state.pending_block
        with st.container(border=True):
            st.warning("⚠️ **SCHEDULING WARNING: Overlapping Time Slots Detected!**")
            st.markdown(f"The proposed block **{pending['taskTitle']} ({pending['startTime']} - {pending['endTime']})** overlaps with existing plans on this date.")
            
            # List overlapping blocks
            st.markdown("**Overlapping Scheduled Blocks:**")
            for b in st.session_state.get("overlapping_blocks_list", []):
                st.markdown(f"- **{b['taskTitle']}** &nbsp;&nbsp;&middot;&nbsp;&nbsp; {b['startTime']} - {b['endTime']}")
            
            st.markdown("What would you like to do?")
            
            btn_col_c1, btn_col_c2 = st.columns(2)
            with btn_col_c1:
                if st.button("⚠️ Continue & Force Schedule", key="force_schedule_block_btn", use_container_width=True):
                    ExcelDB.add_time_block(
                        task_id=pending["taskId"],
                        task_title=pending["taskTitle"],
                        start_time=pending["startTime"],
                        end_time=pending["endTime"],
                        date=pending["date"],
                        category=pending["category"]
                    )
                    st.success("Forced scheduling completed!")
                    st.session_state.pending_block = None
                    st.session_state.show_overlap_warning = False
                    st.rerun()
            with btn_col_c2:
                if st.button("✕ Stop & Cancel", key="cancel_schedule_block_btn", use_container_width=True):
                    st.toast("Scheduling cancelled.")
                    st.session_state.pending_block = None
                    st.session_state.show_overlap_warning = False
                    st.rerun()
                    
    block_cols = st.columns([1, 1.3])
    
    with block_cols[0]:
        # Block Allocation form
        with st.container(border=True):
            st.markdown("#### ➕ Add Time Block")
            with st.form("new_block_form", clear_on_submit=True):
                active_tasks = [t for t in tasks if t["status"] != "done"]
                task_opts = {t["id"]: f"[{CATEGORIES[t['category']]['icon']} {CATEGORIES[t['category']]['label']}] {t['title']}" for t in active_tasks}
                
                b_task_id = st.selectbox("Select Target Task", options=list(task_opts.keys()), format_func=lambda x: task_opts[x])
                
                col_t1, col_e2 = st.columns(2)
                with col_t1:
                    b_start = st.text_input("Start Time (HH:MM)", value="09:00", placeholder="09:00")
                with col_e2:
                    b_end = st.text_input("End Time (HH:MM)", value="10:00", placeholder="10:00")
                    
                b_submitted = st.form_submit_button("Add Block", use_container_width=True)
                if b_submitted:
                    try:
                        datetime.strptime(b_start, "%H:%M")
                        datetime.strptime(b_end, "%H:%M")
                        target_task = next((t for t in tasks if str(t["id"]) == str(b_task_id)), None)
                        if target_task:
                            # 1. Fetch current blocks for the selected date
                            selected_blocks = [b for b in time_blocks if b["date"] == selected_date_str]
                            
                            # 2. Check for overlaps
                            def to_mins(t_str):
                                h, m = map(int, t_str.split(":"))
                                return h * 60 + m
                            
                            new_s_m = to_mins(b_start)
                            new_e_m = to_mins(b_end)
                            
                            overlapping_blocks = []
                            for ob in selected_blocks:
                                try:
                                    ob_s_m = to_mins(ob["startTime"])
                                    ob_e_m = to_mins(ob["endTime"])
                                    if new_s_m < ob_e_m and new_e_m > ob_s_m:
                                        overlapping_blocks.append(ob)
                                except:
                                    pass
                            
                            # 3. Handle Warn vs Direct Save
                            if overlapping_blocks:
                                st.session_state.pending_block = {
                                    "taskId": b_task_id,
                                    "taskTitle": target_task["title"],
                                    "startTime": b_start,
                                    "endTime": b_end,
                                    "date": selected_date_str,
                                    "category": target_task["category"]
                                }
                                st.session_state.show_overlap_warning = True
                                st.session_state.overlapping_blocks_list = overlapping_blocks
                                st.rerun()
                            else:
                                # Save directly
                                ExcelDB.add_time_block(
                                    task_id=b_task_id,
                                    task_title=target_task["title"],
                                    start_time=b_start,
                                    end_time=b_end,
                                    date=selected_date_str,
                                    category=target_task["category"]
                                )
                                st.success("Time block scheduled!")
                                st.rerun()
                    except ValueError:
                        st.error("Invalid time format. Please use HH:MM format (e.g. 09:30)")
                        
        # Interactive Daily Agenda list
        with st.container(border=True):
            st.markdown("#### 📅 Hourly Agenda Timeline")
            selected_blocks = [b for b in time_blocks if b["date"] == selected_date_str]
            
            for h in range(7, 23):
                slot_str = f"{h:02d}:00"
                overlapping = []
                for b in selected_blocks:
                    try:
                        start_h = int(b["startTime"].split(":")[0])
                        end_h = int(b["endTime"].split(":")[0])
                        if start_h <= h <= end_h:
                            overlapping.append(b)
                    except:
                        pass
                
                if overlapping:
                    for b in overlapping:
                        c_info = CATEGORIES.get(b["category"], {})
                        st.markdown(f'''
                        <div style="display:flex; align-items:center; gap:10px; margin-bottom:4px; font-size:13px;">
                            <span style="font-weight:700; color:gray; width:45px;">{slot_str}</span>
                            <div style="flex:1; background-color:{c_info.get("bg")}80; border-left:4px solid {c_info.get("color")}; padding:6px 10px; border-radius:4px; font-weight:600;">
                                {c_info.get("icon")} {b["taskTitle"]} ({b["startTime"]} - {b["endTime"]})
                            </div>
                        </div>
                        ''', unsafe_allow_html=True)
                else:
                    st.markdown(f'''
                    <div style="display:flex; align-items:center; gap:10px; margin-bottom:4px; font-size:13px; opacity:0.4;">
                        <span style="color:gray; width:45px;">{slot_str}</span>
                        <div style="flex:1; border-top:1px dashed rgba(128,128,128,0.2); height:1px;"></div>
                    </div>
                    ''', unsafe_allow_html=True)

    with block_cols[1]:
        # Scheduled Blocks list
        with st.container(border=True):
            st.markdown("#### 🗒 Scheduled Time Blocks")
            selected_blocks = [b for b in time_blocks if b["date"] == selected_date_str]
            if not selected_blocks:
                st.caption("No schedule blocks allocated for this date.")
            else:
                for b in sorted(selected_blocks, key=lambda x: x["startTime"]):
                    c_info = CATEGORIES.get(b["category"], {})
                    
                    b_cols = st.columns([5, 1])
                    with b_cols[0]:
                        st.markdown(schedule_item_html(b["startTime"], b["endTime"], b["taskTitle"], c_info["label"], c_info["icon"], c_info["color"], c_info["bg"]), unsafe_allow_html=True)
                    with b_cols[1]:
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button("✕", key=f"del_block_{b['id']}", help="Remove scheduled block"):
                            ExcelDB.remove_time_block(b["id"])
                            st.warning("Block removed!")
                            st.rerun()

        # Allocations Sum by Category
        with st.container(border=True):
            st.markdown("#### 📊 Allocation by Category")
            selected_blocks = [b for b in time_blocks if b["date"] == selected_date_str]
            cat_minutes = {c: 0 for c in CATEGORIES.keys()}
            
            for b in selected_blocks:
                try:
                    sh, sm = map(int, b["startTime"].split(":"))
                    eh, em = map(int, b["endTime"].split(":"))
                    dur = (eh*60 + em) - (sh*60 + sm)
                    cat_minutes[b["category"]] += dur
                except:
                    pass
            
            has_allocations = False
            for c_id, mins in cat_minutes.items():
                if mins > 0:
                    has_allocations = True
                    c_info = CATEGORIES[c_id]
                    st.markdown(category_list_item(c_info["icon"], c_info["label"], format_minutes(mins), c_info["color"]), unsafe_allow_html=True)
            if not has_allocations:
                st.caption("No category allocations recorded for this date.")

elif nav_route == "notes":
    st.markdown("## 📝 Daily Notes Binder")
    st.caption("Quickly capture hints, reminders, meeting minutes, and tie document attachments")
    
    notes_col1, notes_col2 = st.columns([1, 1.4])
    
    with notes_col1:
        # Note Creator Form
        with st.container(border=True):
            st.markdown("#### ➕ Capture New Note")
            with st.form("new_note_form", clear_on_submit=True):
                n_title = st.text_input("Note Title *", placeholder="Meetingsync, Q3 Idea, Snippet...")
                n_tags = st.text_input("Tags (comma separated)", placeholder="roadmap, idea, backend")
                n_content = st.text_area("Content / Body (Markdown supported) *", placeholder="Type your thoughts, copy paste logs, or record hints here...", height=150)
                
                # Attachment file uploader
                uploaded_docs = st.file_uploader(
                    "📎 Attach Documents (PDF, Images, CSV, TXT, Excel)",
                    accept_multiple_files=True,
                    help="Upload related project documents or reference logs to save them permanently in local storage."
                )
                
                note_submitted = st.form_submit_button("Save Note", use_container_width=True)
                if note_submitted:
                    if not n_title.strip() or not n_content.strip():
                        st.error("Note title and content cannot be empty.")
                    else:
                        ExcelDB.add_note(
                            title=n_title,
                            content=n_content,
                            tags=n_tags,
                            uploaded_files=uploaded_docs
                        )
                        st.success("Note saved to database!")
                        st.rerun()
                        
        # Search note history filter
        with st.container(border=True):
            st.markdown("#### 🔍 Search Saved Notes")
            search_note_q = st.text_input("Search Notes", placeholder="Search by text, tags, or titles...", label_visibility="collapsed")

    with notes_col2:
        st.markdown("#### 📋 Note Feed")
        
        # Apply notes filtering
        filtered_notes = sorted(notes_list, key=lambda x: x["createdAt"], reverse=True)
        if search_note_q:
            q = search_note_q.lower()
            filtered_notes = [
                n for n in filtered_notes 
                if q in n["title"].lower() or q in n["content"].lower() or any(q in t.lower() for t in n["tags"])
            ]
            
        if not filtered_notes:
            st.info("No saved notes found. Create your first note in the left panel to begin capture!")
        else:
            for n in filtered_notes:
                # Create a gorgeous custom card container
                with st.container(border=True):
                    note_h_cols = st.columns([10, 1])
                    with note_h_cols[0]:
                        st.markdown(f"### 📝 {n['title']}")
                    with note_h_cols[1]:
                        if st.button("🗑", key=f"del_note_{n['id']}", help="Delete Note and purge files"):
                            ExcelDB.delete_note(n["id"])
                            st.warning("Note deleted successfully!")
                            st.rerun()
                            
                    # Formatting tags pills
                    note_tag_pills = " ".join([f"<span style='background-color:rgba(128,128,128,0.08); padding:2px 8px; border-radius:4px; font-size:10px; margin-right:4px; font-weight:600;'>#{tag}</span>" for tag in n["tags"]])
                    st.markdown(f'''
                    <div style="display:flex; align-items:center; gap:10px; margin-top:-10px; margin-bottom:12px; font-size:11px; color:gray;">
                        <span>📅 Created: {n['createdAt']}</span>
                        {note_tag_pills}
                    </div>
                    ''', unsafe_allow_html=True)
                    
                    # Content Markdown parser
                    st.markdown(n["content"].replace("\\n", "\n"))
                    
                    # Attachment loader list
                    attachments = ExcelDB.get_note_attachments(n["id"])
                    if attachments:
                        st.markdown("<div style='font-size:12px; font-weight:700; margin-top:12px; color:gray;'>📎 Linked Attachments:</div>", unsafe_allow_html=True)
                        for att in attachments:
                            att_cols = st.columns([4, 1.2])
                            with att_cols[0]:
                                size_kb = round(int(att["fileSize"]) / 1024, 1)
                                st.markdown(f"<span style='font-size:12px; font-weight:600;'>📄 {att['fileName']}</span> &nbsp;&middot;&nbsp; <span style='font-size:10px; color:gray;'>({size_kb} KB)</span>", unsafe_allow_html=True)
                                
                                # Inline visual helper: If file is image, display a clean preview
                                if att["fileType"].startswith("image/") and os.path.exists(att["filePath"]):
                                    st.image(att["filePath"], width=200)
                            with att_cols[1]:
                                if os.path.exists(att["filePath"]):
                                    with open(att["filePath"], "rb") as file_bin:
                                        st.download_button(
                                            label="📥 Download",
                                            data=file_bin.read(),
                                            file_name=att["fileName"],
                                            mime=att["fileType"],
                                            key=f"dl_att_{att['id']}",
                                            use_container_width=True
                                        )
                                else:
                                    st.caption("⚠️ File missing")
                    st.markdown("<br>", unsafe_allow_html=True)

elif nav_route == "activity":
    st.markdown("## 📜 Activity Log")
    st.caption("Full audit stream of operations and focus logs")
    
    st.markdown("#### 🔍 Search Activity Logs")
    act_filt_col1, act_col2 = st.columns([2, 1])
    with act_filt_col1:
        act_search = st.text_input("Search Logs", placeholder="Search activity description or task title...", label_visibility="collapsed")
    
    filt_activities = sorted(activity_logs, key=lambda x: x["timestamp"], reverse=True)
    if act_search:
        filt_activities = [a for a in filt_activities if act_search.lower() in a["taskTitle"].lower() or act_search.lower() in a["action"].lower()]
        
    if not filt_activities:
        st.info("No activity records match your criteria.")
    else:
        # Group list items dynamically by date headers
        grouped_activities = {}
        for act in filt_activities:
            date_str = act["timestamp"].split(" ")[0]
            if date_str not in grouped_activities:
                grouped_activities[date_str] = []
            grouped_activities[date_str].append(act)
            
        for d, acts in grouped_activities.items():
            if d == today_str:
                header = "Today"
            elif d == (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d"):
                header = "Yesterday"
            else:
                header = datetime.strptime(d, "%Y-%m-%d").strftime("%A, %B %d, %Y")
                
            st.markdown(f"##### {header} &nbsp;&nbsp;&middot;&nbsp;&nbsp; <span style='font-size:12px; color:gray;'>{len(acts)} events</span>", unsafe_allow_html=True)
            
            with st.container(border=True):
                for act in acts:
                    task_item = next((t for t in tasks if str(t["id"]) == str(act["taskId"])), None)
                    icon = "📌"
                    color = "#3b82f6"
                    if task_item:
                        c_info = CATEGORIES.get(task_item["category"], {})
                        icon = c_info.get("icon", icon)
                        color = c_info.get("color", color)
                        
                    dt = datetime.strptime(act["timestamp"], "%Y-%m-%d %H:%M:%S")
                    time_str = dt.strftime("%H:%M:%S")
                    
                    st.markdown(f'''
                    <div style="display:flex; align-items:start; gap:12px; padding:8px 0; border-bottom:1px solid rgba(128,128,128,0.05); font-size:13px;">
                        <span style="font-size:16px; background-color:rgba(128,128,128,0.05); padding:6px; border-radius:50%;">{icon}</span>
                        <div style="flex-1;">
                            <span style="font-weight:700;">{act["taskTitle"]}</span>
                            <span style="margin: 0 4px; opacity:0.6;">·</span>
                            <span>{act["action"]}</span>
                            <div style="font-size:10px; color:gray; margin-top:2px;">⏱ {time_str}</div>
                        </div>
                    </div>
                    ''', unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
