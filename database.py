import os
import pandas as pd
import uuid
from datetime import datetime, timedelta

EXCEL_FILE = "workspace_data.xlsx"
ATTACHMENT_DIR = "attachments"

class ExcelDB:
    @staticmethod
    def initialize_db():
        if os.path.exists(EXCEL_FILE):
            return
        
        # Fresh initialization with mock data matching the HTML app
        today = datetime.now().strftime("%Y-%m-%d")
        yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        tomorrow = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")
        in_2_days = (datetime.now() + timedelta(days=2)).strftime("%Y-%m-%d")
        in_3_days = (datetime.now() + timedelta(days=3)).strftime("%Y-%m-%d")
        
        tasks_data = [
            {
                "id": "1",
                "title": "Prepare Q3 Business Report",
                "description": "Compile sales data, analyze performance metrics, and prepare executive summary for Q3.",
                "category": "work",
                "priority": "critical",
                "status": "in-progress",
                "dueDate": today,
                "estimatedMinutes": 240,
                "loggedMinutes": 90,
                "tags": "report, Q3, finance",
                "createdAt": yesterday + " 08:00:00",
                "completedAt": ""
            },
            {
                "id": "2",
                "title": "Team Standup Meeting",
                "description": "Daily team sync to discuss blockers and progress.",
                "category": "meetings",
                "priority": "high",
                "status": "done",
                "dueDate": today,
                "estimatedMinutes": 30,
                "loggedMinutes": 30,
                "tags": "team, daily",
                "createdAt": today + " 07:00:00",
                "completedAt": today + " 09:30:00"
            },
            {
                "id": "3",
                "title": "Complete React Course Module 5",
                "description": "Finish advanced hooks and context API chapter.",
                "category": "learning",
                "priority": "medium",
                "status": "todo",
                "dueDate": in_2_days,
                "estimatedMinutes": 120,
                "loggedMinutes": 0,
                "tags": "react, frontend, course",
                "createdAt": yesterday + " 10:00:00",
                "completedAt": ""
            },
            {
                "id": "4",
                "title": "Review Pull Requests",
                "description": "Review and provide feedback on team PRs before EOD.",
                "category": "work",
                "priority": "high",
                "status": "todo",
                "dueDate": today,
                "estimatedMinutes": 60,
                "loggedMinutes": 0,
                "tags": "code-review, github",
                "createdAt": today + " 08:00:00",
                "completedAt": ""
            },
            {
                "id": "5",
                "title": "Gym Workout",
                "description": "45-minute strength training session.",
                "category": "health",
                "priority": "medium",
                "status": "blocked",
                "dueDate": today,
                "estimatedMinutes": 60,
                "loggedMinutes": 0,
                "tags": "fitness, wellness",
                "createdAt": today + " 06:00:00",
                "completedAt": ""
            },
            {
                "id": "6",
                "title": "Update Project Roadmap",
                "description": "Revise the product roadmap for next quarter based on stakeholder feedback.",
                "category": "project",
                "priority": "high",
                "status": "review",
                "dueDate": tomorrow,
                "estimatedMinutes": 90,
                "loggedMinutes": 75,
                "tags": "roadmap, planning",
                "createdAt": yesterday + " 14:00:00",
                "completedAt": ""
            },
            {
                "id": "7",
                "title": "Monthly Budget Review",
                "description": "Check expenses, categorize transactions, review savings target.",
                "category": "finance",
                "priority": "low",
                "status": "todo",
                "dueDate": in_3_days,
                "estimatedMinutes": 45,
                "loggedMinutes": 0,
                "tags": "budget, personal-finance",
                "createdAt": yesterday + " 20:00:00",
                "completedAt": ""
            }
        ]
        
        time_blocks_data = [
            {
                "id": "tb1",
                "taskId": "2",
                "taskTitle": "Team Standup Meeting",
                "startTime": "09:00",
                "endTime": "09:30",
                "date": today,
                "category": "meetings"
            },
            {
                "id": "tb2",
                "taskId": "1",
                "taskTitle": "Prepare Q3 Business Report",
                "startTime": "10:00",
                "endTime": "11:30",
                "date": today,
                "category": "work"
            },
            {
                "id": "tb3",
                "taskId": "4",
                "taskTitle": "Review Pull Requests",
                "startTime": "14:00",
                "endTime": "15:00",
                "date": today,
                "category": "work"
            },
            {
                "id": "tb4",
                "taskId": "5",
                "taskTitle": "Gym Workout",
                "startTime": "17:30",
                "endTime": "18:30",
                "date": today,
                "category": "health"
            }
        ]
        
        activity_logs_data = [
            {
                "id": "a1",
                "taskId": "1",
                "taskTitle": "Prepare Q3 Business Report",
                "action": "Started task (Logged 90 minutes)",
                "timestamp": today + " 09:00:00",
                "minutesLogged": 90
            },
            {
                "id": "a2",
                "taskId": "2",
                "taskTitle": "Team Standup Meeting",
                "action": "Marked as done (Logged 30 minutes)",
                "timestamp": today + " 09:30:00",
                "minutesLogged": 30
            },
            {
                "id": "a3",
                "taskId": "1",
                "taskTitle": "Prepare Q3 Business Report",
                "action": "Task created",
                "timestamp": yesterday + " 08:00:00",
                "minutesLogged": ""
            },
            {
                "id": "a4",
                "taskId": "2",
                "taskTitle": "Team Standup Meeting",
                "action": "Task created",
                "timestamp": today + " 07:00:00",
                "minutesLogged": ""
            },
            {
                "id": "a5",
                "taskId": "3",
                "taskTitle": "Complete React Course Module 5",
                "action": "Task created",
                "timestamp": yesterday + " 10:00:00",
                "minutesLogged": ""
            },
            {
                "id": "a6",
                "taskId": "4",
                "taskTitle": "Review Pull Requests",
                "action": "Task created",
                "timestamp": today + " 08:00:00",
                "minutesLogged": ""
            },
            {
                "id": "a7",
                "taskId": "5",
                "taskTitle": "Gym Workout",
                "action": "Task created",
                "timestamp": today + " 06:00:00",
                "minutesLogged": ""
            },
            {
                "id": "a8",
                "taskId": "6",
                "taskTitle": "Update Project Roadmap",
                "action": "Task created",
                "timestamp": yesterday + " 14:00:00",
                "minutesLogged": ""
            },
            {
                "id": "a9",
                "taskId": "7",
                "taskTitle": "Monthly Budget Review",
                "action": "Task created",
                "timestamp": yesterday + " 20:00:00",
                "minutesLogged": ""
            }
        ]
        
        comments_data = [
            {
                "id": "c1",
                "taskId": "1",
                "commentText": "Completed sales data compilation. Proceeding to visual analytics.",
                "timestamp": today + " 10:15:00"
            },
            {
                "id": "c2",
                "taskId": "1",
                "commentText": "Drafted performance metrics chart. Ready for review.",
                "timestamp": today + " 11:45:00"
            }
        ]
        
        notes_data = [
            {
                "id": "n1",
                "title": "System Architecture Decisions",
                "content": "Using Streamlit front-end directly backed by a multi-sheet local Excel sheet. This ensures local-first compliance, zero external database latency, and pristine data privacy.\\n\\n*Captured on yesterday at 14:22 PM*",
                "tags": "architecture, system-design",
                "createdAt": yesterday + " 14:22:15",
                "updatedAt": yesterday + " 14:22:15"
            }
        ]
        
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            pd.DataFrame(tasks_data).to_excel(writer, sheet_name="Tasks", index=False)
            pd.DataFrame(time_blocks_data).to_excel(writer, sheet_name="TimeBlocks", index=False)
            pd.DataFrame(activity_logs_data).to_excel(writer, sheet_name="ActivityLogs", index=False)
            pd.DataFrame(comments_data).to_excel(writer, sheet_name="Comments", index=False)
            pd.DataFrame(notes_data).to_excel(writer, sheet_name="Notes", index=False)
            pd.DataFrame(columns=["id", "noteId", "fileName", "filePath", "fileType", "fileSize", "uploadedAt"]).to_excel(writer, sheet_name="Attachments", index=False)

    @classmethod
    def _read_sheet(cls, name):
        cls.initialize_db()
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name=name, dtype={"id": str, "taskId": str, "noteId": str})
            return df.fillna("")
        except Exception:
            return pd.DataFrame()

    @classmethod
    def _write_sheet(cls, name, df):
        all_sheets = {}
        for s in ["Tasks", "TimeBlocks", "ActivityLogs", "Comments", "Notes", "Attachments"]:
            if s == name:
                all_sheets[s] = df
            else:
                all_sheets[s] = cls._read_sheet(s)
        
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            for s, data in all_sheets.items():
                data.to_excel(writer, sheet_name=s, index=False)

    # --- TASKS METHODS ---
    @classmethod
    def get_tasks(cls):
        tasks = cls._read_sheet("Tasks").to_dict(orient="records")
        for t in tasks:
            t["estimatedMinutes"] = int(t["estimatedMinutes"]) if t["estimatedMinutes"] != "" else 0
            t["loggedMinutes"] = int(t["loggedMinutes"]) if t["loggedMinutes"] != "" else 0
            t["tags"] = [x.strip() for x in str(t["tags"]).split(",") if x.strip()] if t["tags"] != "" else []
        return tasks

    @classmethod
    def get_time_blocks(cls):
        return cls._read_sheet("TimeBlocks").to_dict(orient="records")

    @classmethod
    def get_activity_logs(cls):
        logs = cls._read_sheet("ActivityLogs").to_dict(orient="records")
        for l in logs:
            l["minutesLogged"] = int(l["minutesLogged"]) if l["minutesLogged"] != "" else None
        return logs

    @classmethod
    def get_comments(cls, task_id):
        comments = cls._read_sheet("Comments")
        task_id = str(task_id)
        filtered = comments[comments["taskId"] == task_id]
        return filtered.sort_values(by="timestamp", ascending=False).to_dict(orient="records")

    @classmethod
    def add_comment(cls, task_id, comment_text):
        task_id = str(task_id)
        df_tasks = cls._read_sheet("Tasks")
        title = df_tasks.loc[df_tasks["id"] == task_id, "title"].values[0] if task_id in df_tasks["id"].values else "Unknown Task"
        
        df = cls._read_sheet("Comments")
        new_id = str(uuid.uuid4())[:8]
        new_row = {
            "id": new_id,
            "taskId": task_id,
            "commentText": comment_text,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        cls._write_sheet("Comments", df)
        cls.add_activity_log(task_id, title, f"Added comment: '{comment_text[:30]}...'")

    @classmethod
    def add_task(cls, title, description, category, priority, status, dueDate, estimatedMinutes, tags):
        df = cls._read_sheet("Tasks")
        new_id = str(uuid.uuid4())[:8]
        new_row = {
            "id": new_id,
            "title": title,
            "description": description,
            "category": category,
            "priority": priority,
            "status": status,
            "dueDate": dueDate,
            "estimatedMinutes": int(estimatedMinutes),
            "loggedMinutes": 0,
            "tags": tags,
            "createdAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "completedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S") if status == "done" else ""
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        cls._write_sheet("Tasks", df)
        cls.add_activity_log(new_id, title, "Task created")
        return new_id

    @classmethod
    def update_task(cls, task_id, updates):
        df = cls._read_sheet("Tasks")
        task_id = str(task_id)
        if task_id not in df["id"].values:
            return
        
        old_row = df[df["id"] == task_id].iloc[0]
        old_status = old_row["status"]
        old_title = old_row["title"]
        
        for k, v in updates.items():
            if k == "tags":
                v = ", ".join(v) if isinstance(v, list) else v
            df.loc[df["id"] == task_id, k] = v
            
        if "status" in updates and updates["status"] == "done" and old_status != "done":
            df.loc[df["id"] == task_id, "completedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cls.add_activity_log(task_id, old_title, "Marked as done")
        elif "status" in updates and updates["status"] != "done" and old_status == "done":
            df.loc[df["id"] == task_id, "completedAt"] = ""
            cls.add_activity_log(task_id, old_title, f"Status updated to {updates['status']}")
        elif "status" in updates and updates["status"] != old_status:
            cls.add_activity_log(task_id, old_title, f"Status updated to {updates['status']}")
            
        cls._write_sheet("Tasks", df)

    @classmethod
    def delete_task(cls, task_id):
        task_id = str(task_id)
        df_tasks = cls._read_sheet("Tasks")
        title = df_tasks.loc[df_tasks["id"] == task_id, "title"].values[0] if task_id in df_tasks["id"].values else ""
        df_tasks = df_tasks[df_tasks["id"] != task_id]
        cls._write_sheet("Tasks", df_tasks)
        
        df_blocks = cls._read_sheet("TimeBlocks")
        df_blocks = df_blocks[df_blocks["taskId"] != task_id]
        cls._write_sheet("TimeBlocks", df_blocks)
        
        df_comments = cls._read_sheet("Comments")
        df_comments = df_comments[df_comments["taskId"] != task_id]
        cls._write_sheet("Comments", df_comments)
        
        if title:
            cls.add_activity_log(task_id, title, "Task deleted")
            
    @classmethod
    def delete_multiple_tasks(cls, task_ids):
        task_ids_str = [str(tid) for tid in task_ids]
        
        # Tasks
        df_tasks = cls._read_sheet("Tasks")
        df_tasks = df_tasks[~df_tasks["id"].isin(task_ids_str)]
        cls._write_sheet("Tasks", df_tasks)
        
        # TimeBlocks
        df_blocks = cls._read_sheet("TimeBlocks")
        df_blocks = df_blocks[~df_blocks["taskId"].isin(task_ids_str)]
        cls._write_sheet("TimeBlocks", df_blocks)
        
        # Comments
        df_comments = cls._read_sheet("Comments")
        df_comments = df_comments[~df_comments["taskId"].isin(task_ids_str)]
        cls._write_sheet("Comments", df_comments)
        
        # Log bulk action
        cls.add_activity_log("None", "Bulk Action", f"Deleted {len(task_ids_str)} tasks")

    @classmethod
    def log_time(cls, task_id, minutes):
        task_id = str(task_id)
        df_tasks = cls._read_sheet("Tasks")
        if task_id not in df_tasks["id"].values:
            return
        
        old_row = df_tasks[df_tasks["id"] == task_id].iloc[0]
        old_title = old_row["title"]
        old_logged = int(old_row["loggedMinutes"]) if old_row["loggedMinutes"] != "" else 0
        
        df_tasks.loc[df_tasks["id"] == task_id, "loggedMinutes"] = old_logged + int(minutes)
        cls._write_sheet("Tasks", df_tasks)
        cls.add_activity_log(task_id, old_title, f"Logged {minutes} minutes", minutes)

    @classmethod
    def add_time_block(cls, task_id, task_title, start_time, end_time, date, category):
        df = cls._read_sheet("TimeBlocks")
        new_id = str(uuid.uuid4())[:8]
        new_row = {
            "id": new_id,
            "taskId": str(task_id),
            "taskTitle": task_title,
            "startTime": start_time,
            "endTime": end_time,
            "date": date,
            "category": category
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        cls._write_sheet("TimeBlocks", df)
        cls.add_activity_log(task_id, task_title, f"Scheduled time block {start_time} - {end_time}")

    @classmethod
    def remove_time_block(cls, block_id):
        block_id = str(block_id)
        df = cls._read_sheet("TimeBlocks")
        if block_id not in df["id"].values:
            return
        row = df[df["id"] != block_id].iloc[0] if len(df[df["id"] == block_id]) > 0 else None
        df = df[df["id"] != block_id]
        cls._write_sheet("TimeBlocks", df)
        if row is not None:
            cls.add_activity_log(row["taskId"], row["taskTitle"], f"Removed scheduled time block {row['startTime']} - {row['endTime']}")

    @classmethod
    def add_activity_log(cls, task_id, task_title, action, minutes_logged=""):
        df = cls._read_sheet("ActivityLogs")
        new_row = {
            "id": str(uuid.uuid4())[:8],
            "taskId": str(task_id),
            "taskTitle": task_title,
            "action": action,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "minutesLogged": minutes_logged
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        cls._write_sheet("ActivityLogs", df)

    # --- DAILY NOTES & ATTACHMENTS METHODS ---
    @classmethod
    def get_notes(cls):
        notes = cls._read_sheet("Notes").to_dict(orient="records")
        for n in notes:
            n["tags"] = [x.strip() for x in str(n["tags"]).split(",") if x.strip()] if n["tags"] != "" else []
        return notes

    @classmethod
    def get_note_attachments(cls, note_id):
        df_attach = cls._read_sheet("Attachments")
        filtered = df_attach[df_attach["noteId"] == str(note_id)]
        return filtered.to_dict(orient="records")

    @classmethod
    def add_note(cls, title, content, tags, uploaded_files=None):
        df_notes = cls._read_sheet("Notes")
        note_id = str(uuid.uuid4())[:8]
        now_time = datetime.now()
        timestamp_str = now_time.strftime("%Y-%m-%d %H:%M:%S")
        
        # Format human readable append block
        weekday_map = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        day_name = weekday_map[now_time.weekday()]
        time_display = now_time.strftime("%I:%M %p")
        date_display = now_time.strftime("%B %d, %Y")
        
        content_with_stamp = f"{content}\\n\\n*Captured on {day_name}, {date_display} at {time_display}*"
        
        new_row = {
            "id": note_id,
            "title": title,
            "content": content_with_stamp,
            "tags": tags,
            "createdAt": timestamp_str,
            "updatedAt": timestamp_str
        }
        df_notes = pd.concat([df_notes, pd.DataFrame([new_row])], ignore_index=True)
        cls._write_sheet("Notes", df_notes)
        
        # Save attached files physically and logically
        if uploaded_files:
            df_attach = cls._read_sheet("Attachments")
            for f in uploaded_files:
                file_uuid = str(uuid.uuid4())[:8]
                ext = os.path.splitext(f.name)[1]
                local_filename = f"{file_uuid}{ext}"
                local_filepath = os.path.join(ATTACHMENT_DIR, local_filename)
                
                # Write file physically
                with open(local_filepath, "wb") as buffer:
                    buffer.write(f.getbuffer())
                
                attach_row = {
                    "id": file_uuid,
                    "noteId": note_id,
                    "fileName": f.name,
                    "filePath": local_filepath,
                    "fileType": f.type,
                    "fileSize": len(f.getbuffer()),
                    "uploadedAt": timestamp_str
                }
                df_attach = pd.concat([df_attach, pd.DataFrame([attach_row])], ignore_index=True)
            
            cls._write_sheet("Attachments", df_attach)
            
        cls.add_activity_log("None", f"Daily Note: {title}", f"Created note with {len(uploaded_files) if uploaded_files else 0} attachments")
        return note_id

    @classmethod
    def delete_note(cls, note_id):
        note_id = str(note_id)
        
        # Tasks reference check
        df_notes = cls._read_sheet("Notes")
        title = df_notes.loc[df_notes["id"] == note_id, "title"].values[0] if note_id in df_notes["id"].values else ""
        df_notes = df_notes[df_notes["id"] != note_id]
        cls._write_sheet("Notes", df_notes)
        
        # Purge attachments physically and logically
        df_attach = cls._read_sheet("Attachments")
        associated_files = df_attach[df_attach["noteId"] == note_id]
        
        for idx, row in associated_files.iterrows():
            f_path = row["filePath"]
            if os.path.exists(f_path):
                try:
                    os.remove(f_path)
                except Exception:
                    pass
                    
        df_attach = df_attach[df_attach["noteId"] != note_id]
        cls._write_sheet("Attachments", df_attach)
        
        if title:
            cls.add_activity_log("None", f"Daily Note: {title}", "Deleted note and purged physical attachments")

    # --- CALENDAR GENERATOR ---
    @staticmethod
    def generate_ics(task_id, title, description, due_date_str):
        try:
            due_dt = datetime.strptime(due_date_str, "%Y-%m-%d")
            dtstart = f"DTSTART;VALUE=DATE:{due_dt.strftime('%Y%m%d')}"
        except:
            dtstart = f"DTSTART;VALUE=DATE:{datetime.now().strftime('%Y%m%d')}"
            
        dtstamp = datetime.now().strftime('%Y%m%dT%H%M%SZ')
        clean_desc = str(description).replace('\n', ' ').replace(',', r'\,').replace(';', r'\;')
        clean_title = str(title).replace(',', r'\,').replace(';', r'\;')
        
        ics_lines = [
            "BEGIN:VCALENDAR",
            "VERSION:2.0",
            "PRODID:-//Gemini Core//WorkManager//EN",
            "CALSCALE:GREGORIAN",
            "BEGIN:VEVENT",
            f"UID:{task_id}@workmanager.local",
            f"DTSTAMP:{dtstamp}",
            dtstart,
            f"SUMMARY:{clean_title}",
            f"DESCRIPTION:{clean_desc}",
            "PRIORITY:5",
            "CLASS:PUBLIC",
            "END:VEVENT",
            "END:VCALENDAR"
        ]
        return "\n".join(ics_lines)
