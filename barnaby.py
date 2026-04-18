import json
import os
import platform
import random
import shutil
import subprocess
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog
from urllib.parse import quote_plus

try:
    import psutil
except Exception:
    psutil = None

try:
    import pyttsx3
except Exception:
    pyttsx3 = None

APP_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "BarnabyDesktopHelper"
PROFILE_FILE = APP_DIR / "profile.json"
TASKS_FILE = APP_DIR / "tasks.json"
NOTES_FILE = APP_DIR / "notes.json"
QUARANTINE_DIR = APP_DIR / "Quarantine"


class BarnabyVoice:
    def __init__(self):
        self.engine = None
        self.lock = threading.Lock()
        if pyttsx3:
            try:
                self.engine = pyttsx3.init()
                self.engine.setProperty("rate", 165)
                self.engine.setProperty("volume", 0.95)
            except Exception:
                self.engine = None

    def say(self, text):
        threading.Thread(target=self._speak, args=(text,), daemon=True).start()

    def _speak(self, text):
        with self.lock:
            if self.engine:
                try:
                    self.engine.say(text)
                    self.engine.runAndWait()
                    return
                except Exception:
                    pass
            if platform.system() == "Windows":
                try:
                    safe_text = text.replace("'", "''")
                    subprocess.run(
                        [
                            "powershell",
                            "-NoProfile",
                            "-ExecutionPolicy",
                            "Bypass",
                            "-Command",
                            f"Add-Type -AssemblyName System.Speech; $speak=New-Object System.Speech.Synthesis.SpeechSynthesizer; $speak.Rate=0; $speak.Volume=95; $speak.Speak('{safe_text}')",
                        ],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        check=False,
                    )
                except Exception:
                    pass


class BarnabyApp:
    def __init__(self):
        APP_DIR.mkdir(parents=True, exist_ok=True)
        self.root = tk.Tk()
        self.root.title("Barnaby the Octopus")
        self.root.geometry("420x650+120+120")
        self.root.minsize(380, 500)
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.transparent = "#ff00ff"
        self.root.configure(bg=self.transparent)
        try:
            self.root.wm_attributes("-transparentcolor", self.transparent)
        except Exception:
            pass
        self.voice = BarnabyVoice()
        self.drag_x = 0
        self.drag_y = 0
        self.closed = False
        self.name = self.load_name()
        self.tasks = self.load_list(TASKS_FILE)
        self.notes = self.load_list(NOTES_FILE)
        self.walk_enabled = True
        self.silent_mode_enabled = False
        self.walk_phase = 0
        self.last_scan_count = 0
        self.startup_scan_done = False
        self.build_pet()
        self.root.bind("<ButtonPress-1>", self.start_drag)
        self.root.bind("<B1-Motion>", self.drag)
        self.root.bind("<Button-3>", self.toggle_tools)
        self.root.bind("<Escape>", lambda event: self.close())
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.safe_after(700, self.introduce)
        self.safe_after(9000, self.random_walk)
        self.safe_after(14000, self.startup_safety_scan)
        self.safe_after(30000, self.background_safety_check)
        self.safe_after(45000, self.helpful_nudge)

    def build_pet(self):
        self.frame = tk.Frame(self.root, bg=self.transparent)
        self.frame.pack(fill="both", expand=True)
        self.bubble = tk.Label(
            self.frame,
            text="",
            bg="#fff8df",
            fg="#241744",
            font=("Segoe UI", 11),
            wraplength=310,
            justify="left",
            padx=14,
            pady=10,
            bd=2,
            relief="ridge",
        )
        self.bubble.pack(pady=(6, 3))
        self.canvas = tk.Canvas(self.frame, width=310, height=235, bg=self.transparent, highlightthickness=0)
        self.canvas.pack()
        self.draw_octopus()
        self.tools = tk.Frame(self.frame, bg="#f2ecff", bd=2, relief="ridge")
        tk.Button(self.tools, text="Barnaby.ai", command=self.open_barnaby_ai, width=18).grid(row=0, column=0, padx=5, pady=4)
        tk.Button(self.tools, text="Writing helper", command=self.open_writing_helper, width=18).grid(row=0, column=1, padx=5, pady=4)
        tk.Button(self.tools, text="My tasks", command=self.show_tasks, width=18).grid(row=1, column=0, padx=5, pady=4)
        tk.Button(self.tools, text="Running programs", command=self.show_programs, width=18).grid(row=1, column=1, padx=5, pady=4)
        tk.Button(self.tools, text="Check Outlook", command=self.check_outlook, width=18).grid(row=2, column=0, padx=5, pady=4)
        tk.Button(self.tools, text="Suspicious emails", command=self.check_suspicious_emails, width=18).grid(row=2, column=1, padx=5, pady=4)
        tk.Button(self.tools, text="Search internet", command=self.search_internet, width=18).grid(row=3, column=0, padx=5, pady=4)
        tk.Button(self.tools, text="Web ad shield", command=self.open_web_shield, width=18).grid(row=3, column=1, padx=5, pady=4)
        tk.Button(self.tools, text="Organize files", command=self.organize_files, width=18).grid(row=4, column=0, padx=5, pady=4)
        tk.Button(self.tools, text="Clean junk files", command=self.clean_junk_files, width=18).grid(row=4, column=1, padx=5, pady=4)
        tk.Button(self.tools, text="Safety scan", command=self.scan_suspicious_files, width=18).grid(row=5, column=0, padx=5, pady=4)
        tk.Button(self.tools, text="Workspace silent", command=self.open_workspace_silent_mode, width=18).grid(row=5, column=1, padx=5, pady=4)
        tk.Button(self.tools, text="Quick notes", command=self.show_notes, width=18).grid(row=6, column=0, padx=5, pady=4)
        tk.Button(self.tools, text="System info", command=self.show_system_info, width=18).grid(row=6, column=1, padx=5, pady=4)
        tk.Button(self.tools, text="Open data folder", command=self.open_data_folder, width=18).grid(row=7, column=0, padx=5, pady=4)
        tk.Button(self.tools, text="Stop walking", command=self.toggle_walking, width=18).grid(row=7, column=1, padx=5, pady=4)
        tk.Button(self.tools, text="Hide Barnaby", command=self.close, width=18).grid(row=8, column=0, columnspan=2, padx=5, pady=4)
        self.tools_visible = False

    def draw_octopus(self):
        c = self.canvas
        c.delete("all")
        body = "#6f55d9"
        mid = "#8a73f2"
        shadow = "#3f2d91"
        suction = "#f1dcff"
        eye = "#18102f"
        offset = 7 if self.walk_phase % 2 else -7
        c.create_oval(86, 22, 224, 158, fill=body, outline=shadow, width=4)
        c.create_oval(104, 38, 206, 150, fill=mid, outline="")
        c.create_oval(118, 52, 146, 84, fill="white", outline=shadow, width=2)
        c.create_oval(164, 52, 192, 84, fill="white", outline=shadow, width=2)
        c.create_oval(129, 63, 140, 76, fill=eye, outline="")
        c.create_oval(175, 63, 186, 76, fill=eye, outline="")
        c.create_oval(133, 64, 136, 67, fill="white", outline="")
        c.create_oval(179, 64, 182, 67, fill="white", outline="")
        c.create_arc(133, 87, 178, 118, start=200, extent=140, style="arc", outline=eye, width=3)
        c.create_oval(109, 43, 126, 56, fill="#b9adff", outline="")
        tentacles = [
            (82, 140, 42, 208, 78, 194),
            (104, 151, 78, 225, 117, 199),
            (128, 157, 123, 229, 145, 202),
            (154, 160, 158, 230, 171, 202),
            (179, 157, 194, 228, 193, 201),
            (204, 150, 235, 224, 219, 197),
            (226, 140, 271, 210, 235, 195),
            (94, 129, 59, 184, 100, 177),
        ]
        for index, (x1, y1, x2, y2, x3, y3) in enumerate(tentacles):
            wave = offset if index % 2 == 0 else -offset
            c.create_line(x1, y1, x2 + wave, y2, x3 - wave, y3, smooth=True, fill=shadow, width=17, capstyle="round")
            c.create_line(x1, y1, x2 + wave, y2, x3 - wave, y3, smooth=True, fill=body, width=13, capstyle="round")
            for dot in range(2):
                sx = int((x2 + x3) / 2) + dot * 9 - wave // 2
                sy = int((y2 + y3) / 2) + dot * 6
                c.create_oval(sx - 3, sy - 3, sx + 3, sy + 3, fill=suction, outline="")
        c.create_text(155, 214, text="Barnaby", fill=eye, font=("Segoe UI", 15, "bold"))

    def safe_after(self, delay, callback):
        if not self.closed:
            self.root.after(delay, callback)

    def start_drag(self, event):
        self.drag_x = event.x
        self.drag_y = event.y

    def drag(self, event):
        x = self.root.winfo_x() + event.x - self.drag_x
        y = self.root.winfo_y() + event.y - self.drag_y
        self.root.geometry(f"+{x}+{y}")

    def toggle_tools(self, event=None):
        if self.tools_visible:
            self.tools.pack_forget()
            self.tools_visible = False
        else:
            self.tools.pack(pady=(2, 8))
            self.tools_visible = True

    def say(self, text):
        if self.closed:
            return
        self.bubble.config(text=text)
        if not self.silent_mode_enabled:
            self.voice.say(text)

    def introduce(self):
        if self.name:
            self.say(f"Hello again, {self.name}! Barnaby is here to help. Right-click me for tools.")
            self.safe_after(2200, self.capability_intro)
            return
        self.say("hello! i dont think we have been properly introduced")
        self.safe_after(1200, self.ask_name)

    def ask_name(self):
        name = simpledialog.askstring("Barnaby the Octopus", "What is your name?", parent=self.root)
        if name and name.strip():
            self.name = name.strip()
            self.save_json(PROFILE_FILE, {"name": self.name})
            self.say(f"Nice to meet you! {self.name}")
            self.safe_after(1800, self.capability_intro)
        else:
            self.say("That is okay. I will be Barnaby, your desktop helper.")
            self.safe_after(1800, self.capability_intro)

    def capability_intro(self):
        intro = (
            "I can remember tasks and notes, show running programs, check Outlook, search the internet, "
            "organize a folder for you, clean junk files, block many web ad and tracker domains, help with writing, and open Barnaby.ai. "
            "I can also run workspace silent mode, show system info, walk around your screen, scan suspicious emails, and run safety checks for suspicious files and processes. "
            "I can look for names like MEMZ, trojans, RATs, stealers, and risky scripts, and I can launch a Windows Security quick scan when Windows supports it. "
            "If I find something that looks risky, I can warn you and ask before moving files to quarantine. "
            "Does that sound great?"
        )
        self.say(intro)
        self.safe_after(8500, self.ask_if_great)

    def ask_if_great(self):
        if self.closed:
            return
        choice = messagebox.askyesno("Barnaby the Octopus", "Does that sound great?", parent=self.root)
        if choice:
            self.say("Wonderful. Right-click me whenever you need help.")
        else:
            self.say("No worries. I will stay nearby and help only when you ask.")

    def load_name(self):
        data = self.load_json(PROFILE_FILE, {})
        if isinstance(data, dict):
            value = data.get("name", "")
            if isinstance(value, str):
                return value
        return ""

    def load_list(self, path):
        data = self.load_json(path, [])
        if isinstance(data, list):
            return [str(item) for item in data if str(item).strip()]
        return []

    def load_json(self, path, fallback):
        try:
            if path.exists():
                return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return fallback
        return fallback

    def save_json(self, path, data):
        try:
            APP_DIR.mkdir(parents=True, exist_ok=True)
            path.write_text(json.dumps(data, indent=2), encoding="utf-8")
            return True
        except Exception as error:
            messagebox.showwarning("Barnaby save problem", f"Barnaby could not save this yet.\n\n{error}")
            return False

    def show_tasks(self):
        win = self.make_window("Barnaby Tasks", "430x430")
        tk.Label(win, text="Barnaby can remember tasks for you.", font=("Segoe UI", 12, "bold")).pack(pady=8)
        listbox = tk.Listbox(win, font=("Segoe UI", 10), height=12)
        listbox.pack(fill="both", expand=True, padx=12, pady=6)
        for task in self.tasks:
            listbox.insert("end", task)
        entry = tk.Entry(win, font=("Segoe UI", 11))
        entry.pack(fill="x", padx=12, pady=6)

        def add_task():
            value = entry.get().strip()
            if value:
                self.tasks.append(value)
                if self.save_json(TASKS_FILE, self.tasks):
                    listbox.insert("end", value)
                    entry.delete(0, "end")
                    self.say("I added that to your tasks.")

        def finish_task():
            selected = listbox.curselection()
            if selected:
                index = selected[0]
                done = self.tasks.pop(index)
                if self.save_json(TASKS_FILE, self.tasks):
                    listbox.delete(index)
                    self.say(f"Done with {done}. Nice work!")

        buttons = tk.Frame(win)
        buttons.pack(pady=8)
        tk.Button(buttons, text="Add task", command=add_task, width=14).grid(row=0, column=0, padx=5)
        tk.Button(buttons, text="Mark done", command=finish_task, width=14).grid(row=0, column=1, padx=5)

    def show_notes(self):
        win = self.make_window("Barnaby Notes", "460x440")
        tk.Label(win, text="Quick notes Barnaby can keep for you.", font=("Segoe UI", 12, "bold")).pack(pady=8)
        listbox = tk.Listbox(win, font=("Segoe UI", 10), height=10)
        listbox.pack(fill="both", expand=True, padx=12, pady=6)
        for note in self.notes:
            listbox.insert("end", note)
        entry = tk.Entry(win, font=("Segoe UI", 11))
        entry.pack(fill="x", padx=12, pady=6)

        def add_note():
            value = entry.get().strip()
            if value:
                self.notes.append(value)
                if self.save_json(NOTES_FILE, self.notes):
                    listbox.insert("end", value)
                    entry.delete(0, "end")
                    self.say("I saved that note.")

        def remove_note():
            selected = listbox.curselection()
            if selected:
                index = selected[0]
                self.notes.pop(index)
                if self.save_json(NOTES_FILE, self.notes):
                    listbox.delete(index)
                    self.say("I removed that note.")

        buttons = tk.Frame(win)
        buttons.pack(pady=8)
        tk.Button(buttons, text="Add note", command=add_note, width=14).grid(row=0, column=0, padx=5)
        tk.Button(buttons, text="Remove note", command=remove_note, width=14).grid(row=0, column=1, padx=5)

    def show_programs(self):
        names = []
        if psutil:
            try:
                for proc in psutil.process_iter(["name"]):
                    name = proc.info.get("name")
                    if name and name not in names:
                        names.append(name)
            except Exception:
                names = []
        if not names and platform.system() == "Windows":
            try:
                output = subprocess.check_output("tasklist", shell=True, text=True, errors="ignore")
                names = [line.split()[0] for line in output.splitlines()[3:] if line.strip()]
            except Exception:
                names = []
        names = sorted(set(names))[:100]
        win = self.make_window("Running Programs", "390x470")
        tk.Label(win, text="Programs running right now", font=("Segoe UI", 12, "bold")).pack(pady=8)
        box = tk.Listbox(win, font=("Segoe UI", 10))
        box.pack(fill="both", expand=True, padx=12, pady=8)
        if names:
            for name in names:
                box.insert("end", name)
            self.say(f"I found {len(names)} running programs.")
        else:
            box.insert("end", "Barnaby could not read the running program list.")
            self.say("I could not read the running program list yet.")

    def check_outlook(self):
        if platform.system() != "Windows":
            self.say("Outlook checking works when Barnaby is running on Windows with Outlook installed.")
            return
        try:
            import win32com.client

            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.GetDefaultFolder(6)
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)
            unread = int(getattr(inbox, "UnReadItemCount", 0))
            win = self.make_window("Outlook Inbox", "640x430")
            tk.Label(win, text=f"Unread Outlook messages: {unread}", font=("Segoe UI", 12, "bold")).pack(pady=8)
            box = tk.Listbox(win, font=("Segoe UI", 10))
            box.pack(fill="both", expand=True, padx=12, pady=8)
            count = min(20, int(getattr(messages, "Count", 0)))
            if count == 0:
                box.insert("end", "No Outlook messages found in the inbox.")
            for index in range(1, count + 1):
                item = messages.Item(index)
                sender = str(getattr(item, "SenderName", "Unknown") or "Unknown")
                subject = str(getattr(item, "Subject", "No subject") or "No subject")
                box.insert("end", f"{sender}: {subject}")
            self.say(f"You have {unread} unread Outlook messages.")
        except Exception as error:
            messagebox.showwarning(
                "Barnaby Outlook",
                f"Barnaby could not read Outlook yet. Make sure Outlook is installed, open, and signed in.\n\n{error}",
            )
            self.say("I could not reach Outlook yet, but I can try again later.")

    def check_suspicious_emails(self):
        if platform.system() != "Windows":
            self.say("Suspicious email checking works on Windows with Outlook installed.")
            return
        try:
            import win32com.client

            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.GetDefaultFolder(6)
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)
            suspicious = []
            phishing_words = [
                "urgent",
                "verify",
                "password",
                "account locked",
                "gift card",
                "wire transfer",
                "payment failed",
                "invoice attached",
                "security alert",
                "click here",
                "reset your",
                "unusual sign",
                "crypto",
                "lottery",
                "prize",
            ]
            risky_extensions = (".exe", ".scr", ".bat", ".cmd", ".ps1", ".vbs", ".js", ".jse", ".wsf", ".hta", ".jar", ".zip", ".rar", ".7z")
            count = min(50, int(getattr(messages, "Count", 0)))
            for index in range(1, count + 1):
                item = messages.Item(index)
                sender = str(getattr(item, "SenderEmailAddress", "") or getattr(item, "SenderName", "Unknown") or "Unknown")
                subject = str(getattr(item, "Subject", "") or "")
                body = str(getattr(item, "Body", "") or "")[:1200]
                haystack = f"{sender} {subject} {body}".lower()
                reasons = []
                if any(word in haystack for word in phishing_words):
                    reasons.append("phishing words")
                if getattr(item, "Attachments", None):
                    for attachment_index in range(1, int(item.Attachments.Count) + 1):
                        filename = str(item.Attachments.Item(attachment_index).FileName or "").lower()
                        if filename.endswith(risky_extensions):
                            reasons.append(f"risky attachment: {filename}")
                if reasons:
                    suspicious.append(f"{sender}: {subject} ({', '.join(sorted(set(reasons)))})")
            win = self.make_window("Suspicious Outlook Emails", "760x460")
            tk.Label(win, text="Suspicious email check", font=("Segoe UI", 12, "bold")).pack(pady=8)
            tk.Label(win, text="Barnaby checks recent Outlook inbox messages for phishing words and risky attachments.", wraplength=720, justify="left").pack(padx=12)
            box = tk.Listbox(win, font=("Segoe UI", 9))
            box.pack(fill="both", expand=True, padx=12, pady=8)
            if suspicious:
                for email in suspicious[:100]:
                    box.insert("end", email)
                self.say(f"I found {len(suspicious)} suspicious email{'s' if len(suspicious) != 1 else ''}. Please review before clicking links or attachments.")
            else:
                box.insert("end", "No obvious suspicious emails found in the recent Outlook inbox messages Barnaby checked.")
                self.say("I did not find obvious suspicious emails in the recent Outlook inbox messages.")
        except Exception as error:
            messagebox.showwarning("Barnaby Email Safety", f"Barnaby could not check suspicious Outlook emails yet.\n\n{error}", parent=self.root)
            self.say("I could not check suspicious Outlook emails yet.")

    def search_internet(self):
        query = simpledialog.askstring("Barnaby Internet Search", "What should Barnaby search for?", parent=self.root)
        if not query or not query.strip():
            self.say("No search problem. Ask me when you need the internet.")
            return
        text = query.strip()
        url = "https://www.bing.com/search?q=" + quote_plus(text)
        try:
            webbrowser.open(url)
            self.say(f"I am searching the internet for {text}.")
        except Exception as error:
            messagebox.showwarning("Barnaby Search", f"Barnaby could not open your browser.\n\n{error}")
            self.say("I could not open the browser yet.")

    def open_barnaby_ai(self):
        win = self.make_window("Barnaby.ai", "620x520")
        tk.Label(win, text="Ask Barnaby.ai anything...", font=("Segoe UI", 13, "bold")).pack(pady=8)
        prompt = tk.Text(win, font=("Segoe UI", 10), height=6, wrap="word")
        prompt.pack(fill="x", padx=12, pady=6)
        answer = tk.Text(win, font=("Segoe UI", 10), wrap="word")
        answer.pack(fill="both", expand=True, padx=12, pady=6)
        answer.insert("1.0", "Hi, I am Barnaby.ai. I can help with work ideas, writing, PC cleanup, safety, and web searching. Type a question and press Ask.")

        def ask():
            question = prompt.get("1.0", "end").strip()
            if not question:
                self.say("Ask Barnaby dot AI anything first.")
                return
            response = self.barnaby_ai_answer(question)
            answer.delete("1.0", "end")
            answer.insert("1.0", response)
            self.say("Barnaby dot AI answered.")

        def search_web():
            question = prompt.get("1.0", "end").strip()
            if question:
                webbrowser.open("https://www.bing.com/search?q=" + quote_plus(question))
                self.say("I opened a web search for that.")

        buttons = tk.Frame(win)
        buttons.pack(pady=8)
        tk.Button(buttons, text="Ask", command=ask, width=14).grid(row=0, column=0, padx=5)
        tk.Button(buttons, text="Search web", command=search_web, width=14).grid(row=0, column=1, padx=5)

    def barnaby_ai_answer(self, question):
        text = question.lower()
        if any(word in text for word in ["write", "essay", "email", "letter", "paragraph"]):
            return "Barnaby.ai writing help:\n\n1. Start with the main point.\n2. Add 2 or 3 supporting details.\n3. Use short clear sentences.\n4. End with what you want the reader to do or remember.\n\nIf you paste your draft into Writing helper, I can suggest a cleaner version."
        if any(word in text for word in ["clean", "junk", "slow", "storage", "space"]):
            return "Barnaby.ai PC cleanup plan:\n\nUse Clean junk files to remove safe temporary files, then use Running programs or Workspace silent mode to close apps you do not need. Do not delete personal files unless you recognize them."
        if any(word in text for word in ["virus", "trojan", "memz", "rat", "malware", "suspicious"]):
            return "Barnaby.ai safety plan:\n\nRun Safety scan, then run Windows Security quick scan. If Barnaby finds a suspicious file, review it before quarantine. If Windows Security reports a threat, follow Windows Security first."
        if any(word in text for word in ["ad", "tracker", "youtube", "yt", "privacy"]):
            return "Barnaby.ai web shield note:\n\nUse Web ad shield to block many known ad and tracker domains. YouTube changes often, so no desktop helper can promise every YouTube ad will disappear, but Barnaby will block common tracking and ad domains where Windows allows it."
        if any(word in text for word in ["work", "focus", "workspace", "silent", "quiet"]):
            return "Barnaby.ai focus plan:\n\nOpen Workspace silent mode, mute Windows sound if needed, turn off Barnaby walking, and close distracting apps you select. Keep only the apps you need for work."
        return "Barnaby.ai answer:\n\nI can help with PC safety, cleaning junk files, writing, web searches, Outlook checks, tasks, notes, and focus mode. For live facts, press Search web so I can open a browser search for you."

    def open_writing_helper(self):
        win = self.make_window("Barnaby Writing Helper", "680x560")
        tk.Label(win, text="Paste your writing and Barnaby will suggest improvements.", font=("Segoe UI", 12, "bold")).pack(pady=8)
        draft = tk.Text(win, font=("Segoe UI", 10), height=12, wrap="word")
        draft.pack(fill="both", expand=True, padx=12, pady=6)
        result = tk.Text(win, font=("Segoe UI", 10), height=10, wrap="word")
        result.pack(fill="both", expand=True, padx=12, pady=6)

        def suggest():
            text = draft.get("1.0", "end").strip()
            if not text:
                self.say("Paste some writing first.")
                return
            improved = self.suggest_writing(text)
            result.delete("1.0", "end")
            result.insert("1.0", improved)
            self.say("I made writing suggestions.")

        tk.Button(win, text="Suggest improvements", command=suggest, width=22).pack(pady=8)

    def suggest_writing(self, text):
        cleaned = " ".join(text.split())
        replacements = {
            " alot ": " a lot ",
            " u ": " you ",
            " ur ": " your ",
            " i ": " I ",
            " im ": " I am ",
            " dont ": " do not ",
            " cant ": " cannot ",
            " wont ": " will not ",
        }
        improved = f" {cleaned} "
        for old, new in replacements.items():
            improved = improved.replace(old, new)
        improved = improved.strip()
        if improved and improved[-1] not in ".!?":
            improved += "."
        tips = [
            "Keep the first sentence focused on the main idea.",
            "Break long paragraphs into smaller pieces.",
            "Use specific examples when you can.",
            "Remove repeated words or phrases.",
        ]
        return "Suggested cleaner version:\n\n" + improved + "\n\nBarnaby tips:\n- " + "\n- ".join(tips)

    def open_web_shield(self):
        win = self.make_window("Barnaby Web Ad Shield", "700x470")
        tk.Label(win, text="Aggressive ad and tracker shield", font=("Segoe UI", 12, "bold")).pack(pady=8)
        message = (
            "Barnaby can add known ad and tracker domains to the Windows hosts file. This can block many trackers and some ads across browsers. "
            "It may not remove every YouTube ad because YouTube serves ads in changing ways, but it blocks common ad/tracker domains aggressively. "
            "Windows may require you to run Barnaby as Administrator for this to work."
        )
        tk.Label(win, text=message, wraplength=660, justify="left").pack(padx=12, pady=6)
        box = tk.Listbox(win, font=("Segoe UI", 9))
        box.pack(fill="both", expand=True, padx=12, pady=8)
        for domain in self.ad_tracker_domains():
            box.insert("end", domain)

        buttons = tk.Frame(win)
        buttons.pack(pady=8)
        tk.Button(buttons, text="Enable shield", command=self.enable_web_shield, width=18).grid(row=0, column=0, padx=5)
        tk.Button(buttons, text="Disable shield", command=self.disable_web_shield, width=18).grid(row=0, column=1, padx=5)
        tk.Button(buttons, text="Get browser ad blocker", command=self.open_adblock_search, width=22).grid(row=0, column=2, padx=5)

    def ad_tracker_domains(self):
        return sorted(set([
            "ad.doubleclick.net",
            "ads.doubleclick.net",
            "doubleclick.net",
            "googleads.g.doubleclick.net",
            "pagead2.googlesyndication.com",
            "googleadservices.com",
            "googlesyndication.com",
            "adservice.google.com",
            "static.doubleclick.net",
            "www.googleadservices.com",
            "s.youtube.com",
            "ads.youtube.com",
            "ad.youtube.com",
            "pubads.g.doubleclick.net",
            "securepubads.g.doubleclick.net",
            "analytics.google.com",
            "google-analytics.com",
            "www.google-analytics.com",
            "ssl.google-analytics.com",
            "connect.facebook.net",
            "ads.facebook.com",
            "analytics.twitter.com",
            "ads-twitter.com",
            "bat.bing.com",
            "ads.linkedin.com",
            "px.ads.linkedin.com",
            "adsrvr.org",
            "scorecardresearch.com",
            "quantserve.com",
            "hotjar.com",
            "static.hotjar.com",
            "taboola.com",
            "outbrain.com",
            "criteo.com",
            "adnxs.com",
            "rubiconproject.com",
            "openx.net",
            "adform.net",
            "moatads.com",
        ]))

    def enable_web_shield(self):
        if platform.system() != "Windows":
            self.say("Web shield host blocking is made for Windows.")
            return
        confirm = messagebox.askyesno("Barnaby Web Shield", "Enable aggressive ad and tracker blocking in the Windows hosts file? You may need to run Barnaby as Administrator.", parent=self.root)
        if not confirm:
            self.say("I did not change web shield settings.")
            return
        try:
            hosts = Path(os.environ.get("SystemRoot", "C:/Windows")) / "System32" / "drivers" / "etc" / "hosts"
            original = hosts.read_text(encoding="utf-8", errors="ignore")
            backup = APP_DIR / "hosts.barnaby.backup"
            if not backup.exists():
                backup.write_text(original, encoding="utf-8")
            start = "# Barnaby Web Shield Start"
            end = "# Barnaby Web Shield End"
            kept = original
            if start in kept and end in kept:
                before = kept.split(start)[0].rstrip()
                after = kept.split(end, 1)[1].lstrip()
                kept = (before + "\n" + after).strip() + "\n"
            block = [start]
            for domain in self.ad_tracker_domains():
                block.append(f"0.0.0.0 {domain}")
                block.append(f"0.0.0.0 www.{domain}" if not domain.startswith("www.") else f"0.0.0.0 {domain[4:]}")
            block.append(end)
            hosts.write_text(kept.rstrip() + "\n\n" + "\n".join(block) + "\n", encoding="utf-8")
            subprocess.run(["ipconfig", "/flushdns"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
            self.say("Barnaby web shield is enabled.")
            messagebox.showinfo("Barnaby Web Shield", "Web shield enabled. Restart browsers for best results.", parent=self.root)
        except Exception as error:
            messagebox.showwarning("Barnaby Web Shield", f"Barnaby could not change the hosts file. Try running Barnaby as Administrator.\n\n{error}", parent=self.root)
            self.say("I could not enable web shield. I may need Administrator permission.")

    def disable_web_shield(self):
        if platform.system() != "Windows":
            self.say("Web shield host blocking is made for Windows.")
            return
        try:
            hosts = Path(os.environ.get("SystemRoot", "C:/Windows")) / "System32" / "drivers" / "etc" / "hosts"
            original = hosts.read_text(encoding="utf-8", errors="ignore")
            start = "# Barnaby Web Shield Start"
            end = "# Barnaby Web Shield End"
            if start in original and end in original:
                before = original.split(start)[0].rstrip()
                after = original.split(end, 1)[1].lstrip()
                hosts.write_text((before + "\n" + after).strip() + "\n", encoding="utf-8")
                subprocess.run(["ipconfig", "/flushdns"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
            self.say("Barnaby web shield is disabled.")
            messagebox.showinfo("Barnaby Web Shield", "Web shield disabled.", parent=self.root)
        except Exception as error:
            messagebox.showwarning("Barnaby Web Shield", f"Barnaby could not disable the web shield. Try running Barnaby as Administrator.\n\n{error}", parent=self.root)

    def open_adblock_search(self):
        webbrowser.open("https://www.bing.com/search?q=" + quote_plus("uBlock Origin official browser extension"))
        self.say("I opened a search for a strong browser ad blocker.")

    def organize_files(self):
        folder = filedialog.askdirectory(title="Choose a folder for Barnaby to organize", parent=self.root)
        if not folder:
            self.say("Okay, I will not organize anything yet.")
            return
        target = Path(folder)
        if not target.exists() or not target.is_dir():
            self.say("That folder does not look available.")
            return
        confirm = messagebox.askyesno(
            "Barnaby File Organizer",
            f"Barnaby will organize files inside:\n{target}\n\nFiles will be moved into folders like Images, Documents, Videos, Music, Archives, Installers, Code, and Other. Continue?",
            parent=self.root,
        )
        if not confirm:
            self.say("No changes made. I will leave your files alone.")
            return
        moved = 0
        categories = {
            "Images": {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".svg", ".ico"},
            "Documents": {".pdf", ".doc", ".docx", ".txt", ".rtf", ".xls", ".xlsx", ".ppt", ".pptx", ".csv"},
            "Videos": {".mp4", ".mov", ".avi", ".mkv", ".wmv", ".webm"},
            "Music": {".mp3", ".wav", ".flac", ".aac", ".ogg", ".m4a"},
            "Archives": {".zip", ".rar", ".7z", ".tar", ".gz"},
            "Installers": {".exe", ".msi", ".dmg", ".pkg"},
            "Code": {".py", ".js", ".ts", ".html", ".css", ".json", ".bat", ".ps1"},
        }
        try:
            for item in target.iterdir():
                if item.is_dir():
                    continue
                category = "Other"
                suffix = item.suffix.lower()
                for folder_name, suffixes in categories.items():
                    if suffix in suffixes:
                        category = folder_name
                        break
                destination_dir = target / category
                destination_dir.mkdir(exist_ok=True)
                destination = destination_dir / item.name
                counter = 1
                while destination.exists():
                    destination = destination_dir / f"{item.stem}_{counter}{item.suffix}"
                    counter += 1
                shutil.move(str(item), str(destination))
                moved += 1
            self.say(f"I organized {moved} file{'s' if moved != 1 else ''} for you.")
            messagebox.showinfo("Barnaby File Organizer", f"Barnaby organized {moved} file(s).", parent=self.root)
        except Exception as error:
            messagebox.showwarning("Barnaby File Organizer", f"Barnaby could not finish organizing that folder.\n\n{error}", parent=self.root)
            self.say("I could not finish organizing that folder.")

    def clean_junk_files(self):
        targets = []
        home = Path.home()
        temp_locations = [
            Path(os.environ.get("TEMP", str(home / "AppData" / "Local" / "Temp"))),
            Path(os.environ.get("TMP", str(home / "AppData" / "Local" / "Temp"))),
            Path(os.environ.get("SystemRoot", "C:/Windows")) / "Temp",
        ]
        seen = set()
        for folder in temp_locations:
            if folder in seen or not folder.exists():
                continue
            seen.add(folder)
            try:
                for path in folder.rglob("*"):
                    if len(targets) >= 1000:
                        break
                    if path.is_file():
                        targets.append(path)
            except Exception:
                continue
        total_size = 0
        for path in targets:
            try:
                total_size += path.stat().st_size
            except Exception:
                pass
        size_mb = round(total_size / (1024 * 1024), 1)
        if not targets:
            self.say("I did not find safe temporary junk files to clean.")
            messagebox.showinfo("Barnaby Junk Cleaner", "No safe temporary junk files were found.", parent=self.root)
            return
        confirm = messagebox.askyesno(
            "Barnaby Junk Cleaner",
            f"Barnaby found {len(targets)} temporary junk file(s), about {size_mb} MB.\n\nDelete these temporary files? Files in use will be skipped.",
            parent=self.root,
        )
        if not confirm:
            self.say("I did not delete any junk files.")
            return
        deleted = 0
        freed = 0
        for path in targets:
            try:
                size = path.stat().st_size
                path.unlink()
                deleted += 1
                freed += size
            except Exception:
                continue
        freed_mb = round(freed / (1024 * 1024), 1)
        self.say(f"I cleaned {deleted} junk file{'s' if deleted != 1 else ''} and freed about {freed_mb} megabytes.")
        messagebox.showinfo("Barnaby Junk Cleaner", f"Cleaned {deleted} temporary file(s).\nFreed about {freed_mb} MB.", parent=self.root)

    def open_workspace_silent_mode(self):
        self.silent_mode_enabled = True
        self.walk_enabled = False
        win = self.make_window("Barnaby Workspace Silent Mode", "700x520")
        tk.Label(win, text="Workspace silent mode", font=("Segoe UI", 12, "bold")).pack(pady=8)
        tk.Label(
            win,
            text="Barnaby will stay quiet, stop walking, and help you close distracting apps. Select only apps you want to close.",
            wraplength=660,
            justify="left",
        ).pack(padx=12)
        box = tk.Listbox(win, font=("Segoe UI", 10), selectmode="extended")
        box.pack(fill="both", expand=True, padx=12, pady=8)
        processes = self.workspace_process_suggestions()
        if processes:
            for proc in processes:
                box.insert("end", f"PID {proc['pid']}: {proc['name']}")
        else:
            box.insert("end", "No common distracting apps found. Barnaby is quiet and staying still.")

        def close_selected():
            selected = box.curselection()
            if not selected or not processes:
                return
            confirm = messagebox.askyesno("Workspace Silent Mode", "Close selected programs? Unsaved work in those apps could be lost.", parent=win)
            if not confirm:
                return
            closed = 0
            for index in reversed(selected):
                proc_info = processes[index]
                try:
                    if psutil:
                        proc = psutil.Process(proc_info["pid"])
                        proc.terminate()
                    else:
                        subprocess.run(["taskkill", "/PID", str(proc_info["pid"])], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
                    box.delete(index)
                    closed += 1
                except Exception:
                    continue
            self.bubble.config(text=f"Workspace silent mode closed {closed} program{'s' if closed != 1 else ''}.")

        def mute_windows():
            if platform.system() != "Windows":
                self.bubble.config(text="Windows sound muting only works on Windows.")
                return
            try:
                subprocess.Popen(
                    [
                        "powershell",
                        "-NoProfile",
                        "-ExecutionPolicy",
                        "Bypass",
                        "-Command",
                        "(New-Object -ComObject WScript.Shell).SendKeys([char]173)",
                    ],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
                self.bubble.config(text="I toggled Windows mute for workspace silent mode.")
            except Exception as error:
                messagebox.showwarning("Workspace Silent Mode", f"Barnaby could not toggle mute.\n\n{error}", parent=win)

        def exit_silent():
            self.silent_mode_enabled = False
            self.walk_enabled = True
            self.say("Workspace silent mode is off. I can talk and walk again.")

        buttons = tk.Frame(win)
        buttons.pack(pady=8)
        tk.Button(buttons, text="Close selected apps", command=close_selected, width=20).grid(row=0, column=0, padx=5)
        tk.Button(buttons, text="Mute/unmute Windows", command=mute_windows, width=20).grid(row=0, column=1, padx=5)
        tk.Button(buttons, text="Exit silent mode", command=exit_silent, width=18).grid(row=0, column=2, padx=5)
        self.bubble.config(text="Workspace silent mode is on. I will stay quiet and stop walking.")

    def workspace_process_suggestions(self):
        noisy_names = [
            "discord",
            "spotify",
            "steam",
            "epicgameslauncher",
            "battle.net",
            "teams",
            "slack",
            "telegram",
            "whatsapp",
            "zoom",
            "skype",
            "onedrive",
            "chrome",
            "msedge",
            "firefox",
        ]
        suggestions = []
        if psutil:
            try:
                for proc in psutil.process_iter(["pid", "name"]):
                    name = str(proc.info.get("name") or "")
                    lower = name.lower().replace(".exe", "")
                    if any(noisy in lower for noisy in noisy_names):
                        suggestions.append({"pid": int(proc.info.get("pid")), "name": name})
            except Exception:
                pass
        elif platform.system() == "Windows":
            try:
                output = subprocess.check_output("tasklist", shell=True, text=True, errors="ignore")
                for line in output.splitlines()[3:]:
                    parts = line.split()
                    if len(parts) >= 2:
                        name = parts[0]
                        lower = name.lower().replace(".exe", "")
                        if any(noisy in lower for noisy in noisy_names):
                            suggestions.append({"pid": int(parts[1]), "name": name})
            except Exception:
                pass
        return suggestions[:80]

    def scan_suspicious_files(self):
        findings = self.find_suspicious_files()
        processes = self.find_suspicious_processes()
        entries = [("process", item) for item in processes] + [("file", item) for item in findings]
        self.last_scan_count = len(entries)
        win = self.make_window("Barnaby Safety Scan", "760x500")
        tk.Label(win, text="PC safety scan", font=("Segoe UI", 12, "bold")).pack(pady=8)
        tk.Label(
            win,
            text="Barnaby checks suspicious running processes and risky files in common folders, startup folders, and temp folders. This helps, but keep Windows Security enabled too.",
            wraplength=720,
            justify="left",
        ).pack(padx=12)
        box = tk.Listbox(win, font=("Segoe UI", 9), selectmode="extended")
        box.pack(fill="both", expand=True, padx=12, pady=8)
        if not entries:
            box.insert("end", "No suspicious files or running processes found in the places Barnaby checked.")
            self.say("I did not find suspicious files or suspicious running processes in the places I checked.")
        else:
            for entry_type, value in entries:
                label = "RUNNING PROCESS" if entry_type == "process" else "FILE"
                box.insert("end", f"{label}: {value}")
            self.say(f"I found {len(entries)} suspicious item{'s' if len(entries) != 1 else ''}. Please review them before taking action.")

        def quarantine_selected():
            selected = box.curselection()
            if not selected or not entries:
                self.say("Select a suspicious file first if you want me to quarantine it.")
                return
            confirm = messagebox.askyesno(
                "Barnaby Quarantine",
                "Move selected suspicious files to Barnaby quarantine? Running processes cannot be moved. Only do this if you do not trust the files.",
                parent=win,
            )
            if not confirm:
                self.say("I did not move anything.")
                return
            QUARANTINE_DIR.mkdir(parents=True, exist_ok=True)
            moved = 0
            for index in reversed(selected):
                entry_type, source = entries[index]
                if entry_type != "file":
                    continue
                if not source.exists() or source.is_dir():
                    continue
                destination = QUARANTINE_DIR / source.name
                counter = 1
                while destination.exists():
                    destination = QUARANTINE_DIR / f"{source.stem}_{counter}{source.suffix}"
                    counter += 1
                try:
                    shutil.move(str(source), str(destination))
                    box.delete(index)
                    moved += 1
                except Exception:
                    pass
            self.say(f"I moved {moved} suspicious file{'s' if moved != 1 else ''} to quarantine.")

        buttons = tk.Frame(win)
        buttons.pack(pady=(0, 10))
        tk.Button(buttons, text="Quarantine selected files", command=quarantine_selected, width=22).grid(row=0, column=0, padx=5)
        tk.Button(buttons, text="Run Windows Security quick scan", command=self.run_windows_security_quick_scan, width=28).grid(row=0, column=1, padx=5)
        tk.Button(buttons, text="Open Windows Security", command=self.open_windows_security, width=22).grid(row=0, column=2, padx=5)

    def find_suspicious_files(self):
        home = Path.home()
        folders = [
            home / "Desktop",
            home / "Downloads",
            home / "Documents",
            Path(os.environ.get("TEMP", str(home / "AppData" / "Local" / "Temp"))),
            Path(os.environ.get("APPDATA", str(home / "AppData" / "Roaming"))) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup",
            Path(os.environ.get("PROGRAMDATA", "C:/ProgramData")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup",
        ]
        risky_extensions = {".exe", ".scr", ".bat", ".cmd", ".ps1", ".vbs", ".js", ".jse", ".wsf", ".hta", ".dll", ".msi", ".jar", ".lnk", ".com", ".pif"}
        risky_words = self.risky_words()
        findings = []
        for folder in folders:
            if not folder.exists():
                continue
            try:
                for path in folder.rglob("*"):
                    if len(findings) >= 100:
                        return findings
                    if path.is_dir():
                        continue
                    name = path.name.lower()
                    suffix = path.suffix.lower()
                    parts = path.name.lower().split(".")
                    has_double_extension = len(parts) >= 3 and f".{parts[-2]}" in {".pdf", ".doc", ".docx", ".jpg", ".png", ".txt"} and suffix in risky_extensions
                    if suffix in risky_extensions and any(word in name for word in risky_words):
                        findings.append(path)
                    elif has_double_extension:
                        findings.append(path)
                    elif suffix in {".scr", ".hta", ".jse", ".wsf"}:
                        findings.append(path)
            except Exception:
                continue
        return findings

    def risky_words(self):
        return [
            "memz",
            "rat",
            "trojan",
            "stealer",
            "keylogger",
            "miner",
            "payload",
            "crack",
            "hack",
            "backdoor",
            "virus",
            "malware",
            "ransom",
            "worm",
            "spyware",
            "rootkit",
            "botnet",
            "remoteaccess",
            "credential",
            "grabber",
        ]

    def find_suspicious_processes(self):
        risky_words = self.risky_words()
        findings = []
        if psutil:
            try:
                for proc in psutil.process_iter(["pid", "name", "exe", "cmdline"]):
                    info = proc.info
                    name = str(info.get("name") or "")
                    exe = str(info.get("exe") or "")
                    cmdline = " ".join([str(part) for part in info.get("cmdline") or []])
                    haystack = f"{name} {exe} {cmdline}".lower()
                    if any(word in haystack for word in risky_words):
                        findings.append(f"PID {info.get('pid')}: {name} {exe}".strip())
            except Exception:
                pass
        elif platform.system() == "Windows":
            try:
                output = subprocess.check_output("tasklist", shell=True, text=True, errors="ignore")
                for line in output.splitlines():
                    lower = line.lower()
                    if any(word in lower for word in risky_words):
                        findings.append(line.strip())
            except Exception:
                pass
        return sorted(set(findings))[:50]

    def startup_safety_scan(self):
        if self.closed or self.startup_scan_done:
            return
        self.startup_scan_done = True
        findings = self.find_suspicious_files()
        processes = self.find_suspicious_processes()
        total = len(findings) + len(processes)
        self.last_scan_count = total
        if total:
            self.say(f"Safety alert. I found {total} suspicious item{'s' if total != 1 else ''}. Right-click me and choose Safety scan to review them.")
        elif platform.system() == "Windows":
            self.say("I ran my Barnaby safety check and did not spot obvious suspicious items. I can also launch a Windows Security quick scan from my Safety scan button.")
        else:
            self.say("I ran my Barnaby safety check and did not spot obvious suspicious items.")

    def run_windows_security_quick_scan(self):
        if platform.system() != "Windows":
            self.say("Windows Security quick scan is only available on Windows.")
            return
        try:
            subprocess.Popen(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    "Start-MpScan -ScanType QuickScan",
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            self.say("I started a Windows Security quick scan. Keep Windows Security open if it asks for attention.")
            messagebox.showinfo("Barnaby Safety Scan", "Windows Security quick scan was started.", parent=self.root)
        except Exception as error:
            messagebox.showwarning("Barnaby Safety Scan", f"Barnaby could not start Windows Security quick scan.\n\n{error}", parent=self.root)
            self.say("I could not start Windows Security quick scan from here.")

    def open_windows_security(self):
        if platform.system() != "Windows":
            self.say("Windows Security opens only on Windows.")
            return
        try:
            subprocess.Popen(["cmd", "/c", "start", "windowsdefender:"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            self.say("I opened Windows Security for you.")
        except Exception as error:
            messagebox.showwarning("Barnaby Safety Scan", f"Barnaby could not open Windows Security.\n\n{error}", parent=self.root)

    def background_safety_check(self):
        if self.closed:
            return
        findings = self.find_suspicious_files()
        processes = self.find_suspicious_processes()
        total = len(findings) + len(processes)
        if total and total != self.last_scan_count:
            self.last_scan_count = total
            self.say(f"Safety alert. I found {total} suspicious item{'s' if total != 1 else ''}. Right-click me and choose Safety scan to review them.")
        self.safe_after(180000, self.background_safety_check)

    def show_system_info(self):
        lines = [
            f"Computer: {platform.node() or 'Unknown'}",
            f"System: {platform.system()} {platform.release()}",
            f"Python: {platform.python_version()}",
        ]
        if psutil:
            try:
                memory = psutil.virtual_memory()
                battery = psutil.sensors_battery()
                lines.append(f"Memory used: {memory.percent}%")
                if battery:
                    lines.append(f"Battery: {battery.percent}%")
            except Exception:
                lines.append("Extra system details are not available right now.")
        win = self.make_window("Barnaby System Info", "430x280")
        tk.Label(win, text="System info", font=("Segoe UI", 12, "bold")).pack(pady=8)
        text = tk.Text(win, font=("Segoe UI", 10), wrap="word")
        text.pack(fill="both", expand=True, padx=12, pady=8)
        text.insert("1.0", "\n".join(lines))
        text.config(state="disabled")
        self.say("I checked a few system details for you.")

    def open_data_folder(self):
        APP_DIR.mkdir(parents=True, exist_ok=True)
        try:
            if platform.system() == "Windows":
                os.startfile(str(APP_DIR))
            else:
                self.say(f"Barnaby data is saved at {APP_DIR}")
        except Exception as error:
            messagebox.showwarning("Barnaby folder", f"Barnaby could not open his data folder.\n\n{error}")

    def toggle_walking(self):
        self.walk_enabled = not self.walk_enabled
        self.say("I will keep walking around sometimes." if self.walk_enabled else "I will stay put for now.")

    def make_window(self, title, geometry):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.geometry(geometry)
        try:
            win.attributes("-topmost", True)
        except Exception:
            pass
        return win

    def random_walk(self):
        if self.closed:
            return
        if self.walk_enabled:
            try:
                screen_w = self.root.winfo_screenwidth()
                screen_h = self.root.winfo_screenheight()
                width = max(self.root.winfo_width(), 330)
                height = max(self.root.winfo_height(), 420)
                target_x = random.randint(0, max(1, screen_w - width))
                target_y = random.randint(20, max(30, screen_h - height - 40))
                self.glide_to(target_x, target_y, 30)
                if random.random() < 0.45:
                    self.say(random.choice(["Just taking a little walk.", "I am checking the desktop currents.", "Barnaby patrol complete."]))
            except Exception:
                pass
        self.safe_after(random.randint(20000, 45000), self.random_walk)

    def glide_to(self, target_x, target_y, steps):
        start_x = self.root.winfo_x()
        start_y = self.root.winfo_y()
        step = {"value": 0}

        def move():
            if self.closed:
                return
            step["value"] += 1
            self.walk_phase += 1
            self.draw_octopus()
            t = step["value"] / steps
            x = int(start_x + (target_x - start_x) * t)
            y = int(start_y + (target_y - start_y) * t)
            self.root.geometry(f"+{x}+{y}")
            if step["value"] < steps:
                self.safe_after(45, move)

        move()

    def helpful_nudge(self):
        if self.closed:
            return
        choices = []
        if self.tasks:
            choices.append(f"You still have {len(self.tasks)} task{'s' if len(self.tasks) != 1 else ''}. Right-click me if you want to review them.")
        if self.notes:
            choices.append(f"I am holding {len(self.notes)} note{'s' if len(self.notes) != 1 else ''} for you.")
        choices.extend([
            "Right-click me when you need tasks, programs, Outlook, internet search, file organizing, safety scan, notes, or system info.",
            "I can search the internet for you whenever you want.",
            "If you want me out of the way, drag me by holding the mouse button.",
        ])
        self.say(random.choice(choices))
        self.safe_after(random.randint(70000, 120000), self.helpful_nudge)

    def close(self):
        self.closed = True
        try:
            self.root.destroy()
        except Exception:
            pass

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    try:
        BarnabyApp().run()
    except Exception as error:
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Barnaby crashed", f"Barnaby ran into a problem.\n\n{error}")
            root.destroy()
        except Exception:
            raise
