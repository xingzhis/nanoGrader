#!/usr/bin/env python3
import csv
import json
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

try:
    import fitz  # PyMuPDF
    from PIL import Image, ImageTk

    PDF_EMBED_AVAILABLE = True
except Exception:
    fitz = None
    Image = None
    ImageTk = None
    PDF_EMBED_AVAILABLE = False


class QuizGraderApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Quiz Grader")
        self.root.geometry("1400x860")

        self.base_dir = Path.cwd()
        self.default_roster = self.base_dir / "roster.csv"
        self.default_submissions = self.base_dir / "Quiz1"
        self.state_path = self.base_dir / "grading_state.json"
        self.export_path = self.base_dir / "grades_export.csv"

        self.roster_path_var = tk.StringVar(value=str(self.default_roster))
        self.submissions_path_var = tk.StringVar(value=str(self.default_submissions))
        self.full_score_var = tk.StringVar(value="10")

        self.students = []
        self.student_by_netid = {}
        self.grades = {}
        self.submissions = {}
        self.unmatched_files = []
        self.manual_mappings = {}
        self.rubric_items = []
        self.rubric_vars = {}
        self.current_index = 0

        self.current_pdf_path = None
        self.pdf_doc = None
        self.pdf_zoom_multiplier = 1.0
        self.pdf_tk_imgs = []
        self.pdf_page_offsets = []
        self.pdf_total_height = 1
        self.last_canvas_width = 0

        self.unmatched_choice_var = tk.StringVar()
        self.map_student_choice_var = tk.StringVar()
        self.extra_deduction_var = tk.StringVar(value="0")
        self.computed_score_var = tk.StringVar(value="Score: -")
        self.pdf_status_var = tk.StringVar(value="PDF: -")

        self.info_name_var = tk.StringVar(value="Name: -")
        self.info_netid_var = tk.StringVar(value="NetID: -")
        self.info_email_var = tk.StringVar(value="Email: -")
        self.info_submission_var = tk.StringVar(value="Submission: -")
        self.info_progress_var = tk.StringVar(value="Progress: -")

        self.add_rubric_name_var = tk.StringVar()
        self.add_rubric_points_var = tk.StringVar(value="1")

        self.comments_text = None
        self._loading_form = False
        self.rubric_checks_frame = None
        self.rubric_canvas = None
        self.rubric_canvas_window = None
        self.info_box = None
        self.nav_bar = None
        self.rubric_setup_box = None
        self.grade_box = None
        self.map_box = None
        self.output_box = None
        self.unmatched_combo = None
        self.map_student_combo = None
        self.pdf_canvas = None
        self.pdf_canvas_x_scroll = None
        self.pdf_canvas_y_scroll = None

        self._build_ui()
        self._load_data()

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill=tk.X)

        ttk.Label(top, text="Roster CSV").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.roster_path_var, width=58).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Label(top, text="Submissions Folder").grid(row=1, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.submissions_path_var, width=58).grid(row=1, column=1, sticky="we", padx=6)
        ttk.Button(top, text="Reload", command=self._load_data).grid(row=0, column=2, rowspan=2, padx=8)
        ttk.Button(top, text="Reset", width=8, command=self._clear_state_with_confirm).grid(row=0, column=3, rowspan=2, padx=(0, 8))

        ttk.Label(top, text="Full Score").grid(row=0, column=4, sticky="e")
        ttk.Entry(top, textvariable=self.full_score_var, width=8).grid(row=0, column=5, sticky="w")

        top.columnconfigure(1, weight=1)

        body = ttk.Panedwindow(self.root, orient=tk.HORIZONTAL)
        body.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        left = ttk.Frame(body)
        right = ttk.Frame(body)
        body.add(left, weight=3)
        body.add(right, weight=3)

        pdf_box = ttk.LabelFrame(left, text="PDF Viewer", padding=8)
        pdf_box.pack(fill=tk.BOTH, expand=True)

        pdf_controls = ttk.Frame(pdf_box)
        pdf_controls.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(pdf_controls, text="Prev Page", command=self._pdf_prev_page).pack(side=tk.LEFT)
        ttk.Button(pdf_controls, text="Next Page", command=self._pdf_next_page).pack(side=tk.LEFT, padx=6)
        ttk.Button(pdf_controls, text="Zoom -", command=self._pdf_zoom_out).pack(side=tk.LEFT)
        ttk.Button(pdf_controls, text="Zoom +", command=self._pdf_zoom_in).pack(side=tk.LEFT, padx=6)
        ttk.Button(pdf_controls, text="Reset Fit", command=self._pdf_reset_fit).pack(side=tk.LEFT, padx=6)
        ttk.Label(pdf_controls, textvariable=self.pdf_status_var).pack(side=tk.LEFT, padx=10)

        pdf_canvas_wrap = ttk.Frame(pdf_box)
        pdf_canvas_wrap.pack(fill=tk.BOTH, expand=True)
        self.pdf_canvas = tk.Canvas(pdf_canvas_wrap, background="#2b2b2b", highlightthickness=0)
        self.pdf_canvas_y_scroll = ttk.Scrollbar(pdf_canvas_wrap, orient=tk.VERTICAL, command=self.pdf_canvas.yview)
        self.pdf_canvas_x_scroll = ttk.Scrollbar(pdf_canvas_wrap, orient=tk.HORIZONTAL, command=self.pdf_canvas.xview)
        self.pdf_canvas.configure(
            yscrollcommand=self._pdf_on_yscroll,
            xscrollcommand=self.pdf_canvas_x_scroll.set,
        )
        self.pdf_canvas.grid(row=0, column=0, sticky="nsew")
        self.pdf_canvas_y_scroll.grid(row=0, column=1, sticky="ns")
        self.pdf_canvas_x_scroll.grid(row=1, column=0, sticky="ew")
        pdf_canvas_wrap.rowconfigure(0, weight=1)
        pdf_canvas_wrap.columnconfigure(0, weight=1)
        self.pdf_canvas.bind("<MouseWheel>", self._pdf_on_mousewheel)
        self.pdf_canvas.bind("<Configure>", self._pdf_on_canvas_resize)

        self.info_box = ttk.LabelFrame(right, text="Current Student", padding=8)
        self.info_box.pack(fill=tk.X)
        ttk.Label(self.info_box, textvariable=self.info_name_var).pack(anchor="w")
        ttk.Label(self.info_box, textvariable=self.info_netid_var).pack(anchor="w")
        ttk.Label(self.info_box, textvariable=self.info_email_var).pack(anchor="w")
        ttk.Label(self.info_box, textvariable=self.info_submission_var).pack(anchor="w")
        ttk.Label(self.info_box, textvariable=self.info_progress_var).pack(anchor="w", pady=(4, 0))

        self.nav_bar = ttk.Frame(right, padding=(0, 8))
        self.nav_bar.pack(fill=tk.X)
        ttk.Button(self.nav_bar, text="Next Ungraded", command=self._go_next_ungraded).pack(side=tk.LEFT)
        ttk.Button(self.nav_bar, text="Previous", command=self._go_previous).pack(side=tk.LEFT, padx=6)
        ttk.Button(self.nav_bar, text="Next Student", command=self._go_next).pack(side=tk.LEFT)
        ttk.Button(self.nav_bar, text="Export CSV", command=self._export_csv).pack(side=tk.LEFT, padx=6)

        self.rubric_setup_box = ttk.LabelFrame(right, text="Rubric Setup", padding=8)
        self.rubric_setup_box.pack(fill=tk.X)
        add_row = ttk.Frame(self.rubric_setup_box)
        add_row.pack(fill=tk.X)
        ttk.Label(add_row, text="Name").pack(side=tk.LEFT)
        name_entry = ttk.Entry(add_row, textvariable=self.add_rubric_name_var, width=18)
        name_entry.pack(side=tk.LEFT, padx=(4, 8))
        ttk.Label(add_row, text="Pts").pack(side=tk.LEFT)
        points_entry = ttk.Entry(add_row, textvariable=self.add_rubric_points_var, width=6)
        points_entry.pack(side=tk.LEFT, padx=4)
        add_btn = ttk.Button(add_row, text="Add Rubric", command=self._add_rubric_item)
        add_btn.pack(side=tk.LEFT, padx=(6, 0))
        name_entry.bind("<Return>", lambda _e: self._add_rubric_item())
        points_entry.bind("<Return>", lambda _e: self._add_rubric_item())

        self.grade_box = ttk.LabelFrame(right, text="Rubric + Grading", padding=8)
        self.grade_box.pack(fill=tk.X, expand=False, pady=(8, 0))

        rubric_list_wrap = ttk.Frame(self.grade_box)
        rubric_list_wrap.pack(fill=tk.X, pady=(0, 8))
        self.rubric_canvas = tk.Canvas(rubric_list_wrap, height=96, highlightthickness=0)
        rubric_scroll = ttk.Scrollbar(rubric_list_wrap, orient=tk.VERTICAL, command=self.rubric_canvas.yview)
        self.rubric_canvas.configure(yscrollcommand=rubric_scroll.set)
        self.rubric_canvas.grid(row=0, column=0, sticky="nsew")
        rubric_scroll.grid(row=0, column=1, sticky="ns")
        rubric_list_wrap.columnconfigure(0, weight=1)

        self.rubric_checks_frame = ttk.Frame(self.rubric_canvas)
        self.rubric_canvas_window = self.rubric_canvas.create_window((0, 0), window=self.rubric_checks_frame, anchor="nw")
        self.rubric_checks_frame.bind("<Configure>", self._on_rubric_frame_configure)
        self.rubric_canvas.bind("<Configure>", self._on_rubric_canvas_configure)

        extra = ttk.Frame(self.grade_box)
        extra.pack(fill=tk.X)
        ttk.Label(extra, text="Extra deduction").pack(side=tk.LEFT)
        ttk.Entry(extra, textvariable=self.extra_deduction_var, width=8).pack(side=tk.LEFT, padx=6)
        self.extra_deduction_var.trace_add("write", lambda *_: self._on_form_changed())

        ttk.Label(self.grade_box, textvariable=self.computed_score_var, font=("TkDefaultFont", 12, "bold")).pack(
            anchor="w", pady=(8, 6)
        )

        ttk.Label(self.grade_box, text="Comments").pack(anchor="w")
        self.comments_text = tk.Text(self.grade_box, height=3, width=70)
        self.comments_text.pack(fill=tk.X, expand=False)
        self.comments_text.bind("<KeyRelease>", lambda _e: self._on_form_changed())
        self.comments_text.bind("<FocusOut>", lambda _e: self._on_form_changed())

        grade_actions = ttk.Frame(self.grade_box, padding=(0, 8, 0, 0))
        grade_actions.pack(fill=tk.X)
        ttk.Label(grade_actions, text="Auto-saves on edit/navigation").pack(side=tk.LEFT)

        self.map_box = ttk.LabelFrame(right, text="Unmatched PDF Mapping", padding=8)
        self.map_box.pack(fill=tk.X, pady=(10, 0))
        self.unmatched_combo = ttk.Combobox(self.map_box, textvariable=self.unmatched_choice_var, state="readonly", width=28)
        self.unmatched_combo.pack(fill=tk.X, pady=(0, 4))
        self.map_student_combo = ttk.Combobox(self.map_box, textvariable=self.map_student_choice_var, state="readonly", width=28)
        self.map_student_combo.pack(fill=tk.X, pady=(0, 4))
        ttk.Button(self.map_box, text="Assign PDF to Student", command=self._assign_mapping).pack(fill=tk.X)


    def _load_data(self):
        roster_path = Path(self.roster_path_var.get()).expanduser()
        submissions_dir = Path(self.submissions_path_var.get()).expanduser()

        if not roster_path.exists():
            messagebox.showerror("Missing file", f"Roster not found:\n{roster_path}")
            return
        if not submissions_dir.exists():
            messagebox.showerror("Missing folder", f"Submissions folder not found:\n{submissions_dir}")
            return

        self.students = []
        self.student_by_netid = {}
        self.submissions = {}
        self.unmatched_files = []

        with roster_path.open(newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if (row.get("Role", "") or "").strip() != "Student":
                    continue
                netid = (row.get("Net ID", "") or "").strip().lower()
                if not netid:
                    continue
                student = {
                    "netid": netid,
                    "first": (row.get("First Name", "") or "").strip(),
                    "last": (row.get("Last Name", "") or "").strip(),
                    "email": (row.get("Email", "") or "").strip(),
                    "submission": None,
                }
                self.students.append(student)
                self.student_by_netid[netid] = student

        self.students.sort(key=lambda s: (s["last"].lower(), s["first"].lower(), s["netid"]))

        pdf_paths = sorted(submissions_dir.glob("*.pdf"))
        by_stem = {p.stem.lower(): p for p in pdf_paths}

        saved = self._read_saved_state()
        self.rubric_items = saved.get("rubric_items", []) or []
        self.manual_mappings = saved.get("manual_mappings", {}) or {}
        self.grades = saved.get("grades", {}) or {}
        if saved.get("full_score") is not None:
            self.full_score_var.set(str(saved.get("full_score")))

        for netid, student in self.student_by_netid.items():
            if netid in by_stem:
                student["submission"] = str(by_stem[netid])
                self.submissions[netid] = str(by_stem[netid])

        for filename, mapped_netid in self.manual_mappings.items():
            mapped_netid = (mapped_netid or "").strip().lower()
            if mapped_netid in self.student_by_netid:
                path = submissions_dir / filename
                if path.exists():
                    self.student_by_netid[mapped_netid]["submission"] = str(path)
                    self.submissions[mapped_netid] = str(path)

        student_netids = set(self.student_by_netid.keys())
        for p in pdf_paths:
            if p.stem.lower() not in student_netids and p.name not in self.manual_mappings:
                self.unmatched_files.append(p.name)

        self._ensure_grade_defaults()
        self._build_rubric_checkboxes()
        self._refresh_mapping_controls()

        self.current_index = self._first_ungraded_index()
        self._show_current_student()
        self._save_state()

    def _read_saved_state(self):
        if not self.state_path.exists():
            return {}
        try:
            with self.state_path.open("r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def _new_record(self):
        return {
            "selected_rubrics": [],
            "extra_deduction": 0.0,
            "comments": "",
            "graded": False,
            "status": "ungraded",
            "score": None,
            "total_deduction": 0.0,
        }

    def _get_record(self, netid, create=False):
        rec = self.grades.get(netid)
        if not isinstance(rec, dict):
            rec = None
        if rec is None and create:
            rec = self._new_record()
            self.grades[netid] = rec
        return rec

    def _ensure_grade_defaults(self):
        for s in self.students:
            netid = s["netid"]
            rec = self._get_record(netid, create=True)
            for k, v in self._new_record().items():
                if k not in rec:
                    rec[k] = v

            if not s["submission"]:
                rec["graded"] = True
                rec["status"] = "missing"
                rec["score"] = 0.0
            elif rec.get("status") == "missing":
                rec["graded"] = False
                rec["status"] = "ungraded"
                rec["score"] = None

    def _build_rubric_checkboxes(self):
        for child in self.rubric_checks_frame.winfo_children():
            child.destroy()
        self.rubric_vars = {}

        if not self.rubric_items:
            ttk.Label(self.rubric_checks_frame, text="No rubric deductions yet. Add one above.").pack(anchor="w")
            return

        for item in self.rubric_items:
            name = item.get("name", "").strip()
            points = float(item.get("points", 0))
            var = tk.IntVar(value=0)
            self.rubric_vars[name] = var
            cb = ttk.Checkbutton(
                self.rubric_checks_frame,
                text=f"{name} (-{self._fmt(points)})",
                variable=var,
                command=self._on_form_changed,
            )
            cb.pack(anchor="w")

        self.rubric_checks_frame.update_idletasks()
        self._on_rubric_frame_configure(None)

    def _on_rubric_frame_configure(self, _event):
        if self.rubric_canvas is None:
            return
        self.rubric_canvas.configure(scrollregion=self.rubric_canvas.bbox("all"))

    def _on_rubric_canvas_configure(self, event):
        if self.rubric_canvas is None or self.rubric_canvas_window is None:
            return
        self.rubric_canvas.itemconfigure(self.rubric_canvas_window, width=event.width)

    def _refresh_mapping_controls(self):
        self.unmatched_combo["values"] = self.unmatched_files
        if self.unmatched_files and self.unmatched_choice_var.get() not in self.unmatched_files:
            self.unmatched_choice_var.set(self.unmatched_files[0])
        if not self.unmatched_files:
            self.unmatched_choice_var.set("")

        student_labels = [f"{s['netid']} | {s['last']}, {s['first']}" for s in self.students]
        self.map_student_combo["values"] = student_labels
        if student_labels and self.map_student_choice_var.get() not in student_labels:
            self.map_student_choice_var.set(student_labels[0])

    def _assign_mapping(self):
        filename = self.unmatched_choice_var.get().strip()
        student_label = self.map_student_choice_var.get().strip()
        if not filename or not student_label:
            messagebox.showerror("Missing selection", "Pick both an unmatched PDF and a student.")
            return

        netid = student_label.split("|", 1)[0].strip().lower()
        target = self.student_by_netid.get(netid)
        if not target:
            messagebox.showerror("Invalid student", f"Could not find student for '{student_label}'.")
            return

        pdf_path = Path(self.submissions_path_var.get()).expanduser() / filename
        if not pdf_path.exists():
            messagebox.showerror("Missing PDF", f"File not found:\n{pdf_path}")
            return

        target["submission"] = str(pdf_path)
        self.submissions[netid] = str(pdf_path)
        self.manual_mappings[filename] = netid

        rec = self._get_record(netid, create=True)
        if rec.get("status") == "missing":
            rec["graded"] = False
            rec["status"] = "ungraded"
            rec["score"] = None

        self.unmatched_files = [x for x in self.unmatched_files if x != filename]
        self._refresh_mapping_controls()
        self._ensure_grade_defaults()
        self._show_current_student()
        self._save_state()

    def _safe_float(self, value, default=0.0):
        try:
            return float(value)
        except Exception:
            return default

    def _fmt(self, x):
        if abs(x - int(x)) < 1e-9:
            return str(int(x))
        return f"{x:.2f}"

    def _current_student(self):
        if not self.students:
            return None
        return self.students[self.current_index]

    def _compute_score_for_current(self):
        full_score = self._safe_float(self.full_score_var.get(), 10.0)
        selected = []
        total_deduction = 0.0
        rubric_points = {i.get("name", "").strip(): self._safe_float(i.get("points", 0.0), 0.0) for i in self.rubric_items}
        for name, var in self.rubric_vars.items():
            if var.get() == 1:
                selected.append(name)
                total_deduction += rubric_points.get(name, 0.0)
        extra = self._safe_float(self.extra_deduction_var.get(), 0.0)
        total_deduction += extra
        score = max(0.0, full_score - total_deduction)
        return score, total_deduction, selected, extra

    def _update_score_preview(self):
        if not self.students:
            self.computed_score_var.set("Score: -")
            return
        student = self._current_student()
        if not student["submission"]:
            self.computed_score_var.set("Score: 0 (missing submission)")
            return
        score, total_deduction, _, _ = self._compute_score_for_current()
        self.computed_score_var.set(f"Score: {self._fmt(score)}   (total deduction: {self._fmt(total_deduction)})")

    def _load_form_from_record(self, netid):
        rec = self._get_record(netid, create=True)
        selected = set(rec.get("selected_rubrics", []))
        self._loading_form = True
        for name, var in self.rubric_vars.items():
            var.set(1 if name in selected else 0)

        self.extra_deduction_var.set(self._fmt(self._safe_float(rec.get("extra_deduction", 0.0), 0.0)))
        self.comments_text.delete("1.0", tk.END)
        self.comments_text.insert("1.0", rec.get("comments", ""))
        self._loading_form = False
        self._update_score_preview()

    def _show_current_student(self):
        if not self.students:
            self.info_name_var.set("Name: -")
            self.info_netid_var.set("NetID: -")
            self.info_email_var.set("Email: -")
            self.info_submission_var.set("Submission: -")
            self.info_progress_var.set("Progress: No students loaded")
            self.computed_score_var.set("Score: -")
            return

        self._ensure_grade_defaults()

        s = self._current_student()
        rec = self._get_record(s["netid"], create=True)
        display_name = f"{s['last']}, {s['first']}"

        self.info_name_var.set(f"Name: {display_name}")
        self.info_netid_var.set(f"NetID: {s['netid']}")
        self.info_email_var.set(f"Email: {s['email']}")
        if s["submission"]:
            self.info_submission_var.set(f"Submission: {Path(s['submission']).name}")
        else:
            self.info_submission_var.set("Submission: MISSING (auto 0)")

        total = len(self.students)
        submission_total = sum(1 for st in self.students if st["submission"])
        manual_graded = sum(
            1
            for st in self.students
            if st["submission"] and self._get_record(st["netid"], create=True).get("status") == "graded"
        )
        missing_auto_zero = total - submission_total
        self.info_progress_var.set(
            f"student {self.current_index + 1}/{total}, graded {manual_graded}/{submission_total} "
            f"(missing auto-0: {missing_auto_zero}), status={rec.get('status', 'ungraded')}"
        )

        self._load_form_from_record(s["netid"])
        self._load_embedded_pdf_for_current()

    def _persist_current_form(self, mark_graded=False):
        if not self.students or self._loading_form:
            return
        s = self._current_student()
        rec = self._get_record(s["netid"], create=True)

        if not s["submission"]:
            rec["graded"] = True
            rec["status"] = "missing"
            rec["score"] = 0.0
            rec["selected_rubrics"] = []
            rec["extra_deduction"] = 0.0
            rec["total_deduction"] = 0.0
            rec["comments"] = ""
        else:
            score, total_deduction, selected, extra = self._compute_score_for_current()
            rec["selected_rubrics"] = selected
            rec["extra_deduction"] = extra
            rec["total_deduction"] = total_deduction
            rec["score"] = score
            rec["comments"] = self.comments_text.get("1.0", tk.END).strip()
            if mark_graded:
                rec["graded"] = True
                rec["status"] = "graded"

        self._save_state()

    def _on_form_changed(self):
        if self._loading_form:
            return
        self._update_score_preview()
        self._persist_current_form(mark_graded=True)
        if self.students:
            rec = self._get_record(self._current_student()["netid"], create=True)
            total = len(self.students)
            submission_total = sum(1 for st in self.students if st["submission"])
            manual_graded = sum(
                1
                for st in self.students
                if st["submission"] and self._get_record(st["netid"], create=True).get("status") == "graded"
            )
            missing_auto_zero = total - submission_total
            self.info_progress_var.set(
                f"student {self.current_index + 1}/{total}, graded {manual_graded}/{submission_total} "
                f"(missing auto-0: {missing_auto_zero}), status={rec.get('status', 'ungraded')}"
            )

    def _go_previous(self):
        if not self.students:
            return
        self._persist_current_form(mark_graded=True)
        self.current_index = (self.current_index - 1) % len(self.students)
        self._show_current_student()

    def _go_next(self):
        if not self.students:
            return
        self._persist_current_form(mark_graded=True)
        self.current_index = (self.current_index + 1) % len(self.students)
        self._show_current_student()

    def _first_ungraded_index(self):
        for i, s in enumerate(self.students):
            rec = self._get_record(s["netid"], create=True)
            if s["submission"] and not rec.get("graded", False):
                return i
        return 0

    def _go_next_ungraded(self):
        if not self.students:
            return
        self._persist_current_form(mark_graded=True)
        n = len(self.students)
        start = self.current_index
        for offset in range(1, n + 1):
            i = (start + offset) % n
            s = self.students[i]
            rec = self._get_record(s["netid"], create=True)
            if s["submission"] and not rec.get("graded", False):
                self.current_index = i
                self._show_current_student()
                return
        messagebox.showinfo("Done", "No ungraded students with submissions remain.")

    def _pdf_on_yscroll(self, first, last):
        self.pdf_canvas_y_scroll.set(first, last)
        self._update_pdf_status_from_view()

    def _pdf_set_status(self, text):
        self.pdf_status_var.set(text)

    def _pdf_on_mousewheel(self, event):
        if not self.pdf_canvas:
            return
        step = -1 * int(event.delta / 120)
        if step == 0:
            step = -1 if event.delta > 0 else 1
        self.pdf_canvas.yview_scroll(step, "units")
        self._update_pdf_status_from_view()

    def _pdf_on_canvas_resize(self, _event):
        w = self.pdf_canvas.winfo_width()
        if self.pdf_doc is None or w <= 10:
            return
        if abs(w - self.last_canvas_width) < 8:
            return
        self.last_canvas_width = w
        self._render_pdf_document(preserve_view=True)

    def _pdf_reset_fit(self):
        self.pdf_zoom_multiplier = 1.0
        self._render_pdf_document(preserve_view=True)

    def _close_pdf_doc(self):
        if self.pdf_doc is not None:
            try:
                self.pdf_doc.close()
            except Exception:
                pass
        self.pdf_doc = None
        self.current_pdf_path = None
        self.pdf_tk_imgs = []
        self.pdf_page_offsets = []
        self.pdf_total_height = 1

    def _clear_pdf_canvas(self):
        if self.pdf_canvas:
            self.pdf_canvas.delete("all")
            self.pdf_canvas.configure(scrollregion=(0, 0, 0, 0))

    def _load_embedded_pdf_for_current(self):
        s = self._current_student()
        if not s or not s.get("submission"):
            self._close_pdf_doc()
            self._clear_pdf_canvas()
            self._pdf_set_status("PDF: missing submission")
            return

        if not PDF_EMBED_AVAILABLE:
            self._close_pdf_doc()
            self._clear_pdf_canvas()
            self._pdf_set_status("PDF: install pymupdf + pillow")
            return

        path = s["submission"]
        if self.current_pdf_path == path and self.pdf_doc is not None:
            self._render_pdf_document(preserve_view=True)
            return

        self._close_pdf_doc()
        self.pdf_zoom_multiplier = 1.0
        self.last_canvas_width = 0
        try:
            self.pdf_doc = fitz.open(path)
            self.current_pdf_path = path
            self._render_pdf_document(preserve_view=False)
        except Exception as exc:
            self._close_pdf_doc()
            self._clear_pdf_canvas()
            self._pdf_set_status(f"PDF load failed: {exc}")

    def _render_pdf_document(self, preserve_view=True):
        if not PDF_EMBED_AVAILABLE or self.pdf_doc is None:
            return
        page_count = len(self.pdf_doc)
        if page_count == 0:
            self._clear_pdf_canvas()
            self._pdf_set_status("PDF: empty file")
            return

        prev_top = 0.0
        if preserve_view:
            prev_top = self.pdf_canvas.yview()[0]

        canvas_w = max(100, self.pdf_canvas.winfo_width() - 16)
        gap = 10
        y = gap
        offsets = []
        imgs = []

        self._clear_pdf_canvas()

        for i in range(page_count):
            page = self.pdf_doc.load_page(i)
            page_w = max(1.0, float(page.rect.width))
            base_scale = canvas_w / page_w
            scale = max(0.2, min(5.0, base_scale * self.pdf_zoom_multiplier))
            matrix = fitz.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            mode = "RGB" if pix.n < 4 else "RGBA"
            img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
            if mode == "RGBA":
                img = img.convert("RGB")

            tk_img = ImageTk.PhotoImage(img)
            imgs.append(tk_img)
            x = max(0, (canvas_w - tk_img.width()) // 2 + 4)
            self.pdf_canvas.create_image(x, y, image=tk_img, anchor="nw")
            offsets.append(y)
            y += tk_img.height() + gap

        self.pdf_tk_imgs = imgs
        self.pdf_page_offsets = offsets
        self.pdf_total_height = max(y, 1)
        self.pdf_canvas.configure(scrollregion=(0, 0, max(canvas_w + 8, 100), self.pdf_total_height))

        if preserve_view:
            self.pdf_canvas.yview_moveto(prev_top)
        else:
            self.pdf_canvas.yview_moveto(0.0)

        self._update_pdf_status_from_view()

    def _current_page_from_view(self):
        if not self.pdf_page_offsets:
            return 0
        top = self.pdf_canvas.canvasy(0)
        for i, start_y in enumerate(self.pdf_page_offsets):
            next_start = self.pdf_page_offsets[i + 1] if i + 1 < len(self.pdf_page_offsets) else self.pdf_total_height + 1
            if top < next_start:
                return i
        return len(self.pdf_page_offsets) - 1

    def _scroll_to_page(self, page_idx):
        if not self.pdf_page_offsets:
            return
        page_idx = max(0, min(page_idx, len(self.pdf_page_offsets) - 1))
        y = self.pdf_page_offsets[page_idx]
        frac = 0.0 if self.pdf_total_height <= 0 else min(1.0, max(0.0, y / self.pdf_total_height))
        self.pdf_canvas.yview_moveto(frac)
        self._update_pdf_status_from_view()

    def _update_pdf_status_from_view(self):
        if self.pdf_doc is None:
            return
        page_count = len(self.pdf_doc)
        page_idx = self._current_page_from_view()
        name = Path(self.current_pdf_path).name if self.current_pdf_path else "-"
        self._pdf_set_status(
            f"PDF: {name}  page {page_idx + 1}/{page_count}  zoom {self.pdf_zoom_multiplier:.2f}x fit"
        )

    def _pdf_prev_page(self):
        if self.pdf_doc is None:
            return
        self._scroll_to_page(self._current_page_from_view() - 1)

    def _pdf_next_page(self):
        if self.pdf_doc is None:
            return
        self._scroll_to_page(self._current_page_from_view() + 1)

    def _pdf_zoom_in(self):
        if self.pdf_doc is None:
            return
        self.pdf_zoom_multiplier = min(3.0, self.pdf_zoom_multiplier * 1.15)
        self._render_pdf_document(preserve_view=True)

    def _pdf_zoom_out(self):
        if self.pdf_doc is None:
            return
        self.pdf_zoom_multiplier = max(0.4, self.pdf_zoom_multiplier / 1.15)
        self._render_pdf_document(preserve_view=True)

    def _add_rubric_item(self):
        name = self.add_rubric_name_var.get().strip()
        points = self._safe_float(self.add_rubric_points_var.get(), None)
        if not name:
            messagebox.showerror("Invalid rubric", "Enter a rubric name.")
            return
        if points is None or points < 0:
            messagebox.showerror("Invalid rubric", "Deduction points must be a non-negative number.")
            return
        if any((item.get("name", "").strip() == name) for item in self.rubric_items):
            messagebox.showerror("Duplicate rubric", f"Rubric '{name}' already exists.")
            return
        self.rubric_items.append({"name": name, "points": points})
        self.add_rubric_name_var.set("")
        self._build_rubric_checkboxes()
        self._show_current_student()
        self._save_state()

    def _save_state(self):
        payload = {
            "full_score": self._safe_float(self.full_score_var.get(), 10.0),
            "rubric_items": self.rubric_items,
            "manual_mappings": self.manual_mappings,
            "grades": self.grades,
        }
        with self.state_path.open("w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)

    def _clear_state_with_confirm(self):
        confirmed = messagebox.askyesno(
            "Reset State",
            "This will permanently clear saved grades, rubric items, and manual PDF mappings.\n\nContinue?",
            icon="warning",
        )
        if not confirmed:
            return

        self.rubric_items = []
        self.manual_mappings = {}
        self.grades = {}
        try:
            if self.state_path.exists():
                self.state_path.unlink()
        except Exception as exc:
            messagebox.showerror("Reset Failed", f"Could not remove state file:\n{exc}")
            return
        self._load_data()
        messagebox.showinfo("State Reset", "Saved grading state was reset.")

    def _export_csv(self):
        self._persist_current_form(mark_graded=True)
        out_path = self.export_path
        fieldnames = [
            "Net ID",
            "First Name",
            "Last Name",
            "Email",
            "Submission File",
            "Status",
            "Score",
            "Selected Rubrics",
            "Extra Deduction",
            "Comments",
        ]
        with out_path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for s in self.students:
                rec = self._get_record(s["netid"], create=True)
                score = rec.get("score")
                writer.writerow(
                    {
                        "Net ID": s["netid"],
                        "First Name": s["first"],
                        "Last Name": s["last"],
                        "Email": s["email"],
                        "Submission File": Path(s["submission"]).name if s["submission"] else "",
                        "Status": rec.get("status", "ungraded"),
                        "Score": "" if score is None else self._fmt(float(score)),
                        "Selected Rubrics": "; ".join(rec.get("selected_rubrics", [])),
                        "Extra Deduction": self._fmt(self._safe_float(rec.get("extra_deduction", 0.0), 0.0)),
                        "Comments": rec.get("comments", ""),
                    }
                )
        messagebox.showinfo("Export complete", f"Wrote:\n{out_path}")


def main():
    root = tk.Tk()
    QuizGraderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
