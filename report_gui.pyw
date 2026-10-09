"""
Report Helper - a window for the supervisor reports.

Double-click this file (or run: pythonw report_gui.pyw) to:
    - pick the hours workbook, the levels file and the output folder
    - change the words used to find each column in the Excel
    - add, remove, rename and reorder the columns of the Word table, and pick
      which Excel column fills each one
    - ignore the supervisor column and type the supervisor's name instead
    - check how each sheet is read before making the reports

Settings are kept per profile ("Weekly", "Monthly", or your own copies) in
report_settings.json next to this file.

Needs: pip install pandas openpyxl python-docx
"""

import copy
import os
import queue
import sys
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk

HERE = Path(__file__).resolve().parent
SETTINGS_FILE = HERE / "report_settings.json"
sys.path.insert(0, str(HERE))          # so double-clicking works from anywhere

try:
    import report_engine as eng
    import make_levels
except ImportError as e:
    # Double-clicking starts the main Python, not a virtual environment. If
    # there is one next to this file (or one folder up), restart with it.
    if not os.environ.get("REPORT_HELPER_RELAUNCHED"):
        import subprocess
        for folder in (HERE, HERE.parent):
            for venv in (".venv", "venv", "env"):
                for exe in ("Scripts/pythonw.exe", "Scripts/python.exe", "bin/python"):
                    candidate = folder / venv / exe
                    if candidate.exists() and candidate.resolve() != Path(sys.executable).resolve():
                        subprocess.Popen([str(candidate), str(Path(__file__).resolve())],
                                         env={**os.environ, "REPORT_HELPER_RELAUNCHED": "1"})
                        sys.exit(0)
    # Otherwise say which Python is running and how to install into exactly that one
    _root = tk.Tk()
    _root.withdraw()
    package = {"docx": "python-docx"}.get(e.name, e.name)
    messagebox.showerror("Report Helper",
                         f"The package '{package}' is missing in this Python:\n"
                         f"    {sys.executable}\n\n"
                         "Install the packages into it with:\n"
                         f'    "{sys.executable}" -m pip install pandas openpyxl python-docx\n\n'
                         "If the packages are in a virtual environment, put it in this "
                         "folder as '.venv' or 'venv' (it is then used automatically), "
                         "or start the window from it:\n"
                         r"    <venv>\Scripts\python report_gui.pyw")
    sys.exit(1)

LAYOUTS = {"weekly": "One page per week (weekly report)",
           "monthly": "Title box, then every week (monthly report)"}
WEEKS_FROM = {"sheet": "Each sheet is one week (named after the sheet)",
              "date": "From the Date column (Monday to Sunday)"}
LEVEL_TEXT = "Service level (Basic / Intermediate / Senior)"
BLANK_TEXT = "Empty - filled in by hand"
EXCEL_PREFIX = "Excel: "
DONE = object()                         # end-of-task marker in the log queue


def split_words(text):
    return [w.strip() for w in text.split(",") if w.strip()]


def source_text(source):
    if source == eng.LEVEL_SOURCE:
        return LEVEL_TEXT
    if source == eng.BLANK_SOURCE:
        return BLANK_TEXT
    return EXCEL_PREFIX + source


def source_from_text(text):
    if text == LEVEL_TEXT:
        return eng.LEVEL_SOURCE
    if text == BLANK_TEXT:
        return eng.BLANK_SOURCE
    return text[len(EXCEL_PREFIX):]


# ---------------------------------------------------------------- dialogs
class Dialog(tk.Toplevel):
    """Small modal window with OK / Cancel; subclasses fill self.body."""

    def __init__(self, parent, title):
        super().__init__(parent)
        self.withdraw()
        self.transient(parent)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        self.body = ttk.Frame(self, padding=12)
        self.body.pack(fill="both", expand=True)
        buttons = ttk.Frame(self, padding=(12, 0, 12, 12))
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(buttons, text="OK", command=self._ok).pack(side="right", padx=6)
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())

    def show(self):
        self.update_idletasks()
        parent = self.master
        x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_reqwidth()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_reqheight()) // 3
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        self.deiconify()
        self.grab_set()
        self.wait_window()
        return self.result

    def _ok(self):
        result = self.validate()
        if result is not None:
            self.result = result
            self.destroy()

    def validate(self):
        raise NotImplementedError


class FieldDialog(Dialog):
    """Add / edit an Excel column: its name and the words used to find it."""

    def __init__(self, parent, field, taken_names):
        super().__init__(parent, "Excel column")
        self.field = field
        self.taken = taken_names
        built_in = field["name"] in eng.BUILT_IN
        b = self.body

        ttk.Label(b, text="Name:").grid(padx=(0, 8), row=0, column=0, sticky="w", pady=3)
        self.name = tk.StringVar(value=field["name"])
        e = ttk.Entry(b, textvariable=self.name, width=50)
        e.grid(row=0, column=1, sticky="we", pady=3)
        if built_in:
            e.state(["disabled"])

        ttk.Label(b, text="Words to look for:").grid(padx=(0, 8), row=1, column=0, sticky="w", pady=3)
        self.words = tk.StringVar(value=", ".join(field["words"]))
        words_entry = ttk.Entry(b, textvariable=self.words, width=50)
        words_entry.grid(row=1, column=1, sticky="we", pady=3)

        ttk.Label(b, text="Skip headers containing:").grid(padx=(0, 8), row=2, column=0, sticky="w", pady=3)
        self.ignore = tk.StringVar(value=", ".join(field.get("ignore", [])))
        ttk.Entry(b, textvariable=self.ignore, width=50).grid(row=2, column=1, sticky="we", pady=3)

        self.required = tk.BooleanVar(value=eng.is_required(field))
        cb = ttk.Checkbutton(b, variable=self.required,
                             text="Required - skip a sheet that doesn't have this column")
        cb.grid(row=3, column=1, sticky="w", pady=(6, 3))
        if built_in:
            cb.state(["disabled"])

        ttk.Label(b, foreground="gray", wraplength=430, justify="left", text=(
            "Separate several words with commas. Not case-sensitive. A header "
            "matches if it is exactly one of the words, or contains one of them "
            "(the first word is tried first). Example: 'jira' finds "
            "'JIRA ID/Test Rail Test Run'.\nUse 'Skip headers containing' to "
            "avoid look-alikes, e.g. Name skips 'Supervisor Name'.")
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))
        (words_entry if built_in else e).focus_set()

    def validate(self):
        name = self.name.get().strip()
        words = split_words(self.words.get())
        if not name:
            messagebox.showwarning("Excel column", "Give the column a name.", parent=self)
            return None
        if name.startswith("@") or name.lower() in {n.lower() for n in self.taken}:
            messagebox.showwarning("Excel column",
                                   f"There is already a column called '{name}'.", parent=self)
            return None
        if not words:
            messagebox.showwarning("Excel column",
                                   "Write at least one word to look for.", parent=self)
            return None
        return {"name": name, "words": words, "ignore": split_words(self.ignore.get()),
                "required": self.required.get()}


class ColumnDialog(Dialog):
    """Add / edit a Word table column: header, where it's filled from, width."""

    def __init__(self, parent, column, choices):
        super().__init__(parent, "Word table column")
        b = self.body

        ttk.Label(b, text="Header in Word:").grid(padx=(0, 8), row=0, column=0, sticky="w", pady=3)
        self.header = tk.StringVar(value=column["header"])
        e = ttk.Entry(b, textvariable=self.header, width=60)
        e.grid(row=0, column=1, sticky="we", pady=3)

        ttk.Label(b, text="Filled from:").grid(padx=(0, 8), row=1, column=0, sticky="w", pady=3)
        self.source = tk.StringVar(value=source_text(column["source"]))
        ttk.Combobox(b, textvariable=self.source, values=choices, state="readonly",
                     width=57).grid(row=1, column=1, sticky="we", pady=3)

        ttk.Label(b, text="Width (cm):").grid(padx=(0, 8), row=2, column=0, sticky="w", pady=3)
        self.width = tk.StringVar(value=str(column["width"]))
        ttk.Spinbox(b, textvariable=self.width, from_=0.5, to=25, increment=0.1,
                    width=8).grid(row=2, column=1, sticky="w", pady=3)

        ttk.Label(b, foreground="gray", wraplength=470, justify="left", text=(
            "Each level (Basic / Intermediate / Senior) gets one row; a column "
            "filled from Excel lists one bullet per activity. Excel columns are "
            "set up in the 'Excel columns' tab.")
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))
        e.focus_set()

    def validate(self):
        header = self.header.get().strip()
        if not header:
            messagebox.showwarning("Word column", "Write the column header.", parent=self)
            return None
        try:
            width = round(float(self.width.get().replace(",", ".")), 2)
            if width <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Word column", "The width must be a number above 0.",
                                   parent=self)
            return None
        return {"header": header, "source": source_from_text(self.source.get()),
                "width": width}


# ------------------------------------------------------------- main window
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Report Helper")
        self.geometry("980x860")
        self.minsize(840, 700)
        self.settings = eng.load_settings(SETTINGS_FILE)
        self.profile = None                    # the profile dict being edited
        self.log_queue = queue.Queue()
        self.busy = False

        self.v = {k: tk.StringVar() for k in
                  ("hours_file", "levels_file", "out_dir", "file_name",
                   "supervisor_mode", "supervisor_name", "first_week")}
        self.layout_text = tk.StringVar()
        self.weeks_text = tk.StringVar()

        self._build()
        self._show_profile(self.settings["profile"])
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._drain_log)

    # ------------------------------------------------------------ layout
    def _build(self):
        bar = ttk.Frame(self, padding=(10, 10, 10, 4))
        bar.pack(fill="x")
        ttk.Label(bar, text="Profile:").pack(side="left")
        self.profile_box = ttk.Combobox(bar, state="readonly", width=26)
        self.profile_box.pack(side="left", padx=6)
        self.profile_box.bind("<<ComboboxSelected>>",
                              lambda e: self._switch_profile(self.profile_box.get()))
        for text, cmd in (("New copy…", self._new_profile), ("Rename…", self._rename_profile),
                          ("Delete", self._delete_profile),
                          ("Restore default columns", self._restore_defaults)):
            ttk.Button(bar, text=text, command=cmd).pack(side="left", padx=2)

        paned = ttk.PanedWindow(self, orient="vertical")
        paned.pack(fill="both", expand=True, padx=10, pady=4)
        nb = ttk.Notebook(paned)
        paned.add(nb, weight=4)
        nb.add(self._build_files_tab(nb), text="  Files & supervisor  ")
        nb.add(self._build_fields_tab(nb), text="  Excel columns  ")
        nb.add(self._build_columns_tab(nb), text="  Word table  ")

        bottom = ttk.Frame(paned)
        paned.add(bottom, weight=1)
        actions = ttk.Frame(bottom, padding=(0, 6))
        actions.pack(fill="x")
        style = ttk.Style(self)
        style.configure("Big.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 4))
        self.make_btn = ttk.Button(actions, text="Make reports", style="Big.TButton",
                                   command=self._make_reports)
        self.make_btn.pack(side="left")
        ttk.Button(actions, text="Open output folder",
                   command=self._open_output).pack(side="left", padx=6)
        ttk.Button(actions, text="Save settings",
                   command=self._save).pack(side="left")
        ttk.Button(actions, text="Clear log",
                   command=lambda: self.log.delete("1.0", "end")).pack(side="right")
        self.status = ttk.Label(actions, foreground="gray")
        self.status.pack(side="left", padx=12)

        self.log = scrolledtext.ScrolledText(bottom, height=8, font=("Consolas", 9),
                                             wrap="word")
        self.log.pack(fill="both", expand=True)
        self.log.tag_configure("ok", foreground="#1a7f37")
        self.log.tag_configure("warn", foreground="#b35900")
        self.log.tag_configure("err", foreground="#c62828")
        self.log.tag_configure("head", font=("Consolas", 9, "bold"))

    def _build_files_tab(self, parent):
        tab = ttk.Frame(parent, padding=10)
        tab.columnconfigure(0, weight=1)

        files = ttk.LabelFrame(tab, text="Files", padding=8)
        files.grid(row=0, column=0, sticky="we")
        files.columnconfigure(1, weight=1)

        def row(r, label, key, browse, extra=()):
            ttk.Label(files, text=label).grid(row=r, column=0, sticky="w", pady=3)
            ttk.Entry(files, textvariable=self.v[key]).grid(row=r, column=1, sticky="we",
                                                            padx=6, pady=3)
            ttk.Button(files, text="Browse…", command=browse).grid(row=r, column=2, pady=3)
            for i, (text, cmd) in enumerate(extra):
                ttk.Button(files, text=text, command=cmd).grid(row=r, column=3 + i,
                                                               padx=(4, 0), pady=3)

        row(0, "Hours workbook (Excel):", "hours_file", self._browse_hours,
            [("Check columns", self._check_excel)])
        row(1, "Levels file (Name | Level):", "levels_file", self._browse_levels,
            [("Create / update", self._update_levels), ("Open", self._open_levels)])
        row(2, "Output folder:", "out_dir", self._browse_out)
        ttk.Label(files, text="Word file name:").grid(row=3, column=0, sticky="w", pady=3)
        ttk.Entry(files, textvariable=self.v["file_name"]).grid(row=3, column=1, sticky="we",
                                                                padx=6, pady=3)
        ttk.Label(files, text="{supervisor} = the supervisor's name",
                  foreground="gray").grid(row=3, column=2, columnspan=3, sticky="w")

        sup = ttk.LabelFrame(tab, text="Supervisor", padding=8)
        sup.grid(row=1, column=0, sticky="we", pady=10)
        sup.columnconfigure(1, weight=1)
        ttk.Radiobutton(sup, text="Read it from the supervisor column in the Excel "
                                  "(one report per supervisor)",
                        variable=self.v["supervisor_mode"], value="column",
                        command=self._mode_changed).grid(row=0, column=0, columnspan=2,
                                                         sticky="w", pady=2)
        ttk.Radiobutton(sup, text="Ignore the supervisor column - everything goes to:",
                        variable=self.v["supervisor_mode"], value="fixed",
                        command=self._mode_changed).grid(row=1, column=0, sticky="w", pady=2)
        self.sup_entry = ttk.Entry(sup, textvariable=self.v["supervisor_name"], width=40)
        self.sup_entry.grid(row=1, column=1, sticky="w", padx=6)

        weeks = ttk.LabelFrame(tab, text="Weeks and page layout", padding=8)
        weeks.grid(row=2, column=0, sticky="we")
        ttk.Label(weeks, text="Word layout:").grid(row=0, column=0, sticky="w", pady=3)
        box = ttk.Combobox(weeks, textvariable=self.layout_text, state="readonly",
                           values=list(LAYOUTS.values()), width=48)
        box.grid(row=0, column=1, columnspan=2, sticky="w", padx=6, pady=3)
        box.bind("<<ComboboxSelected>>", lambda e: self._mode_changed())
        ttk.Label(weeks, text="Weeks come from:").grid(row=1, column=0, sticky="w", pady=3)
        box = ttk.Combobox(weeks, textvariable=self.weeks_text, state="readonly",
                           values=list(WEEKS_FROM.values()), width=48)
        box.grid(row=1, column=1, columnspan=2, sticky="w", padx=6, pady=3)
        box.bind("<<ComboboxSelected>>", lambda e: self._mode_changed())
        ttk.Label(weeks, text="First week number:").grid(row=2, column=0, sticky="w", pady=3)
        self.first_week_entry = ttk.Entry(weeks, textvariable=self.v["first_week"], width=8)
        self.first_week_entry.grid(row=2, column=1, sticky="w", padx=6, pady=3)
        ttk.Label(weeks, text="empty = calendar (ISO) week numbers",
                  foreground="gray").grid(row=2, column=2, sticky="w")
        return tab

    def _tree_with_buttons(self, parent, columns, buttons):
        frame = ttk.Frame(parent)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        tree = ttk.Treeview(frame, columns=[c for c, _, _ in columns], show="headings",
                            selectmode="browse")
        for key, text, width in columns:
            tree.heading(key, text=text, anchor="w")
            tree.column(key, width=width, stretch=width > 150, anchor="w")
        tree.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=sb.set)
        side = ttk.Frame(frame, padding=(8, 0, 0, 0))
        side.grid(row=0, column=2, sticky="n")
        for text, cmd in buttons:
            if text is None:
                ttk.Separator(side).pack(fill="x", pady=6)
            else:
                ttk.Button(side, text=text, command=cmd, width=16).pack(pady=2)
        return frame, tree

    def _build_fields_tab(self, parent):
        tab = ttk.Frame(parent, padding=10)
        ttk.Label(tab, wraplength=860, justify="left", text=(
            "The columns to find in every Excel sheet and the words used to find them. "
            "The header row can be anywhere in the first 60 rows. Name, Supervisor and "
            "Date are always there; add your own for anything you want in the Word "
            "table. Double-click a row to edit it.")).pack(fill="x", pady=(0, 8))
        frame, self.fields_tree = self._tree_with_buttons(
            tab,
            [("name", "Excel column", 150), ("words", "Words to look for", 260),
             ("ignore", "Skip headers containing", 230), ("req", "Required", 160)],
            [("Add…", self._add_field), ("Edit…", self._edit_field),
             ("Remove", self._remove_field), (None, None),
             ("Move up", lambda: self._move_field(-1)),
             ("Move down", lambda: self._move_field(1)), (None, None),
             ("Check Excel file", self._check_excel)])
        frame.pack(fill="both", expand=True)
        self.fields_tree.bind("<Double-1>", lambda e: self._edit_field())
        ttk.Label(tab, foreground="gray", text=(
            "Columns higher in the list pick their header first. 'Check Excel file' "
            "shows which header each column found, sheet by sheet.")
        ).pack(fill="x", pady=(6, 0))
        return tab

    def _build_columns_tab(self, parent):
        tab = ttk.Frame(parent, padding=10)
        ttk.Label(tab, wraplength=860, justify="left", text=(
            "The Word table, from left to right. Each one has a header, what fills it "
            "and its width. Double-click a row to edit it.")).pack(fill="x", pady=(0, 8))
        frame, self.columns_tree = self._tree_with_buttons(
            tab,
            [("n", "#", 36), ("header", "Header in Word", 380),
             ("source", "Filled from", 260), ("width", "Width (cm)", 80)],
            [("Add…", self._add_column), ("Edit…", self._edit_column),
             ("Remove", self._remove_column), (None, None),
             ("Move left (up)", lambda: self._move_column(-1)),
             ("Move right (down)", lambda: self._move_column(1))])
        frame.pack(fill="both", expand=True)
        self.columns_tree.bind("<Double-1>", lambda e: self._edit_column())
        self.width_label = ttk.Label(tab)
        self.width_label.pack(fill="x", pady=(6, 0))
        return tab

    # ---------------------------------------------------------- profiles
    def _show_profile(self, name):
        self.settings["profile"] = name
        self.profile = self.settings["profiles"][name]
        self.profile_box["values"] = list(self.settings["profiles"])
        self.profile_box.set(name)
        p = self.profile
        for key in ("hours_file", "levels_file", "out_dir", "file_name",
                    "supervisor_mode", "supervisor_name"):
            self.v[key].set(p.get(key) or "")
        self.v["first_week"].set("" if p.get("first_week") is None else str(p["first_week"]))
        self.layout_text.set(LAYOUTS[p["layout"]])
        self.weeks_text.set(WEEKS_FROM[p["weeks_from"]])
        self._mode_changed()

    def _store(self, strict=True):
        """Copy the entry boxes into the profile. Returns False if something's invalid."""
        p = self.profile
        for key in ("hours_file", "levels_file", "out_dir", "file_name",
                    "supervisor_mode", "supervisor_name"):
            p[key] = self.v[key].get().strip()
        p["layout"] = next(k for k, t in LAYOUTS.items() if t == self.layout_text.get())
        p["weeks_from"] = next(k for k, t in WEEKS_FROM.items() if t == self.weeks_text.get())
        text = self.v["first_week"].get().strip()
        if not text:
            p["first_week"] = None
        elif text.isdigit():
            p["first_week"] = int(text)
        elif strict:
            messagebox.showwarning("Report Helper",
                                   "The first week number must be a whole number "
                                   "(or empty for calendar week numbers).")
            return False
        return True

    def _switch_profile(self, name):
        self._store(strict=False)
        self._show_profile(name)

    def _ask_profile_name(self, title, initial):
        name = simpledialog.askstring(title, "Profile name:", initialvalue=initial,
                                      parent=self)
        name = (name or "").strip()
        if name and name in self.settings["profiles"]:
            messagebox.showwarning(title, f"There is already a profile called '{name}'.")
            return None
        return name or None

    def _new_profile(self):
        self._store(strict=False)
        name = self._ask_profile_name("New profile", f"{self.profile_box.get()} copy")
        if name:
            self.settings["profiles"][name] = copy.deepcopy(self.profile)
            self._show_profile(name)

    def _rename_profile(self):
        old = self.profile_box.get()
        name = self._ask_profile_name("Rename profile", old)
        if name:
            self._store(strict=False)
            self.settings["profiles"] = {(name if k == old else k): v
                                         for k, v in self.settings["profiles"].items()}
            self._show_profile(name)

    def _delete_profile(self):
        if len(self.settings["profiles"]) == 1:
            messagebox.showinfo("Delete profile", "This is the only profile.")
            return
        name = self.profile_box.get()
        if messagebox.askyesno("Delete profile", f"Delete the profile '{name}'?"):
            del self.settings["profiles"][name]
            self._show_profile(next(iter(self.settings["profiles"])))

    def _restore_defaults(self):
        self._store(strict=False)
        layout = self.profile["layout"]
        default = eng.default_profile(layout)
        if not messagebox.askyesno(
                "Restore defaults",
                f"Put back the default {layout} Excel columns, Word table, file name "
                "and week settings?\nYour file choices and supervisor stay as they are."):
            return
        for key in ("fields", "columns", "file_name", "weeks_from", "first_week"):
            self.profile[key] = default[key]
        self._show_profile(self.profile_box.get())

    # ---------------------------------------------------- refresh / state
    def _mode_changed(self):
        fixed = self.v["supervisor_mode"].get() == "fixed"
        self.sup_entry.state(["!disabled"] if fixed else ["disabled"])
        dated = self.weeks_text.get() == WEEKS_FROM["date"]
        self.first_week_entry.state(["!disabled"] if dated else ["disabled"])
        if self.profile is not None:
            self._store(strict=False)
            self._refresh_fields()
            self._refresh_columns()

    def _refresh_fields(self, select=None):
        tree = self.fields_tree
        tree.delete(*tree.get_children())
        active = {f["name"] for f in eng.active_fields(self.profile)}
        for i, f in enumerate(self.profile["fields"]):
            if f["name"] not in active:
                req = ("not used (typed)" if f["name"] == eng.SUPERVISOR
                       else "not used (sheets)")
            elif f["name"] == eng.DATE:
                req = "no"
            else:
                req = "yes" if eng.is_required(f) else "no"
            tree.insert("", "end", iid=str(i), values=(
                f["name"], ", ".join(f["words"]), ", ".join(f.get("ignore", [])), req))
        if select is not None:
            tree.selection_set(str(select))
            tree.see(str(select))

    def _refresh_columns(self, select=None):
        tree = self.columns_tree
        tree.delete(*tree.get_children())
        for i, c in enumerate(self.profile["columns"]):
            tree.insert("", "end", iid=str(i), values=(
                i + 1, c["header"], source_text(c["source"]), f"{float(c['width']):g}"))
        if select is not None:
            tree.selection_set(str(select))
            tree.see(str(select))
        total = sum(float(c["width"]) for c in self.profile["columns"])
        room = eng.usable_width_cm(self.profile["layout"])
        over = total > room + 0.05
        self.width_label.configure(
            foreground="#c62828" if over else "gray",
            text=f"Total width {total:.1f} cm - the page has room for about {room:.1f} cm"
                 + ("  (too wide: make some columns narrower)" if over else ""))

    def _selected(self, tree):
        sel = tree.selection()
        return int(sel[0]) if sel else None

    # ------------------------------------------------------ Excel fields
    def _source_choices(self):
        """What a Word column can be filled from."""
        names = [f["name"] for f in self.profile["fields"]
                 if f["name"] not in (eng.SUPERVISOR, eng.DATE)]
        return [LEVEL_TEXT, BLANK_TEXT] + [EXCEL_PREFIX + n for n in names]

    def _add_field(self):
        field = {"name": "", "words": [], "ignore": [], "required": True}
        taken = [f["name"] for f in self.profile["fields"]]
        result = FieldDialog(self, field, taken).show()
        if result:
            self.profile["fields"].append(result)
            self._refresh_fields(len(self.profile["fields"]) - 1)

    def _edit_field(self):
        i = self._selected(self.fields_tree)
        if i is None:
            return
        field = self.profile["fields"][i]
        taken = [f["name"] for j, f in enumerate(self.profile["fields"]) if j != i]
        result = FieldDialog(self, field, taken).show()
        if not result:
            return
        old = field["name"]
        if old in eng.BUILT_IN:
            result["name"], result["required"] = old, field.get("required", True)
        self.profile["fields"][i] = result
        if result["name"] != old:                 # renamed: keep the Word columns linked
            for c in self.profile["columns"]:
                if c["source"] == old:
                    c["source"] = result["name"]
        self._refresh_fields(i)
        self._refresh_columns()

    def _remove_field(self):
        i = self._selected(self.fields_tree)
        if i is None:
            return
        name = self.profile["fields"][i]["name"]
        if name in eng.BUILT_IN:
            messagebox.showinfo("Remove", f"'{name}' is built in and can't be removed "
                                          "(you can change its words).")
            return
        used = [c["header"] for c in self.profile["columns"] if c["source"] == name]
        msg = f"Remove the Excel column '{name}'?"
        if used:
            msg += ("\n\nThese Word columns are filled from it and will be left empty:\n  - "
                    + "\n  - ".join(used))
        if not messagebox.askyesno("Remove", msg):
            return
        del self.profile["fields"][i]
        for c in self.profile["columns"]:
            if c["source"] == name:
                c["source"] = eng.BLANK_SOURCE
        self._refresh_fields()
        self._refresh_columns()

    def _move_field(self, step):
        i = self._selected(self.fields_tree)
        fields = self.profile["fields"]
        if i is None or not 0 <= i + step < len(fields):
            return
        fields[i], fields[i + step] = fields[i + step], fields[i]
        self._refresh_fields(i + step)

    # ------------------------------------------------------ Word columns
    def _add_column(self):
        column = {"header": "", "source": eng.BLANK_SOURCE, "width": 2.5}
        result = ColumnDialog(self, column, self._source_choices()).show()
        if result:
            i = self._selected(self.columns_tree)
            cols = self.profile["columns"]
            pos = len(cols) if i is None else i + 1     # after the selected one
            cols.insert(pos, result)
            self._refresh_columns(pos)

    def _edit_column(self):
        i = self._selected(self.columns_tree)
        if i is None:
            return
        result = ColumnDialog(self, self.profile["columns"][i], self._source_choices()).show()
        if result:
            self.profile["columns"][i] = result
            self._refresh_columns(i)

    def _remove_column(self):
        i = self._selected(self.columns_tree)
        if i is None:
            return
        header = self.profile["columns"][i]["header"]
        if messagebox.askyesno("Remove", f"Remove the Word column '{header}'?"):
            del self.profile["columns"][i]
            self._refresh_columns(min(i, len(self.profile["columns"]) - 1)
                                  if self.profile["columns"] else None)

    def _move_column(self, step):
        i = self._selected(self.columns_tree)
        cols = self.profile["columns"]
        if i is None or not 0 <= i + step < len(cols):
            return
        cols[i], cols[i + step] = cols[i + step], cols[i]
        self._refresh_columns(i + step)

    # -------------------------------------------------------------- files
    def _initial_dir(self, key):
        value = self.v[key].get().strip()
        if value:
            path = self._resolve(value)
            folder = path if path.is_dir() else path.parent
            if folder.exists():
                return str(folder)
        return str(HERE)

    def _browse_hours(self):
        path = filedialog.askopenfilename(
            parent=self, title="Excel workbook with the hours",
            initialdir=self._initial_dir("hours_file"),
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")])
        if path:
            self.v["hours_file"].set(os.path.normpath(path))

    def _browse_levels(self):
        # A save dialog, so a new levels file can be named too
        path = filedialog.asksaveasfilename(
            parent=self, title="Levels file (pick one, or type a new name)",
            initialdir=self._initial_dir("levels_file"), confirmoverwrite=False,
            defaultextension=".xlsx", initialfile="engineer_levels.xlsx",
            filetypes=[("Excel workbooks", "*.xlsx"), ("All files", "*.*")])
        if path:
            self.v["levels_file"].set(os.path.normpath(path))

    def _browse_out(self):
        path = filedialog.askdirectory(parent=self, title="Output folder",
                                       initialdir=self._initial_dir("out_dir"))
        if path:
            self.v["out_dir"].set(os.path.normpath(path))

    @staticmethod
    def _resolve(value):
        """Relative paths are relative to this program's folder."""
        path = Path(value)
        return path if path.is_absolute() else HERE / path

    def _resolved_profile(self):
        p = copy.deepcopy(self.profile)
        for key in ("hours_file", "levels_file"):
            if p[key]:
                p[key] = str(self._resolve(p[key]))
        p["out_dir"] = str(self._resolve(p["out_dir"] or "reports"))
        return p

    def _open_path(self, path, what):
        if not path.exists():
            messagebox.showinfo("Report Helper", f"The {what} doesn't exist yet:\n{path}")
            return
        try:
            os.startfile(path)
        except OSError as e:
            messagebox.showerror("Report Helper", f"Can't open {path}:\n{e}")

    def _open_output(self):
        self._store(strict=False)
        self._open_path(Path(self._resolved_profile()["out_dir"]), "output folder")

    def _open_levels(self):
        self._store(strict=False)
        levels = self._resolved_profile()["levels_file"]
        if not levels:
            messagebox.showinfo("Report Helper", "Choose the levels file first.")
            return
        self._open_path(Path(levels), "levels file")

    # -------------------------------------------------------------- tasks
    def _run_task(self, title, func):
        """Run func(profile, log) in the background, logging into the window."""
        if self.busy or not self._store():
            return
        self._save()
        profile = self._resolved_profile()
        self._set_busy(True, f"{title}…")
        self.log_queue.put(("head", f"\n=== {title} ({self.profile_box.get()}) ==="))

        def log(text=""):
            self.log_queue.put((None, str(text)))

        def work():
            try:
                func(profile, log)
            except eng.ReportError as e:
                self.log_queue.put(("err", f"  ✗ {e}"))
            except Exception:
                self.log_queue.put(("err", "  ✗ Unexpected error:\n" + traceback.format_exc()))
            finally:
                self.log_queue.put(DONE)

        threading.Thread(target=work, daemon=True).start()

    def _make_reports(self):
        self._run_task("Making reports", lambda p, log: eng.build_reports(p, log))

    def _check_excel(self):
        self._run_task("Checking Excel columns", lambda p, log: eng.inspect_workbook(p, log))

    def _update_levels(self):
        def run(p, log):
            make_levels.update_levels(p["hours_file"], p["levels_file"], p, log)
        if not self.v["levels_file"].get().strip():
            self.v["levels_file"].set("engineer_levels.xlsx")
        self._run_task("Creating / updating the levels file", run)

    def _set_busy(self, busy, text=""):
        self.busy = busy
        self.make_btn.state(["disabled"] if busy else ["!disabled"])
        self.status.configure(text=text)
        self.configure(cursor="watch" if busy else "")

    def _drain_log(self):
        try:
            while True:
                item = self.log_queue.get_nowait()
                if item is DONE:
                    self._set_busy(False, "Done.")
                    continue
                tag, text = item
                for line in text.split("\n"):
                    t = tag
                    if t is None:
                        s = line.strip()
                        t = ("ok" if s.startswith("✓") else
                             "warn" if s.startswith(("!", "- Skipping")) else None)
                    self.log.insert("end", line + "\n", t or ())
                self.log.see("end")
        except queue.Empty:
            pass
        self.after(100, self._drain_log)

    # ------------------------------------------------------------ settings
    def _save(self):
        self._store(strict=False)
        try:
            eng.save_settings(SETTINGS_FILE, self.settings)
            self.status.configure(text="Settings saved.")
        except OSError as e:
            messagebox.showerror("Report Helper", f"Can't save the settings:\n{e}")

    def _on_close(self):
        self._store(strict=False)
        try:
            eng.save_settings(SETTINGS_FILE, self.settings)
        except OSError:
            pass
        self.destroy()

    def report_callback_exception(self, exc, val, tb):
        # pythonw has no console, so show errors instead of losing them
        messagebox.showerror("Report Helper",
                             "".join(traceback.format_exception(exc, val, tb))[-2000:])


def main():
    try:                                   # sharp text on high-DPI screens
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App().mainloop()


if __name__ == "__main__":
    main()
