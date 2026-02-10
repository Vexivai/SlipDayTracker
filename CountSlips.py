

"""
SlipDayCounter / CountSlips.py

Purpose
-------
Count earned "slip days" from a Canvas gradebook export CSV.
- Attendance columns: "MM-DD (#####)" with values like 1.00, 0.00, EX.
  (EX is counted as attendance only if included in CONFIRMED_ATTENDANCE_VALUES.)
- Every 5 confirmed attendances earns 1 slip day.

Workflow
--------
- Load latest CSV from ./import (or --csv)
- Load save.txt if it exists
- If roster changed, choose tracked students
- Show GUI table of tracked students
- Save and update per-student state
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import time
import threading
import sys
from dataclasses import dataclass, field
from pathlib import Path

from typing import Any


# Optional GUI (Tkinter is stdlib). We import lazily in gui_main() to avoid issues in headless runs.

def _maybe_reexec_into_venv(reason: str) -> None:
    """If we're not in a venv and a project venv exists, re-exec into it.

    This prevents "No module named _tkinter" (Homebrew python) and missing pandas issues.
    """
    script_dir = Path(__file__).resolve().parent
    venv_py = script_dir / ".venv" / "bin" / "python"
    in_venv = (sys.prefix != sys.base_prefix)
    if (not in_venv) and venv_py.exists():
        env = dict(os.environ)
        # Prevent any other wrappers from trying to delegate again.
        env["SLIPDAYCOUNTER_SKIP_MAKE"] = "1"
        env["SLIPDAYCOUNTER_REEXEC_REASON"] = reason
        os.execve(str(venv_py), [str(venv_py), *sys.argv], env)


# -----------------------------
# Configuration
# -----------------------------

# Values in attendance cells that count as "confirmed attendance".
# Based on the sample CSV: 1.00 is present, 0.00 is absent, EX is excused.
# Default: "1" counts as present.
# If your policy treats EX (excused) as present for slip-day accrual, include it here.
CONFIRMED_ATTENDANCE_VALUES = {"1", "1.0", "1.00", "1.000", "1.0000", "EX"}

# How many attendances are required to earn 1 slip day.
ATTENDANCES_PER_SLIP = 5

# Save file name (kept next to this script by default).
DEFAULT_SAVE_NAME = "save.txt"

# Save esc hold time
SAVEESCTIME = 800

# -----------------------------
# Data Models
# -----------------------------

@dataclass
class StudentState:
    key: str  # stable id: sid:<id> or name:<name>#<n>
    name: str
    student_id: str = ""  # Canvas column: "SIS User ID"
    tracked: bool = True

    # Consumable slip days
    used: int = 0

    # Optional metadata
    notes: str = ""
    dropped: bool = False

    # Cached (recomputed from CSV each run)
    attendance_confirmed: int = 0
    earned: int = 0

    @property
    def available(self) -> int:
        return max(0, self.earned - self.used)


@dataclass
class SaveState:
    """Persisted UI state + per-student overrides."""
    roster_keys_in_order: list[str] = field(default_factory=list)  # Roster order from CSV
    students: dict[str, StudentState] = field(default_factory=dict)  # Per-student state by key
    last_csv_name: str = ""


# -----------------------------
# CSV Parsing
# -----------------------------

_ATTENDANCE_COL_RE = re.compile(r"^\s*(\d{1,2})-(\d{1,2})\s*\(")  # "1-12 (7448946)"


def find_latest_csv(import_dir: Path) -> Path | None:
    if not import_dir.exists() or not import_dir.is_dir():
        return None
    csvs = sorted(import_dir.glob("*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)
    return csvs[0] if csvs else None


def read_canvas_gradebook_csv(
    csv_path: Path,
) -> tuple[list[str], dict[str, int], dict[str, dict[str, str]]]:
    """
    Returns:
      (roster_keys_in_order, confirmed_attendance, student_meta)
      Keys: sid:<id> if available, else name:<name>#<n>; attendance columns auto-detected.
    """
    try:
        import pandas as pd  # type: ignore
    except ImportError as e:
        _maybe_reexec_into_venv("pandas")
        raise SystemExit(
            "Missing dependency: pandas.\n"
            "Run: make run (recommended)\n"
            "Or: ./.venv/bin/python -m pip install pandas"
        ) from e

    df = pd.read_csv(csv_path)

    if "Student" not in df.columns:
        raise ValueError("CSV does not appear to be a Canvas gradebook export (missing 'Student' column).")

    # Detect attendance columns
    attendance_cols = [c for c in df.columns if _ATTENDANCE_COL_RE.match(str(c))]
    if not attendance_cols:
        raise ValueError("No attendance columns detected (expected headers like 'MM-DD (#####)').")

    # Filter out header rows and NaN
    def is_real_student(val: Any) -> bool:
        if val is None:
            return False
        s = str(val).strip()
        if not s:
            return False
        sl = s.lower()
        if sl in {"nan", "points possible"}:
            return False
        return True
    df_students = df[df["Student"].apply(is_real_student)].copy()

    # Normalize name: "Last, First" -> "First Last"
    def normalize_name(raw: Any) -> str:
        s = str(raw).strip()
        if "," in s:
            parts = [p.strip() for p in s.split(",", 1)]
            if len(parts) == 2 and parts[1] and parts[0]:
                return f"{parts[1]} {parts[0]}".strip()
        return s
    df_students["__name__"] = df_students["Student"].apply(normalize_name)

    # Generate stable keys: prefer SIS ID, else name+occurrence
    student_meta: dict[str, dict[str, str]] = {}
    roster_keys_in_order: list[str] = []
    has_sis = "SIS User ID" in df_students.columns
    name_occurrence: dict[str, int] = {}
    def normalize_sis(raw_id: Any) -> str:
        if raw_id is None:
            return ""
        sid = str(raw_id).strip()
        if sid.lower() == "nan":
            sid = ""
        if sid.endswith(".0") and sid[:-2].isdigit():
            sid = sid[:-2]
        return sid
    for _, row in df_students.iterrows():
        name = str(row["__name__"])
        sid = normalize_sis(row.get("SIS User ID")) if has_sis else ""
        if sid:
            key = f"sid:{sid}"
        else:
            name_occurrence[name] = name_occurrence.get(name, 0) + 1
            key = f"name:{name}#{name_occurrence[name]}"
        roster_keys_in_order.append(key)
        student_meta[key] = {"name": name, "student_id": sid}

    # Sort keys alphabetically by LAST name (then first/middle), then student_id, then key.
    def last_name_sort_key(display_name: str) -> tuple[str, str]:
        parts = (display_name or "").strip().split()
        if not parts:
            return ("", "")
        last = parts[-1].lower()
        rest = " ".join(parts[:-1]).lower()
        return (last, rest)

    roster_keys_in_order.sort(
        key=lambda k: (
            last_name_sort_key(student_meta.get(k, {}).get("name", "")),
            student_meta.get(k, {}).get("student_id", ""),
            k,
        )
    )

    confirmed_attendance: dict[str, int] = {}
    # Count attendance per student, independent of sorted order
    name_occurrence_for_att: dict[str, int] = {}
    for _, row in df_students.iterrows():
        name = str(row["__name__"])
        sid = normalize_sis(row.get("SIS User ID")) if has_sis else ""
        if sid:
            key = f"sid:{sid}"
        else:
            name_occurrence_for_att[name] = name_occurrence_for_att.get(name, 0) + 1
            key = f"name:{name}#{name_occurrence_for_att[name]}"
        count = 0
        for col in attendance_cols:
            cell = row.get(col)
            if cell is None:
                continue
            cell_s = str(cell).strip()
            if cell_s in CONFIRMED_ATTENDANCE_VALUES:
                count += 1
        confirmed_attendance[key] = count
    return roster_keys_in_order, confirmed_attendance, student_meta

# -----------------------------
# Save / Load
# -----------------------------

def get_default_save_path(script_path: Path) -> Path:
    return script_path.with_name(DEFAULT_SAVE_NAME)


def load_save(save_path: Path) -> SaveState | None:
    if not save_path.exists():
        return None
    try:
        raw = save_path.read_text(encoding="utf-8")
        data = json.loads(raw)
        state = SaveState(
            roster_keys_in_order=list(data.get("roster_keys_in_order", [])),
            last_csv_name=str(data.get("last_csv_name", "")),
        )

        students_raw = data.get("students", {})

        # Back-compat: older saves keyed by name.
        for key_or_name, sdata in students_raw.items():
            # Determine stable key
            key = str(sdata.get("key", ""))
            if not key:
                # legacy: dict key was the name
                legacy_name = str(key_or_name)
                sid = str(sdata.get("student_id", "")).strip()
                if sid:
                    key = f"sid:{sid}"
                else:
                    key = f"name:{legacy_name}#1"
            name = str(sdata.get("name", key_or_name))

            st = StudentState(
                key=key,
                name=name,
                student_id=str(sdata.get("student_id", "")),
                tracked=bool(sdata.get("tracked", True)),
                used=int(sdata.get("used", 0)),
                notes=str(sdata.get("notes", "")),
                dropped=bool(sdata.get("dropped", False)),
            )
            state.students[key] = st


        return state
    except Exception as e:
        raise SystemExit(f"Failed to read {save_path.name}: {e}") from e


def save_state(save_path: Path, state: SaveState) -> None:
    payload: dict[str, Any] = {
        "roster_keys_in_order": state.roster_keys_in_order,
        "last_csv_name": state.last_csv_name,
        "students": {
            key: {
                "key": st.key,
                "name": st.name,
                "student_id": st.student_id,
                "tracked": st.tracked,
                "used": st.used,
                "notes": st.notes,
                "dropped": st.dropped,
            }
            for key, st in state.students.items()
        },
    }
    save_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")

# -----------------------------
# Slip Calculation + Reconciliation
# -----------------------------

def compute_earned(attendance_confirmed: int) -> int:
    return attendance_confirmed // ATTENDANCES_PER_SLIP


def reconcile_state_with_csv(
    roster_keys_in_order: list[str],
    attendance_confirmed: dict[str, int],
    student_meta: dict[str, dict[str, str]],
    prior: SaveState | None,
    selected_tracked: set[str] | None,
    csv_name: str,
) -> SaveState:
    """
    Build a new SaveState for this run.
    - If prior exists, carry over used/notes/dropped/tracked where possible
    - Apply selected_tracked if provided (new roster selection step)
    - Recompute attendance_confirmed + earned (source of truth = CSV)
    """
    state = SaveState(
        roster_keys_in_order=roster_keys_in_order,
        last_csv_name=csv_name,
    )

    prior_students = prior.students if prior else {}

    # Index prior by SIS ID for cross-run matching if keys changed
    prior_by_sid: dict[str, StudentState] = {}
    prior_by_name: dict[str, list[StudentState]] = {}
    for pst in prior_students.values():
        sid = (pst.student_id or "").strip()
        if sid:
            prior_by_sid[sid] = pst
        prior_by_name.setdefault(pst.name, []).append(pst)

    for key in roster_keys_in_order:
        meta = student_meta.get(key, {"name": "", "student_id": ""})
        name = str(meta.get("name", ""))
        sid = str(meta.get("student_id", ""))

        prev = prior_students.get(key)
        if prev is None and sid:
            prev = prior_by_sid.get(sid)
        if prev is None and name:
            # If unique name in prior, carry it over; otherwise treat as new.
            matches = prior_by_name.get(name, [])
            if len(matches) == 1:
                prev = matches[0]

        tracked = prev.tracked if prev else False
        used = prev.used if prev else 0
        notes = prev.notes if prev else ""
        dropped = prev.dropped if prev else False

        if selected_tracked is not None:
            tracked = key in selected_tracked

        st = StudentState(
            key=key,
            name=name,
            student_id=sid,
            tracked=tracked,
            used=used,
            notes=notes,
            dropped=dropped,
        )
        st.attendance_confirmed = int(attendance_confirmed.get(key, 0))
        st.earned = compute_earned(st.attendance_confirmed)

        if st.used > st.earned:
            st.used = st.earned

        state.students[key] = st

    return state


# -----------------------------
# Optional GUI (Tkinter)
# -----------------------------

def _tracked_sorted(state: SaveState) -> list[StudentState]:
    """Tracked students in roster order; dropped grouped last."""
    tracked = [s for s in state.students.values() if s.tracked]
    key_to_pos: dict[str, int] = {k: i for i, k in enumerate(state.roster_keys_in_order)}
    tracked.sort(key=lambda s: (s.dropped, key_to_pos.get(s.key, 10**9)))
    return tracked


def gui_main(
    csv_path: Path | None,
    save_path: Path,
    import_dir: Path,
) -> None:
    """GUI for slip day tracking: shows table, lets you use/refund slips, drop, edit notes, save. Main controls: buttons, Enter for notes."""

    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, simpledialog
        from tkinter import ttk
    except Exception as e:
        # Common on macOS when launched with Homebrew python: "No module named _tkinter".
        _maybe_reexec_into_venv("tkinter")
        raise SystemExit(
            "GUI requested, but Tkinter could not be imported.\n"
            "Try running via the project venv (make run) or reinstall Python with Tk.\n"
            f"Error: {e}"
        ) from e

    sort_col: str = "#"        # Treeview column identifier
    sort_desc: bool = False    # False = ascending

    def _coerce_for_sort(col: str, st: StudentState, roster_idx: int) -> object:
        """Return a comparable sort key for the given column."""
        def last_name_sort_key(display_name: str) -> tuple[str, str]:
            parts = (display_name or "").strip().split()
            if not parts:
                return ("", "")
            last = parts[-1].lower()
            rest = " ".join(parts[:-1]).lower()
            return (last, rest)
        if col == "#":
            return roster_idx
        if col == "Name":
            return last_name_sort_key(st.name or "")
        if col == "Student ID":
            return (st.student_id or "").lower()
        if col == "Attended":
            return int(st.attendance_confirmed)
        if col == "Earned Slips":
            return int(st.earned)
        if col == "Used Slips":
            return int(st.used)
        if col == "Available Slips":
            return int(st.available)
        if col == "Notes":
            return (st.notes or "").lower()
        return ""

    def _on_sort(col: str) -> None:
        nonlocal sort_col, sort_desc
        if sort_col == col:
            sort_desc = not sort_desc
        else:
            sort_col = col
            sort_desc = False
        refresh_table()

    def load_or_init_state(active_csv: Path) -> SaveState:
        roster_keys_in_order, attendance_confirmed, student_meta = read_canvas_gradebook_csv(active_csv)
        prior = load_save(save_path)

        roster_same = False
        if prior is not None and prior.roster_keys_in_order:
            roster_same = prior.roster_keys_in_order == roster_keys_in_order

        selected_tracked: set[str] | None = None

        if prior is None or not roster_same:
            roster_items = [
                (i + 1, k, student_meta.get(k, {}).get("name", ""), student_meta.get(k, {}).get("student_id", ""))
                for i, k in enumerate(roster_keys_in_order)
            ]
            selected_tracked = gui_select_students(root, roster_items)
            if selected_tracked is None:
                # If this is the first-ever run (no prior), cancel should exit the app gracefully.
                if prior is None:
                    messagebox.showinfo("Cancelled", "No students selected. Exiting.")
                    root.destroy()
                    raise SystemExit(0)
                # If there is prior state, keep existing tracked selection.
                selected_tracked = None
        else:
            selected_tracked = None

        state = reconcile_state_with_csv(
            roster_keys_in_order=roster_keys_in_order,
            attendance_confirmed=attendance_confirmed,
            student_meta=student_meta,
            prior=prior,
            selected_tracked=selected_tracked,
            csv_name=active_csv.name,
        )

        if prior is None or selected_tracked is not None:
            save_state(save_path, state)

        return state

    def gui_select_students(parent: "tk.Tk", roster_items: list[tuple[int, str, str, str]]) -> set[str] | None:
        win = tk.Toplevel(parent)
        win.title("Select tracked students")
        win.geometry("520x640")
        win.transient(parent)
        win.grab_set()

        info = tk.Label(
            win,
            text=(
                "New roster detected. Select the students you want to TRACK.\n"
                "Tip: Ctrl/Shift click to multi-select."
            ),
            justify="left",
            anchor="w",
        )
        info.pack(fill="x", padx=12, pady=(12, 6))

        frame = tk.Frame(win)
        frame.pack(fill="both", expand=True, padx=12, pady=6)

        lb = tk.Listbox(frame, selectmode=tk.EXTENDED)
        vsb = tk.Scrollbar(frame, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=vsb.set)

        # Ensure roster order (CSV order, now sorted alphabetically)
        roster_items_sorted = sorted(roster_items, key=lambda t: t[0])

        keys: list[str] = []
        for idx, key, name, sid in roster_items_sorted:
            keys.append(key)

            sid_clean = (sid or "").strip()
            if sid_clean.lower() == "nan":
                sid_clean = ""

            label = f"{idx:>3}  {name}"
            if sid_clean:
                label = f"{idx:>3}  {name} ({sid_clean})"
            lb.insert(tk.END, label)

        lb.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        btns = tk.Frame(win)
        btns.pack(fill="x", padx=12, pady=12)

        selected: set[str] | None = set()

        def select_all() -> None:
            lb.select_set(0, tk.END)

        def select_none() -> None:
            lb.select_clear(0, tk.END)

        def on_save() -> None:
            nonlocal selected
            idxs = lb.curselection()
            if not idxs:
                messagebox.showwarning("Nothing selected", "Select at least one student (or cancel).")
                return
            selected = {keys[i] for i in idxs}
            win.destroy()

        def on_cancel() -> None:
            nonlocal selected
            selected = None
            win.destroy()

        tk.Button(btns, text="All", command=select_all).pack(side="left")
        tk.Button(btns, text="None", command=select_none).pack(side="left", padx=(8, 0))
        tk.Button(btns, text="Cancel", command=on_cancel).pack(side="right")
        tk.Button(btns, text="Save selection", command=on_save).pack(side="right", padx=(0, 8))

        parent.wait_window(win)
        return selected

    def set_actions_enabled(enabled: bool) -> None:
        for b in action_buttons:
            b.configure(state=(tk.NORMAL if enabled else tk.DISABLED))

    def _student_key(st: StudentState) -> str:
        """Stable key for Treeview row IDs."""
        return st.key

    def _get_student_by_key(key: str) -> StudentState | None:
        """Find a student in state by the stable key used as the Treeview iid."""
        return state.students.get(key)

    def refresh_table() -> None:
        # Preserve selection across refreshes
        prev_sel: str | None = None
        sel = tree.selection()
        if sel:
            iid = sel[0]
            if iid != "divider":
                prev_sel = iid

        for item in tree.get_children():
            tree.delete(item)

        if active_csv is None:
            status_var.set("○ CSV")
            set_actions_enabled(False)
            return

        tracked = _tracked_sorted(state)
        # Sort active and dropped separately for clarity
        active_students = [s for s in tracked if not s.dropped]
        dropped_students = [s for s in tracked if s.dropped]

        key_to_roster_idx: dict[str, int] = {
            key: i + 1 for i, key in enumerate(state.roster_keys_in_order)
        }
        def sort_key(st: StudentState) -> object:
            ridx = int(key_to_roster_idx.get(st.key, 10**9))
            return _coerce_for_sort(sort_col, st, ridx)
        active_students.sort(key=sort_key, reverse=sort_desc)
        dropped_students.sort(key=sort_key, reverse=sort_desc)

        arrow = "▼" if sort_desc else "▲"
        for c in cols:
            label = c
            if c == sort_col:
                label = f"{c} {arrow}"
            tree.heading(c, text=label, command=lambda cc=c: _on_sort(cc))

        for st in active_students:
            tree.insert(
                "",
                tk.END,
                iid=_student_key(st),
                values=(
                    key_to_roster_idx.get(st.key, ""),
                    st.name,
                    st.student_id,
                    st.attendance_confirmed,
                    st.earned,
                    st.used,
                    st.available,
                    (st.notes[:40] + "…") if len(st.notes) > 41 else st.notes,
                ),
            )

        # Divider row for dropped students
        if dropped_students:
            divider_iid = "divider"
            tree.insert(
                "",
                tk.END,
                iid=divider_iid,
                values=("", "──────── Dropped ────────", "", "", "", "", "", ""),
            )

        for st in dropped_students:
            tree.insert(
                "",
                tk.END,
                iid=_student_key(st),
                values=(
                    key_to_roster_idx.get(st.key, ""),
                    st.name,
                    st.student_id,
                    st.attendance_confirmed,
                    st.earned,
                    st.used,
                    st.available,
                    (st.notes[:40] + "…") if len(st.notes) > 41 else st.notes,
                ),
            )

        # Restore selection if possible
        if prev_sel is not None:
            try:
                tree.selection_set(prev_sel)
                tree.focus(prev_sel)
                tree.see(prev_sel)
            except Exception:
                pass

        status_var.set(f"● CSV: {active_csv.name}")
        set_actions_enabled(True)

    def get_selected_student() -> StudentState | None:
        if active_csv is None:
            return None
        sel = tree.selection()
        if not sel:
            return None

        iid = sel[0]
        if iid == "divider":
            return None

        return _get_student_by_key(iid)

    def do_use(n: int = 1) -> None:
        st = get_selected_student()
        if st is None:
            return
        if st.dropped:
            messagebox.showinfo("Dropped", "Student is marked dropped; undrop to track usage.")
            return
        if st.available < n:
            messagebox.showwarning("Not enough", f"Not enough available slip days. Available={st.available}.")
            return
        st.used += n
        refresh_table()

    def do_refund(n: int = 1) -> None:
        st = get_selected_student()
        if st is None:
            return
        st.used = max(0, st.used - n)
        refresh_table()

    def do_toggle_drop() -> None:
        st = get_selected_student()
        if st is None:
            return
        st.dropped = not st.dropped
        if st.dropped:
            # Optional: don’t let dropped students keep “used” slip days lingering
            st.used = min(st.used, st.earned)
        refresh_table()

    def do_edit_note() -> None:
        st = get_selected_student()
        if st is None:
            return
        new_note = simpledialog.askstring("Edit note", f"Notes for {st.name}:", initialvalue=st.notes)
        if new_note is None:
            return
        st.notes = new_note.strip()
        refresh_table()

    def do_save() -> None:
        save_state(save_path, state)
        messagebox.showinfo("Saved", f"Saved to {save_path}")

    def do_change_tracked() -> None:
        nonlocal state
        roster_keys_in_order, attendance_confirmed, student_meta = read_canvas_gradebook_csv(active_csv)
        roster_items = [
            (i + 1, k, student_meta.get(k, {}).get("name", ""), student_meta.get(k, {}).get("student_id", ""))
            for i, k in enumerate(roster_keys_in_order)
        ]
        selected = gui_select_students(root, roster_items)
        if selected is None:
            return
        prior = load_save(save_path)
        state = reconcile_state_with_csv(
            roster_keys_in_order=roster_keys_in_order,
            attendance_confirmed=attendance_confirmed,
            student_meta=student_meta,
            prior=prior,
            selected_tracked=selected,
            csv_name=active_csv.name,
        )
        save_state(save_path, state)
        refresh_table()


    def do_open_csv() -> None:
        nonlocal active_csv, state
        path = filedialog.askopenfilename(
            title="Select Canvas gradebook CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*")],
            initialdir=str(import_dir) if import_dir.exists() else None,
        )
        if not path:
            return
        active_csv = Path(path).expanduser().resolve()
        # Replace the contents of /import with the selected CSV so future runs can auto-detect.
        try:
            import_dir.mkdir(parents=True, exist_ok=True)

            # Remove existing CSVs in /import (treat it as a single-file cache).
            for p in import_dir.glob("*.csv"):
                try:
                    p.unlink()
                except Exception:
                    pass

            dest = import_dir / active_csv.name
            if active_csv.resolve() != dest.resolve():
                shutil.copy2(active_csv, dest)
            active_csv = dest.resolve()
        except Exception:
            # Non-fatal: still proceed using the selected CSV.
            pass
        state = load_or_init_state(active_csv)
        refresh_table()

    # ---- root window ----
    root = tk.Tk()
    # Hide immediately to avoid the "corner flash" before we finish layout/centering.
    root.withdraw()

    root.title("Slip Day Counter")
    root.geometry("1200x560")

    # Center the window on screen (use the requested size, not the transient default size).
    root.update_idletasks()

    width = 1200
    height = 560

    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()

    x = (screen_w // 2) - (width // 2)
    y = (screen_h // 2) - (height // 2) - 80

    root.geometry(f"{width}x{height}+{x}+{y}")

    # Now show it once geometry is correct.
    root.deiconify()
    root.update_idletasks()

    # Bring to front briefly so it doesn't spawn behind VSCode.
    try:
        root.lift()
        root.attributes("-topmost", True)
        root.after(300, lambda: root.attributes("-topmost", False))
    except Exception:
        pass

    # Top area (toolbar OR exit progress) stays above the table.
    top_area = tk.Frame(root)
    top_area.pack(fill="x")

    # top toolbar
    bar = tk.Frame(top_area)
    bar.pack(fill="x", padx=10, pady=10)

    # HStack: (CSV Loaded Status) (Open CSV) (Change Students) ...space... (Save) (Edit note) (Toggle drop) (Refund Slip Day) (Use Slip Day)
    status_var = tk.StringVar(value="○ CSV")
    status = tk.Label(bar, textvariable=status_var, anchor="w")
    status.pack(side="left")

    tk.Button(bar, text="Open CSV", command=do_open_csv).pack(side="left", padx=(10, 0))
    btn_change = tk.Button(bar, text="Change Students", command=do_change_tracked)
    btn_change.pack(side="left", padx=(8, 0))

    btn_save = tk.Button(bar, text="Save", command=do_save)
    btn_use = tk.Button(bar, text="Use Slip Day", command=lambda: do_use(1))
    btn_refund = tk.Button(bar, text="Refund Slip Day", command=lambda: do_refund(1))
    btn_note = tk.Button(bar, text="Edit note", command=do_edit_note)
    btn_drop = tk.Button(bar, text="Toggle drop", command=do_toggle_drop)

    # List of buttons that should be disabled until a CSV is loaded
    action_buttons = [btn_use, btn_refund, btn_drop, btn_note, btn_save, btn_change]

    btn_save.pack(side="right", padx=(0, 8))    
    btn_use.pack(side="right")
    btn_refund.pack(side="right", padx=(0, 8))
    btn_note.pack(side="right", padx=(0, 8))
    btn_drop.pack(side="right", padx=(0, 8))

    # Hold-Esc to save & exit (3s). While holding, replace the toolbar with a large progress bar.
    esc_frame = tk.Frame(top_area)

    # Slightly thicker progress bar (tk/ttk supports thickness via style in many builds)
    style = ttk.Style()
    try:
        # Thicker + green to be more visible (note: some macOS ttk themes ignore colors).
        style.configure(
            "Exit.Horizontal.TProgressbar",
            thickness=28,
            background="#2ecc71",  # green fill
            troughcolor="#2b2b2b",  # dark track
        )
        style.map(
            "Exit.Horizontal.TProgressbar",
            background=[("active", "#2ecc71"), ("!disabled", "#2ecc71")],
        )
    except Exception:
        pass

    esc_label_var = tk.StringVar(value="")
    esc_label = tk.Label(esc_frame, textvariable=esc_label_var, anchor="w", font=("TkDefaultFont", 13, "bold"))
    esc_label.pack(side="top", fill="x", padx=12, pady=(10, 4))

    # Use millisecond scale for smoother determinate updates
    esc_bar = ttk.Progressbar(
        esc_frame,
        orient="horizontal",
        mode="determinate",
        maximum=SAVEESCTIME,
        style="Exit.Horizontal.TProgressbar",
    )
    esc_bar.pack(side="top", fill="x", expand=True, padx=12, pady=(0, 12))

    esc_active: bool = False
    esc_done: bool = False
    esc_start: float = 0.0
    esc_after_id: str | None = None
    esc_is_down: bool = False
    esc_release_after_id: str | None = None

    esc_save_started: bool = False
    esc_save_done: bool = False
    esc_save_error: str = ""
    esc_finish_after_id: str | None = None

    def _esc_start_save() -> None:
        nonlocal esc_save_started, esc_save_done, esc_save_error
        if esc_save_started:
            return
        esc_save_started = True
        esc_save_done = False
        esc_save_error = ""

        # Nothing to save if no CSV/state yet.
        if active_csv is None:
            esc_save_done = True
            return

        def worker() -> None:
            nonlocal esc_save_done, esc_save_error
            try:
                save_state(save_path, state)
            except Exception as e:
                esc_save_error = str(e)
            finally:
                esc_save_done = True

        t = threading.Thread(target=worker, daemon=True)
        t.start()

    def _esc_hide() -> None:
        nonlocal esc_active, esc_after_id, esc_release_after_id, esc_finish_after_id
        esc_active = False
        if esc_after_id is not None:
            try:
                root.after_cancel(esc_after_id)
            except Exception:
                pass
            esc_after_id = None
        if esc_release_after_id is not None:
            try:
                root.after_cancel(esc_release_after_id)
            except Exception:
                pass
            esc_release_after_id = None
        if esc_finish_after_id is not None:
            try:
                root.after_cancel(esc_finish_after_id)
            except Exception:
                pass
            esc_finish_after_id = None
        esc_label_var.set("")
        esc_bar["value"] = 0
        esc_frame.pack_forget()
        # Restore toolbar
        bar.pack(fill="x", padx=10, pady=10)

    def _esc_finish() -> None:
        """Exit only after save is complete (safe exit)."""
        nonlocal esc_done, esc_finish_after_id
        if esc_done:
            return

        # If save is still running, wait with the bar full.
        if esc_save_started and not esc_save_done:
            esc_label_var.set("Finishing save...")
            esc_bar["value"] = SAVEESCTIME

            def _poll() -> None:
                nonlocal esc_finish_after_id
                esc_finish_after_id = None
                if esc_done:
                    return
                if esc_save_done:
                    _esc_finish()
                    return
                esc_finish_after_id = root.after(50, _poll)

            if esc_finish_after_id is None:
                esc_finish_after_id = root.after(50, _poll)
            return

        # Save is done (or there was nothing to save): exit now.
        esc_done = True
        try:
            root.destroy()
        finally:
            raise SystemExit(0)

    def _esc_tick() -> None:
        nonlocal esc_after_id
        if not esc_active:
            return
        elapsed_ms = int((time.monotonic() - esc_start) * 1000)
        if elapsed_ms < 0:
            elapsed_ms = 0
        # Optionally, show error in saving
        if esc_save_error and "(save error" not in esc_label_var.get().lower():
            esc_label_var.set(f"Saving and exiting... (save error: {esc_save_error})")
        if elapsed_ms >= SAVEESCTIME:
            esc_bar["value"] = SAVEESCTIME
            _esc_finish()
            return
        esc_bar["value"] = elapsed_ms
        # ~60fps updates for smoother animation
        esc_after_id = root.after(1, _esc_tick)

    def _esc_start_hold() -> None:
        nonlocal esc_active, esc_start, esc_done
        if esc_active or esc_done or esc_frame.winfo_ismapped():
            return
        esc_active = True
        esc_start = time.monotonic()
        esc_label_var.set("Saving and exiting...")
        _esc_start_save()
        # Replace toolbar with exit progress UI
        bar.pack_forget()
        esc_frame.pack(fill="x", padx=10, pady=10)
        esc_bar["value"] = 0
        _esc_tick()

    def _esc_maybe_cancel() -> None:
        nonlocal esc_release_after_id
        esc_release_after_id = None
        if esc_done:
            return
        # Only cancel if the key is still up.
        if not esc_is_down:
            _esc_hide()

    # table
    cols = ("#", "Name", "Student ID", "Attended", "Earned Slips", "Used Slips", "Available Slips", "Notes")
    tree = ttk.Treeview(root, columns=cols, show="headings", height=18)

    for c in cols:
        tree.heading(c, text=c, command=lambda cc=c: _on_sort(cc))

    tree.column("#", width=20, anchor="e")
    tree.column("Name", width=200, anchor="w")
    tree.column("Student ID", width=70, anchor="w")
    tree.column("Attended", width=80, anchor="e")
    tree.column("Earned Slips", width=95, anchor="e")
    tree.column("Used Slips", width=85, anchor="e")
    tree.column("Available Slips", width=100, anchor="e")
    tree.column("Notes", width=320, anchor="w")

    yscroll = ttk.Scrollbar(root, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=yscroll.set)

    tree.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=(0, 10))
    yscroll.pack(side="right", fill="y", padx=(0, 10), pady=(0, 10))

    # Enter edits notes; double-click does nothing.
    def on_enter_key(_evt: Any) -> None:
        st = get_selected_student()
        if st is not None:
            do_edit_note()
    tree.bind("<Return>", on_enter_key)
    tree.bind("<KP_Enter>", on_enter_key)

    # Up/Down navigation with wrap-around (skip divider row).
    def _selectable_iids() -> list[str]:
        return [iid for iid in tree.get_children("") if iid != "divider"]

    def _set_selection(iid: str) -> None:
        try:
            tree.selection_set(iid)
            tree.focus(iid)
            tree.see(iid)
        except Exception:
            pass

    def on_up_key(_evt: Any) -> str:
        iids = _selectable_iids()
        if not iids:
            return "break"
        sel = tree.selection()
        if not sel or sel[0] == "divider":
            _set_selection(iids[-1])
            return "break"
        cur = sel[0]
        try:
            idx = iids.index(cur)
        except ValueError:
            _set_selection(iids[-1])
            return "break"
        next_iid = iids[idx - 1] if idx > 0 else iids[-1]
        _set_selection(next_iid)
        return "break"

    def on_down_key(_evt: Any) -> str:
        iids = _selectable_iids()
        if not iids:
            return "break"
        sel = tree.selection()
        if not sel or sel[0] == "divider":
            _set_selection(iids[0])
            return "break"
        cur = sel[0]
        try:
            idx = iids.index(cur)
        except ValueError:
            _set_selection(iids[0])
            return "break"
        next_iid = iids[idx + 1] if idx < (len(iids) - 1) else iids[0]
        _set_selection(next_iid)
        return "break"

    tree.bind("<Up>", on_up_key)
    tree.bind("<Down>", on_down_key)

    def on_delete_key(_evt: Any) -> None:
        st = get_selected_student()
        if st is not None:
            do_toggle_drop()

    tree.bind("<Delete>", on_delete_key)
    tree.bind("<BackSpace>", on_delete_key)

    def on_left_bracket(_evt: Any) -> None:
        st = get_selected_student()
        if st is not None:
            do_use(1)

    def on_right_bracket(_evt: Any) -> None:
        st = get_selected_student()
        if st is not None:
            do_refund(1)

    tree.bind("[", on_left_bracket)
    tree.bind("]", on_right_bracket)

    def on_esc_press(_evt: Any) -> None:
        nonlocal esc_is_down, esc_release_after_id
        esc_is_down = True
        # If a release-cancel was scheduled (auto-repeat), cancel it.
        if esc_release_after_id is not None:
            try:
                root.after_cancel(esc_release_after_id)
            except Exception:
                pass
            esc_release_after_id = None
        _esc_start_hold()

    def on_esc_release(_evt: Any) -> None:
        nonlocal esc_is_down, esc_release_after_id
        esc_is_down = False
        if esc_done:
            return
        # Debounce: some systems emit release/press pairs during auto-repeat.
        if esc_release_after_id is None:
            esc_release_after_id = root.after(140, _esc_maybe_cancel)

    root.bind_all("<KeyPress-Escape>", on_esc_press)
    root.bind_all("<KeyRelease-Escape>", on_esc_release)

    def on_select(_evt: Any) -> None:
        sel = tree.selection()
        if sel and sel[0] == "divider":
            tree.selection_remove("divider")

    tree.bind("<<TreeviewSelect>>", on_select)

    # initial load
    active_csv: Path | None = csv_path
    state: SaveState
    if active_csv is not None and active_csv.exists():
        state = load_or_init_state(active_csv)
    else:
        active_csv = None
        state = SaveState()

    refresh_table()

    root.mainloop()


# -----------------------------
# Entrypoint
# -----------------------------

def main() -> None:
    script_path = Path(__file__).resolve()
    default_save = get_default_save_path(script_path)

    parser = argparse.ArgumentParser(description="Count and track Canvas slip days (attendance-based).")
    parser.add_argument("--csv", type=str, default="", help="Path to a Canvas gradebook export CSV.")
    parser.add_argument("--import-dir", type=str, default="import", help="Directory containing CSV imports.")
    parser.add_argument("--save", type=str, default=str(default_save), help="Path to save.txt (JSON).")
    parser.add_argument("--ui", action="store_true", help="Launch the GUI.")
    args = parser.parse_args()

    save_path = Path(args.save).expanduser().resolve()

    # Resolve import dir relative to the script folder by default (so it lives in /SlipDayTool/import).
    import_dir_raw = Path(args.import_dir).expanduser()
    if import_dir_raw.is_absolute():
        import_dir = import_dir_raw
    else:
        import_dir = script_path.parent / import_dir_raw
    import_dir = import_dir.resolve()

    if args.csv:
        csv_path = Path(args.csv).expanduser().resolve()
    else:
        csv_path = find_latest_csv(import_dir)

    # If a CSV was provided explicitly, replace the contents of /import with it so future runs can auto-detect.
    if csv_path is not None and csv_path.exists() and args.csv:
        try:
            import_dir.mkdir(parents=True, exist_ok=True)

            # Remove existing CSVs in /import (treat it as a single-file cache).
            for p in import_dir.glob("*.csv"):
                try:
                    p.unlink()
                except Exception:
                    pass

            dest = import_dir / csv_path.name
            if csv_path.resolve() != dest.resolve():
                shutil.copy2(csv_path, dest)
            csv_path = dest.resolve()
        except Exception:
            # Non-fatal: importing should still work even if copy fails
            pass

    # Always launch the GUI (legacy behavior); --ui is accepted for compatibility
    gui_main(
        csv_path=(csv_path if (csv_path is not None and csv_path.exists()) else None),
        save_path=save_path,
        import_dir=import_dir,
    )

if __name__ == "__main__":
    main()