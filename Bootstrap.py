
from __future__ import annotations

import argparse
import os
import shlex
import subprocess
import sys
import tkinter as tk
from pathlib import Path


def _center(win: tk.Tk, width: int, height: int) -> None:
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = max(0, int((sw - width) / 2))
    y = max(0, int((sh - height) / 2))
    win.geometry(f"{width}x{height}+{x}+{y}")


def _launch_countslips(script: Path, import_dir: Path, extra_args: list[str]) -> None:
    """Launch CountSlips in a detached process using the current interpreter (venv)."""
    py = Path(sys.executable)
    cmd = [str(py), str(script), "--import-dir", str(import_dir), "--ui", *extra_args]
    env = dict(os.environ)
    # Prevent any in-file delegation logic from spawning make again.
    env["SLIPDAYCOUNTER_SKIP_MAKE"] = "1"

    subprocess.Popen(
        cmd,
        cwd=str(script.parent),
        env=env,
        start_new_session=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Slip Day Counter Bootstrap")

    # Prefer env vars (set by make) but allow flags too.
    parser.add_argument("--script", type=str, default=os.environ.get("SCRIPT", "CountSlips.py"))
    parser.add_argument("--import-dir", type=str, default=os.environ.get("IMPORT_DIR", "import"))
    parser.add_argument("--args", type=str, default=os.environ.get("ARGS", ""))

    args = parser.parse_args()

    root_dir = Path(__file__).resolve().parent
    script_path = (root_dir / args.script).resolve()
    import_dir = Path(args.import_dir).expanduser().resolve()

    try:
        extra_args = shlex.split(args.args) if str(args.args).strip() else []
    except Exception:
        extra_args = []

    # --- Splash window ---
    win = tk.Tk()
    win.title("SlipDayCounter")
    win.resizable(False, False)
    _center(win, 640, 220)

    # Make sure it shows up (VSCode can spawn behind).
    try:
        win.lift()
        win.attributes("-topmost", True)
        win.after(300, lambda: win.attributes("-topmost", False))
    except Exception:
        pass

    frame = tk.Frame(win, padx=18, pady=18)
    frame.pack(fill="both", expand=True)

    title = tk.Label(frame, text="Slip Day Counter", font=("Helvetica", 18, "bold"))
    title.pack(anchor="w")

    subtitle = tk.Label(
        frame,
        text="Launching…",
        font=("Helvetica", 12),
    )
    subtitle.pack(anchor="w", pady=(6, 14))

    bar = tk.Canvas(frame, height=16, highlightthickness=0)
    bar.pack(fill="x")

    # Simple smooth-ish fill animation for ~700ms.
    total_ms = 700
    step_ms = 20
    steps = max(1, total_ms // step_ms)
    state = {"i": 0}

    def tick() -> None:
        w = max(1, bar.winfo_width())
        h = 16
        i = state["i"]
        frac = min(1.0, i / steps)
        bar.delete("all")
        # background
        bar.create_rectangle(0, 0, w, h, outline="", fill="#2b2b2b")
        # fill
        bar.create_rectangle(0, 0, int(w * frac), h, outline="", fill="#2ecc71")

        if i >= steps:
            # Launch app at the start of the completed state.
            try:
                _launch_countslips(script_path, import_dir, extra_args)
            finally:
                win.after(150, win.destroy)
            return

        state["i"] = i + 1
        win.after(step_ms, tick)

    # Start animation ASAP.
    win.after(30, tick)
    win.mainloop()


if __name__ == "__main__":
    main()
