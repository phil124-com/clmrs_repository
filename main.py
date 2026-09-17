"""
CLMRS Data Analyzer - entry point.
Double-click target once packaged; launches the desktop GUI.
"""
import sys
import os
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    try:
        from gui.app import main as gui_main
        gui_main()
    except Exception:
        # Never let the app vanish silently on a double-click launch —
        # write a crash log next to the exe and show a message box.
        import tkinter as tk
        from tkinter import messagebox
        tb = traceback.format_exc()
        try:
            log_dir = os.path.join(os.environ.get("APPDATA", os.getcwd()), "CLMRS_Analyzer")
            os.makedirs(log_dir, exist_ok=True)
            with open(os.path.join(log_dir, "crash.log"), "w", encoding="utf-8") as f:
                f.write(tb)
        except Exception:
            pass
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("CLMRS Data Analyzer",
                                  f"The application failed to start:\n\n{tb[-800:]}")
        except Exception:
            print(tb, file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
