"""Freight tab / GL coding mini mockup (Merz POC - Process New Non-PO Invoice Request).

Super tiny stand-in for a narrow-window Excel view showing the "Freight" sheet
tab: a GL coding list (GL / Facilities / Marketing accounts) next to an
"Approval Email To" field. Deliberately small and mostly decorative - no
cell editing here, just a faithful tiny snapshot of this one section, using
the narrow/tablet-style ribbon (single-icon compact groups, truncated tab
labels) rather than the full desktop ribbon used in the vendor-list mockup.

The GL/Facilities/Marketing numbers and the approval email are made-up demo
values - the reference screenshot had them redacted.
"""
import json
import os
import tkinter as tk
from pathlib import Path

FONT_FAMILY = "Calibri"
GRIDLINE = "#D9D9D9"
HEADER_BG = "#F3F2F1"
RIBBON_GREEN = "#217346"
BLUE_FILL = "#DCE6F1"
TOPBAR_BG = "#F3F2F1"

ROW_H = 24
COL_W = 130


class FreightGLApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Vendor Invoice Tracker.xlsx - Excel")
        self.root.geometry("460x340")
        self.root.configure(bg="white")

        self._build_top_bar()
        self._build_ribbon()
        self._build_grid()
        self._build_bottom_bar()

        if AUTOMATION_ENABLED:
            # Opt-in automation bridge (see AUTOMATION BRIDGE section below) --
            # same pattern as the SAP/Excel mirrors: polled from inside this
            # same Tk event loop via self.root.after(), no threads.
            self.root.after(BRIDGE_POLL_MS, lambda: _bridge_poll(self))

    # -- top strip: autosave / filename / icons ------------------------------

    def _build_top_bar(self):
        bar = tk.Frame(self.root, bg=TOPBAR_BG, height=26)
        bar.pack(side="top", fill="x")
        tk.Label(bar, text="X", bg=RIBBON_GREEN, fg="white", font=(FONT_FAMILY, 10, "bold"),
                 padx=6, pady=2).pack(side="left", padx=(4, 6), pady=3)
        pill = tk.Label(bar, text="●  On", bg="#E1E1E1", fg="#333333", font=(FONT_FAMILY, 8),
                         padx=6, pady=1)
        pill.pack(side="left", pady=5)
        tk.Label(bar, text="Vendor Invoice Tracker.xlsx", bg=TOPBAR_BG, fg="#333333",
                 font=(FONT_FAMILY, 9), padx=8).pack(side="left")
        right = tk.Frame(bar, bg=TOPBAR_BG)
        right.pack(side="right", padx=6)
        for sym in ["\U0001F4C4", "\U0001F50D"]:
            tk.Label(right, text=sym, bg=TOPBAR_BG, fg="#555555", font=(FONT_FAMILY, 10)).pack(side="left", padx=3)

    # -- narrow/tablet-style ribbon -------------------------------------------

    def _build_ribbon(self):
        ribbon = tk.Frame(self.root, bg="white")
        ribbon.pack(side="top", fill="x")

        tabs_bar = tk.Frame(ribbon, bg="white", highlightthickness=1, highlightbackground=GRIDLINE)
        tabs_bar.pack(side="top", fill="x")
        for tab in ["File", "Home", "Draw", "Page I", "Form", "Data", "Revie", "View"]:
            active = tab == "Home"
            lbl = tk.Label(tabs_bar, text=tab, bg="white", fg=RIBBON_GREEN if active else "#333333",
                            font=(FONT_FAMILY, 9, "bold" if active else "normal"), padx=6, pady=4)
            lbl.pack(side="left")
            lbl.bind("<Button-1>", lambda e, t=tab: self._flash(f"'{t}' is not available in this demo."))
        underline = tk.Frame(ribbon, bg=RIBBON_GREEN, height=2)
        underline.pack(side="top", fill="x")

        groups_bar = tk.Frame(ribbon, bg=HEADER_BG, height=64, highlightthickness=1,
                               highlightbackground=GRIDLINE)
        groups_bar.pack(side="top", fill="x")

        def compact_group(name, icon):
            g = tk.Frame(groups_bar, bg=HEADER_BG, padx=6)
            g.pack(side="left", fill="y")
            b = tk.Label(g, text=f"{icon}\n{name} ⌄", bg=HEADER_BG, fg="#333333",
                         font=(FONT_FAMILY, 8), justify="center", pady=4)
            b.pack(side="top")
            b.bind("<Button-1>", lambda e, n=name: self._flash(f"'{n}' is not available in this demo."))
            sep = tk.Frame(groups_bar, bg=GRIDLINE, width=1)
            sep.pack(side="left", fill="y", pady=6)

        compact_group("Clipboard", "\U0001F4CB")
        compact_group("Font", "A")
        compact_group("Alignment", "≡")
        compact_group("Number", "%")

        styles = tk.Frame(groups_bar, bg=HEADER_BG, padx=6)
        styles.pack(side="left", fill="y")
        for item in ["Conditional Formatting", "Format as Table ⌄", "Cell Styles ⌄"]:
            b = tk.Label(styles, text=item, bg=HEADER_BG, fg="#333333", font=(FONT_FAMILY, 8),
                         anchor="w", justify="left")
            b.pack(side="top", anchor="w", pady=1)
            b.bind("<Button-1>", lambda e, i=item: self._flash(f"'{i}' is not available in this demo."))
        tk.Label(styles, text="Styles", bg=HEADER_BG, fg="#555555", font=(FONT_FAMILY, 7)).pack(side="bottom")

    # -- grid: columns E, F, G -------------------------------------------------

    def _build_grid(self):
        grid = tk.Frame(self.root, bg="white")
        grid.pack(side="top", fill="both", expand=True)

        header_widths = {"E": 18, "F": 16, "G": 10}
        tk.Label(grid, text="", bg=HEADER_BG, width=3, highlightthickness=1,
                 highlightbackground=GRIDLINE).grid(row=0, column=0)
        for i, letter in enumerate(["E", "F", "G"]):
            tk.Label(grid, text=letter, bg=HEADER_BG, fg="#444444", font=(FONT_FAMILY, 9),
                     width=header_widths[letter], highlightthickness=1,
                     highlightbackground=GRIDLINE).grid(row=0, column=i + 1)

        # Label and number share one cell, like a real spreadsheet value - these are real
        # Entry widgets (not Labels), so the text is genuinely selectable: double-clicking
        # the digits selects just "4200" (the "-"/space around it breaks the word
        # boundary), without needing a separate cell for the number.
        rows = [
            {"row": "1", "E": ("Approver Approval Email To", "white", "black", True, False),
             "F": ("GL-4200", "white", "black", True, False)},
            {"row": "2", "E": ("", BLUE_FILL, "black", False, True),
             "F": ("Facilities -1100", "white", "black", False, False)},
            {"row": "3", "F": ("Marketing -2200", "white", "black", False, False)},
        ]
        col_widths = {"E": 18, "F": 16, "G": 10}
        # Opt-in automation bridge target cells -- see AUTOMATION BRIDGE
        # section below: (row, col) -> Tk widget name, so Locate Widget can
        # find these two specific cells directly (no search-by-text needed,
        # this mockup always shows the same fixed reference values).
        BRIDGE_CELL_NAMES = {(1, "F"): "freight_gl_code_cell", (2, "F"): "freight_facility_cell"}
        self.cell_entries = {}
        for i, letter in enumerate(["E", "F", "G"]):
            grid.grid_columnconfigure(i + 1, minsize=0)
        for r, row_data in enumerate(rows, start=1):
            tk.Label(grid, text=row_data["row"], bg=HEADER_BG, fg="#444444", font=(FONT_FAMILY, 9),
                     width=3, highlightthickness=1, highlightbackground=GRIDLINE).grid(row=r, column=0, sticky="nsew")
            for i, col in enumerate(["E", "F", "G"]):
                text, bg, fg, bold, editable = row_data.get(col, ("", "white", "black", False, False))
                var = tk.StringVar(value=text)
                entry_kwargs = {}
                name = BRIDGE_CELL_NAMES.get((r, col))
                if name:
                    entry_kwargs["name"] = name
                entry = tk.Entry(grid, textvariable=var, bg=bg, fg=fg, disabledforeground=fg,
                                  readonlybackground=bg, relief="flat", bd=0, insertbackground=fg,
                                  font=(FONT_FAMILY, 9, "bold" if bold else "normal"),
                                  highlightthickness=1, highlightbackground=GRIDLINE,
                                  width=col_widths[col], **entry_kwargs)
                if not editable:
                    entry.configure(state="readonly")
                entry.grid(row=r, column=i + 1, sticky="nsew", ipady=3)
                self.cell_entries[(r, col)] = entry

    # -- bottom bar: sheet tabs + status --------------------------------------

    def _build_bottom_bar(self):
        tabs_bar = tk.Frame(self.root, bg=HEADER_BG, height=24)
        tabs_bar.pack(side="bottom", fill="x")
        tk.Label(tabs_bar, text="‹  ›", bg=HEADER_BG, fg="#555555",
                 font=(FONT_FAMILY, 9), padx=4).pack(side="left")
        for name, active in [("PO #", False), ("Freight", True)]:
            lbl = tk.Label(tabs_bar, text=name, bg=HEADER_BG,
                            fg=RIBBON_GREEN if active else "#555555",
                            font=(FONT_FAMILY, 9, "bold" if active else "normal"), padx=8, pady=3)
            lbl.pack(side="left")
            if active:
                tk.Frame(tabs_bar, bg=RIBBON_GREEN, height=2, width=lbl.winfo_reqwidth()).place(
                    in_=lbl, relx=0, rely=1.0, anchor="nw")
            lbl.bind("<Button-1>", lambda e, n=name: self._flash(f"'{n}' tab is not available in this demo."))
        tk.Label(tabs_bar, text="+", bg=HEADER_BG, fg="#555555", font=(FONT_FAMILY, 10),
                 padx=8).pack(side="left")

        status_bar = tk.Frame(self.root, bg=HEADER_BG, height=20)
        status_bar.pack(side="bottom", fill="x")
        self.status_var = tk.StringVar(value="Edit")
        tk.Label(status_bar, textvariable=self.status_var, bg=HEADER_BG, fg="#333333",
                 font=(FONT_FAMILY, 8), padx=8).pack(side="left")
        tk.Label(status_bar, text="♿ Accessibility: Investigate", bg=HEADER_BG, fg="#555555",
                 font=(FONT_FAMILY, 8)).pack(side="left", padx=10)

    def _flash(self, message):
        self.status_var.set(message)


# =========================================================================
# AUTOMATION BRIDGE -- strictly opt-in, additive, and read-only with respect
# to app behavior. Gated behind FREIGHT_GL_AUTOMATION=1; when that env var is
# unset (normal `python3 freight_gl_mirror.py`), none of this code runs.
# Same design as the SAP/Excel mirrors' own bridges (see main.py's
# "AUTOMATION BRIDGE" section for the full rationale) -- this module only
# *reports* geometry/state; it never triggers actions.
#
# Only two locator kinds needed here: this mockup always shows the same
# fixed reference values (no search-by-text like the vendor list), so a
# plain name-based widget lookup (for clicking) + a direct text read (for
# the resulting value) is enough -- no formula bar to click-then-read, these
# ARE the display widgets themselves.
# =========================================================================

AUTOMATION_ENABLED = os.environ.get("FREIGHT_GL_AUTOMATION") == "1"
BRIDGE_REQUEST_FILE = Path(os.environ.get("FREIGHT_GL_REQUEST_FILE", "/tmp/freight_gl_bridge_request.json"))
BRIDGE_RESPONSE_FILE = Path(os.environ.get("FREIGHT_GL_RESPONSE_FILE", "/tmp/freight_gl_bridge_response.json"))
BRIDGE_POLL_MS = 130


def _bridge_find_widget(root, name):
    stack = list(root.winfo_children())
    while stack:
        w = stack.pop()
        try:
            if w.winfo_name() == name:
                return w
        except tk.TclError:
            continue
        try:
            stack.extend(w.winfo_children())
        except tk.TclError:
            continue
    return None


def _bridge_bbox(w):
    w.update_idletasks()
    x, y = w.winfo_rootx(), w.winfo_rooty()
    width, height = w.winfo_width(), w.winfo_height()
    return {
        "ok": True, "x": x, "y": y, "width": width, "height": height,
        "center_x": x + width // 2, "center_y": y + height // 2,
    }


def _bridge_resolve_widget(app, req):
    name = req.get("name")
    w = _bridge_find_widget(app.root, name)
    if w is None:
        return {"ok": False, "error": f"no widget named {name!r} found"}
    return _bridge_bbox(w)


def _bridge_resolve_widget_text(app, req):
    """Read a named Entry cell's current text directly -- this mockup's
    cells have no separate formula bar to click-then-read, the Entry IS the
    display, so this is the equivalent of Excel mirror's Get Excel State."""
    name = req.get("name")
    w = _bridge_find_widget(app.root, name)
    if w is None:
        return {"ok": False, "error": f"no widget named {name!r} found"}
    try:
        return {"ok": True, "text": w.get()}
    except tk.TclError as exc:
        return {"ok": False, "error": str(exc)}


_BRIDGE_RESOLVERS = {
    "widget": _bridge_resolve_widget,
    "widget_text": _bridge_resolve_widget_text,
}


def _bridge_poll(app):
    try:
        if BRIDGE_REQUEST_FILE.exists():
            try:
                req = json.loads(BRIDGE_REQUEST_FILE.read_text())
            except (ValueError, OSError) as exc:
                req = None
                result = {"ok": False, "error": f"could not read/parse request file: {exc}"}
            if req is not None:
                kind = req.get("kind")
                resolver = _BRIDGE_RESOLVERS.get(kind)
                if resolver is None:
                    result = {"ok": False, "error": f"unknown locator kind {kind!r}"}
                else:
                    try:
                        result = resolver(app, req)
                    except Exception as exc:  # never let a bad/mistimed request crash the app
                        result = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
            try:
                BRIDGE_RESPONSE_FILE.write_text(json.dumps(result))
            finally:
                try:
                    BRIDGE_REQUEST_FILE.unlink()
                except OSError:
                    pass
    finally:
        app.root.after(BRIDGE_POLL_MS, lambda: _bridge_poll(app))


def main():
    root = tk.Tk()
    FreightGLApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
