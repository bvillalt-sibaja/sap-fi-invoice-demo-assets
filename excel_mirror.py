"""Vendor list Excel mirror demo (Merz POC - Process New Non-PO Invoice Request).

A desktop stand-in for Excel showing the vendor master list a user would check
to find a vendor's SAP vendor code before processing a non-PO invoice. Same
look as real Excel (native app menu, ribbon with tabs and grouped controls,
name box + formula bar, column/row headers, AutoFilter arrows, colored
"owner" legend cells, sheet tabs, zoom/status bar). Real behavior: click a
cell to select it (name box + formula bar update), edit via the formula bar.
Every ribbon/menu control is decorative and safely no-ops via the status bar
message, matching this project's other mirror-demo apps.

Single self-contained file, no data file / companion widget modules - this
replaces an earlier version of this file that imported grid_widget.py and
ribbon.py and read data/invoices.json, none of which existed anywhere in this
project (that version described a completely different mockup - a prepaid
amortization schedule with typing animations and an external control-file
API - and would not run). Rebuilt from scratch for the vendor-list screen
actually needed here.

The exact vendor names/owner names in the reference screenshot were redacted
by the user, so the filler rows below are generic placeholders, not a
reconstruction of the real list - only the "Meridian Freight Lines" / MFL250
row is real, matching the vendor on the invoice PDF used elsewhere in this
POC (Invoice_MFL119284.pdf).
"""
import json
import os
import tkinter as tk
from pathlib import Path
from tkinter import ttk

FONT_FAMILY = "Calibri"
GRIDLINE = "#D9D9D9"
HEADER_BG = "#E9E9E9"
HEADER_BORDER = "#BFBFBF"
YELLOW = "#FFFF00"
PINK_BG = "#FFC7CE"
PINK_FG = "#9C0006"
STATUS_BG = "#E7E6E5"
RIBBON_GREEN = "#217346"
SELECT_BORDER = "#1A7A3C"
ROW_H = 20
NAME_COL_W = 40

# (header label, width, has_autofilter_arrow)
COLUMNS = [
    ("NAME", 210, True),
    ("SUPPLIER/VI", 110, True),
    ("SYSTEM", 100, True),
    ("", 36, False),
    ("CURRENT PO", 110, True),
    ("SAP VENDOR CODE", 110, True),
    ("COST", 85, True),
    ("COST", 85, True),
    ("NOTES", 210, False),
    ("", 50, False),
    ("", 50, False),
    ("", 150, False),
    ("", 160, False),
    ("", 50, False),
    ("A-K", 90, False),
    ("L-R", 90, False),
    ("S-Z", 90, False),
]
N_COLS = len(COLUMNS)

# Legend/"owner" cells that live in the header row, columns O/P/Q (indices 14-16).
LEGEND_CELLS = {
    14: ("A-K", "#1F4E79", "white"),
    15: ("L-R", "#C00CB6", "white"),
    16: ("S-Z", "#FFFF00", "#333333"),
}

# Each row: {col_index: value}. Column 0 (NAME) is always yellow-filled, matching
# the reference screenshot. A couple of rows carry a pink "over budget" style
# highlight on a COST cell and a note, matching the screenshot's conditional
# formatting look.
ROWS = [
    {0: "Meridian Freight Lines", 1: "MFL250", 2: "SAP", 4: "PO# ON FILE", 5: "SAP-10234",
     6: "$50,000.00", 7: "$2,829.84", 8: "Freight carrier - non-PO invoice"},
    {0: "Waste Mgmt", 1: "WMG-102", 2: "SAP", 4: "PO# ON FILE", 5: "SAP-10871",
     6: "$18,000.00", 7: "$12,430.55", 8: "Recurring monthly service"},
    {0: "We Energies", 1: "WEN-330", 2: "SAP", 3: "UV", 5: "SAP-11290",
     6: "$9,500.00", 7: ("$11,820.10", PINK_BG, PINK_FG), 8: "Over budget - utilities",
     12: ("Over $2,500 = PO#", YELLOW, "black")},
    {0: "Wheels Inc", 1: "WHL-410", 2: "SAP non PO-Upload", 5: "SAP-12045",
     6: "$6,200.00", 7: "$3,140.00", 8: "Fleet maintenance"},
    {0: "World Courier, Inc", 1: "WCI-558", 2: "SAP", 4: "PO# ON FILE", 5: "SAP-13387",
     6: "$27,750.00", 7: "$19,875.40", 8: "RON to confirm invoice #"},
    {0: "Sewer Utility District #1", 1: "SWD-771", 2: "SAP", 5: "SAP-14002",
     6: ("$41,900.00", PINK_BG, PINK_FG), 7: "$38,220.00", 8: "Over budget - review coding",
     12: ("Over $2,500 = PO#", YELLOW, "black")},
    {0: "Water Utility District #1", 1: "WTD-772", 2: "SAP", 5: "SAP-14003",
     6: ("$39,600.00", PINK_BG, PINK_FG), 7: "$36,410.00", 8: "Over budget - review coding",
     12: ("Over $2,500 = PO#", YELLOW, "black")},
    {0: "Packaging Solutions Inc", 1: "PKS-905", 2: "SAP", 4: "PO# ON FILE", 5: "SAP-15230",
     6: "$14,200.00", 7: "$8,960.00", 8: "Standard packaging supplier"},
    {0: "Environmental Services Corp", 1: "ENV-140", 2: "SAP & EDI", 3: "EDI", 5: "SAP-16044",
     6: "$22,000.00", 7: "$17,300.00", 8: "Tariff-related charges (?)"},
    {0: "International Freight Corp", 1: "IFC-260", 2: "SAP & EDI", 3: "EDI", 4: "PO# ON FILE",
     5: "SAP-17182", 6: "$31,500.00", 7: "$28,940.00",
     11: ("Coding for Credit Memo - (Process using EDI)", None, "#444444")},
    {0: "Regional Logistics Group", 1: "RLG-334", 2: "SAP", 5: "SAP-18099",
     6: "$12,750.00", 7: "$9,410.00", 8: "Standard freight vendor"},
    {0: "Southern Distribution Co", 1: "SDC-489", 2: "SAP", 4: "PO# ON FILE", 5: "SAP-19255",
     6: "$16,300.00", 7: "$13,120.00", 8: "Standard freight vendor"},
]
DATA_ROW_COUNT = len(ROWS)
BLANK_ROWS_BELOW = 18


def col_letter(idx):
    letters = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


class VendorListApp:
    def __init__(self, root):
        self.root = root
        self.root.title("VENDOR LIST.xlsx - Excel")
        self.root.geometry("1500x820")
        self.root.configure(bg="white")

        self.cell_labels = {}   # (col, row) -> tk.Label
        self.cell_values = {}   # (col, row) -> str
        self.selected = None    # (col, row)

        self._build_native_menu()
        self._build_ribbon()
        self._build_formula_bar()
        self._build_grid()
        self._build_bottom_bar()

        if AUTOMATION_ENABLED:
            # Opt-in automation bridge (see AUTOMATION BRIDGE section below) --
            # same pattern as the SAP mirror's main.py: polled from inside
            # this same Tk event loop via self.root.after(), no threads.
            self.root.after(BRIDGE_POLL_MS, lambda: _bridge_poll(self))

    # -- native app menu ----------------------------------------------------

    def _build_native_menu(self):
        menubar = tk.Menu(self.root)
        for tab in ["File", "Edit", "View", "Insert", "Format", "Data", "Help"]:
            menu = tk.Menu(menubar, tearoff=0)
            if tab == "File":
                menu.add_command(label="Save", command=lambda: self._flash("Save is not available in this demo."))
                menu.add_command(label="Close", command=self.root.quit)
            else:
                menu.add_command(label=f"{tab} options",
                                  command=lambda t=tab: self._flash(f"{t} is not available in this demo."))
            menubar.add_cascade(label=tab, menu=menu)
        self.root.config(menu=menubar)

    # -- ribbon ---------------------------------------------------------------

    def _build_ribbon(self):
        ribbon = tk.Frame(self.root, bg="white")
        ribbon.pack(side="top", fill="x")

        tabs_bar = tk.Frame(ribbon, bg=RIBBON_GREEN, height=28)
        tabs_bar.pack(side="top", fill="x")
        tk.Label(tabs_bar, text="X", bg=RIBBON_GREEN, fg="white", font=(FONT_FAMILY, 11, "bold"),
                 padx=10).pack(side="left")
        for tab in ["File", "Home", "Insert", "Page Layout", "Formulas", "Data",
                    "Review", "View", "Automate", "Help", "Acrobat"]:
            active = tab == "Home"
            lbl = tk.Label(tabs_bar, text=tab, bg="white" if active else RIBBON_GREEN,
                            fg="#1a1a1a" if active else "white", font=(FONT_FAMILY, 10),
                            padx=8, pady=4)
            lbl.pack(side="left")
            lbl.bind("<Button-1>", lambda e, t=tab: self._flash(f"'{t}' tab is not available in this demo."))
        right = tk.Frame(tabs_bar, bg=RIBBON_GREEN)
        right.pack(side="right", padx=8)
        for label in ["Comments", "Share"]:
            b = tk.Label(right, text=label, bg="#3b8a5e", fg="white", font=(FONT_FAMILY, 9), padx=8, pady=2)
            b.pack(side="left", padx=3)
            b.bind("<Button-1>", lambda e, l=label: self._flash(f"'{l}' is not available in this demo."))

        groups_bar = tk.Frame(ribbon, bg="#F3F2F1", height=86, highlightthickness=1,
                               highlightbackground=GRIDLINE)
        groups_bar.pack(side="top", fill="x")

        def make_group(parent, name, items):
            group = tk.Frame(parent, bg="#F3F2F1", padx=6)
            group.pack(side="left", fill="y", padx=(0, 4))
            body = tk.Frame(group, bg="#F3F2F1")
            body.pack(side="top", fill="both", expand=True, pady=(4, 0))
            for item in items:
                b = tk.Label(body, text=item, bg="#F3F2F1", fg="#333333", font=(FONT_FAMILY, 9),
                             padx=6, pady=2)
                b.pack(side="left", anchor="n")
                b.bind("<Button-1>", lambda e, i=item: self._flash(f"'{i}' is not available in this demo."))
            tk.Label(group, text=name, bg="#F3F2F1", fg="#555555", font=(FONT_FAMILY, 8)).pack(side="bottom")
            sep = tk.Frame(parent, bg=GRIDLINE, width=1)
            sep.pack(side="left", fill="y", pady=6)
            return group

        make_group(groups_bar, "Clipboard", ["Paste", "Cut", "Copy", "Format Painter"])
        make_group(groups_bar, "Font", ["Calibri", "B", "I", "U", "Fill", "Color"])
        make_group(groups_bar, "Alignment", ["≡", "Wrap Text", "Merge & Center"])
        make_group(groups_bar, "Number", ["General", "$", "%", ","])
        make_group(groups_bar, "Styles", ["Conditional Formatting", "Format as Table", "Cell Styles"])
        make_group(groups_bar, "Cells", ["Insert", "Delete", "Format"])
        make_group(groups_bar, "Editing", ["Fill", "Clear", "Sort & Filter", "Find & Select"])
        make_group(groups_bar, "Sensitivity", ["Sensitivity"])
        make_group(groups_bar, "Add-ins", ["Add-ins"])

    # -- name box / formula bar ----------------------------------------------

    def _build_formula_bar(self):
        bar = tk.Frame(self.root, bg="white", highlightthickness=1, highlightbackground=GRIDLINE)
        bar.pack(side="top", fill="x")
        self.name_box_var = tk.StringVar(value="A1")
        name_box = tk.Entry(bar, textvariable=self.name_box_var, width=10, font=(FONT_FAMILY, 10),
                             relief="flat", highlightthickness=1, highlightbackground=GRIDLINE,
                             justify="center", state="readonly", readonlybackground="white")
        name_box.pack(side="left", padx=(6, 0), pady=3)
        sep = tk.Frame(bar, bg=GRIDLINE, width=1)
        sep.pack(side="left", fill="y", pady=3, padx=6)
        tk.Label(bar, text="fx", fg=RIBBON_GREEN, bg="white", font=(FONT_FAMILY, 10, "italic")).pack(side="left")

        self.formula_var = tk.StringVar(value="")
        formula_entry = tk.Entry(bar, textvariable=self.formula_var, font=(FONT_FAMILY, 11),
                                  relief="flat", highlightthickness=0, bg="white")
        formula_entry.pack(side="left", fill="x", expand=True, padx=(6, 6), pady=3)
        formula_entry.bind("<Return>", self._commit_formula_bar)

    def _commit_formula_bar(self, event=None):
        if self.selected is None:
            return
        c, r = self.selected
        value = self.formula_var.get()
        self.cell_values[(c, r)] = value
        self.cell_labels[(c, r)].config(text=value)

    # -- grid -----------------------------------------------------------------

    def _build_grid(self):
        container = tk.Frame(self.root, bg="white")
        container.pack(side="top", fill="both", expand=True)

        canvas = tk.Canvas(container, bg="white", highlightthickness=0)
        vsb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(canvas, bg="white")
        canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 40), "units"))

        # Column-letter bar (row above the sheet's own header row).
        letter_row = tk.Frame(inner, bg=HEADER_BG)
        letter_row.grid(row=0, column=0, columnspan=N_COLS + 1, sticky="ew")
        tk.Label(letter_row, text="", width=NAME_COL_W // 8, bg=HEADER_BG,
                 highlightthickness=1, highlightbackground=HEADER_BORDER).grid(row=0, column=0)
        for c, (_, width, _) in enumerate(COLUMNS):
            tk.Label(letter_row, text=col_letter(c), bg=HEADER_BG, fg="#444444",
                     font=(FONT_FAMILY, 9), width=max(4, width // 8),
                     highlightthickness=1, highlightbackground=HEADER_BORDER).grid(row=0, column=c + 1)

        total_rows = 1 + DATA_ROW_COUNT + BLANK_ROWS_BELOW  # row 1 = sheet header row
        for r in range(1, total_rows + 1):
            tk.Label(inner, text=str(r), bg=HEADER_BG, fg="#444444", font=(FONT_FAMILY, 9),
                     width=max(4, NAME_COL_W // 8), highlightthickness=1,
                     highlightbackground=HEADER_BORDER).grid(row=r, column=0, sticky="nsew")
            for c, (header, width, has_arrow) in enumerate(COLUMNS):
                if r == 1:
                    value, bg, fg = self._header_cell(c, header, has_arrow)
                    font = (FONT_FAMILY, 9, "bold")
                elif r - 1 <= DATA_ROW_COUNT:
                    value, bg, fg = self._data_cell(c, r - 2)
                    font = (FONT_FAMILY, 10)
                else:
                    value, bg, fg = "", "white", "black"
                    font = (FONT_FAMILY, 10)
                lbl = tk.Label(inner, text=value, bg=bg, fg=fg, font=font, anchor="w",
                               width=max(4, width // 8), highlightthickness=1,
                               highlightbackground=GRIDLINE, padx=3)
                lbl.grid(row=r, column=c + 1, sticky="nsew")
                self.cell_labels[(c, r)] = lbl
                self.cell_values[(c, r)] = value
                lbl.bind("<Button-1>", lambda e, c=c, r=r: self._select(c, r))
                if has_arrow and r == 1:
                    lbl.bind("<Button-1>", lambda e, letter=col_letter(c): self._filter_clicked(letter))

    def _header_cell(self, c, header, has_arrow):
        if c in LEGEND_CELLS:
            text, bg, fg = LEGEND_CELLS[c]
            return text, bg, fg
        text = f"{header} ▾" if has_arrow and header else header
        return text, HEADER_BG, "black"

    def _data_cell(self, c, row_idx):
        row = ROWS[row_idx]
        raw = row.get(c, "")
        if isinstance(raw, tuple):
            text, bg, fg = raw
            bg = bg or "white"
            fg = fg or "black"
        else:
            text, bg, fg = raw, "white", "black"
        if c == 0:
            bg = YELLOW
        return text, bg, fg

    def _select(self, c, r):
        if self.selected is not None and self.selected in self.cell_labels:
            self.cell_labels[self.selected].config(highlightbackground=GRIDLINE, highlightthickness=1)
        self.selected = (c, r)
        self.cell_labels[(c, r)].config(highlightbackground=SELECT_BORDER, highlightthickness=2)
        ref = f"{col_letter(c)}{r}"
        self.name_box_var.set(ref)
        self.formula_var.set(self.cell_values.get((c, r), ""))
        self.status_var.set("Ready")

    def _filter_clicked(self, letter):
        self._flash(f"AutoFilter on column {letter} is not available in this demo.")

    # -- bottom bar: sheet tabs (left) + status/view/zoom (right) -----------

    def _build_bottom_bar(self):
        bar = tk.Frame(self.root, bg=STATUS_BG, height=26)
        bar.pack(side="bottom", fill="x")

        tabs = tk.Frame(bar, bg=STATUS_BG)
        tabs.pack(side="left", padx=4)
        active = tk.Label(tabs, text="Vendor List (A-Z)", font=(FONT_FAMILY, 9, "bold"), bg="white",
                           fg=RIBBON_GREEN, padx=10, pady=4, highlightthickness=1,
                           highlightbackground=RIBBON_GREEN)
        active.pack(side="left", pady=(2, 0))
        active.bind("<Button-1>", lambda e: self._flash("'Vendor List (A-Z)' is already active."))
        for name in ["FY2026-2027 Worksheet", "Notes"]:
            lbl = tk.Label(tabs, text=name, font=(FONT_FAMILY, 9), bg=STATUS_BG, fg="#444444",
                           padx=10, pady=4)
            lbl.pack(side="left", pady=(2, 0))
            lbl.bind("<Button-1>", lambda e, n=name: self._flash(f"'{n}' is not available in this demo."))
        plus = tk.Label(tabs, text="+", font=(FONT_FAMILY, 10), bg=STATUS_BG, fg="#444444",
                        padx=8, pady=4, cursor="hand2")
        plus.pack(side="left", pady=(2, 0))
        plus.bind("<Button-1>", lambda e: self._flash("Adding sheets is not available in this demo."))

        self.status_var = tk.StringVar(value="Ready")
        tk.Label(bar, textvariable=self.status_var, font=(FONT_FAMILY, 9), bg=STATUS_BG,
                 fg="#333333").pack(side="left", padx=14)

        right = tk.Frame(bar, bg=STATUS_BG)
        right.pack(side="right", padx=8)
        self.zoom_var = tk.IntVar(value=100)
        self.zoom_label = tk.Label(right, text="100%", font=(FONT_FAMILY, 9), bg=STATUS_BG, fg="#333333")
        self.zoom_label.pack(side="right", padx=(4, 0))
        zoom_scale = ttk.Scale(right, from_=10, to=400, orient="horizontal", length=90,
                              variable=self.zoom_var, command=self._zoom_changed)
        zoom_scale.pack(side="right")
        for sym, name in [("▦", "Normal"), ("▤", "Page Break Preview"), ("▥", "Page Layout")]:
            lbl = tk.Label(right, text=sym, font=(FONT_FAMILY, 10), bg=STATUS_BG, fg="#444444",
                          padx=4, cursor="hand2")
            lbl.pack(side="right")
            lbl.bind("<Button-1>", lambda e, n=name: self._flash(f"{n} view is not available in this demo."))

    def _zoom_changed(self, value):
        self.zoom_label.config(text=f"{int(float(value))}%")

    def _flash(self, message):
        self.status_var.set(message)

    # -- programmatic access for headless verification -----------------------

    def get_cell(self, ref):
        c, r = self._ref_to_index(ref)
        return self.cell_values.get((c, r), "")

    def _ref_to_index(self, ref):
        letters = "".join(ch for ch in ref if ch.isalpha())
        digits = "".join(ch for ch in ref if ch.isdigit())
        c = 0
        for ch in letters.upper():
            c = c * 26 + (ord(ch) - ord("A") + 1)
        return c - 1, int(digits)


# =========================================================================
# AUTOMATION BRIDGE -- strictly opt-in, additive, and read-only with respect
# to app behavior. Gated behind EXCEL_MIRROR_AUTOMATION=1; when that env var
# is unset (normal `python3 excel_mirror.py`), none of this code runs.
#
# Purpose: report live, real, absolute screen coordinates + app state,
# computed fresh from the actual running Tk widgets each time, so an
# external OS-level automation (RPA.Desktop) can click/type into this app
# the same way a human would -- no hardcoded pixels, no image matching.
# Same design as the SAP mirror's main.py (see its own "AUTOMATION BRIDGE"
# section for the full rationale) -- this module only *reports*
# geometry/state; it never triggers actions.
#
# Transport: a tiny file-based request/response IPC, polled from inside the
# existing Tk event loop via self.root.after(...) -- no threads, no server.
# =========================================================================

AUTOMATION_ENABLED = os.environ.get("EXCEL_MIRROR_AUTOMATION") == "1"
BRIDGE_REQUEST_FILE = Path(os.environ.get("EXCEL_MIRROR_REQUEST_FILE", "/tmp/excel_mirror_bridge_request.json"))
BRIDGE_RESPONSE_FILE = Path(os.environ.get("EXCEL_MIRROR_RESPONSE_FILE", "/tmp/excel_mirror_bridge_response.json"))
BRIDGE_POLL_MS = 130


def _bridge_bbox(w):
    w.update_idletasks()
    x, y = w.winfo_rootx(), w.winfo_rooty()
    width, height = w.winfo_width(), w.winfo_height()
    return {
        "ok": True, "x": x, "y": y, "width": width, "height": height,
        "center_x": x + width // 2, "center_y": y + height // 2,
    }


def _bridge_resolve_find_cell(app, req):
    """Find the data row whose `match_column` cell contains `contains`
    (case-insensitive substring), then return the bbox of that same row's
    `target_column` cell (defaults to `match_column` itself if omitted) --
    e.g. find the row where NAME contains "Meridian", then hand back the
    SAP VENDOR CODE cell in that row for a real click."""
    # int(): the .robot side sends these through Robot Framework keyword
    # arguments, which arrive as plain strings ("0") unless explicitly cast
    # there -- cast defensively here instead, since cell_values/cell_labels
    # are keyed by real (int, int) tuples and a string-keyed lookup would
    # silently miss every row (dict.get returns its default, never raises).
    match_column = int(req.get("match_column"))
    contains = (req.get("contains") or "").strip().lower()
    target_column = int(req.get("target_column", match_column))
    match_row = None
    for r in range(2, DATA_ROW_COUNT + 2):
        value = str(app.cell_values.get((match_column, r), "")).strip().lower()
        # Bidirectional: a legal-entity suffix (", LLC"/", Inc.") often appears
        # on an invoice but not in a shorthand Excel vendor list, or vice versa
        # -- match if either string contains the other, not just one direction.
        if value and (contains in value or value in contains):
            match_row = r
            break
    if match_row is None:
        return {"ok": False, "error": f"no row with column {match_column} containing {contains!r} found"}
    label = app.cell_labels.get((target_column, match_row))
    if label is None:
        return {"ok": False, "error": f"column {target_column} not found in row {match_row}"}
    result = _bridge_bbox(label)
    result["row"] = match_row
    return result


def _bridge_resolve_state(app, req):
    return {
        "ok": True,
        "name_box": app.name_box_var.get(),
        "formula_bar": app.formula_var.get(),
        "status": app.status_var.get(),
    }


_BRIDGE_RESOLVERS = {
    "find_cell": _bridge_resolve_find_cell,
    "state": _bridge_resolve_state,
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
    VendorListApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
