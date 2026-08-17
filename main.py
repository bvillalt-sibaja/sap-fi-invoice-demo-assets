"""SAP FI/AP Invoice Entry mirror demo -- single-file app.

Mirrors a trimmed SAP GUI FI/AP invoice workflow -- FB60 invoice entry plus
the object-services attachment flow -- for RPA-automation demo purposes. The
app launches straight into FB60 (Enter Supplier Invoice: Company Code 5300);
there is no SAP Easy Access shell, Company Code entry dialog, FBL1N vendor
line-item display, or Connections/SAP Logon dialog. Everything -- screens,
dialogs/popups, theme/icon helpers, and all dummy data -- lives in this one
file; there is no local package import and no external data/*.json files.
Resets on restart (in-memory only).

Run with: python3 main.py

Visual reference: hand-matched against reference screenshots of the real SAP
GUI windows (title bar/menu/toolbar chrome, field label coloring for
required-entry fields, dialog layouts, grid column order). See the project
notes for the handful of things a Tkinter app can't reproduce pixel-for-pixel
(native Windows dialog chrome, the real SAP GUI's continuous-scroll calendar,
and the Tahoma font substitution on macOS).
"""

import copy
import datetime
import calendar as _calendar_mod
import datetime as _dt
import json
import os
import tkinter as tk
from pathlib import Path
from tkinter import ttk, messagebox
from decimal import Decimal, InvalidOperation

# =========================================================================
# THEME -- visual constants approximating the classic SAP GUI look, hand
# matched against the reference screenshots (title bar slate-blue, white
# toolbar/menu strip, blue required-entry field labels, etc).
# =========================================================================

WINDOW_BG = "#ECECE4"
TOOLBAR_BG = "#F2F2F0"
MENUBAR_BG = "#FFFFFF"
CONTENT_BG = "#FFFFFF"
TREE_BG = "#FFFFFF"
STATUSBAR_BG = "#FFFFFF"
STATUSBAR_FG = "#000000"
TITLE_BG = "#3F5468"          # slate blue-gray title bar (FB60 window title)
TITLE_FG = "#FFFFFF"
HEADER_BG = "#FFFFFF"
HEADER_FG = "#000000"
LINK_FG = "#0000CC"
BUTTON_RAISED = "#F2F2F0"
GRID_HEADER_BG = "#DCE6EF"
GRID_ROW_ALT_BG = "#F2F6F9"
FIELD_BG = "#FFFFFF"
READONLY_BG = "#ECECE6"
ERROR_FG = "#B00000"
OK_FG = "#1B6B1B"
DIALOG_BG = "#F7F7F5"

# Required-entry field labels (Supplier, Invoice date, Reference, Amount)
# render in a distinctive blue in real SAP GUI classic theme -- everything
# else stays black.
REQUIRED_LABEL_FG = "#1B4F8C"
TAB_ACTIVE_FG = "#0F5FA8"

# Real SAP GUI on Windows renders in Tahoma; macOS doesn't ship that font so
# Tk silently substitutes a default sans-serif here (verified: still legible,
# just not pixel-identical) -- on the actual Windows target this renders
# correctly.
FONT_FAMILY = "Tahoma"
FONT_NORMAL = (FONT_FAMILY, 11)
FONT_SMALL = (FONT_FAMILY, 10)
FONT_TINY = (FONT_FAMILY, 9)
FONT_HEADER = (FONT_FAMILY, 12, "bold")
FONT_SMALL_BOLD = (FONT_FAMILY, 10, "bold")
FONT_MONO = ("Courier New", 11)

APP_TITLE = "SAP"


def style_treeview(style, name, heading_name=None):
    """Force explicit fg/bg on a ttk Treeview style.

    The native aqua theme derives default text color from the OS light/dark
    appearance, which leaves text invisible against an explicitly forced
    light background when macOS is in Dark Mode. Setting foreground/selected
    colors explicitly here keeps the look consistent regardless of OS theme.
    """
    style.configure(
        name,
        background=TREE_BG,
        fieldbackground=TREE_BG,
        foreground=HEADER_FG,
        font=FONT_NORMAL,
        rowheight=22,
    )
    style.map(
        name,
        background=[("selected", TITLE_BG)],
        foreground=[("selected", "#FFFFFF")],
    )
    style.configure(
        heading_name or f"{name}.Heading",
        background=GRID_HEADER_BG,
        foreground=HEADER_FG,
        font=FONT_SMALL,
    )


def style_combobox(style):
    """Explicit fg/bg for Combobox popdown lists (same aqua dark-mode issue)."""
    style.configure(
        "Sap.TCombobox",
        fieldbackground=FIELD_BG,
        background=FIELD_BG,
        foreground=HEADER_FG,
    )
    style.map(
        "Sap.TCombobox",
        fieldbackground=[("readonly", FIELD_BG)],
        foreground=[("readonly", HEADER_FG)],
        selectbackground=[("readonly", FIELD_BG)],
        selectforeground=[("readonly", HEADER_FG)],
    )


# =========================================================================
# ICONS -- small hand-drawn PhotoImage icons for the tree, toolbar, status
# bar, and dialogs. Drawn as raw pixel grids rather than Unicode emoji so the
# look is identical on macOS (dev/test) and Windows (where the RPA demo
# actually runs) instead of depending on whichever emoji font each OS
# happens to substitute, and to avoid astral-plane emoji (e.g. folder/
# magnifier glyphs) rendering as blank tofu on some Tk/Tcl builds.
# =========================================================================

_SIZE = 16
_FOLDER_BODY = "#F0C040"
_FOLDER_EDGE = "#8A6D1D"
_STAR_FILL = "#2B62B3"

_icon_cache = {}


def _blank_grid(size=_SIZE, bg=None):
    bg = bg or TREE_BG
    return [[bg] * size for _ in range(size)]


def _rect(grid, x0, y0, x1, y1, color):
    h, w = len(grid), len(grid[0])
    for y in range(int(y0), int(y1) + 1):
        for x in range(int(x0), int(x1) + 1):
            if 0 <= y < h and 0 <= x < w:
                grid[y][x] = color


def _rect_outline(grid, x0, y0, x1, y1, color):
    for x in range(int(x0), int(x1) + 1):
        _rect(grid, x, y0, x, y0, color)
        _rect(grid, x, y1, x, y1, color)
    for y in range(int(y0), int(y1) + 1):
        _rect(grid, x0, y, x0, y, color)
        _rect(grid, x1, y, x1, y, color)


def _line(grid, x0, y0, x1, y1, color, thickness=1):
    h, w = len(grid), len(grid[0])
    steps = int(round(max(abs(x1 - x0), abs(y1 - y0), 1)))
    half = thickness // 2
    for i in range(steps + 1):
        t = i / steps
        cx = round(x0 + (x1 - x0) * t)
        cy = round(y0 + (y1 - y0) * t)
        for dy in range(-half, thickness - half):
            for dx in range(-half, thickness - half):
                x, y = cx + dx, cy + dy
                if 0 <= y < h and 0 <= x < w:
                    grid[y][x] = color


def _circle(grid, cx, cy, r, color):
    h, w = len(grid), len(grid[0])
    for y in range(h):
        for x in range(w):
            if (x - cx) ** 2 + (y - cy) ** 2 <= r * r:
                grid[y][x] = color


def _ring(grid, cx, cy, r, thickness, color):
    h, w = len(grid), len(grid[0])
    for y in range(h):
        for x in range(w):
            d2 = (x - cx) ** 2 + (y - cy) ** 2
            if (r - thickness) ** 2 <= d2 <= r * r:
                grid[y][x] = color


def _checkmark(grid, color, thickness=2, x0=0.16, y0=0.55, x1=0.42, y1=0.82, x2=0.86, y2=0.2):
    size = len(grid)
    _line(grid, size * x0, size * y0, size * x1, size * y1, color, thickness)
    _line(grid, size * x1, size * y1, size * x2, size * y2, color, thickness)


def _to_photoimage(grid):
    img = tk.PhotoImage(width=len(grid[0]), height=len(grid))
    img.put(grid)
    return img


def _cached(name, builder):
    if name not in _icon_cache:
        _icon_cache[name] = _to_photoimage(builder())
    return _icon_cache[name]


# -------------------------------------------------------------- tree icons
def _folder_pixels():
    grid = _blank_grid()
    _rect(grid, 1, 2, 6, 4, _FOLDER_BODY)
    _rect(grid, 1, 5, 14, 12, _FOLDER_BODY)
    _rect(grid, 1, 4, 14, 4, _FOLDER_EDGE)
    _rect(grid, 1, 12, 14, 12, _FOLDER_EDGE)
    return grid


_STAR_PATTERN = [
    "....X....",
    "....X....",
    "...XXX...",
    "XXXXXXXXX",
    ".XXXXXXX.",
    "..XXXXX..",
    ".XX...XX.",
    "XX.....XX",
    ".........",
]


def _star_pixels():
    grid = _blank_grid()
    offset_y = (_SIZE - len(_STAR_PATTERN)) // 2
    offset_x = (_SIZE - len(_STAR_PATTERN[0])) // 2
    for j, row in enumerate(_STAR_PATTERN):
        for i, ch in enumerate(row):
            if ch == "X":
                grid[offset_y + j][offset_x + i] = _STAR_FILL
    return grid


def folder_icon():
    return _cached("folder", _folder_pixels)


def star_icon():
    return _cached("star", _star_pixels)


# ------------------------------------------------------- line-item row status
def _checkbox_empty_pixels():
    grid = _blank_grid(size=14, bg=TREE_BG)
    _rect_outline(grid, 2, 2, 11, 11, "#888888")
    return grid


def _checkmark_row_pixels():
    grid = _blank_grid(size=14, bg=TREE_BG)
    _checkmark(grid, "#1B8A1B", thickness=2, x0=0.1, y0=0.55, x1=0.38, y1=0.85, x2=0.9, y2=0.15)
    return grid


def checkbox_empty_icon():
    return _cached("checkbox_empty", _checkbox_empty_pixels)


def checkmark_row_icon():
    return _cached("checkmark_row", _checkmark_row_pixels)


# ----------------------------------------------------------- toolbar icons
def _enter_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    _checkmark(grid, "#1B8A1B", thickness=2)
    return grid


def _save_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    _rect(grid, 2, 2, 13, 13, "#3B6EA5")
    _rect_outline(grid, 2, 2, 13, 13, "#1F3B57")
    _rect(grid, 4, 3, 10, 6, "#DCE6F1")
    _rect(grid, 3, 9, 12, 12, "#1F3B57")
    return grid


def _cancel_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#B00020"
    _line(grid, 3, 3, 12, 12, color, thickness=2)
    _line(grid, 12, 3, 3, 12, color, thickness=2)
    return grid


def _exit_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#B00020"
    _rect_outline(grid, 3, 2, 9, 13, color)
    _line(grid, 8, 8, 14, 8, color, thickness=2)
    _line(grid, 14, 8, 11, 5, color, thickness=2)
    _line(grid, 14, 8, 11, 11, color, thickness=2)
    return grid


def _print_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    _rect(grid, 4, 2, 11, 7, "#FFFFFF")
    _rect_outline(grid, 4, 2, 11, 7, "#666666")
    _rect(grid, 2, 7, 13, 11, "#808080")
    _rect_outline(grid, 2, 7, 13, 11, "#4D4D4D")
    _rect(grid, 4, 12, 11, 13, "#FFFFFF")
    _rect(grid, 10, 9, 11, 9, "#3B6EA5")
    return grid


def _hold_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#8A6D1D"
    _rect(grid, 4, 2, 6, 13, color)
    _rect(grid, 9, 2, 11, 13, color)
    return grid


def _simulate_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#3B6EA5"
    _rect_outline(grid, 2, 2, 13, 13, color)
    _line(grid, 5, 5, 5, 11, color, thickness=1)
    _line(grid, 5, 5, 11, 8, color, thickness=1)
    _line(grid, 5, 11, 11, 8, color, thickness=1)
    return grid


def _park_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#3B6EA5"
    _rect_outline(grid, 3, 4, 12, 12, color)
    _line(grid, 6, 4, 6, 2, color, thickness=2)
    _line(grid, 10, 4, 10, 2, color, thickness=2)
    return grid


def _post_pixels():
    # Real SAP GUI's Document Overview 'Post' toolbar icon is a red plus.
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#C0392B"
    _line(grid, 8, 2, 8, 13, color, thickness=2)
    _line(grid, 2, 8, 13, 8, color, thickness=2)
    return grid


def _tree_toggle_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#2B2B2B"
    _rect_outline(grid, 2, 2, 13, 13, color)
    _line(grid, 5, 5, 10, 5, color, thickness=1)
    _line(grid, 5, 8, 10, 8, color, thickness=1)
    _line(grid, 5, 11, 8, 11, color, thickness=1)
    return grid


def _options_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#555555"
    _circle(grid, 8, 8, 2, color)
    for angle_pts in [(8, 2), (8, 14), (2, 8), (14, 8)]:
        _rect(grid, angle_pts[0] - 1, angle_pts[1] - 1, angle_pts[0] + 1, angle_pts[1] + 1, color)
    return grid


def enter_icon():
    return _cached("enter", _enter_pixels)


def save_icon():
    return _cached("save", _save_pixels)


def cancel_icon():
    return _cached("cancel", _cancel_pixels)


def exit_icon():
    return _cached("exit", _exit_pixels)


def print_icon():
    return _cached("print", _print_pixels)


def hold_icon():
    return _cached("hold", _hold_pixels)


def simulate_icon():
    return _cached("simulate", _simulate_pixels)


def park_icon():
    return _cached("park", _park_pixels)


def post_icon():
    return _cached("post", _post_pixels)


def tree_toggle_icon():
    return _cached("tree_toggle", _tree_toggle_pixels)


def options_icon():
    return _cached("options", _options_pixels)


def _new_session_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#2B2B2B"
    _rect_outline(grid, 3, 2, 11, 13, color)
    _line(grid, 3, 5, 11, 5, color, thickness=1)
    _line(grid, 8, 8, 8, 12, color, thickness=1)
    _line(grid, 6, 10, 10, 10, color, thickness=1)
    return grid


def _forward_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#2B2B2B"
    _line(grid, 3, 8, 11, 8, color, thickness=2)
    _line(grid, 11, 8, 8, 4, color, thickness=2)
    _line(grid, 11, 8, 8, 12, color, thickness=2)
    return grid


def new_session_icon():
    return _cached("new_session", _new_session_pixels)


def forward_icon():
    return _cached("forward", _forward_pixels)


# --------------------------------------------------------------- status dots
def _dot_pixels(color, size=12, bg=None):
    grid = _blank_grid(size=size, bg=bg or CONTENT_BG)
    _circle(grid, size / 2, size / 2, size / 2 - 1, color)
    return grid


def red_dot_icon():
    return _cached("red_dot", lambda: _dot_pixels("#C0392B"))


def green_dot_icon():
    return _cached("green_dot", lambda: _dot_pixels("#1B8A1B"))


# ------------------------------------------------------- balance indicator
# Real FB60 shows a small 3-segment indicator next to "Bal.": unbalanced is
# a filled red circle + two hollow circles; balanced is two hollow circles
# + a filled green square.
def _balance_indicator_pixels(balanced):
    w, h = 34, 12
    grid = [[CONTENT_BG] * w for _ in range(h)]
    cy = h // 2
    centers = [6, 17, 28]
    color = "#555555"
    if not balanced:
        _circle(grid, centers[0], cy, 4, "#C0392B")
        _ring(grid, centers[1], cy, 4, 1.4, color)
        _ring(grid, centers[2], cy, 4, 1.4, color)
    else:
        _ring(grid, centers[0], cy, 4, 1.4, color)
        _ring(grid, centers[1], cy, 4, 1.4, color)
        _rect(grid, centers[2] - 3, cy - 3, centers[2] + 3, cy + 3, "#1B8A1B")
    return grid


def balance_indicator_icon(balanced):
    return _cached(f"balance_{balanced}", lambda: _balance_indicator_pixels(balanced))


# --------------------------------------------------------------- status bar icons
def _status_circle_pixels():
    grid = _blank_grid(size=12, bg=STATUSBAR_BG)
    _circle(grid, 5.5, 5.5, 5, "#2E7D32")
    return grid


def _lock_pixels():
    grid = _blank_grid(size=12, bg=STATUSBAR_BG)
    color = "#555555"
    _rect_outline(grid, 3, 5, 8, 10, color)
    _rect_outline(grid, 4, 1, 7, 5, color)
    return grid


def status_circle_icon():
    return _cached("status_circle", _status_circle_pixels)


def lock_icon():
    return _cached("lock", _lock_pixels)


# -------------------------------------------------------------- dialog icons
def _success_pixels():
    size = 32
    grid = _blank_grid(size=size, bg=CONTENT_BG)
    _circle(grid, size / 2, size / 2, size / 2 - 1, "#2E7D32")
    _checkmark(
        grid, "#FFFFFF", thickness=3,
        x0=0.22, y0=0.55, x1=0.42, y1=0.75, x2=0.82, y2=0.28,
    )
    return grid


def _error_pixels():
    size = 32
    grid = _blank_grid(size=size, bg=CONTENT_BG)
    _circle(grid, size / 2, size / 2, size / 2 - 1, "#B00020")
    _line(grid, size * 0.28, size * 0.28, size * 0.72, size * 0.72, "#FFFFFF", thickness=3)
    _line(grid, size * 0.72, size * 0.28, size * 0.28, size * 0.72, "#FFFFFF", thickness=3)
    return grid


def success_icon():
    return _cached("success", _success_pixels)


def error_icon():
    return _cached("error", _error_pixels)


# ------------------------------------------------------------- misc icons
def _magnifier_pixels():
    grid = _blank_grid(bg=FIELD_BG)
    color = "#2B2B2B"
    _ring(grid, 6.5, 6.5, 5, 1.6, color)
    _line(grid, 10, 10, 14.5, 14.5, color, thickness=2)
    return grid


def magnifier_icon():
    return _cached("magnifier", _magnifier_pixels)


def _calendar_pixels():
    grid = _blank_grid(bg=FIELD_BG)
    color = "#2B2B2B"
    _rect_outline(grid, 1, 3, 14, 14, color)
    _line(grid, 1, 6, 14, 6, color, thickness=1)
    _rect(grid, 3, 1, 4, 3, color)
    _rect(grid, 11, 1, 12, 3, color)
    _rect(grid, 3, 8, 4, 9, "#3B6EA5")
    _rect(grid, 7, 8, 8, 9, "#3B6EA5")
    _rect(grid, 11, 8, 12, 9, "#3B6EA5")
    _rect(grid, 3, 11, 4, 12, "#3B6EA5")
    _rect(grid, 7, 11, 8, 12, "#3B6EA5")
    return grid


def calendar_icon():
    return _cached("calendar", _calendar_pixels)


def _paperclip_pixels():
    grid = _blank_grid(bg=FIELD_BG)
    color = "#555555"
    _line(grid, 10, 2, 5, 9, color, thickness=2)
    _ring(grid, 6, 10, 3, 1.4, color)
    _line(grid, 10, 2, 13, 5, color, thickness=2)
    return grid


def paperclip_icon():
    return _cached("paperclip", _paperclip_pixels)


def _new_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#1B6B1B"
    _line(grid, 8, 2, 8, 13, color, thickness=2)
    _line(grid, 2, 8, 13, 8, color, thickness=2)
    return grid


def new_icon():
    return _cached("new", _new_pixels)


def _objects_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#3B6EA5"
    _rect(grid, 2, 2, 6, 6, color)
    _rect(grid, 9, 2, 13, 6, color)
    _rect(grid, 2, 9, 6, 13, color)
    _rect(grid, 9, 9, 13, 13, color)
    return grid


def objects_icon():
    return _cached("objects", _objects_pixels)


def _sap_logo_pixels():
    grid = _blank_grid(size=24, bg=TITLE_BG)
    _rect(grid, 1, 4, 22, 19, "#FFFFFF")
    _rect(grid, 3, 7, 20, 16, TITLE_BG)
    return grid


def sap_logo_icon():
    return _cached("sap_logo", _sap_logo_pixels)


def _phone_pixels():
    grid = _blank_grid(bg=CONTENT_BG)
    color = "#2B6EA5"
    _rect(grid, 4, 2, 8, 5, color)
    _rect(grid, 6, 5, 10, 9, color)
    _rect(grid, 8, 9, 12, 12, color)
    _rect(grid, 10, 12, 14, 14, color)
    return grid


def phone_icon():
    return _cached("phone", _phone_pixels)


def _sync_pixels():
    grid = _blank_grid(bg=CONTENT_BG)
    color = "#2B6EA5"
    _ring(grid, 8, 8, 6, 1.6, color)
    _line(grid, 12, 3, 14, 5, color, thickness=2)
    _line(grid, 4, 13, 2, 11, color, thickness=2)
    return grid


def sync_icon():
    return _cached("sync", _sync_pixels)


def _grid_small_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#2B6EA5"
    _rect_outline(grid, 2, 2, 13, 13, color)
    _line(grid, 2, 7, 13, 7, color, thickness=1)
    _line(grid, 7, 2, 7, 13, color, thickness=1)
    return grid


def grid_small_icon():
    return _cached("grid_small", _grid_small_pixels)


# ------------------------------------------------- Document Overview toolbar
def _reset_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#2B6EA5"
    _ring(grid, 8, 8, 5, 1.6, color)
    _line(grid, 12, 4, 14, 4, color, thickness=2)
    _line(grid, 14, 4, 14, 6, color, thickness=2)
    return grid


def _funnel_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#555555"
    _line(grid, 2, 3, 14, 3, color, thickness=2)
    _line(grid, 2, 3, 8, 9, color, thickness=2)
    _line(grid, 14, 3, 8, 9, color, thickness=2)
    _line(grid, 8, 9, 8, 14, color, thickness=2)
    return grid


def _sort_asc_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#2B2B2B"
    _rect(grid, 3, 11, 5, 13, color)
    _rect(grid, 7, 8, 9, 13, color)
    _rect(grid, 11, 4, 13, 13, color)
    return grid


def _sort_desc_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#2B2B2B"
    _rect(grid, 3, 4, 5, 13, color)
    _rect(grid, 7, 8, 9, 13, color)
    _rect(grid, 11, 11, 13, 13, color)
    return grid


def _copy_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#555555"
    _rect_outline(grid, 2, 4, 10, 12, color)
    _rect_outline(grid, 5, 2, 13, 10, color)
    return grid


def _envelope_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#2B6EA5"
    _rect_outline(grid, 1, 3, 14, 12, color)
    _line(grid, 1, 3, 7, 8, color, thickness=1)
    _line(grid, 14, 3, 7, 8, color, thickness=1)
    return grid


def _info_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#2B6EA5"
    _ring(grid, 8, 8, 6, 1.4, color)
    _rect(grid, 7, 4, 8, 5, color)
    _rect(grid, 7, 7, 8, 11, color)
    return grid


def _resize_grip_pixels():
    grid = _blank_grid(size=10, bg=DIALOG_BG)
    color = "#999999"
    for i in range(3):
        _line(grid, 1 + i * 3, 8, 8, 1 + i * 3, color, thickness=1)
    return grid


def reset_icon():
    return _cached("reset", _reset_pixels)


def funnel_icon():
    return _cached("funnel", _funnel_pixels)


def sort_asc_icon():
    return _cached("sort_asc", _sort_asc_pixels)


def sort_desc_icon():
    return _cached("sort_desc", _sort_desc_pixels)


def copy_icon():
    return _cached("copy", _copy_pixels)


def envelope_icon():
    return _cached("envelope", _envelope_pixels)


def info_icon():
    return _cached("info", _info_pixels)


def resize_grip_icon():
    return _cached("resize_grip", _resize_grip_pixels)


# ------------------------------------------------- Attachment list toolbar
def _refresh_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#2B6EA5"
    _ring(grid, 8, 8, 6, 1.6, color)
    _line(grid, 13, 4, 13, 8, color, thickness=2)
    _line(grid, 13, 4, 9, 4, color, thickness=2)
    return grid


def _trash_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#555555"
    _rect_outline(grid, 3, 5, 12, 14, color)
    _line(grid, 2, 5, 13, 5, color, thickness=1)
    _rect(grid, 6, 2, 9, 4, color)
    for x in (5, 8, 10):
        _line(grid, x, 6, x, 13, color, thickness=1)
    return grid


def _printer_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#555555"
    _rect(grid, 4, 1, 11, 5, "#FFFFFF")
    _rect_outline(grid, 4, 1, 11, 5, color)
    _rect(grid, 1, 5, 14, 10, color)
    _rect(grid, 4, 10, 11, 14, "#FFFFFF")
    _rect_outline(grid, 4, 10, 11, 14, color)
    return grid


def _export_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#2B6EA5"
    _rect_outline(grid, 2, 6, 10, 14, color)
    _line(grid, 11, 2, 11, 9, color, thickness=2)
    _line(grid, 11, 2, 8, 5, color, thickness=2)
    _line(grid, 11, 2, 14, 5, color, thickness=2)
    return grid


def refresh_icon():
    return _cached("refresh", _refresh_pixels)


def trash_icon():
    return _cached("trash", _trash_pixels)


def printer_icon():
    return _cached("printer", _printer_pixels)


def export_icon():
    return _cached("export", _export_pixels)


def _pencil_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#2B6EA5"
    _line(grid, 3, 13, 6, 10, color, thickness=2)
    _line(grid, 6, 10, 12, 4, color, thickness=2)
    _line(grid, 12, 4, 14, 6, color, thickness=2)
    _line(grid, 14, 6, 8, 12, color, thickness=2)
    _rect(grid, 2, 12, 4, 14, color)
    return grid


def pencil_icon():
    return _cached("pencil", _pencil_pixels)


# ------------------------------------------------- Vendor search toolbar
def _find_next_pixels():
    grid = _blank_grid(bg=FIELD_BG)
    color = "#2B2B2B"
    _ring(grid, 6, 6, 4, 1.4, color)
    _line(grid, 9, 9, 12, 12, color, thickness=2)
    _line(grid, 12, 5, 14, 5, color, thickness=1)
    _line(grid, 14, 5, 12, 3, color, thickness=1)
    _line(grid, 14, 5, 12, 7, color, thickness=1)
    return grid


def _help_pixels():
    grid = _blank_grid(bg=TOOLBAR_BG)
    color = "#2B6EA5"
    _ring(grid, 8, 8, 6, 1.4, color)
    _rect(grid, 7, 11, 8, 12, color)
    _rect(grid, 6, 4, 10, 5, color)
    _rect(grid, 9, 5, 10, 8, color)
    _rect(grid, 7, 8, 9, 9, color)
    return grid


def find_next_icon():
    return _cached("find_next", _find_next_pixels)


def help_icon():
    return _cached("help", _help_pixels)


# =========================================================================
# WIDGETS -- small reusable widget helpers.
# =========================================================================

class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, event=None):
        if self.tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 6
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(
            self.tip,
            text=self.text,
            bg="#FFFFE0",
            fg="black",
            relief="solid",
            borderwidth=1,
            font=FONT_SMALL,
            padx=4,
            pady=2,
        ).pack()

    def hide(self, event=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None


def add_tooltip(widget, text):
    return Tooltip(widget, text)


def center_over(dlg, parent):
    dlg.update_idletasks()
    px, py = parent.winfo_rootx(), parent.winfo_rooty()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    dw, dh = dlg.winfo_width(), dlg.winfo_height()
    x = px + max(0, (pw - dw) // 2)
    y = py + max(0, (ph - dh) // 2)
    dlg.geometry(f"+{x}+{y}")


def toolbar_text_button(parent, image, text, name, tooltip, command, bg=None):
    """SAP GUI toolbar buttons show an icon *and* a text label side by side
    (e.g. 'Tree on', 'Company Code', 'Hold', 'Simulate', 'Park', 'Exit')."""
    bg = bg or TOOLBAR_BG
    b = tk.Button(
        parent, image=image, text=f" {text}", compound="left",
        bg=bg, relief="raised", bd=1, font=FONT_SMALL, name=name, command=command,
    )
    add_tooltip(b, tooltip)
    return b


def toolbar_button(parent, image, name, tooltip, command, bg=None):
    b = tk.Button(
        parent,
        image=image,
        bg=bg or TOOLBAR_BG,
        relief="raised",
        bd=1,
        name=name,
        command=command,
    )
    add_tooltip(b, tooltip)
    return b


class CustomMenu(tk.Toplevel):
    """Cross-platform stand-in for `tk.Menu`'s `tk_popup()`: renders a
    dropdown as plain Frame/Label rows in a borderless Toplevel instead of
    relying on the OS-native menu (NSMenu on macOS). Confirmed via a direct
    repro that a native `tk.Menu` popup can fail to actually post at all in
    a constrained/headless window-server context -- `winfo_ismapped()`
    stays False and its geometry stays a degenerate `1x1+0+0` -- even though
    this app's *other* windows/buttons/dialogs (plain `Toplevel`s) render
    and receive clicks there just fine; the failure is specific to native
    menu-tracking, not to Tk windows in general, and reproduces identically
    across two different macOS Python builds (Homebrew and python.org). A
    plain Toplevel doesn't depend on that native menu-tracking machinery at
    all, so it posts exactly as reliably as every other popup in this app,
    on the automation host and on a real desktop session alike -- and its
    items are ordinary widgets the automation bridge can locate the exact
    same way it already locates any other widget (bbox), instead of needing
    `tk.Menu`-specific `yposition()`/`entrycget()` introspection.

    `items` uses the same (label, command) / (label, command, accelerator) /
    (None, None) tuple shape the old `tk.Menu`-based menus used, plus an
    optional 4th element -- a nested items list -- to render a cascade
    (opened on hover, SAP's own object-services submenus). Modal like every
    other popup in this app (`grab_set`); Escape closes it without invoking
    anything.
    """

    HILITE = "#DCE6F1"

    def __init__(self, parent, app, items):
        super().__init__(parent)
        self.app = app
        self.overrideredirect(True)
        self.configure(bg="#8F8F8F")  # 1px border, via padding around a white body
        try:
            self.attributes("-topmost", True)
        except tk.TclError:
            pass
        self.rows = []
        self._submenu = None
        body = tk.Frame(self, bg="white")
        body.pack(padx=1, pady=1)
        for item in items:
            self._add_item(body, item)
        self.withdraw()
        self.bind("<Escape>", lambda e: self.close())

    def _add_item(self, body, item):
        if item[0] is None:
            tk.Frame(body, height=1, bg="#D9D9D9").pack(fill="x", padx=1, pady=3)
            return
        label, command = item[0], item[1]
        accelerator = item[2] if len(item) > 2 else ""
        submenu_items = item[3] if len(item) > 3 else None
        row = tk.Frame(body, bg="white")
        row.pack(fill="x")
        lbl = tk.Label(row, text=label, bg="white", anchor="w", font=FONT_SMALL, padx=16, pady=4)
        lbl.pack(side="left", fill="x", expand=True)
        tail = tk.Label(
            row, text=("▸" if submenu_items else accelerator),
            bg="white", fg="#8A8A8A", font=FONT_SMALL, padx=10,
        )
        tail.pack(side="right")
        entry = {
            "label": label, "frame": row, "widgets": (row, lbl, tail),
            "command": command, "submenu_items": submenu_items,
        }
        for w in entry["widgets"]:
            w.bind("<Enter>", lambda e, en=entry: self._enter(en))
            w.bind("<Leave>", lambda e, en=entry: self._leave(en))
            w.bind("<ButtonRelease-1>", lambda e, en=entry: self._click(en))
        self.rows.append(entry)

    def _enter(self, entry):
        for w in entry["widgets"]:
            w.configure(bg=self.HILITE)
        self._close_submenu()
        if entry["submenu_items"]:
            sub = CustomMenu(self.master, self.app, entry["submenu_items"])
            x = entry["frame"].winfo_rootx() + entry["frame"].winfo_width()
            y = entry["frame"].winfo_rooty()
            sub.update_idletasks()
            sub.geometry(f"+{x}+{y}")
            sub.deiconify()
            sub.lift()
            self._submenu = sub

    def _leave(self, entry):
        for w in entry["widgets"]:
            w.configure(bg="white")

    def _close_submenu(self):
        if self._submenu is not None:
            self._submenu.destroy()
            self._submenu = None

    def _click(self, entry):
        if entry["submenu_items"]:
            return
        command = entry["command"]
        self.close()
        if command:
            command()

    def find_item(self, label):
        for entry in self.rows:
            if entry["label"] == label:
                return entry["frame"]
        return None

    def popup(self, x, y):
        self.update_idletasks()
        self.geometry(f"+{x}+{y}")
        self.deiconify()
        self.lift()
        self.app._active_menu = self  # opt-in automation bridge: currently-posted menu
        self.focus_force()
        self.grab_set()

    def close(self):
        self._close_submenu()
        if getattr(self.app, "_active_menu", None) is self:
            self.app._active_menu = None
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()


def build_menu_strip(parent, defs, bg=None):
    """Renders a per-screen SAP-style transaction menu bar (Document/Edit/Goto/...)
    as a row of buttons, each posting a `CustomMenu` (see above) rather than a
    native `tk.Menu`. `defs` is a list of (label, items) where items is a list
    of (item_label, command) or (item_label, command, accelerator) tuples;
    item_label of None inserts a separator; command of None disables the item.
    `accelerator`, if given, renders as SAP's right-aligned keyboard-shortcut
    label (e.g. 'Ctrl+F2') -- purely a visual label here, the real key isn't
    bound.
    """
    bg = bg or MENUBAR_BG
    strip = tk.Frame(parent, bg=bg, name="menu_strip")
    app = parent.controller

    def open_menu(items, mb):
        menu = CustomMenu(strip, app, items)
        menu.popup(mb.winfo_rootx(), mb.winfo_rooty() + mb.winfo_height())

    for label, items in defs:
        mb = tk.Button(
            strip, text=label, bg=bg, font=FONT_SMALL, padx=6, pady=2,
            name=f"menu_{label.lower().replace(' ', '_')}",
        )
        mb.configure(command=lambda items=items, mb=mb: open_menu(items, mb))
        mb.pack(side="left")
    return strip


# =========================================================================
# DATA -- dummy SAP FI/AP master + transactional data, inline as Python
# literals, held in memory for the session (no disk persistence -- resets
# on restart).
# =========================================================================

DEFAULT_LEDGER_GROUP = "0L"
DOC_TYPE_CODE = "KR"
DOC_TYPE_DESC = "KR (Vendor Invoice) Normal document"

VENDORS = [
    {
        "number": "50713082",
        "search_term": "FEDEX",
        "name1": "FEDEX",
        "country": "US",
        "postal_code": "15250-7461",
        "city": "PITTSBURGH",
        "state": "PA",
        "street": "PO Box 7221",
        "phone": "800 463 3339",
        "bank_account": "9988776655",
        "bank_number": "043000096",
        "swift": "PNBPUS3NNYC",
        "payment_terms": "30 Days net",
        "deletion_flag": "",
    },
    {
        "number": "50716153",
        "search_term": "BLUECREST",
        "name1": "Bluecrest Holdings, LLC",
        "country": "US",
        "postal_code": "75001",
        "city": "Meadowbrook",
        "state": "TX",
        "street": "4821 Sunset Ridge Drive, Ste 200",
        "phone": "800 555 0142",
        "bank_account": "5502341190",
        "bank_number": "073145293",
        "swift": "FICTUS20XXX",
        "payment_terms": "30 Days net",
        "deletion_flag": "",
    },
    {
        # Matches the real demo invoice fetched from Google Drive in Part 1 of
        # the "Process New Non-PO Invoice Request" .robot (Invoice_MFL119284.pdf)
        # -- address/phone/payment terms copied verbatim from that PDF's own
        # "Please Remit To" block, since the Acrobat-read iteration is supposed
        # to find and select THIS exact vendor. "number" matches the "SAP
        # VENDOR CODE" column excel_mirror.py's real Meridian Freight Lines
        # row shows, since the Excel-read iteration types this value directly
        # into the Supplier field expecting it to resolve to this same vendor.
        "number": "SAP-10234",
        "search_term": "MERIDIAN",
        "name1": "Meridian Freight Lines, LLC",
        "country": "US",
        "postal_code": "28241",
        "city": "Charlotte",
        "state": "NC",
        "street": "P.O. Box 8842",
        "phone": "704 555 0148",
        "bank_account": "6613098420",
        "bank_number": "053112345",
        "swift": "FICTUS21XXX",
        "payment_terms": "Net 30 Days",
        "deletion_flag": "",
    },
    {
        # Matches the 2nd real demo invoice fetched from Google Drive in Part 1
        # of the "Process New Non-PO Invoice Request" .robot
        # (Invoice_WM2201947.pdf) -- address copied verbatim from that PDF's
        # own "Remit Payment To" block, for the Acrobat-read iteration's
        # search-by-name to find and select THIS exact vendor. This one is
        # NOT tied to excel_mirror.py's own "SAP-10871" code for the same
        # vendor (that row is only used by the Excel-lookup iteration, which
        # processes Meridian Freight, not this invoice) -- "number" is just a
        # plausible SAP-style vendor number of its own.
        "number": "50719560",
        "search_term": "WASTE MGMT",
        "name1": "Waste Mgmt",
        "country": "US",
        "postal_code": "60566",
        "city": "Naperville",
        "state": "IL",
        "street": "P.O. Box 9410",
        "phone": "855 555 0170",
        "bank_account": "7724501893",
        "bank_number": "071925577",
        "swift": "FICTUS22XXX",
        "payment_terms": "Net 30 Days",
        "deletion_flag": "",
    },
]

DEFAULT_COMPANY_CODE = "5300"

COMPANY_CODES = [
    {"code": "5300", "name": "Northwind Traders, Inc.", "city": "Fairview", "currency": "USD"},
]

GL_ACCOUNTS = [
    {"acct": "5521020000", "text": "Logistics Services"},
    {"acct": "600100", "text": "Freight Expense"},
    {"acct": "601200", "text": "Office Supplies"},
    {"acct": "602300", "text": "Legal & Professional Fees"},
    {"acct": "603400", "text": "IT Services"},
    {"acct": "605100", "text": "Waste & Environmental Services"},
    # Matches the GL code shown on the "Freight" tab of the Vendor Invoice
    # Tracker mockup (freight_gl_mirror.py) -- "GL-4200" -- for the Acrobat
    # iteration's GL-Coding-source branch (map: "Read: GL Code in Freight
    # Sheet in Excel" -> "Write: GL Acct in SAP").
    {"acct": "4200", "text": "Freight - GL Coding"},
    {"acct": "210000", "text": "Accounts Payable - Trade"},
]

# "code" is the actual Cost Center master-data key (what gets typed into the
# grid's Cost center cell); "segment"/"segment_sequence"/"segment_desc" are
# the derived-on-Simulate segment fields shown in the status message and the
# Document Overview/Display "Segment" column -- these are NOT the same value
# as the cost center code itself (real SAP derives a segment from a cost
# center's assignment, they just happen to look similar here).
COST_CENTERS = [
    {
        "code": "CU1X24005",
        "profit_center": "BU00000001",
        "segment": "CNTRY_US",
        "segment_sequence": "0549",
        "segment_desc": "SEGMENT USA COST OBJECT PRINCIPLE",
    },
    # Matches "Facilities -1100" on the same Freight-tab mockup's GL-coding
    # list -- the Acrobat iteration's Cost-Center-source branch (map: "Read:
    # Facilities in Freight Sheet in Excel" -> "Write: Cost Center in SAP").
    {
        "code": "1100",
        "profit_center": "BU00000002",
        "segment": "CNTRY_US",
        "segment_sequence": "0550",
        "segment_desc": "SEGMENT USA FACILITIES PRINCIPLE",
    },
]

TAX_CODES = [
    "D0 (~O_DE_0,00%_Output VAT not taxable)",
    "D4 (~O_DE_0,00%_Output VAT not taxable)",
    "D5 (~O_DE_19,0%_Standard rated output VAT)",
    "DA (~I_DE_19,0%_Standard input rated VAT)",
    "DE (~I_DE_19,0%_Standard Import VAT)",
    "DZ (~I_DE_0,00%_Input VAT not taxable)",
    "S2 (~O_SG_00.0%_Non-taxable supplies)",
    "SY (~I_SG_00.0%_Non-taxable purchases)",
    "U0 (~O_US_00.0%_Tax Exempt_No External Calculation)",
    "U5 (~O_US_100.0%_Taxable_External Tax Determination)",
    "UA (~I_US_100.0%_Taxable_Vendor Billed)",
    "UB (~I_US_100.0%_Use Tax Accrual)",
    "UC (~I_US_100.0%_Taxable_Outside Vendor Tolerance)",
    "UV (~I_US_00.0%_Tax Exempt_No External Calculation)",
]

def fmt_amount(value):
    """Format a Decimal/float as SAP-style '1,234.56' (no sign)."""
    try:
        d = Decimal(str(value))
    except InvalidOperation:
        return str(value)
    negative = d < 0
    d = abs(d)
    text = f"{d:,.2f}"
    return f"{text}-" if negative else text


def parse_amount(text):
    """Parse a plain or SAP-style trailing-minus amount string into a Decimal."""
    text = (text or "").strip().replace(",", "")
    if not text:
        raise InvalidOperation("empty amount")
    negative = text.endswith("-")
    if negative:
        text = text[:-1]
    if text.startswith("-"):
        negative = True
        text = text[1:]
    value = Decimal(text)
    return -value if negative else value


def today_mmddyyyy():
    return datetime.date.today().strftime("%m/%d/%Y")


class DemoData:
    """In-memory store seeded from the literals above. Resets on app restart."""

    def __init__(self):
        self.vendors = copy.deepcopy(VENDORS)
        self.company_codes = copy.deepcopy(COMPANY_CODES)
        self.gl_accounts = copy.deepcopy(GL_ACCOUNTS)
        self.cost_centers = copy.deepcopy(COST_CENTERS)
        self.tax_codes = list(TAX_CODES)
        # attachments: {doc_number: [ {title, creator, created_on}, ... ]}
        self.attachments = {}

        # documents: {doc_number: {header + lines}} -- used by Display Document
        # and by the attachment list.
        self.documents = {}

        # Live posting counter, seeded at the SAP-given base document number.
        self._next_doc_number = 1900000084

    # ------------------------------------------------------------- lookups
    def get_vendor(self, number):
        number = (number or "").strip()
        for v in self.vendors:
            if v["number"] == number:
                return v
        return None

    def search_vendors(self, name=None, search_term=None, country=None,
                        postal_code=None, city=None, supplier=None, max_hits=None):
        def _match(field, needle):
            if not needle:
                return True
            needle = needle.strip().lower()
            field = (field or "").lower()
            if needle.startswith("*") and needle.endswith("*") and len(needle) > 1:
                return needle[1:-1] in field
            if needle.startswith("*"):
                return field.endswith(needle[1:])
            if needle.endswith("*"):
                return field.startswith(needle[:-1])
            return needle in field

        results = []
        for v in self.vendors:
            if supplier and supplier.strip() and v["number"] != supplier.strip():
                continue
            if not _match(v.get("name1"), name):
                continue
            if not _match(v.get("search_term"), search_term):
                continue
            if not _match(v.get("country"), country):
                continue
            if not _match(v.get("postal_code"), postal_code):
                continue
            if not _match(v.get("city"), city):
                continue
            results.append(v)
        if max_hits:
            try:
                results = results[: int(max_hits)]
            except ValueError:
                pass
        return results

    def get_company_code(self, code):
        code = (code or "").strip()
        for cc in self.company_codes:
            if cc["code"] == code:
                return cc
        return None

    def get_gl_account(self, acct):
        acct = (acct or "").strip()
        for g in self.gl_accounts:
            if g["acct"] == acct:
                return g
        return None

    def get_cost_center(self, code):
        code = (code or "").strip()
        for c in self.cost_centers:
            if c["code"] == code:
                return c
        return None

    # ----------------------------------------------------------- documents
    def get_document(self, doc_number):
        return self.documents.get(doc_number)

    def post_invoice(self, vendor_number, header, lines):
        """Simulates FB60 'Post': assigns a new sequential document number
        and stores a full Display Document record for it."""
        doc_number = str(self._next_doc_number)
        self._next_doc_number += 1

        posting_date = header.get("posting_date") or today_mmddyyyy()
        document_date = header.get("invoice_date") or posting_date
        try:
            year = posting_date.split("/")[-1]
            month = int(posting_date.split("/")[0])
        except (IndexError, ValueError):
            year = str(datetime.date.today().year)
            month = datetime.date.today().month

        doc = {
            "doc_number": doc_number,
            "company_code": header.get("company_code", "5300"),
            "fiscal_year": year,
            "document_date": document_date,
            "posting_date": posting_date,
            "period": str(month),
            "reference": header.get("reference", ""),
            "cross_comp_no": "",
            "currency": header.get("currency", "USD"),
            "texts_exist": bool(header.get("text")),
            "ledger_group": DEFAULT_LEDGER_GROUP,
            "vendor_number": vendor_number,
            "vendor_name": header.get("vendor_name", ""),
            "doc_type": DOC_TYPE_DESC,
            "calculate_tax": header.get("calculate_tax", False),
            "lines": copy.deepcopy(lines),
        }
        self.documents[doc_number] = doc
        self.attachments.setdefault(doc_number, [])
        return doc

    # -------------------------------------------------------------- attachments
    def get_attachments(self, doc_number):
        return self.attachments.setdefault(doc_number, [])

    def add_attachment(self, doc_number, filename):
        title = filename[:-4] if filename.lower().endswith(".pdf") else filename
        entry = {
            "title": title,
            "creator": "Alex Carter",
            "created_on": today_mmddyyyy(),
        }
        self.attachments.setdefault(doc_number, []).append(entry)
        return entry


# =========================================================================
# POPUPS -- modal dialogs shared across screens: vendor search, calendar,
# object services, attachment list, the simulated file picker, and the SAP
# GUI Security dialog.
#
# Each popup is a small tk.Toplevel subclass exposing plain public methods
# (`search()`, `select_result(i)`, `allow()`, ...) that mirror its interactive
# actions one-for-one. That lets a headless verification script drive the
# exact same code path a mouse click would, without needing a live display
# event loop.
# =========================================================================

FILE_PICKER_FILES = [
    "Bluecrest Holdings, LLC 113380 $1442.75.pdf",
    "Meridian Freight Lines, LLC MFL-119284 $2,829.84.pdf",
    "Waste Mgmt WM-2201947 $616.18.pdf",
    "FedEx Invoice 998812 $89.50.pdf",
    "Office Depot Receipt 44712.pdf",
    "IT Services Statement August.pdf",
]


# --------------------------------------------------------------- vendor search
class VendorSearchPopup(tk.Toplevel):
    """'Account or Matchcode for the Next Line Item' vendor search."""

    def __init__(self, app, data, on_select):
        super().__init__(app)
        self.app = app
        self.data = data
        self.on_select = on_select
        self.title("Account or Matchcode for the Next Line Item (1)")
        self.resizable(True, True)
        self.transient(app)
        self.configure(bg=DIALOG_BG)

        tabs_row = tk.Frame(self, bg=GRID_HEADER_BG)
        tabs_row.pack(fill="x")
        tk.Label(
            tabs_row, text="A: Vendors (General)", bg="#FFFFFF", fg=TAB_ACTIVE_FG,
            font=FONT_SMALL, anchor="w", padx=8, pady=3, relief="raised", bd=1, name="tab_label",
        ).pack(side="left")
        for extra in (
            "I: Vendors by Country/Company Code", "K: Vendors by Company Code",
            "L: Vendors by Country",
        ):
            tk.Label(
                tabs_row, text=extra, bg=GRID_HEADER_BG, fg="#666666",
                font=FONT_SMALL, anchor="w", padx=8, pady=3,
            ).pack(side="left")
        tk.Label(
            tabs_row, text="›  ▾", bg=GRID_HEADER_BG, fg="#666666",
            font=FONT_SMALL, padx=8, pady=3,
        ).pack(side="right")

        form = tk.Frame(self, bg=DIALOG_BG, padx=14, pady=10, name="search_form")
        form.pack(fill="x")

        self.vars = {}
        fields = [
            ("search_term", "Search term"),
            ("country", "Country"),
            ("postal_code", "Postal Code"),
            ("city", "City"),
            ("name", "Name"),
            ("supplier", "Supplier"),
            ("deletion_flag", "Deletion Flag"),
        ]
        for i, (key, label) in enumerate(fields):
            tk.Label(form, text=label, bg=DIALOG_BG, font=FONT_NORMAL, width=14, anchor="w").grid(
                row=i, column=0, sticky="w", pady=2
            )
            var = tk.StringVar()
            tk.Entry(form, textvariable=var, width=26, font=FONT_NORMAL, name=f"vs_{key}").grid(
                row=i, column=1, sticky="w", pady=2, padx=(6, 0)
            )
            self.vars[key] = var

        tk.Label(form, text="Maximum No. of Hits", bg=DIALOG_BG, font=FONT_NORMAL, width=16, anchor="w").grid(
            row=len(fields), column=0, sticky="w", pady=(8, 2)
        )
        self.max_hits_var = tk.StringVar(value="2000")
        tk.Entry(form, textvariable=self.max_hits_var, width=10, font=FONT_NORMAL, name="vs_max_hits").grid(
            row=len(fields), column=1, sticky="w", pady=(8, 2), padx=(6, 0)
        )

        def stub(label):
            return lambda: self.app.set_status(f"'{label}' is not implemented in this demo.")

        btn_row = tk.Frame(self, bg=DIALOG_BG, padx=14, pady=6, name="vs_toolbar")
        btn_row.pack(fill="x")
        confirm_btn = tk.Button(
            btn_row, image=enter_icon(), bg=BUTTON_RAISED, relief="raised", bd=1,
            name="vs_start_search_btn", command=self.start_search,
        )
        confirm_btn.pack(side="left", padx=2)
        add_tooltip(confirm_btn, "Confirm")
        cancel_btn = tk.Button(
            btn_row, image=cancel_icon(), bg=BUTTON_RAISED, relief="raised", bd=1,
            name="vs_close_btn", command=self.close,
        )
        cancel_btn.pack(side="left", padx=2)
        add_tooltip(cancel_btn, "Cancel")
        for icon_fn, name, tip in (
            (magnifier_icon, "vs_find_btn", "Find"),
            (find_next_icon, "vs_find_next_btn", "Find next"),
            (star_icon, "vs_favorites_btn", "Favorites"),
            (help_icon, "vs_help_search_btn", "Help search"),
            (printer_icon, "vs_print_btn", "Print"),
            (export_icon, "vs_export_btn", "Export"),
        ):
            b = tk.Button(btn_row, image=icon_fn(), bg=BUTTON_RAISED, relief="raised", bd=1, name=name, command=stub(tip))
            b.pack(side="left", padx=2)
            add_tooltip(b, tip)

        self.entries_found_var = tk.StringVar(value="")
        tk.Label(
            self, textvariable=self.entries_found_var, bg=DIALOG_BG, font=FONT_TINY,
            fg="#666666", anchor="w", padx=14,
        ).pack(fill="x")

        columns = ("search_term", "country", "postal_code", "city", "name1", "supplier", "delf")
        headers = {
            "search_term": "SearchTerm", "country": "Cty", "postal_code": "PostalCo...",
            "city": "City", "name1": "Name 1", "supplier": "Supplier", "delf": "DelF",
        }
        results_frame = tk.Frame(self, bg=DIALOG_BG, padx=14)
        results_frame.pack(fill="both", expand=True, pady=(0, 12))
        style = ttk.Style()
        style_treeview(style, "Sap.Treeview")
        self.results = ttk.Treeview(
            results_frame, columns=columns, show="headings", height=8, style="Sap.Treeview",
            name="vs_results",
        )
        for col in columns:
            self.results.heading(col, text=headers[col])
            self.results.column(col, width=95)
        self.results.pack(fill="both", expand=True)
        self.results.bind("<Double-1>", lambda e: self._on_row_activate())

        self._row_vendors = []
        center_over(self, app)
        self.grab_set()

    def set_field(self, key, value):
        if key in self.vars:
            self.vars[key].set(value)
        elif key == "max_hits":
            self.max_hits_var.set(value)

    def start_search(self):
        kwargs = {k: v.get() for k, v in self.vars.items()}
        matches = self.data.search_vendors(
            name=kwargs.get("name"),
            search_term=kwargs.get("search_term"),
            country=kwargs.get("country"),
            postal_code=kwargs.get("postal_code"),
            city=kwargs.get("city"),
            supplier=kwargs.get("supplier"),
            max_hits=self.max_hits_var.get(),
        )
        self.results.delete(*self.results.get_children())
        self._row_vendors = []
        for v in matches:
            self.results.insert("", "end", values=(
                v["search_term"], v["country"], v["postal_code"], v["city"], v["name1"],
                v["number"], v.get("deletion_flag", ""),
            ))
            self._row_vendors.append(v)
        self.entries_found_var.set(f"{len(matches)} Entries found")
        return matches

    def _on_row_activate(self):
        sel = self.results.selection()
        if not sel:
            return
        index = self.results.index(sel[0])
        self.select_result(index)

    def select_result(self, index):
        if 0 <= index < len(self._row_vendors):
            vendor = self._row_vendors[index]
            callback = self.on_select
            self.destroy()
            callback(vendor)

    def close(self):
        self.destroy()


def open_vendor_search_popup(app, data, on_select):
    return VendorSearchPopup(app, data, on_select)


# ------------------------------------------------------------------ calendar
CAL_SELECTED_OUTLINE = "#7A4FBF"  # violet outline the real SAP calendar boxes the selected day in
CAL_MUTED_FG = "#A8A8A8"          # leading/trailing days from adjacent months


class CalendarPopup(tk.Toplevel):
    """SAP GUI's calendar popup: a continuous, vertically-scrolling list of
    several consecutive months (not a single month with prev/next arrows),
    with a small spinner date-input box at top and the selected day boxed in
    a violet outline. Each month's day grid uses MO/TU/WE/TH/FR/SA/SU column
    headers, with leading/trailing days from the adjacent month shown grayed
    out in the same row as the transition -- exactly what
    calendar.Calendar.monthdatescalendar() already gives us."""

    MONTHS_SPAN = 3

    def __init__(self, app, on_select, initial=None):
        super().__init__(app)
        self.app = app
        self.on_select = on_select

        try:
            dt = _dt.datetime.strptime(initial, "%m/%d/%Y").date() if initial else _dt.date.today()
        except ValueError:
            dt = _dt.date.today()
        self.selected_date = dt
        self._anchor_year, self._anchor_month = dt.year, dt.month

        self.title("Calendar")
        self.resizable(False, False)
        self.transient(app)
        self.configure(bg=DIALOG_BG)

        titlebar = tk.Frame(self, bg=DIALOG_BG)
        titlebar.pack(fill="x")
        tk.Label(titlebar, text="\u2261", bg=DIALOG_BG, fg="#2B6EA5", font=FONT_NORMAL).pack(side="left", padx=(8, 2), pady=2)

        spin_row = tk.Frame(self, bg=DIALOG_BG, padx=8)
        spin_row.pack(fill="x", pady=(4, 6))
        self.date_field_var = tk.StringVar(value=dt.strftime("%m/%d/%Y"))
        tk.Entry(
            spin_row, textvariable=self.date_field_var, font=FONT_NORMAL, width=12, name="cal_date_field",
        ).pack(side="left")
        spin_col = tk.Frame(spin_row, bg=DIALOG_BG)
        spin_col.pack(side="left", padx=(2, 0))
        tk.Button(
            spin_col, text="\u25b2", width=2, font=FONT_TINY, name="cal_spin_up",
            command=lambda: self._spin_day(+1),
        ).pack()
        tk.Button(
            spin_col, text="\u25bc", width=2, font=FONT_TINY, name="cal_spin_down",
            command=lambda: self._spin_day(-1),
        ).pack()

        canvas_frame = tk.Frame(self, bg=DIALOG_BG)
        canvas_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.canvas = tk.Canvas(
            canvas_frame, bg=CONTENT_BG, width=284, height=320,
            highlightthickness=1, highlightbackground="#999999", name="cal_canvas",
        )
        vsb = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._months_frame = tk.Frame(self.canvas, bg=CONTENT_BG)
        self.canvas.create_window((0, 0), window=self._months_frame, anchor="nw")
        self._months_frame.bind(
            "<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Button-4>", lambda e: self.canvas.yview_scroll(-1, "units"))
        self.canvas.bind("<Button-5>", lambda e: self.canvas.yview_scroll(1, "units"))

        btn_row = tk.Frame(self, bg=DIALOG_BG)
        btn_row.pack(pady=(0, 8))
        tk.Button(
            btn_row, text="\u2714", bg=BUTTON_RAISED, relief="raised", bd=1,
            font=FONT_NORMAL, name="cal_ok_btn", command=self._confirm_typed_date,
        ).pack(side="left", padx=6)
        tk.Button(
            btn_row, text="\u2718", bg=BUTTON_RAISED, relief="raised", bd=1,
            font=FONT_NORMAL, name="cal_cancel_btn", command=self.destroy,
        ).pack(side="left", padx=6)

        self._day_cells = {}
        self._month_blocks = []
        self._render()
        center_over(self, app)
        self.grab_set()

    # ------------------------------------------------------------- rendering
    def visible_months(self):
        """(year, month) tuples for every month block currently rendered --
        used by headless verification to confirm the 3+ month scroll span."""
        return list(self._month_blocks)

    def _month_sequence(self):
        months = []
        y, m = self._anchor_year, self._anchor_month - 1
        if m < 1:
            m, y = 12, y - 1
        for _ in range(self.MONTHS_SPAN):
            months.append((y, m))
            m += 1
            if m > 12:
                m, y = 1, y + 1
        return months

    def _render(self):
        for w in self._months_frame.winfo_children():
            w.destroy()
        self._day_cells = {}
        self._month_blocks = self._month_sequence()
        for year, month in self._month_blocks:
            self._render_month_block(year, month)
        self.canvas.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self._scroll_to_selected()

    def _render_month_block(self, year, month):
        block = tk.Frame(self._months_frame, bg=CONTENT_BG)
        block.pack(fill="x")
        tk.Label(
            block, text=f"{_calendar_mod.month_name[month]} {year}", bg=CONTENT_BG,
            fg=HEADER_FG, font=FONT_SMALL_BOLD, anchor="w",
        ).pack(fill="x", padx=6, pady=(8, 2))

        grid = tk.Frame(block, bg=CONTENT_BG)
        grid.pack(padx=6)
        for col, wd in enumerate(["MO", "TU", "WE", "TH", "FR", "SA", "SU"]):
            tk.Label(grid, text=wd, bg=CONTENT_BG, fg="#666666", font=FONT_TINY, width=3).grid(
                row=0, column=col, padx=1, pady=1
            )

        cal = _calendar_mod.Calendar(firstweekday=0)
        for r, week in enumerate(cal.monthdatescalendar(year, month), start=1):
            for c, day_date in enumerate(week):
                self._render_day_cell(grid, day_date, month, r, c)

    def _render_day_cell(self, grid, day_date, block_month, row, col):
        in_month = day_date.month == block_month
        is_selected = day_date == self.selected_date
        cell = tk.Label(
            grid, text=str(day_date.day), bg=CONTENT_BG, width=3, font=FONT_SMALL,
            fg=HEADER_FG if in_month else CAL_MUTED_FG,
            highlightthickness=2 if is_selected else 0,
            highlightbackground=CAL_SELECTED_OUTLINE, highlightcolor=CAL_SELECTED_OUTLINE,
        )
        cell.grid(row=row, column=col, padx=1, pady=1)
        cell.bind("<Button-1>", lambda e, d=day_date: self.select_day(d))
        self._day_cells[(day_date.year, day_date.month, day_date.day)] = cell

    def _scroll_to_selected(self):
        cell = self._day_cells.get((self.selected_date.year, self.selected_date.month, self.selected_date.day))
        if cell is None:
            return
        try:
            self.canvas.update_idletasks()
            bbox = self.canvas.bbox("all")
            if not bbox or (bbox[3] - bbox[1]) <= 0:
                return
            frac = max(0.0, min(1.0, cell.winfo_y() / (bbox[3] - bbox[1])))
            self.canvas.yview_moveto(frac)
        except tk.TclError:
            pass

    def _on_mousewheel(self, event):
        step = -1 if event.delta > 0 else 1
        self.canvas.yview_scroll(step, "units")

    # ------------------------------------------------------------ selection
    def _spin_day(self, delta):
        try:
            dt = _dt.datetime.strptime(self.date_field_var.get(), "%m/%d/%Y").date()
        except ValueError:
            dt = _dt.date.today()
        dt += _dt.timedelta(days=delta)
        self.date_field_var.set(dt.strftime("%m/%d/%Y"))
        self.selected_date = dt
        self._anchor_year, self._anchor_month = dt.year, dt.month
        self._render()

    def _confirm_typed_date(self):
        try:
            dt = _dt.datetime.strptime(self.date_field_var.get(), "%m/%d/%Y").date()
        except ValueError:
            self.app.set_status("Please enter a valid date (MM/DD/YYYY).", error=True)
            return
        self.select_day(dt)

    def select_day(self, day_date):
        date_str = day_date.strftime("%m/%d/%Y")
        callback = self.on_select
        self.destroy()
        callback(date_str)

    def select_date(self, date_str):
        """Programmatic equivalent of clicking a specific day cell, so
        headless verification can pick a date without pixel coordinates."""
        self.select_day(_dt.datetime.strptime(date_str, "%m/%d/%Y").date())

    def locate_date_cell(self, date_str):
        """Opt-in automation bridge helper: returns the real tk.Label day
        cell for an arbitrary date string (MM/DD/YYYY), re-anchoring and
        re-rendering the popup if that date isn't in the currently visible
        3-month span, then scrolling it into view. Returns None if the date
        string itself is unparseable. Never called by normal app code."""
        target = _dt.datetime.strptime(date_str, "%m/%d/%Y").date()
        key = (target.year, target.month, target.day)
        if key not in self._day_cells:
            self._anchor_year, self._anchor_month = target.year, target.month
            self._render()
        cell = self._day_cells.get(key)
        if cell is None:
            return None
        try:
            self.canvas.update_idletasks()
            bbox = self.canvas.bbox("all")
            if bbox and (bbox[3] - bbox[1]) > 0:
                frac = max(0.0, min(1.0, cell.winfo_y() / (bbox[3] - bbox[1])))
                self.canvas.yview_moveto(frac)
            self.update_idletasks()
        except tk.TclError:
            pass
        return cell


def open_calendar_popup(app, on_select, initial=None):
    return CalendarPopup(app, on_select, initial=initial)


# ---------------------------------------------------------- object services
def open_object_services_menu(app, x, y, on_attachment_list):
    """Object-services context menu. Only 'Attachment list' is wired."""

    def stub(label):
        return lambda: app.set_status(f"'{label}' is not implemented in this demo.")

    items = [
        ("Create", None, "", [
            ("Create note", stub("Create note")),
            ("Create external document (URL)", stub("Create external document (URL)")),
            ("Store business document", stub("Store business document")),
        ]),
        ("Attachment list", on_attachment_list),
        ("Private Note", stub("Private Note")),
        ("Send", None, "", [("Email", stub("Send > Email"))]),
        ("Relationships", stub("Relationships")),
        ("Workflow", None, "", [("Start Workflow", stub("Workflow > Start Workflow"))]),
        ("My Objects", None, "", [("Add to My Objects", stub("My Objects > Add"))]),
        ("Help for object services", stub("Help for object services")),
    ]
    menu = CustomMenu(app, app, items)
    menu.popup(x, y)
    return menu


# ------------------------------------------------------------- attachment list
class AttachmentListPopup(tk.Toplevel):
    """'Service: Attachment list'."""

    def __init__(self, app, data, doc_number, status_target=None):
        super().__init__(app)
        self.app = app
        self.data = data
        self.doc_number = doc_number
        self.status_target = status_target or app
        self.title("Service: Attachment list")
        self.resizable(True, True)
        self.transient(app)
        self.configure(bg=DIALOG_BG)
        self.geometry("620x340")

        titlebar = tk.Frame(self, bg=DIALOG_BG)
        titlebar.pack(fill="x")
        tk.Label(titlebar, text="\u2261", bg=DIALOG_BG, fg="#2B6EA5", font=FONT_NORMAL).pack(side="left", padx=(8, 2), pady=2)

        def stub(label):
            return lambda: self.app.set_status(f"'{label}' is not implemented in this demo.")

        toolbar = tk.Frame(self, bg=TOOLBAR_BG, pady=3, name="att_toolbar")
        toolbar.pack(fill="x")
        new_btn = tk.Button(
            toolbar, image=new_icon(), text=" New \u25be", compound="left", bg=TOOLBAR_BG,
            relief="raised", bd=1, font=FONT_SMALL, name="att_new_btn", command=self.open_new_menu,
        )
        new_btn.pack(side="left", padx=4)
        add_tooltip(new_btn, "New")
        self._new_btn = new_btn
        for icon_fn, name, tip in (
            (export_icon, "att_open_btn", "Open"),
            (pencil_icon, "att_edit_btn", "Edit"),
            (trash_icon, "att_delete_btn", "Delete"),
            (refresh_icon, "att_refresh_btn", "Refresh"),
            (magnifier_icon, "att_find_btn", "Find"),
            (find_next_icon, "att_find_next_btn", "Find next"),
            (copy_icon, "att_copy_btn", "Copy"),
            (funnel_icon, "att_filter_btn", "Filter"),
            (printer_icon, "att_print_btn", "Print"),
            (export_icon, "att_export_btn", "Export"),
        ):
            b = tk.Button(toolbar, image=icon_fn(), bg=TOOLBAR_BG, relief="raised", bd=1, name=name, command=stub(tip))
            b.pack(side="left", padx=2)
            add_tooltip(b, tip)

        label_row = tk.Frame(self, bg=DIALOG_BG, padx=8)
        label_row.pack(fill="x", pady=(4, 0))
        tk.Label(
            label_row, text=f"AttachmentFor53001{doc_number}",
            bg=DIALOG_BG, font=FONT_SMALL, anchor="w",
        ).pack(side="left")

        columns = ("icon", "title", "creator", "created_on")
        headers = {"icon": "Icon", "title": "Title", "creator": "CreatrName", "created_on": "Created On"}
        style = ttk.Style()
        style_treeview(style, "Sap.Treeview")
        self.grid = ttk.Treeview(
            self, columns=columns, show="headings", style="Sap.Treeview", name="att_grid", height=8,
        )
        for col in columns:
            self.grid.heading(col, text=headers[col])
            self.grid.column(col, width=40 if col == "icon" else 150)
        self.grid.pack(fill="both", expand=True, padx=8, pady=8)

        btn_row = tk.Frame(self, bg=DIALOG_BG, padx=8, pady=6)
        btn_row.pack(fill="x")
        tk.Button(
            btn_row, text="\u2714", bg=BUTTON_RAISED, relief="raised", bd=1,
            font=FONT_NORMAL, name="att_confirm_btn", command=self.close,
        ).pack(side="right", padx=4)
        tk.Button(
            btn_row, text="\u2718", bg=BUTTON_RAISED, relief="raised", bd=1,
            font=FONT_NORMAL, name="att_cancel_btn", command=self.close,
        ).pack(side="right", padx=4)

        self.refresh()
        center_over(self, app)
        self.grab_set()

    def refresh(self):
        self.grid.delete(*self.grid.get_children())
        for row in self.data.get_attachments(self.doc_number):
            self.grid.insert("", "end", values=("\U0001f4ce", row["title"], row["creator"], row["created_on"]))

    def rows(self):
        return list(self.data.get_attachments(self.doc_number))

    def open_new_menu(self):
        """'New' toolbar dropdown. Only 'Create Attachment' is wired."""

        def stub(label):
            return lambda: self.app.set_status(f"'{label}' is not implemented in this demo.")

        items = [
            ("Create Attachment", self.start_create_attachment),
            ("Create note", stub("Create note")),
            ("Create external document (URL)", stub("Create external document (URL)")),
            ("Store business document", stub("Store business document")),
        ]
        menu = CustomMenu(self, self.app, items)
        x = self._new_btn.winfo_rootx()
        y = self._new_btn.winfo_rooty() + self._new_btn.winfo_height()
        menu.popup(x, y)
        return menu

    def start_create_attachment(self):
        """Opens the simulated file picker; returns the popup so a headless
        caller can drive it directly."""
        return FilePickerPopup(self.app, on_select=self._on_file_selected)

    def _on_file_selected(self, filename):
        SecurityDialog(
            self.app, filename,
            on_allow=lambda: self._on_allow(filename),
            on_deny=self._on_deny,
        )

    def _on_allow(self, filename):
        self.data.add_attachment(self.doc_number, filename)
        self.refresh()
        self.status_target.set_status("The attachment was successfully created", ok=True)

    def _on_deny(self):
        # Cancels attachment creation; attachment list is left unchanged.
        pass

    def close(self):
        self.destroy()


def open_attachment_list_popup(app, data, doc_number, status_target=None):
    return AttachmentListPopup(app, data, doc_number, status_target=status_target)


# ---------------------------------------------------------- document overview
class DocumentOverviewPopup(tk.Toplevel):
    """'Document Overview', reached via FB60's Document > Post menu after a
    successful Simulate. Its own toolbar 'Post' button performs the actual
    posting and returns to FB60."""

    def __init__(self, app, data, vendor_number, header, lines, on_posted):
        super().__init__(app)
        self.app = app
        self.data = data
        self.vendor_number = vendor_number
        self.header = header
        self.lines = lines
        self.on_posted = on_posted
        self.posted_doc = None

        self.title("Document Overview")
        self.resizable(True, True)
        self.transient(app)
        self.configure(bg=DIALOG_BG)
        self.geometry("980x420")

        self._build_toolbar()

        content_pane = tk.Frame(self, bg=DIALOG_BG, name="doc_ov_content_pane")
        content_pane.pack(fill="both", expand=True)
        for anchor, corner in (("nw", (0, 0)), ("se", (1.0, 1.0))):
            tk.Label(
                content_pane, image=resize_grip_icon(), bg=DIALOG_BG,
            ).place(relx=corner[0], rely=corner[1], anchor=anchor)

        head = tk.Frame(content_pane, bg=DIALOG_BG, padx=14, pady=10, name="doc_ov_header")
        head.pack(fill="x")

        try:
            fy = header["posting_date"].split("/")[-1]
        except (KeyError, IndexError):
            fy = ""
        try:
            period = str(int(header["posting_date"].split("/")[0]))
        except (KeyError, ValueError, IndexError):
            period = ""

        self.doc_number_var = tk.StringVar(value="(not yet assigned)")
        fields = [
            ("Doc.Type", f"{DOC_TYPE_CODE} ( Vendor Invoice ) Normal document"),
            ("Doc. Number", None),
            ("Company Code", header.get("company_code", "")),
            ("Doc. Date", header.get("invoice_date", "")),
            ("Posting Date", header.get("posting_date", "")),
            ("Fiscal Year", fy),
            ("Period", period),
            ("Calculate Tax", "Yes" if header.get("calculate_tax") else "No"),
            ("Tax Report Date", header.get("posting_date", "")),
            ("Ref.Doc.", header.get("reference", "")),
            ("Doc. Currency", header.get("currency", "USD")),
        ]
        for i, (label, value) in enumerate(fields):
            r, c = divmod(i, 3)
            cell = tk.Frame(head, bg=DIALOG_BG)
            cell.grid(row=r, column=c, sticky="w", padx=8, pady=3)
            tk.Label(cell, text=f"{label}:", bg=DIALOG_BG, font=FONT_SMALL, width=13, anchor="w").pack(side="left")
            if value is None:
                tk.Label(cell, textvariable=self.doc_number_var, bg=DIALOG_BG, font=FONT_SMALL, name="doc_ov_docnum").pack(side="left")
            else:
                tk.Label(cell, text=value, bg=DIALOG_BG, font=FONT_SMALL).pack(side="left")

        columns = ("item", "cocd", "pk", "account", "description", "segment",
                   "crcy", "amount", "lcurr", "amount_lc", "tx", "cost_ctr", "order", "profit_ctr", "assignm")
        headers = {
            "item": "Itm", "cocd": "CoCd", "pk": "PK", "account": "Account",
            "description": "Account short text", "segment": "Segment", "crcy": "Crcy",
            "amount": "Amount", "lcurr": "LCurr", "amount_lc": "Amount in LC", "tx": "Tx",
            "cost_ctr": "Cost Ctr", "order": "Order", "profit_ctr": "Profit Ctr", "assignm": "Assignm",
        }
        style = ttk.Style()
        style_treeview(style, "Sap.Treeview")
        self.table = ttk.Treeview(
            content_pane, columns=columns, show="headings", style="Sap.Treeview", name="doc_ov_table", height=4,
        )
        for col in columns:
            self.table.heading(col, text=headers[col])
            width = 60 if col in ("item", "cocd", "pk", "tx") else 90
            self.table.column(col, width=width)
        self.table.pack(fill="both", expand=True, padx=14, pady=(6, 0))

        total = Decimal("0")
        company_code = header.get("company_code", "")
        for i, line in enumerate(lines, start=1):
            try:
                amt = Decimal(line["amount"].rstrip("-")) * (-1 if line["amount"].endswith("-") else 1)
            except (InvalidOperation, AttributeError):
                amt = Decimal("0")
            total += amt
            self.table.insert("", "end", values=(
                i, company_code, line.get("posting_key", ""), line["account"], line["description"],
                line.get("segment", ""), line.get("currency", "USD"), line["amount"],
                line.get("currency", "USD"), line.get("lcurr_amount", ""), line.get("tax_code", ""),
                line.get("cost_center", ""), line.get("order", ""), line.get("profit_center", ""),
                line.get("assignment", ""),
            ))
        self.table.insert("", "end", values=(
            "*", "", "", "", "", "", "USD", fmt_amount(total), "USD", fmt_amount(total), "", "", "", "", "",
        ))

        center_over(self, app)
        self.grab_set()

    def _build_toolbar(self):
        def stub(label):
            return lambda: self.app.set_status(f"'{label}' is not implemented in this demo.")

        toolbar = tk.Frame(self, bg=TOOLBAR_BG, pady=3, name="doc_ov_toolbar")
        toolbar.pack(fill="x")

        check_btn = tk.Button(toolbar, image=enter_icon(), bg=TOOLBAR_BG, relief="raised", bd=1, name="doc_ov_check_btn")
        check_btn.pack(side="left", padx=(4, 0), pady=3)
        add_tooltip(check_btn, "Enter")
        tk.Button(
            toolbar, text="▾", width=1, bg=TOOLBAR_BG, relief="raised", bd=1, font=FONT_TINY,
            name="doc_ov_check_dropdown",
        ).pack(side="left", padx=(0, 4), pady=3)

        post_btn = tk.Button(
            toolbar, image=post_icon(), text=" Post", compound="left", bg=TOOLBAR_BG,
            relief="raised", bd=1, font=FONT_SMALL, name="doc_ov_post_btn", command=self.do_post,
        )
        post_btn.pack(side="left", padx=2, pady=3)
        add_tooltip(post_btn, "Post")

        find_btn = tk.Button(
            toolbar, image=magnifier_icon(), bg=TOOLBAR_BG, relief="raised", bd=1,
            name="doc_ov_find_btn", command=stub("Find"),
        )
        find_btn.pack(side="left", padx=2, pady=3)
        add_tooltip(find_btn, "Find")

        for image, name, label in (
            (reset_icon(), "doc_ov_reset_btn", "Reset"),
            (info_icon(), "doc_ov_taxes_btn", "Taxes"),
            (save_icon(), "doc_ov_park_btn", "Park"),
            # Real SAP GUI renders 'Complete' with the same floppy-disk glyph
            # as 'Park', not a checkmark.
            (save_icon(), "doc_ov_complete_btn", "Complete"),
        ):
            b = toolbar_text_button(toolbar, image, label, name, label, stub(label))
            b.pack(side="left", padx=2, pady=3)

        for image, name, tip in (
            (funnel_icon(), "doc_ov_filter_btn", "Filter"),
            (sort_asc_icon(), "doc_ov_sort_asc_btn", "Sort ascending"),
            (sort_desc_icon(), "doc_ov_sort_desc_btn", "Sort descending"),
            (copy_icon(), "doc_ov_copy_btn", "Copy"),
        ):
            b = tk.Button(toolbar, image=image, bg=TOOLBAR_BG, relief="raised", bd=1, name=name, command=stub(tip))
            b.pack(side="left", padx=1, pady=3)
            add_tooltip(b, tip)

        # Real SAP GUI renders 'Choose' with the same two-overlapping-pages
        # glyph as 'Copy', not a folder.
        choose_btn = toolbar_text_button(toolbar, copy_icon(), "Choose", "doc_ov_choose_btn", "Choose", stub("Choose"))
        choose_btn.pack(side="left", padx=2, pady=3)
        save_btn = toolbar_text_button(toolbar, save_icon(), "Save", "doc_ov_save_btn", "Save", stub("Save"))
        save_btn.pack(side="left", padx=2, pady=3)

        for text, name, tip in (("Σ", "doc_ov_sum_btn", "Total"), ("½", "doc_ov_half_btn", "Currency units")):
            b = tk.Button(toolbar, text=text, bg=TOOLBAR_BG, relief="raised", bd=1, font=FONT_SMALL, name=name, command=stub(tip))
            b.pack(side="left", padx=1, pady=3)
            add_tooltip(b, tip)

        for image, name, tip in (
            (envelope_icon(), "doc_ov_mail_btn", "Mail"),
            (export_icon(), "doc_ov_extra1_btn", "Export"),
            (pencil_icon(), "doc_ov_extra2_btn", "Edit"),
            (grid_small_icon(), "doc_ov_grid_btn", "Display"),
        ):
            b = tk.Button(toolbar, image=image, bg=TOOLBAR_BG, relief="raised", bd=1, name=name, command=stub(tip))
            b.pack(side="left", padx=1, pady=3)
            add_tooltip(b, tip)

        abc_btn = tk.Button(
            toolbar, text="ABC", bg=TOOLBAR_BG, relief="raised", bd=1, font=FONT_TINY,
            name="doc_ov_abc_btn", command=stub("ABC"),
        )
        abc_btn.pack(side="left", padx=2, pady=3)

        info_btn = tk.Button(
            toolbar, image=info_icon(), bg=TOOLBAR_BG, relief="raised", bd=1,
            name="doc_ov_info_btn", command=stub("Info"),
        )
        info_btn.pack(side="left", padx=2, pady=3)
        add_tooltip(info_btn, "Info")

        cancel_btn = tk.Button(
            toolbar, text="Cancel", bg=TOOLBAR_BG, relief="raised", bd=1, font=FONT_SMALL,
            name="doc_ov_cancel_btn", command=self.destroy,
        )
        cancel_btn.pack(side="left", padx=4, pady=3)

        right_cluster = tk.Frame(toolbar, bg=TOOLBAR_BG, name="doc_ov_right_cluster")
        right_cluster.pack(side="right", padx=4, pady=3)
        for image, name, tip in (
            (new_session_icon(), "doc_ov_new_btn", "Create new session"),
            (forward_icon(), "doc_ov_forward_btn", "Go to session"),
            (options_icon(), "doc_ov_settings_btn", "Settings"),
        ):
            b = tk.Button(right_cluster, image=image, bg=TOOLBAR_BG, relief="raised", bd=1, name=name, command=stub(tip))
            b.pack(side="left", padx=1)
            add_tooltip(b, tip)

    def do_post(self):
        doc = self.data.post_invoice(self.vendor_number, self.header, self.lines)
        self.posted_doc = doc
        self.doc_number_var.set(doc["doc_number"])
        self.on_posted(doc)
        self.app.set_status(
            f"\u2714 Document {doc['doc_number']} was posted in company code {doc['company_code']}", ok=True,
        )
        self.destroy()


# ------------------------------------------------------------------ file picker
class FilePickerPopup(tk.Toplevel):
    """Simulated OS file picker."""

    def __init__(self, app, on_select, files=None):
        super().__init__(app)
        self.app = app
        self.on_select = on_select
        self.files = files or list(FILE_PICKER_FILES)
        self.title("Choose File to Attach")
        self.resizable(False, False)
        self.transient(app)
        self.configure(bg=DIALOG_BG)

        self.listbox = tk.Listbox(
            self, height=len(self.files), width=52, font=FONT_NORMAL,
            bg=CONTENT_BG, fg=HEADER_FG, name="file_picker_listbox",
        )
        for f in self.files:
            self.listbox.insert("end", f)
        self.listbox.pack(padx=10, pady=10)
        self.listbox.bind("<Double-1>", lambda e: self._on_activate())
        self.listbox.bind("<Return>", lambda e: self._on_activate())

        center_over(self, app)
        self.grab_set()

    def _on_activate(self):
        sel = self.listbox.curselection()
        if sel:
            self.select_file_at(sel[0])

    def select_file_at(self, index):
        filename = self.files[index]
        self.select_file(filename)

    def select_file(self, filename):
        callback = self.on_select
        self.destroy()
        callback(filename)


# ------------------------------------------------------------- security dialog
class SecurityDialog(tk.Toplevel):
    """'SAP GUI Security' file-access prompt."""

    def __init__(self, app, filename, on_allow, on_deny):
        super().__init__(app)
        self.app = app
        self.filename = filename
        self.on_allow = on_allow
        self.on_deny = on_deny
        self.title("SAP GUI Security")
        self.resizable(False, False)
        self.transient(app)
        self.configure(bg=DIALOG_BG)

        path = f"C:\\Users\\demo\\Documents\\Invoices\\{filename}"
        body = tk.Frame(self, bg=DIALOG_BG, padx=18, pady=16)
        body.pack(fill="both", expand=True)

        tk.Label(
            body, text="The system is attempting to access the following file:",
            bg=DIALOG_BG, font=FONT_NORMAL, justify="left", wraplength=420,
        ).pack(anchor="w")
        tk.Label(
            body, text=path, bg=DIALOG_BG, font=FONT_MONO, justify="left",
            wraplength=420, name="sec_path_label",
        ).pack(anchor="w", pady=(4, 8))
        tk.Label(
            body, text="Do you want to grant access to this file?",
            bg=DIALOG_BG, font=FONT_NORMAL, justify="left", wraplength=420,
        ).pack(anchor="w")

        self.remember_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            body, text="Remember My Decision", variable=self.remember_var, bg=DIALOG_BG,
            font=FONT_SMALL, name="sec_remember_check",
        ).pack(anchor="w", pady=(8, 8))

        # Real dialog's left-to-right button order is Allow, Deny, Help.
        btn_row = tk.Frame(body, bg=DIALOG_BG)
        btn_row.pack(anchor="w")
        tk.Button(
            btn_row, text="Allow", bg=BUTTON_RAISED, relief="raised", bd=1,
            font=FONT_NORMAL, name="sec_allow_btn", command=self.allow, width=10,
        ).pack(side="left", padx=4)
        tk.Button(
            btn_row, text="Deny", bg=BUTTON_RAISED, relief="raised", bd=1,
            font=FONT_NORMAL, name="sec_deny_btn", command=self.deny, width=10,
        ).pack(side="left", padx=4)
        tk.Button(
            btn_row, text="Help", bg=BUTTON_RAISED, relief="raised", bd=1,
            font=FONT_NORMAL, name="sec_help_btn", width=10,
            command=lambda: self.app.set_status("Help is not available in this demo."),
        ).pack(side="left", padx=(24, 4))

        center_over(self, app)
        self.grab_set()

    def allow(self):
        callback = self.on_allow
        self.destroy()
        callback()

    def deny(self):
        callback = self.on_deny
        self.destroy()
        if callback:
            callback()


# =========================================================================
# SCREENS
# =========================================================================

# --------------------------------------------------------------------- FB60
COLUMNS = [
    "chk", "gl_acct", "short_text", "dc", "amount_doc", "amount_loc",
    "tax_type", "tax_jur", "withhold", "assignment", "cost_center", "order", "profit_ctr",
]
HEADERS = {
    "chk": "S.", "gl_acct": "G/L acct", "short_text": "Short Text", "dc": "D/C",
    "amount_doc": "Amount in doc.curr.", "amount_loc": "Loc.curr.amount",
    "tax_type": "T..", "tax_jur": "Tax jurisdictn code", "withhold": "W",
    "assignment": "Assignment", "cost_center": "Cost center", "order": "Order", "profit_ctr": "Profit Ctr.",
}
WIDTHS = {
    "chk": 28, "gl_acct": 80, "short_text": 110, "dc": 60,
    "amount_doc": 110, "amount_loc": 100, "tax_type": 28, "tax_jur": 110, "withhold": 28,
    "assignment": 90, "cost_center": 90, "order": 70, "profit_ctr": 80,
}


def _blank_line():
    return {
        "chk": False, "gl_acct": "", "short_text": "", "dc": "Debit",
        "amount_doc": "", "amount_loc": "", "tax_type": False, "tax_jur": "", "withhold": False,
        "assignment": "", "cost_center": "", "order": "", "profit_ctr": "",
        "validated": False,
    }


class FB60Screen(tk.Frame):
    TCODE = "FB60"

    def __init__(self, parent, controller, data, context=None):
        super().__init__(parent, bg=CONTENT_BG, name="fb60_screen")
        self.controller = controller
        self.data = data
        self.company_code = DEFAULT_COMPANY_CODE
        self.vendor = None
        self.balanced = False
        self.last_doc_number = None
        self.lines = [_blank_line() for _ in range(10)]

        self._build_menu_strip()
        self._build_toolbar()

        body = tk.Frame(self, bg=CONTENT_BG, name="fb60_body")
        body.pack(fill="both", expand=True, padx=8, pady=(4, 8))

        self._build_header_row(body)

        content_row = tk.Frame(body, bg=CONTENT_BG, name="fb60_content_row")
        content_row.pack(fill="x")
        self._left_col = tk.Frame(content_row, bg=CONTENT_BG)
        self._left_col.pack(side="left", fill="y", anchor="n")

        self._build_basic_data(self._left_col)
        self._build_address_panel(content_row)
        self._build_grid(body)

        self._update_balance_display()

    # --------------------------------------------------------------- chrome
    def _build_menu_strip(self):
        def stub(label):
            return lambda: self.controller.set_status(f"'{label}' is not implemented in this demo.")

        document_items = [
            ("Change", stub("Change"), "Ctrl+F1"),
            ("Display", self.on_document_display, "Ctrl+F2"),
            (None, None),
            ("Post", self.on_document_post, "Ctrl+S"),
            ("Save as Completed", stub("Save as Completed"), "Ctrl+Shift+F6"),
            ("Park", stub("Park"), "F8"),
            ("Hold", stub("Hold"), "F5"),
            (None, None),
            ("Simulate", self.on_simulate, "F9"),
            ("Simulate General Ledger", stub("Simulate General Ledger"), "Ctrl+F12"),
            (None, None),
            ("Exit", self.on_exit, "Shift+F3"),
        ]
        defs = [("Document", document_items)]
        for label in ("Edit", "Goto", "Extras", "Settings", "Environment", "System", "Help"):
            defs.append((label, [("(demo placeholder)", stub(label))]))
        strip = build_menu_strip(self, defs)
        strip.pack(fill="x", side="top")
        self.menu_strip = strip  # exposed for the (opt-in) automation bridge

    def _build_toolbar(self):
        bar = tk.Frame(self, bg=TOOLBAR_BG, name="fb60_toolbar")
        bar.pack(fill="x", side="top")

        def stub(label):
            return lambda: self.controller.set_status(f"'{label}' is not implemented in this demo.")

        check_btn = tk.Button(
            bar, image=enter_icon(), bg=TOOLBAR_BG, relief="raised", bd=1, name="btn_enter",
            command=lambda: self.controller.set_status(""),
        )
        check_btn.pack(side="left", padx=2, pady=3)
        add_tooltip(check_btn, "Enter")

        cmd_field = ttk.Combobox(bar, width=20, font=FONT_SMALL, style="Sap.TCombobox", name="fb60_cmd_field")
        cmd_field.pack(side="left", padx=(0, 4), pady=3)

        save_btn = tk.Button(bar, image=save_icon(), bg=TOOLBAR_BG, relief="raised", bd=1, name="btn_save",
                              command=lambda: self.controller.set_status("Saved (demo).", ok=True))
        save_btn.pack(side="left", padx=2, pady=3)
        add_tooltip(save_btn, "Save")

        text_buttons = [
            (tree_toggle_icon(), "btn_tree_on", "Tree on", stub("Tree on")),
            (folder_icon(), "btn_company_code", "Company Code",
             lambda: self.controller.set_status(f"Company code {self.company_code} is already set for this document.")),
            (hold_icon(), "btn_hold", "Hold", stub("Hold")),
            (simulate_icon(), "btn_simulate", "Simulate", self.on_simulate),
            (park_icon(), "btn_park", "Park", stub("Park")),
            (options_icon(), "btn_editing_options", "Editing Options", stub("Editing Options")),
        ]
        for image, name, label, cmd in text_buttons:
            b = toolbar_text_button(bar, image, label, name, label, cmd)
            b.pack(side="left", padx=2, pady=3)

        cancel_btn = tk.Button(
            bar, text="Cancel", bg=TOOLBAR_BG, relief="raised", bd=1, font=FONT_SMALL,
            name="btn_cancel", command=self.on_exit,
        )
        cancel_btn.pack(side="left", padx=2, pady=3)

        # Further-right cluster of plain (icon-only) buttons, ending in the
        # text+icon Exit button -- packed together in one right-anchored
        # frame so left-to-right pack order reads naturally.
        right_cluster = tk.Frame(bar, bg=TOOLBAR_BG, name="fb60_right_cluster")
        right_cluster.pack(side="right", padx=2, pady=3)
        # Real SAP GUI shows [new-session/star] [go-to-session/arrow]
        # [settings/gear] here -- not a yellow folder.
        for image, name, tooltip in (
            (new_session_icon(), "btn_new_session", "Create new session"),
            (forward_icon(), "btn_forward", "Go to session"),
            (options_icon(), "btn_settings", "Settings"),
        ):
            b = tk.Button(right_cluster, image=image, bg=TOOLBAR_BG, relief="raised", bd=1, name=name)
            b.pack(side="left", padx=1)
            add_tooltip(b, tooltip)

        exit_btn = tk.Button(
            right_cluster, image=exit_icon(), text=" Exit", compound="left", bg=TOOLBAR_BG, relief="raised", bd=1,
            font=FONT_SMALL, name="btn_exit", command=self.on_exit,
        )
        exit_btn.pack(side="left", padx=(6, 0))

    # --------------------------------------------------------------- header
    def _build_header_row(self, parent):
        row = tk.Frame(parent, bg=CONTENT_BG, name="header_row")
        row.pack(fill="x", pady=(2, 6))

        tk.Label(row, text="Transactn", bg=CONTENT_BG, font=FONT_NORMAL).pack(side="left")
        self.transactn_var = tk.StringVar(value="Invoice")
        ttk.Combobox(
            row, textvariable=self.transactn_var, values=["Invoice", "Credit Memo"], width=14,
            state="readonly", font=FONT_NORMAL, name="transactn_field", style="Sap.TCombobox",
        ).pack(side="left", padx=(6, 40))

        tk.Label(row, text="Bal.", bg=CONTENT_BG, font=FONT_NORMAL).pack(side="left")
        self.balance_var = tk.StringVar(value="0.00")
        tk.Entry(
            row, textvariable=self.balance_var, width=14, font=FONT_NORMAL, state="readonly",
            name="balance_field", readonlybackground=READONLY_BG,
        ).pack(side="left", padx=(6, 6))
        self.balance_dot_label = tk.Label(row, image=balance_indicator_icon(False), bg=CONTENT_BG, name="balance_dot")
        self.balance_dot_label.pack(side="left")

    def _build_address_panel(self, parent):
        self.address_panel = tk.Frame(parent, bg=CONTENT_BG, name="address_panel", padx=20)
        tk.Label(self.address_panel, text="Vendor", bg=CONTENT_BG, font=FONT_SMALL, fg="#555555").pack(anchor="w")
        box = tk.LabelFrame(self.address_panel, text="Address", bg=CONTENT_BG, font=FONT_SMALL)
        box.pack(fill="both", expand=True, pady=(2, 0))

        self._addr_vars = {
            key: tk.StringVar() for key in (
                "name", "street", "city_state_zip", "phone", "bank_account", "bank_number", "swift",
            )
        }
        tk.Label(box, textvariable=self._addr_vars["name"], bg=CONTENT_BG, font=FONT_NORMAL, anchor="w",
                 name="addr_name").pack(anchor="w", padx=10, pady=(6, 0))
        tk.Label(box, textvariable=self._addr_vars["street"], bg=CONTENT_BG, font=FONT_SMALL, anchor="w",
                 name="addr_street").pack(anchor="w", padx=10)
        tk.Label(box, textvariable=self._addr_vars["city_state_zip"], bg=CONTENT_BG, font=FONT_SMALL, anchor="w",
                 name="addr_city_state_zip").pack(anchor="w", padx=10, pady=(0, 6))

        phone_row = tk.Frame(box, bg=CONTENT_BG)
        phone_row.pack(fill="x", padx=10, pady=(0, 6))
        tk.Label(phone_row, image=phone_icon(), bg=CONTENT_BG).pack(side="left")
        tk.Label(phone_row, textvariable=self._addr_vars["phone"], bg=CONTENT_BG, font=FONT_SMALL,
                 name="addr_phone").pack(side="left", padx=(4, 0))
        tk.Label(phone_row, image=sync_icon(), bg=CONTENT_BG).pack(side="right")

        for key, label in (("bank_account", "Bank account"), ("bank_number", "Bank Number"), ("swift", "SWIFT")):
            r = tk.Frame(box, bg=CONTENT_BG)
            r.pack(fill="x", padx=10)
            tk.Label(r, text=label, bg=CONTENT_BG, font=FONT_SMALL, width=13, anchor="w").pack(side="left")
            tk.Label(r, textvariable=self._addr_vars[key], bg=CONTENT_BG, font=FONT_SMALL, anchor="w",
                     name=f"addr_{key}").pack(side="left")

        ols_row = tk.Frame(box, bg=CONTENT_BG)
        ols_row.pack(fill="x", padx=10, pady=6)
        tk.Button(
            ols_row, image=grid_small_icon(), text=" OIs", compound="left", bg=BUTTON_RAISED,
            relief="raised", bd=1, font=FONT_SMALL,
            command=lambda: self.controller.set_status("'OIs' is not implemented in this demo."),
        ).pack(side="right")

        # Hidden until a vendor is resolved.
        self.address_panel.pack_forget()

    def _build_basic_data(self, parent):
        tabs = tk.Frame(parent, bg=GRID_HEADER_BG, name="tab_strip")
        tabs.pack(fill="x")
        self._tab_strip = tabs
        tk.Label(
            tabs, text="Basic data", bg=CONTENT_BG, fg=TAB_ACTIVE_FG, font=FONT_SMALL,
            padx=10, pady=3, relief="raised", bd=1,
        ).pack(side="left")
        for extra in ("Payment", "Details", "Tax", "Withholding tax"):
            tk.Label(tabs, text=extra, bg=GRID_HEADER_BG, fg="#555555", font=FONT_SMALL, padx=10, pady=3).pack(side="left")
        tk.Label(tabs, text="\u203a  \u25be", bg=GRID_HEADER_BG, fg="#555555", font=FONT_SMALL, padx=10, pady=3).pack(side="right")

        form = tk.Frame(parent, bg=CONTENT_BG, name="basic_data_form", pady=8)
        form.pack(fill="x")

        def field_row(label_text, required=False):
            row = tk.Frame(form, bg=CONTENT_BG)
            row.pack(fill="x", pady=3)
            fg = REQUIRED_LABEL_FG if required else HEADER_FG
            tk.Label(row, text=label_text, bg=CONTENT_BG, fg=fg, font=FONT_NORMAL, width=14, anchor="w").pack(side="left")
            return row

        # ---- Supplier | SGL Ind
        row = field_row("Supplier", required=True)
        self.supplier_var = tk.StringVar()
        supplier_entry = tk.Entry(row, textvariable=self.supplier_var, width=18, font=FONT_NORMAL, name="supplier_field")
        supplier_entry.pack(side="left")
        supplier_entry.bind("<Return>", self._on_supplier_entry_commit)
        supplier_entry.bind("<FocusOut>", self._on_supplier_entry_commit)
        search_btn = tk.Button(
            row, image=magnifier_icon(), bg=BUTTON_RAISED, relief="raised", bd=1,
            name="supplier_search_btn", command=self.open_vendor_search,
        )
        search_btn.pack(side="left", padx=(2, 0))
        add_tooltip(search_btn, "Search")
        tk.Label(row, text="SGL Ind", bg=CONTENT_BG, font=FONT_NORMAL, width=10, anchor="w").pack(side="left", padx=(24, 0))
        self.sgl_ind_var = tk.BooleanVar(value=False)
        tk.Checkbutton(row, variable=self.sgl_ind_var, bg=CONTENT_BG, name="sgl_ind_check").pack(side="left")

        # ---- Invoice date | Reference
        row = field_row("Invoice date", required=True)
        self.invoice_date_var = tk.StringVar()
        invoice_entry = tk.Entry(row, textvariable=self.invoice_date_var, width=18, font=FONT_NORMAL, name="invoice_date_field")
        invoice_entry.pack(side="left")
        invoice_entry.bind("<Button-1>", lambda e: self.open_calendar())
        cal_btn = tk.Button(
            row, image=calendar_icon(), bg=BUTTON_RAISED, relief="raised", bd=1,
            name="invoice_date_cal_btn", command=self.open_calendar,
        )
        cal_btn.pack(side="left", padx=(2, 0))
        tk.Label(row, text="Reference", bg=CONTENT_BG, fg=REQUIRED_LABEL_FG, font=FONT_NORMAL, width=10, anchor="w").pack(side="left", padx=(24, 0))
        self.reference_var = tk.StringVar()
        tk.Entry(row, textvariable=self.reference_var, width=18, font=FONT_NORMAL, name="reference_field").pack(side="left")

        # ---- Posting Date
        row = field_row("Posting Date")
        self.posting_date_var = tk.StringVar(value=today_mmddyyyy())
        tk.Entry(row, textvariable=self.posting_date_var, width=18, font=FONT_NORMAL, name="posting_date_field").pack(side="left")

        # ---- Amount | USD | Calculate tax
        row = field_row("Amount", required=True)
        self.amount_var = tk.StringVar()
        tk.Entry(row, textvariable=self.amount_var, width=18, font=FONT_NORMAL, name="amount_field").pack(side="left")
        self.currency_var = tk.StringVar(value="USD")
        tk.Entry(row, textvariable=self.currency_var, width=5, font=FONT_NORMAL, name="currency_field").pack(side="left", padx=(4, 12))
        self.calc_tax_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            row, text="Calculate tax", variable=self.calc_tax_var, bg=CONTENT_BG, font=FONT_NORMAL,
            name="calc_tax_check",
        ).pack(side="left")
        self.amount_var.trace_add("write", self._on_amount_changed)

        # ---- Tax Amount | Tax code
        row = field_row("Tax Amount")
        self.tax_amount_var = tk.StringVar()
        tk.Entry(row, textvariable=self.tax_amount_var, width=14, font=FONT_NORMAL, name="tax_amount_field").pack(side="left")
        self.tax_code_var = tk.StringVar()
        ttk.Combobox(
            row, textvariable=self.tax_code_var, values=self.data.tax_codes, width=32,
            font=FONT_SMALL, name="tax_code_field", style="Sap.TCombobox",
        ).pack(side="left", padx=(10, 0))

        # ---- Text
        row = field_row("Text")
        self.text_var = tk.StringVar()
        tk.Entry(row, textvariable=self.text_var, width=50, font=FONT_NORMAL, name="text_field").pack(side="left")

        # ---- Paymt Terms / Baseline Date (only shown once a vendor resolves)
        self._vendor_extra_frame = tk.Frame(form, bg=CONTENT_BG)
        pt_row = tk.Frame(self._vendor_extra_frame, bg=CONTENT_BG)
        pt_row.pack(fill="x", pady=3)
        tk.Label(pt_row, text="Paymt Terms", bg=CONTENT_BG, font=FONT_NORMAL, width=14, anchor="w").pack(side="left")
        self.paymt_terms_var = tk.StringVar()
        tk.Label(pt_row, textvariable=self.paymt_terms_var, bg=CONTENT_BG, font=FONT_NORMAL, anchor="w").pack(side="left")
        bd_row = tk.Frame(self._vendor_extra_frame, bg=CONTENT_BG)
        bd_row.pack(fill="x", pady=3)
        tk.Label(bd_row, text="Baseline Date", bg=CONTENT_BG, font=FONT_NORMAL, width=14, anchor="w").pack(side="left")
        self.baseline_date_var = tk.StringVar()
        tk.Label(bd_row, textvariable=self.baseline_date_var, bg=CONTENT_BG, font=FONT_NORMAL, anchor="w").pack(side="left")
        # not packed until a vendor is resolved

        # ---- Company Code (always visible, readonly)
        row = field_row("Company Code")
        self._company_code_row = row
        self.company_code_var = tk.StringVar()
        tk.Label(row, textvariable=self.company_code_var, bg=CONTENT_BG, font=FONT_NORMAL, anchor="w").pack(side="left")

        # initialize company code display
        cc = self.data.get_company_code(self.company_code)
        if cc:
            self.company_code_var.set(f"{cc['code']} {cc['name']}, {cc['city']}")

    def _build_grid(self, parent):
        grid_frame = tk.Frame(parent, bg=CONTENT_BG, name="line_item_frame", pady=8)
        grid_frame.pack(fill="both", expand=True)

        item_count_row = tk.Frame(grid_frame, bg=CONTENT_BG)
        item_count_row.pack(fill="x")
        self.item_count_var = tk.StringVar(value="0 Items ( No entry variant selected )")
        tk.Label(item_count_row, textvariable=self.item_count_var, bg=CONTENT_BG, font=FONT_SMALL, fg="#555555").pack(anchor="w")

        style = ttk.Style()
        style_treeview(style, "Sap.Treeview")

        self.tree = ttk.Treeview(
            grid_frame, columns=COLUMNS, show="tree headings", style="Sap.Treeview", height=10,
            name="line_item_grid",
        )
        self.tree.heading("#0", text="")
        self.tree.column("#0", width=26, anchor="center", stretch=False)
        for col in COLUMNS:
            self.tree.heading(col, text=HEADERS[col])
            self.tree.column(col, width=WIDTHS[col], anchor="w")
        vsb = ttk.Scrollbar(grid_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        for i in range(10):
            self.tree.insert(
                "", "end", iid=str(i), image=checkbox_empty_icon(), values=self._row_values(i),
            )

        self.tree.bind("<Double-1>", self._on_grid_double_click)

    # ---------------------------------------------------------------- vendor
    def open_vendor_search(self):
        return open_vendor_search_popup(self.controller, self.data, on_select=self._apply_vendor)

    def _on_supplier_entry_commit(self, event=None):
        text = self.supplier_var.get().strip()
        if not text or (self.vendor and self.vendor["number"] == text):
            return
        vendor = self.data.get_vendor(text)
        if vendor:
            self._apply_vendor(vendor)
        else:
            self.controller.set_status(f"No supplier found for {text}.", error=True)

    def _apply_vendor(self, vendor):
        self.vendor = vendor
        self.supplier_var.set(vendor["number"])
        self._addr_vars["name"].set(vendor["name1"])
        self._addr_vars["street"].set(vendor["street"])
        self._addr_vars["city_state_zip"].set(f"{vendor['city']}, {vendor['state']} {vendor['postal_code']}")
        self._addr_vars["phone"].set(vendor["phone"])
        self._addr_vars["bank_account"].set(vendor["bank_account"])
        self._addr_vars["bank_number"].set(vendor["bank_number"])
        self._addr_vars["swift"].set(vendor["swift"])
        self.address_panel.pack(side="left", fill="y", anchor="n", padx=(20, 0))
        self.paymt_terms_var.set(vendor["payment_terms"])
        self.baseline_date_var.set(self.invoice_date_var.get() or self.posting_date_var.get())
        self._vendor_extra_frame.pack(fill="x", before=self._company_code_row)
        self.balanced = False
        if self.lines[0]["validated"] or self.lines[1]["validated"]:
            self.lines[0]["validated"] = False
            self.lines[1]["validated"] = False
            self._refresh_grid_row(0)
            self._refresh_grid_row(1)
        self._update_balance_display()
        self.controller.set_status(f"Supplier {vendor['number']} - {vendor['name1']} selected.", ok=True)

    # --------------------------------------------------------------- calendar
    def open_calendar(self):
        return open_calendar_popup(self.controller, on_select=self._on_invoice_date_selected, initial=self.invoice_date_var.get())

    def _on_invoice_date_selected(self, date_str):
        self.invoice_date_var.set(date_str)
        if self.vendor:
            self.baseline_date_var.set(date_str)

    # ----------------------------------------------------------------- grid
    def _row_values(self, row):
        line = self.lines[row]
        chk_display = "\u2611" if line["chk"] else "\u2610"
        tax_type_display = "\u2611" if line["tax_type"] else "\u2610"
        withhold_display = "\u2611" if line["withhold"] else "\u2610"
        return (
            chk_display, line["gl_acct"], line["short_text"], f"{line['dc']} \u2304",
            line["amount_doc"], line["amount_loc"], tax_type_display, line["tax_jur"], withhold_display,
            line["assignment"], line["cost_center"], line["order"], line["profit_ctr"],
        )

    def _refresh_grid_row(self, row):
        line = self.lines[row]
        row_icon = checkmark_row_icon() if line["validated"] else checkbox_empty_icon()
        self.tree.item(str(row), image=row_icon, values=self._row_values(row))
        filled = sum(1 for ln in self.lines if ln["gl_acct"])
        self.item_count_var.set(f"{filled} Items ( No entry variant selected )")

    def set_line_field(self, row, field, value):
        """Sets a single line-item cell. Shared by the inline grid editor and
        headless verification -- both drive exactly this method."""
        self.lines[row][field] = value
        if field == "gl_acct":
            gl = self.data.get_gl_account(value)
            if gl:
                self.lines[row]["short_text"] = gl["text"]
        if field == "amount_doc":
            try:
                Decimal(value)
                self.lines[row]["amount_loc"] = value
            except InvalidOperation:
                pass
        if field in ("amount_doc", "dc", "gl_acct", "cost_center"):
            self.balanced = False
            self.lines[row]["validated"] = False
        self._refresh_grid_row(row)
        self._update_balance_display()

    def _on_grid_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col_id = self.tree.identify_column(event.x)
        row_iid = self.tree.identify_row(event.y)
        if not row_iid:
            return
        col_index = int(col_id.replace("#", "")) - 1
        if col_index < 0 or col_index >= len(COLUMNS):
            return
        field = COLUMNS[col_index]
        row = int(row_iid)

        if field in ("chk", "withhold", "tax_type"):
            self.set_line_field(row, field, not self.lines[row][field])
            return
        if field == "dc":
            current = self.lines[row][field]
            self.set_line_field(row, field, "Credit" if current == "Debit" else "Debit")
            return

        bbox = self.tree.bbox(row_iid, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        edit_var = tk.StringVar(value=str(self.lines[row][field]))
        entry = tk.Entry(self.tree, textvariable=edit_var, font=FONT_SMALL)
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        entry.select_range(0, "end")

        def commit(event=None):
            if entry.winfo_exists():
                self.set_line_field(row, field, edit_var.get())
                entry.destroy()

        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)
        entry.bind("<Escape>", lambda e: entry.destroy())

    # ------------------------------------------------------------- balance
    def _on_amount_changed(self, *args):
        self.balanced = False
        self._update_balance_display()

    def _update_balance_display(self):
        if self.balanced:
            value = Decimal("0.00")
        else:
            try:
                value = parse_amount(self.amount_var.get())
            except InvalidOperation:
                value = Decimal("0")
        self.balance_var.set(fmt_amount(value))
        dot_img = balance_indicator_icon(self.balanced and value == 0)
        self.balance_dot_label.configure(image=dot_img)
        self.balance_dot_label.image = dot_img

    # ------------------------------------------------------------- simulate
    def on_simulate(self):
        if not self.vendor:
            self.controller.set_status("Please enter a valid supplier before simulating.", error=True)
            return None
        try:
            header_amount = parse_amount(self.amount_var.get())
        except InvalidOperation:
            self.controller.set_status("Please enter a valid invoice amount.", error=True)
            return None
        if header_amount <= 0:
            self.controller.set_status("Invoice amount must be greater than zero.", error=True)
            return None

        line0 = self.lines[0]
        if not line0["gl_acct"] or not self.data.get_gl_account(line0["gl_acct"]):
            self.controller.set_status("Enter a valid G/L account in line item 1.", error=True)
            return None
        try:
            line0_amount = parse_amount(line0["amount_doc"])
        except InvalidOperation:
            self.controller.set_status("Enter a valid amount in line item 1.", error=True)
            return None
        if line0_amount <= 0:
            self.controller.set_status("Line item 1 amount must be greater than zero.", error=True)
            return None

        cost_center_code = (line0["cost_center"] or "").strip()
        cc = self.data.get_cost_center(cost_center_code) or self.data.cost_centers[0]
        line0["profit_ctr"] = cc.get("profit_center", "")

        dc0 = line0["dc"]
        dc1 = "Credit" if dc0 == "Debit" else "Debit"
        ap_account = self.data.get_gl_account("210000")
        self.lines[1].update({
            "chk": False,
            "gl_acct": "210000",
            "short_text": ap_account["text"] if ap_account else "Accounts Payable - Trade",
            "dc": dc1,
            "amount_doc": str(line0_amount),
            "amount_loc": str(line0_amount),
            "tax_type": False,
            "tax_jur": "",
            "withhold": False,
            "assignment": "",
            "cost_center": "",
            "order": "",
            "profit_ctr": "",
            "validated": True,
        })
        line0["validated"] = True
        self._refresh_grid_row(0)
        self._refresh_grid_row(1)
        self.balanced = True
        self._update_balance_display()

        message = f"Segment {cc['segment']} derived as sequence {cc['segment_sequence']} {cc['segment_desc']}"
        self.controller.set_status(message, ok=True)
        return message

    # ---------------------------------------------------------------- post
    def _line_to_doc_format(self, line, tax_code=""):
        try:
            amt = Decimal(line["amount_doc"] or "0")
        except InvalidOperation:
            amt = Decimal("0")
        signed = amt if line["dc"] == "Debit" else -amt
        return {
            "account": line["gl_acct"],
            "description": line["short_text"],
            "profit_center": line["profit_ctr"],
            "cost_center": line["cost_center"],
            "clearing_doc": "",
            "order": line["order"],
            "currency": self.currency_var.get(),
            "amount": fmt_amount(signed),
            "lcurr_amount": fmt_amount(signed),
            "tax_code": tax_code,
            "tax_type": "",
            "assignment": line["assignment"],
        }

    def _build_post_payload(self):
        header = {
            "company_code": self.company_code,
            "posting_date": self.posting_date_var.get().strip() or today_mmddyyyy(),
            "invoice_date": self.invoice_date_var.get().strip(),
            "reference": self.reference_var.get().strip(),
            "currency": self.currency_var.get(),
            "amount": self.amount_var.get().strip(),
            "text": self.text_var.get().strip(),
            "calculate_tax": bool(self.calc_tax_var.get()),
            "vendor_name": self.vendor["name1"] if self.vendor else "",
        }
        gl_line = self._line_to_doc_format(self.lines[0], tax_code=self.tax_code_var.get().strip())
        gl_line["posting_key"] = "40"
        cost_center_code = (self.lines[0]["cost_center"] or "").strip()
        cc = self.data.get_cost_center(cost_center_code) or (self.data.cost_centers[0] if self.data.cost_centers else None)
        if cc:
            gl_line["segment"] = cc["segment"]

        # Real FB60 posts the vendor (payables) line as document item 1 with
        # the vendor's own account number/name -- it's implicit from the
        # header, not something typed into the line-item grid -- and the
        # entered GL line as item 2. This only affects the *display* lines
        # handed to Document Overview/Display; fb60.lines (the grid data
        # model / balance calc) is left untouched.
        ap_line = self._line_to_doc_format(self.lines[1])
        ap_line["posting_key"] = "31"
        if self.vendor:
            ap_line["account"] = self.vendor["number"]
            ap_line["description"] = self.vendor["name1"]

        lines = [ap_line, gl_line]
        return header, lines

    def on_document_post(self):
        if not self.vendor:
            self.controller.set_status("Please enter a supplier before posting.", error=True)
            return None
        if not self.balanced:
            self.controller.set_status("Please simulate the document before posting.", error=True)
            return None
        header, lines = self._build_post_payload()
        popup = DocumentOverviewPopup(
            self.controller, self.data, self.vendor["number"], header, lines, on_posted=self._on_posted,
        )
        return popup

    def _on_posted(self, doc):
        self.last_doc_number = doc["doc_number"]

    def on_document_display(self):
        if not self.last_doc_number:
            self.controller.set_status("No document to display. Post a document first.", error=True)
            return
        self.controller.show_screen(
            "DOC_DISPLAY",
            context={"doc_number": self.last_doc_number, "return_to": "FB60", "return_context": None},
        )

    def on_exit(self):
        """FB60 is the app's home screen in this trimmed scope (no SAP Easy
        Access to return to) -- Exit/Cancel resets to a fresh, blank invoice
        entry screen, ready for the next vendor search."""
        self.controller.show_screen("FB60", record_history=False)


# ------------------------------------------------------------- Doc Display
DOC_DISPLAY_COLUMNS = [
    "item", "pk", "account", "description", "profit_center", "cost_center", "clearing_doc",
    "order", "tr_prt", "currency", "amount", "lcurr", "lcurr_amount", "tax_code", "tax_type", "assignment",
]
DOC_DISPLAY_HEADERS = {
    "item": "Item", "pk": "K", "account": "Account", "description": "Description",
    "profit_center": "Profit Center", "cost_center": "Cost Center", "clearing_doc": "Clrng doc.",
    "order": "Order", "tr_prt": "Tr...", "currency": "Curr.", "amount": "Amount", "lcurr": "LCurr",
    "lcurr_amount": "Amount", "tax_code": "Tx", "tax_type": "TTy", "assignment": "Assig",
}


class DocumentDisplayScreen(tk.Frame):
    TCODE = "DOC_DISPLAY"

    def __init__(self, parent, controller, data, context=None):
        super().__init__(parent, bg=CONTENT_BG, name="document_display_screen")
        self.controller = controller
        self.data = data
        self.context = context or {}
        self.doc_number = self.context.get("doc_number")
        self.doc = data.get_document(self.doc_number) if self.doc_number else None

        self._build_menu_strip()
        self._build_toolbar()

        body = tk.Frame(self, bg=CONTENT_BG, name="doc_display_body", padx=8, pady=6)
        body.pack(fill="both", expand=True)

        self._build_header(body)
        self._build_grid(body)

    def _build_menu_strip(self):
        def stub(label):
            return lambda: self.controller.set_status(f"'{label}' is not implemented in this demo.")

        defs = [
            ("Document", [("Print", stub("Print")), (None, None), ("Exit", self.on_exit)]),
        ]
        for label in ("Edit", "Goto", "Extras", "Environment", "System", "Help"):
            defs.append((label, [("(demo placeholder)", stub(label))]))
        strip = build_menu_strip(self, defs)
        strip.pack(fill="x", side="top")
        self.menu_strip = strip  # exposed for the (opt-in) automation bridge

    def _build_toolbar(self):
        bar = tk.Frame(self, bg=TOOLBAR_BG, name="doc_display_toolbar")
        bar.pack(fill="x", side="top")

        def stub(label):
            return lambda: self.controller.set_status(f"'{label}' is not implemented in this demo.")

        check_btn = tk.Button(bar, image=enter_icon(), bg=TOOLBAR_BG, relief="raised", bd=1)
        check_btn.pack(side="left", padx=2, pady=3)

        for image, tip, cmd in (
            (grid_small_icon, "Display", stub("Display")),
            (magnifier_icon, "Find", stub("Find")),
            (paperclip_icon, "Attachments", stub("Attachments")),
        ):
            b = tk.Button(bar, image=image(), bg=TOOLBAR_BG, relief="raised", bd=1, command=cmd)
            b.pack(side="left", padx=2, pady=3)
            add_tooltip(b, tip)

        for label, tip, cmd in (
            ("Taxes", "Taxes", stub("Taxes")),
            ("Display Currency", "Display Currency", stub("Display Currency")),
            ("General Ledger View", "General Ledger View", stub("General Ledger View")),
        ):
            b = tk.Button(bar, text=f" {label}", bg=TOOLBAR_BG, relief="raised", bd=1, font=FONT_SMALL, fg=TAB_ACTIVE_FG)
            b.pack(side="left", padx=4, pady=3)
            add_tooltip(b, tip)

        cancel_btn = tk.Button(
            bar, text="Cancel", bg=TOOLBAR_BG, relief="raised", bd=1, font=FONT_SMALL, name="doc_display_cancel_btn",
            command=self.on_exit,
        )
        cancel_btn.pack(side="left", padx=4, pady=3)

        obj_btn = tk.Button(
            bar, image=objects_icon(), bg=TOOLBAR_BG, relief="raised", bd=1,
            name="doc_display_object_services_btn", command=self.open_object_services,
        )
        obj_btn.pack(side="right", padx=6, pady=3)
        add_tooltip(obj_btn, "Object Services")

        exit_btn = tk.Button(
            bar, image=exit_icon(), text=" Exit", compound="left", bg=TOOLBAR_BG, relief="raised", bd=1,
            font=FONT_SMALL, name="doc_display_exit_btn", command=self.on_exit,
        )
        exit_btn.pack(side="right", padx=4, pady=3)

    def _build_header(self, parent):
        head = tk.Frame(parent, bg=CONTENT_BG, name="doc_display_header")
        head.pack(fill="x", pady=(0, 8))
        tk.Label(head, text="Data Entry View", bg=CONTENT_BG, font=FONT_HEADER).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))

        doc = self.doc or {}
        fields = [
            ("Document Number", doc.get("doc_number", ""), "Company Code", doc.get("company_code", ""), "Fiscal Year", doc.get("fiscal_year", "")),
            ("Document Date", doc.get("document_date", ""), "Posting Date", doc.get("posting_date", ""), "Period", doc.get("period", "")),
            ("Reference", doc.get("reference", ""), "Cross-Comp.No.", doc.get("cross_comp_no", ""), None, None),
            ("Currency", doc.get("currency", ""), "Texts exist", None, "Ledger Group", doc.get("ledger_group", "")),
        ]
        for r, (l1, v1, l2, v2, l3, v3) in enumerate(fields, start=1):
            cell1 = tk.Frame(head, bg=CONTENT_BG)
            cell1.grid(row=r, column=0, sticky="w", padx=10, pady=3)
            tk.Label(cell1, text=f"{l1}:", bg=CONTENT_BG, font=FONT_SMALL, width=15, anchor="w").pack(side="left")
            tk.Label(cell1, text=v1, bg=CONTENT_BG, font=FONT_SMALL, anchor="w").pack(side="left")

            if l2:
                cell2 = tk.Frame(head, bg=CONTENT_BG)
                cell2.grid(row=r, column=1, sticky="w", padx=10, pady=3)
                tk.Label(cell2, text=f"{l2}:", bg=CONTENT_BG, font=FONT_SMALL, width=15, anchor="w").pack(side="left")
                if v2 is None:
                    texts_var = tk.BooleanVar(value=bool(doc.get("texts_exist")))
                    tk.Checkbutton(cell2, variable=texts_var, state="disabled", bg=CONTENT_BG, name="texts_exist_check").pack(side="left")
                    self._texts_var = texts_var
                else:
                    tk.Label(cell2, text=v2, bg=CONTENT_BG, font=FONT_SMALL, anchor="w").pack(side="left")

            if l3:
                cell3 = tk.Frame(head, bg=CONTENT_BG)
                cell3.grid(row=r, column=2, sticky="w", padx=10, pady=3)
                tk.Label(cell3, text=f"{l3}:", bg=CONTENT_BG, font=FONT_SMALL, width=15, anchor="w").pack(side="left")
                tk.Label(cell3, text=v3, bg=CONTENT_BG, font=FONT_SMALL, anchor="w").pack(side="left")

    def _build_grid(self, parent):
        style = ttk.Style()
        style_treeview(style, "Sap.Treeview")
        self.table = ttk.Treeview(
            parent, columns=DOC_DISPLAY_COLUMNS, show="headings", style="Sap.Treeview", name="doc_display_grid", height=6,
        )
        for col in DOC_DISPLAY_COLUMNS:
            self.table.heading(col, text=DOC_DISPLAY_HEADERS[col])
            width = 45 if col in ("item", "pk") else 80
            self.table.column(col, width=width, anchor="w")
        self.table.pack(fill="both", expand=True)

        for i, line in enumerate((self.doc or {}).get("lines", []), start=1):
            values = (
                i, line.get("posting_key", ""), line.get("account", ""), line.get("description", ""),
                line.get("profit_center", ""), line.get("cost_center", ""), line.get("clearing_doc", ""),
                line.get("order", ""), "", line.get("currency", ""), line.get("amount", ""), "USD",
                line.get("lcurr_amount", ""), line.get("tax_code", ""), line.get("tax_type", ""),
                line.get("assignment", ""),
            )
            self.table.insert("", "end", values=values)

    # ----------------------------------------------------------- interactions
    def open_object_services(self, x=None, y=None):
        if x is None or y is None:
            x = self.winfo_rootx() + self.winfo_width() - 40
            y = self.winfo_rooty() + 60
        return open_object_services_menu(self.controller, x, y, on_attachment_list=self.open_attachment_list)

    def open_attachment_list(self):
        return open_attachment_list_popup(
            self.controller, self.data, self.doc_number, status_target=self.controller,
        )

    def on_exit(self):
        return_to = self.context.get("return_to", "FB60")
        return_context = self.context.get("return_context")
        self.controller.show_screen(return_to, context=return_context, record_history=False)


# =========================================================================
# APP -- main application window: menu bar, toolbar/command field, content
# area, status bar. Launches straight into FB60 (Enter Supplier Invoice:
# Company Code 5300) -- there is no SAP Easy Access shell in this trimmed
# scope. Mirrors the chrome of a real SAP GUI session so downstream RPA
# automation (rpaframework RPA.Desktop) can target it the same way: type a
# t-code + Enter in the command field, or click named toolbar buttons.
# =========================================================================

SCREENS = {
    "FB60": (FB60Screen, "Enter Supplier Invoice: Company Code 5300"),
    "DOC_DISPLAY": (DocumentDisplayScreen, "Display Document: Data Entry View"),
}

HOME_TCODE = "FB60"


# =========================================================================
# AUTOMATION BRIDGE -- strictly opt-in, additive, and read-only with respect
# to app behavior. Gated behind SAP_MIRROR_AUTOMATION=1; when that env var
# is unset (normal `python3 main.py`, `verify.py`, `live_demo.py`), none of
# this code runs and nothing below is ever imported/scheduled/touched.
#
# Purpose: report live, real, absolute screen coordinates + app state,
# computed fresh from the actual running Tk widgets each time, so an
# external OS-level automation (RPA.Desktop) can click/type into this app
# the same way a human would -- no hardcoded pixels, no image matching.
# This module only *reports* geometry/state; it never triggers actions.
#
# Transport: a tiny file-based request/response IPC, polled from inside the
# existing Tk event loop via self.after(...) -- no threads, no server, so
# there is no cross-thread Tkinter safety concern.
# =========================================================================

AUTOMATION_ENABLED = os.environ.get("SAP_MIRROR_AUTOMATION") == "1"
BRIDGE_REQUEST_FILE = Path(os.environ.get("SAP_MIRROR_REQUEST_FILE", "/tmp/sap_mirror_bridge_request.json"))
BRIDGE_RESPONSE_FILE = Path(os.environ.get("SAP_MIRROR_RESPONSE_FILE", "/tmp/sap_mirror_bridge_response.json"))
BRIDGE_POLL_MS = 130


def _bridge_find_widget(root, name):
    """Recursively search the widget tree -- including any currently open
    Toplevel popups/dialogs, since those are real children of `root` in
    Tk's window hierarchy -- for a widget whose winfo_name() == name."""
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
    w = _bridge_find_widget(app, name)
    if w is None:
        return {"ok": False, "error": f"no widget named {name!r} found"}
    return _bridge_bbox(w)


def _bridge_resolve_grid_cell(app, req):
    tree_name = req.get("tree_name") or "line_item_grid"
    tree = _bridge_find_widget(app, tree_name)
    if tree is None:
        return {"ok": False, "error": f"no grid named {tree_name!r} found (is its screen open?)"}
    row = str(req.get("row"))
    column = req.get("column")
    try:
        tree.see(row)
        tree.update_idletasks()
        bbox = tree.bbox(row, column)
    except tk.TclError as exc:
        return {"ok": False, "error": str(exc)}
    if not bbox:
        return {"ok": False, "error": f"row {row!r} / column {column!r} not visible in {tree_name!r}"}
    x, y, w, h = bbox
    rx, ry = tree.winfo_rootx(), tree.winfo_rooty()
    abs_x, abs_y = rx + x, ry + y
    return {
        "ok": True, "x": abs_x, "y": abs_y, "width": w, "height": h,
        "center_x": abs_x + w // 2, "center_y": abs_y + h // 2,
    }


def _bridge_resolve_tree_row_by_text(app, req):
    """Like _bridge_resolve_grid_cell, but the caller doesn't know the row
    index ahead of time (e.g. vendor-search results, whose order depends on
    the search) -- finds the first row whose given column's value contains
    the requested text instead."""
    tree_name = req.get("tree_name")
    tree = _bridge_find_widget(app, tree_name)
    if tree is None:
        return {"ok": False, "error": f"no tree named {tree_name!r} found (is its screen open?)"}
    column = req.get("column")
    text = (req.get("text") or "").strip().lower()
    match_row = None
    for iid in tree.get_children():
        value = str(tree.set(iid, column) or "").strip().lower()
        if text in value:
            match_row = iid
            break
    if match_row is None:
        return {"ok": False, "error": f"no row with {column!r} containing {text!r} found in {tree_name!r}"}
    try:
        tree.see(match_row)
        tree.update_idletasks()
        bbox = tree.bbox(match_row, column)
    except tk.TclError as exc:
        return {"ok": False, "error": str(exc)}
    if not bbox:
        return {"ok": False, "error": f"row {match_row!r} not visible in {tree_name!r}"}
    x, y, w, h = bbox
    rx, ry = tree.winfo_rootx(), tree.winfo_rooty()
    abs_x, abs_y = rx + x, ry + y
    return {
        "ok": True, "x": abs_x, "y": abs_y, "width": w, "height": h,
        "center_x": abs_x + w // 2, "center_y": abs_y + h // 2,
    }


def _bridge_resolve_listbox_item(app, req):
    widget_name = req.get("widget_name") or "file_picker_listbox"
    lb = _bridge_find_widget(app, widget_name)
    if lb is None:
        return {"ok": False, "error": f"no listbox named {widget_name!r} found (is it open?)"}
    text = req.get("text")
    items = list(lb.get(0, "end"))
    if text not in items:
        return {"ok": False, "error": f"{text!r} not found in {widget_name!r} ({items})"}
    index = items.index(text)
    try:
        lb.see(index)
        lb.update_idletasks()
        bbox = lb.bbox(index)
    except tk.TclError as exc:
        return {"ok": False, "error": str(exc)}
    if not bbox:
        return {"ok": False, "error": f"item {index} not visible in {widget_name!r}"}
    x, y, w, h = bbox
    rx, ry = lb.winfo_rootx(), lb.winfo_rooty()
    abs_x, abs_y = rx + x, ry + y
    return {
        "ok": True, "x": abs_x, "y": abs_y, "width": w, "height": h,
        "center_x": abs_x + w // 2, "center_y": abs_y + h // 2,
    }


def _bridge_resolve_calendar_date(app, req):
    cal_popups = [w for w in app.winfo_children() if isinstance(w, CalendarPopup)]
    if not cal_popups:
        return {"ok": False, "error": "calendar popup is not open"}
    cal = cal_popups[-1]
    date_str = req.get("date")
    try:
        cell = cal.locate_date_cell(date_str)
    except (ValueError, TypeError):
        return {"ok": False, "error": f"invalid date {date_str!r}, expected MM/DD/YYYY"}
    if cell is None:
        return {"ok": False, "error": f"date {date_str} could not be rendered in the calendar"}
    return _bridge_bbox(cell)


def _bridge_resolve_menu_item(app, req):
    # All menus in this app (the Document/Edit/... pulldown strip and the
    # transient object-services/"New" context menus) are `CustomMenu`
    # instances now -- only one is ever posted at a time, tracked directly
    # on the app as `_active_menu` -- so a single lookup covers every
    # `menu` value the .robot passes ('document' or 'context'), used here
    # only for a clearer error message.
    menu_id = req.get("menu")
    label = req.get("label")
    menu = getattr(app, "_active_menu", None)
    if menu is None:
        return {"ok": False, "error": f"menu {menu_id!r} is not currently posted"}
    frame = menu.find_item(label)
    if frame is None:
        return {"ok": False, "error": f"no item labeled {label!r} in menu {menu_id!r}"}
    return _bridge_bbox(frame)


def _bridge_resolve_state(app, req):
    return {
        "ok": True,
        "current_tcode": app.current_tcode,
        "status": app.status_var.get(),
    }


_BRIDGE_RESOLVERS = {
    "widget": _bridge_resolve_widget,
    "grid_cell": _bridge_resolve_grid_cell,
    "tree_row_by_text": _bridge_resolve_tree_row_by_text,
    "listbox_item": _bridge_resolve_listbox_item,
    "calendar_date": _bridge_resolve_calendar_date,
    "menu_item": _bridge_resolve_menu_item,
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
        app.after(BRIDGE_POLL_MS, lambda: _bridge_poll(app))


class SAPApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.data = DemoData()
        self.history = []
        self.current_tcode = None
        self.current_context = None
        self.current_screen = None

        style = ttk.Style()
        style_combobox(style)

        self.title(f"{APP_TITLE} - {SCREENS[HOME_TCODE][1]}")
        self.geometry("1180x760")
        self.configure(bg=WINDOW_BG)

        def _debug_click(event):
            try:
                cur = self.grab_current()
            except tk.TclError:
                cur = "ERR"
            try:
                with open("/tmp/menu_debug.log", "a") as f:
                    f.write(
                        f"CLICK x_root={event.x_root} y_root={event.y_root} "
                        f"target={event.widget} grab_current={cur}\n"
                    )
            except Exception:
                pass
        self.bind_all("<Button-1>", _debug_click, add="+")

        self._build_menubar()
        self._build_toolbar()
        self._build_title_bar()
        self._content_container = tk.Frame(self, bg=CONTENT_BG, name="content_area")
        self._content_container.pack(side="top", fill="both", expand=True)
        self._build_statusbar()

        self.protocol("WM_DELETE_WINDOW", self.on_exit)
        self.show_screen(HOME_TCODE)

        if AUTOMATION_ENABLED:
            # Opt-in automation bridge (see AUTOMATION BRIDGE section above)
            # -- reports geometry/state to an external RPA driver over a
            # small file-based IPC. No-op when SAP_MIRROR_AUTOMATION is unset.
            self.after(BRIDGE_POLL_MS, lambda: _bridge_poll(self))

    # ------------------------------------------------------------------ menu
    def _build_menubar(self):
        menubar = tk.Menu(self)

        def stub(label):
            return lambda: self.set_status(f"'{label}' is not implemented in this demo.")

        m_menu = tk.Menu(menubar, tearoff=0)
        m_menu.add_command(label="New session", command=stub("New session"))
        m_menu.add_separator()
        m_menu.add_command(label="Exit", command=self.on_exit)
        menubar.add_cascade(label="Menu", menu=m_menu)

        for label in ("Edit", "Favorites", "Extras"):
            m = tk.Menu(menubar, tearoff=0)
            m.add_command(label="(demo placeholder)", command=stub(label))
            menubar.add_cascade(label=label, menu=m)

        sys_menu = tk.Menu(menubar, tearoff=0)
        sys_menu.add_command(label="Create session", command=stub("Create session"))
        sys_menu.add_separator()
        sys_menu.add_command(label="Log off", command=self.on_exit)
        menubar.add_cascade(label="System", menu=sys_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self.on_about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.config(menu=menubar)

    # --------------------------------------------------------------- toolbar
    def _build_toolbar(self):
        bar = tk.Frame(self, bg=TOOLBAR_BG, name="toolbar")
        bar.pack(side="top", fill="x")

        inner = tk.Frame(bar, bg=TOOLBAR_BG)
        inner.pack(side="top", fill="x", padx=4, pady=3)

        check_btn = tk.Button(
            inner, image=enter_icon(), bg=TOOLBAR_BG, relief="raised", bd=1,
            name="btn_check", command=self.on_enter_command,
        )
        check_btn.pack(side="left", padx=(2, 4))
        add_tooltip(check_btn, "Enter")

        self.command_var = tk.StringVar()
        self.command_entry = ttk.Combobox(
            inner,
            textvariable=self.command_var,
            width=28,
            font=FONT_NORMAL,
            name="command_field",
            values=["FB60"],
            style="Sap.TCombobox",
        )
        self.command_entry.pack(side="left", padx=(0, 4))
        self.command_entry.bind("<Return>", lambda e: self.on_enter_command())

        toolbar_buttons = [
            (folder_icon(), "btn_open", "Open"),
            (save_icon(), "btn_save2", "Save"),
            (star_icon(), "btn_favorite", "Favorites"),
            (options_icon(), "btn_more_favs", "More favorites"),
        ]
        for image, name, tooltip in toolbar_buttons:
            b = tk.Button(
                inner, image=image, bg=BUTTON_RAISED, relief="raised", bd=1, name=name,
                command=lambda t=tooltip: self.set_status(f"'{t}' is not implemented in this demo."),
            )
            b.pack(side="left", padx=2)
            add_tooltip(b, tooltip)

        right_buttons = [
            (magnifier_icon(), "btn_search", "Find"),
            (print_icon(), "btn_print", "Print"),
        ]
        for image, name, tooltip in right_buttons:
            b = tk.Button(
                inner, image=image, bg=BUTTON_RAISED, relief="raised", bd=1, name=name,
                command=lambda t=tooltip: self.set_status(f"'{t}' is not implemented in this demo."),
            )
            b.pack(side="right", padx=2)
            add_tooltip(b, tooltip)

        logoff_btn = tk.Button(
            inner, text="Log Off", bg=TOOLBAR_BG, relief="flat", bd=0, font=FONT_SMALL,
            fg=LINK_FG, name="btn_logoff", command=self.on_exit,
        )
        logoff_btn.pack(side="right", padx=6)

        self.bind("<F3>", lambda e: self.on_back())

    def _build_title_bar(self):
        bar = tk.Frame(self, bg=TITLE_BG, name="app_title_bar", height=40)
        bar.pack(side="top", fill="x")
        logo = sap_logo_icon()
        logo_label = tk.Label(bar, image=logo, bg=TITLE_BG, name="sap_logo")
        logo_label.image = logo
        logo_label.pack(side="left", padx=(10, 8), pady=6)
        self.title_bar_var = tk.StringVar(value=SCREENS[HOME_TCODE][1])
        tk.Label(
            bar, textvariable=self.title_bar_var, bg=TITLE_BG, fg=TITLE_FG,
            font=FONT_HEADER, name="app_title_label",
        ).pack(side="left", expand=True)

    # ------------------------------------------------------------- statusbar
    def _build_statusbar(self):
        bar = tk.Frame(self, bg=STATUSBAR_BG, name="statusbar", height=24)
        bar.pack(side="bottom", fill="x")

        self.status_var = tk.StringVar(value="")
        self.status_label = tk.Label(
            bar,
            textvariable=self.status_var,
            bg=STATUSBAR_BG,
            fg=STATUSBAR_FG,
            font=FONT_SMALL,
            anchor="w",
            name="status_label",
        )
        self.status_label.pack(side="left", padx=6, fill="x", expand=True)

        tk.Label(
            bar, image=lock_icon(), bg=STATUSBAR_BG, name="lock_icon"
        ).pack(side="right", padx=(0, 6))

        tk.Label(
            bar,
            text="SESSION_MANAGER   |   ussdprdecap01   |   INS",
            bg=STATUSBAR_BG,
            fg=STATUSBAR_FG,
            font=FONT_SMALL,
            name="session_info",
        ).pack(side="right", padx=6)

        tk.Label(
            bar, image=status_circle_icon(), bg=STATUSBAR_BG, name="status_circle"
        ).pack(side="right", padx=(6, 0))

    # ------------------------------------------------------------- behavior
    def set_status(self, message, ok=False, error=False):
        self.status_var.set(message)
        if error:
            self.status_label.configure(fg=ERROR_FG)
        elif ok:
            self.status_label.configure(fg=OK_FG)
        else:
            self.status_label.configure(fg=STATUSBAR_FG)

    def on_enter_command(self):
        tcode = self.command_var.get().strip().upper()
        if not tcode:
            return
        self.navigate_to(tcode)

    def navigate_to(self, tcode, context=None):
        """Central navigation entry point used by the command field and
        internal screen->screen links alike."""
        if tcode not in SCREENS:
            self.set_status(f"Transaction code {tcode} does not exist.", error=True)
            return
        self.show_screen(tcode, context=context)

    def on_back(self):
        if self.history:
            previous_tcode, previous_context = self.history.pop()
            self.show_screen(previous_tcode, context=previous_context, record_history=False)
        elif self.current_tcode != HOME_TCODE:
            self.show_screen(HOME_TCODE, record_history=False)

    def on_exit(self):
        try:
            self.destroy()
        except tk.TclError:
            pass

    def on_about(self):
        messagebox.showinfo(
            "About",
            "SAP FI/AP Invoice Entry Mirror (Demo)\nDummy application for RPA automation demos.\nNot affiliated with SAP SE.",
        )

    def show_screen(self, tcode, context=None, record_history=True):
        if tcode not in SCREENS:
            self.set_status(f"Transaction code {tcode} does not exist.", error=True)
            return

        if record_history and self.current_tcode and self.current_tcode != tcode:
            self.history.append((self.current_tcode, self.current_context))

        if self.current_screen is not None:
            self.current_screen.destroy()

        # Set the generic per-transaction status first -- a screen's own
        # __init__ may set something more specific, which should win rather
        # than being clobbered afterward.
        self.set_status(f"Transaction {tcode} started.")

        screen_cls, title_suffix = SCREENS[tcode]
        self.current_screen = screen_cls(self._content_container, self, self.data, context=context)
        self.current_screen.pack(fill="both", expand=True)

        self.current_tcode = tcode
        self.current_context = context
        self.command_var.set(tcode)
        self.title(f"{APP_TITLE} - {title_suffix}")
        self.title_bar_var.set(title_suffix)


def main():
    app = SAPApp()
    app.mainloop()


if __name__ == "__main__":
    main()
