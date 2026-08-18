"""Bot Progress window: a small always-on-top notification window that
narrates what an unattended automation is doing, step by step, for a human
watching it run. Styled to match the Maker Player desktop app's actual brand
tokens (hosts/src/MakerPlayer.Desktop/Themes/Colors.axaml + Brushes.axaml:
dark background #150B27, white/secondary text, Inter font) rather than a
generic look, so it reads as the same product family.

Runs as its own OS process (not a thread inside the automation): Tkinter's
Cocoa backend requires its event loop on a process's main thread, which a
background thread inside the Robot Framework/Python process cannot
guarantee across platforms. Decoupling via a process + a polled state file
mirrors this plugin's own monitor-ui (a local server the automation posts
to) rather than introducing threading.

Usage from a `.robot` task (see build-rpa-automation.md section 10). Verify
`${PROGRESS_PYTHON}` actually has Tkinter before wiring this in
(`${PROGRESS_PYTHON} -c "import tkinter"`) - a project's dedicated RPA venv
can lack Tcl/Tk bindings even though the automation's own libraries work
fine in it, in which case point this at a different interpreter that does
(confirmed the same way) rather than assuming they're interchangeable:
    *** Settings ***
    Library    Process
    Library    OperatingSystem

    *** Variables ***
    ${PROGRESS_PYTHON}    python3

    *** Tasks ***
    Run With Progress
        Create File    ${STATE_PATH}    content={}
        Start Process    ${PROGRESS_PYTHON}    ${CURDIR}/bot_progress_window.py    ${STATE_PATH}
        ...    alias=bot_progress
        Update Bot Progress    ${STATE_PATH}    Looking Up the Claim
        ...    The bot opened Salesforce, found the case, and read the claim number.
        # ... real automation steps ...
        Terminate Process    bot_progress

    *** Keywords ***
    Update Bot Progress
        [Arguments]    ${state_path}    ${headline}    ${body}
        ${json}=    Evaluate    json.dumps({"headline": $headline, "body": $body})    modules=json
        Create File    ${state_path}    content=${json}

Standalone usage: `python3 bot_progress_window.py <state_file_path>`. The
state file is a small JSON object `{"headline": "...", "body": "..."}`;
whichever process owns the automation rewrites it whenever the step changes,
and this script polls it (every 200ms) and updates the window in place. The
window exits when the automation kills this process (e.g. RF's
`Terminate Process`) or when the human closes it by hand.
"""
import json
import sys

try:
    import tkinter as tk
    import tkinter.font as tkfont
except ImportError:
    # Fails loudly instead of a silent no-op: Process.Start Process (RF) does
    # not surface a child's stderr by default, so a swallowed ImportError here
    # would look like "nothing happened" from the automation's side rather
    # than a clear diagnosis. Confirmed live: a Homebrew-built Python used for
    # an RPA venv can lack Tcl/Tk bindings even though every rpaframework
    # library works fine in it - point Start Process at a different
    # interpreter (verified via `<interpreter> -c "import tkinter"`) rather
    # than assuming the automation's own interpreter doubles as GUI-capable.
    print(
        "bot_progress_window: this Python has no working tkinter (no Tcl/Tk "
        "bindings). Point Start Process at a different interpreter that has "
        "it - verify with `<interpreter> -c \"import tkinter\"` first.",
        file=sys.stderr,
    )
    sys.exit(1)

BACKGROUND = "#150B27"
TEXT = "#FFFFFF"
TEXT_SECONDARY = "#DFDEE0"

WINDOW_WIDTH = 412
WINDOW_HEIGHT = 200
SCREEN_MARGIN = 24
POLL_MS = 200

# 32x32 PNG render of the Mimica mark (mimica.ico), embedded so this file has
# no external asset to lose track of. tk.PhotoImage decodes PNG natively on
# Tk 8.6+; regenerate from hosts/src/MakerPlayer.Desktop/Assets/mimica.ico if
# the brand mark ever changes.
_ICON_PNG_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAEEklEQVR4nO1Wa2wUVRT+5s6dmW3H7jYUKggqikiM"
    "CQmK+FpFLUYJSgIaDUk1Ka+wVttE+kNjbGJAomhCSSyWpBCjqJGghlgCtFQLJSrEB0YlRiQQQozp7pbuq3V3Z+6Y"
    "e/bRKWx3W/nRPz3J/TH3nDnnO895zCjTzHkOJlDYRAafBDDJwLgYUBTlqvRX14QOIIQAU0c3l3q1iP7/A3BAgTnn"
    "iMcS4FwtaKbrGqLROIEYDxusmFJVVcSG4njgoXvQ8/1+3Ou/C32RMN3LIPIk/03ixtmzcOzkV1hVuxLBaDjjmI2N"
    "DVbKwHEceL0VmHfbLejo/Bj1gdUIR/vpXrJhCwHTNDFn7k1o/6gFb21pRmJoEKlUalS2xgVAinAcqjHXObbv2IF3"
    "338HllgYHByCxjUIkdHL0/RqPfbt3w1vZQUikSiVrpiwolq3IWNwhAPbtrF2Qy06Oj/BzOtnYCB9CZrOSS+PZVl4"
    "fFkNunu/xMK7FyAYCRGI0fqCYRwincj6yyD3+RfhSO8XqHlwMYJ9obyNDGZZNubMnY1D3+zFmrpaBCNhYqcQCFY6"
    "c4XqLbN3B5FMTJ9RjYNdn6Fh43pqxmG9SgHLyjxo2/0uWlo2UyTpB1DGBkDNzvSQSNKIKUyBbQuXXiWHmq5RSQyP"
    "QUGHgTPSp9MW6hvX4JXXGjGQiIBzVhoAYwyxaJwc3HrDzfj11Gn8deYcAZH050RSKm2ELZBMpug9CYISzerlkcCv"
    "nV5dmGEUyDyeiOPhR/3oOv45fvzta7S2b8W2rW3oPNSTpz9D5/CGfHvTduzauSefubtk0mc6nS4IgLsfaLEkU5g5"
    "6zrs2dsGX6WX7u9YOJ+ayn/nMgQa6vBCw+p88Fy2sVgCazcE8PfFf/D6pia6k0DdvksywBhDIjmI+/2LKHgOdSqV"
    "hs/nxZLHFqO+8UVsfKmZ7qS9bWdKIkfRULx4c/M2PPVkHULB/vzEFBN2RQmg4lL/QP7ZtoazkPflbCpa39uFJ5as"
    "woXzF2EYBmVqW5INB9VVU9HRcRg1/hX44eQpeDxSn1lSJQHYto0K08TRnm/R3XkMmqZB5So134nvfsLBA90wPWWo"
    "qpyC470n8Ih/BboOH6VMDUOHgKCMq3xTcO7sBSyteRYftH9KPWCa5XDkl+0y4VfcZGv13DMBrAs8j/kLbscfp//E"
    "3rYPU2kqfT6E+sJ4enkdmt9ows8//06d69m1bFNAabd+3cs4e+Y8QqH+/NoeEW5agd/yzOjY9CVUoECAQcVnGlow"
    "ue4n+rJjF0lE4eEeotutz43hQCwCTeEolyw4YwAgRb5IPyDZ5SVn/fKXM3YSiAoRSNGGrJ5bckurUB/wYp9hdwOO"
    "bjdy3ApJMT3DBAubBIDJEmBi5T/ty9ZlOJ90YQAAAABJRU5ErkJggg=="
)


def _pick_font(root, *candidates, size=13, weight="normal"):
    available = set(tkfont.families(root))
    for name in candidates:
        if name in available:
            return tkfont.Font(family=name, size=size, weight=weight)
    return tkfont.Font(size=size, weight=weight)


def main(state_path):
    root = tk.Tk()
    root.title("Bot Progress")
    root.configure(bg=BACKGROUND)

    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    x = screen_w - WINDOW_WIDTH - SCREEN_MARGIN
    y = screen_h - WINDOW_HEIGHT - SCREEN_MARGIN
    root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{x}+{y}")
    root.minsize(320, 160)
    root.attributes("-topmost", True)

    try:
        icon = tk.PhotoImage(data=_ICON_PNG_B64, format="png")
        root.iconphoto(True, icon)
    except tk.TclError:
        pass

    headline_font = _pick_font(
        root, "Inter SemiBold", "Inter", "Segoe UI Semibold", "SF Pro Text",
        size=15, weight="bold",
    )
    body_font = _pick_font(root, "Inter", "Segoe UI", "SF Pro Text", size=12)

    outer = tk.Frame(root, bg=BACKGROUND, padx=20, pady=18)
    outer.pack(fill="both", expand=True)

    headline_label = tk.Label(
        outer, text="", font=headline_font, fg=TEXT, bg=BACKGROUND,
        anchor="w", justify="left", wraplength=WINDOW_WIDTH - 40,
    )
    headline_label.pack(fill="x", pady=(0, 8))

    body_label = tk.Label(
        outer, text="", font=body_font, fg=TEXT_SECONDARY, bg=BACKGROUND,
        anchor="nw", justify="left", wraplength=WINDOW_WIDTH - 40,
    )
    body_label.pack(fill="both", expand=True)

    last_raw = None

    def poll():
        nonlocal last_raw
        try:
            with open(state_path, "r") as f:
                raw = f.read()
        except OSError:
            root.after(POLL_MS, poll)
            return
        if raw != last_raw:
            last_raw = raw
            try:
                state = json.loads(raw) if raw.strip() else {}
            except json.JSONDecodeError:
                state = {}
            headline_label.config(text=state.get("headline", ""))
            body_label.config(text=state.get("body", ""))
        root.after(POLL_MS, poll)

    root.after(POLL_MS, poll)
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("usage: bot_progress_window.py <state_file_path>", file=sys.stderr)
        sys.exit(2)
    main(sys.argv[1])
