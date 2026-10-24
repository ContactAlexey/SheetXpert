#!/bin/sh
# SheetXpert - portable launcher (no installation needed).
# Uses your system Python; if fpdf2 is missing it creates a private virtualenv in ./.venv
set -eu
HERE=$(cd "$(dirname "$0")" && pwd)
PY=$(command -v python3 || command -v python || true)
if [ -z "$PY" ]; then
    echo "Python 3 was not found. Run ./install.sh or install python3 with your package manager." >&2
    exit 1
fi

has_fpdf2() {
    "$1" -c 'import sys, fpdf; sys.exit(0 if int(getattr(fpdf, "__version__", "0").split(".")[0]) >= 2 else 1)' >/dev/null 2>&1
}

if ! has_fpdf2 "$PY"; then
    VENV="$HERE/.venv"
    if [ ! -x "$VENV/bin/python" ]; then
        echo "First run: creating a private Python environment (internet needed once)..." >&2
        if "$PY" -m venv "$VENV" >/dev/null 2>&1 && "$VENV/bin/python" -m pip install --quiet -r "$HERE/requirements.txt"; then
            :
        else
            rm -rf "$VENV"
            echo "Warning: could not install fpdf2 - PDF export will be unavailable." >&2
            echo "         Run ./install.sh, or install python3-venv/python3-pip and try again." >&2
        fi
    fi
    [ -x "$VENV/bin/python" ] && PY="$VENV/bin/python"
fi

if ! "$PY" -c 'import tkinter' >/dev/null 2>&1; then
    echo "Note: tkinter is not installed, so the GUI is unavailable (text mode will be used)." >&2
    echo "      Debian/Ubuntu: python3-tk | Fedora: python3-tkinter | Arch: tk | openSUSE: python3-tk | Alpine: py3-tkinter" >&2
fi
exec "$PY" "$HERE/sheetxpert.py" "$@"
