#!/bin/sh
# Builds a single-file executable (dist/sheetxpert) with PyInstaller.
# TIP: for the widest compatibility build it on an OLD distro (Ubuntu 20.04/22.04, Debian 11):
# the binary then runs on that glibc version and anything newer. Requires python3 with tkinter.
set -eu
cd "$(dirname "$0")"
python3 -m venv .build-venv
. ./.build-venv/bin/activate
pip install --quiet -r requirements.txt pyinstaller
pyinstaller --noconfirm --clean --onefile --name sheetxpert --add-data "assets:assets" sheetxpert.py
echo
echo "Done: $(pwd)/dist/sheetxpert"
