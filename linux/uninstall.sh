#!/bin/sh
# Removes SheetXpert. Use "sudo ./uninstall.sh --system" for a system-wide install.
set -eu
if [ "${1:-}" = "--system" ]; then
    [ "$(id -u)" -eq 0 ] || { echo "Run with sudo." >&2; exit 1; }
    DEST=/opt/sheetxpert; BIN=/usr/local/bin; APPS=/usr/share/applications; ICONROOT=/usr/share/icons/hicolor
else
    DATA=${XDG_DATA_HOME:-$HOME/.local/share}
    DEST=$DATA/sheetxpert; BIN=$HOME/.local/bin; APPS=$DATA/applications; ICONROOT=$DATA/icons/hicolor
fi
rm -rf "$DEST"
rm -f "$BIN/sheetxpert" "$APPS/sheetxpert.desktop" "$ICONROOT/256x256/apps/sheetxpert.png"
command -v update-desktop-database >/dev/null 2>&1 && update-desktop-database "$APPS" >/dev/null 2>&1 || true
echo "SheetXpert removed."
