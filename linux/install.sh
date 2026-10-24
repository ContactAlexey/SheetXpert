#!/bin/sh
# SheetXpert installer for Linux.
# Supports apt (Debian/Ubuntu/Mint/Pop!_OS), dnf/yum (Fedora/RHEL/Rocky/Alma), pacman (Arch/Manjaro),
# zypper (openSUSE), apk (Alpine) and xbps (Void). Other distros: install python3 + tkinter yourself.
#
#   ./install.sh              per-user install  (~/.local, no root needed except for system packages)
#   sudo ./install.sh --system  system-wide install (/opt, /usr/local/bin)
#   ./install.sh --no-deps    do not touch the package manager
set -eu

HERE=$(cd "$(dirname "$0")" && pwd)
MODE=user
DEPS=yes

usage() { sed -n '2,8p' "$0" | sed 's/^# \{0,1\}//'; }

for arg in "$@"; do
    case "$arg" in
        --system)  MODE=system ;;
        --no-deps) DEPS=no ;;
        -h|--help) usage; exit 0 ;;
        *) echo "Unknown option: $arg" >&2; usage; exit 1 ;;
    esac
done

if [ "$MODE" = system ]; then
    [ "$(id -u)" -eq 0 ] || { echo "--system needs root: sudo ./install.sh --system" >&2; exit 1; }
    DEST=/opt/sheetxpert
    BIN=/usr/local/bin
    APPS=/usr/share/applications
    ICONROOT=/usr/share/icons/hicolor
else
    DATA=${XDG_DATA_HOME:-$HOME/.local/share}
    DEST=$DATA/sheetxpert
    BIN=$HOME/.local/bin
    APPS=$DATA/applications
    ICONROOT=$DATA/icons/hicolor
fi
ICONS=$ICONROOT/256x256/apps

as_root() {
    if [ "$(id -u)" -eq 0 ]; then "$@"
    elif command -v sudo >/dev/null 2>&1; then sudo "$@"
    elif command -v doas >/dev/null 2>&1; then doas "$@"
    else return 1
    fi
}

PYTHON=$(command -v python3 || true)
have_tk()   { [ -n "$PYTHON" ] && "$PYTHON" -c 'import tkinter' >/dev/null 2>&1; }
have_venv() { [ -n "$PYTHON" ] && "$PYTHON" -c 'import venv, ensurepip' >/dev/null 2>&1; }

install_deps() {
    if   command -v apt-get >/dev/null 2>&1; then as_root apt-get install -y python3 python3-venv python3-tk
    elif command -v dnf     >/dev/null 2>&1; then as_root dnf install -y python3 python3-tkinter
    elif command -v yum     >/dev/null 2>&1; then as_root yum install -y python3 python3-tkinter
    elif command -v zypper  >/dev/null 2>&1; then as_root zypper --non-interactive install python3 python3-pip python3-tk
    elif command -v pacman  >/dev/null 2>&1; then as_root pacman -S --needed --noconfirm python tk
    elif command -v apk     >/dev/null 2>&1; then
        as_root apk add python3 py3-pip python3-tkinter || as_root apk add python3 py3-pip py3-tkinter
    elif command -v xbps-install >/dev/null 2>&1; then as_root xbps-install -Sy python3 python3-tkinter
    elif command -v emerge  >/dev/null 2>&1; then
        echo "Gentoo: run  emerge --ask dev-lang/python dev-lang/tk  (Python must be built with USE=tk)" >&2; return 1
    else
        echo "Unknown package manager." >&2; return 1
    fi
}

if [ "$DEPS" = yes ] && { [ -z "$PYTHON" ] || ! have_tk || ! have_venv; }; then
    echo "==> Installing system packages (python3, venv, tkinter)..."
    install_deps || echo "!! Could not install packages automatically. Install python3, python3-venv and tkinter manually."
    PYTHON=$(command -v python3 || true)
fi

[ -n "$PYTHON" ] || { echo "Python 3 is required but was not found." >&2; exit 1; }
have_venv || { echo "The Python 'venv' module is missing (Debian/Ubuntu: apt install python3-venv)." >&2; exit 1; }
have_tk   || echo "!! tkinter is missing: the GUI will not start, but 'sheetxpert --cli' will still work."

echo "==> Installing SheetXpert to $DEST"
mkdir -p "$DEST" "$BIN" "$APPS" "$ICONS"
cp "$HERE/sheetxpert.py" "$HERE/requirements.txt" "$DEST/"
rm -rf "$DEST/assets"
cp -R "$HERE/assets" "$DEST/assets"

echo "==> Creating virtual environment and installing fpdf2 (internet needed)..."
rm -rf "$DEST/venv"
"$PYTHON" -m venv "$DEST/venv"
"$DEST/venv/bin/python" -m pip install --quiet -r "$DEST/requirements.txt" \
    || { echo "pip could not install the dependencies. Check your internet connection." >&2; exit 1; }

cat > "$BIN/sheetxpert" <<LAUNCHER
#!/bin/sh
exec "$DEST/venv/bin/python" "$DEST/sheetxpert.py" "\$@"
LAUNCHER
chmod 755 "$BIN/sheetxpert"

cp "$HERE/assets/sheetxpert_256.png" "$ICONS/sheetxpert.png"
cat > "$APPS/sheetxpert.desktop" <<DESKTOP
[Desktop Entry]
Type=Application
Name=SheetXpert
GenericName=System Technical Sheet
GenericName[es]=Ficha técnica del equipo
Comment=Generate a technical sheet of this computer and export it to PDF
Comment[es]=Genera una ficha técnica del equipo y expórtala a PDF
Exec=$BIN/sheetxpert
Icon=sheetxpert
Terminal=false
Categories=System;Utility;
Keywords=system;hardware;specs;pdf;ficha;
StartupWMClass=SheetXpert
DESKTOP

command -v update-desktop-database >/dev/null 2>&1 && update-desktop-database "$APPS" >/dev/null 2>&1 || true
command -v gtk-update-icon-cache   >/dev/null 2>&1 && gtk-update-icon-cache -q -t -f "$ICONROOT" >/dev/null 2>&1 || true

echo
echo "Done! Launch it from your applications menu, or run:  sheetxpert"
case ":$PATH:" in
    *":$BIN:"*) ;;
    *) echo "Note: $BIN is not in your PATH. Add this to your shell profile:  export PATH=\"$BIN:\$PATH\"" ;;
esac
echo "Command line mode:  sheetxpert --cli | --json | --pdf report.pdf"
