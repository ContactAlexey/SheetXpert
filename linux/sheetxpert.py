# Copyright (c) 2025 Alexey
# All rights reserved.
# Strictly personal and educational use.
# Modification, redistribution, or commercial use without express permission is prohibited.
# See LICENSE file for more information.
#
# SheetXpert 2.0.0 - Linux edition
# Requires Python 3.8+  |  pip install -r requirements.txt

import argparse
import json
import os
import queue
import re
import shutil
import socket
import subprocess
import sys
import threading
import platform
from datetime import datetime

try:  # GUI is optional so the tool also works on headless servers
    import tkinter as tk
    from tkinter import ttk, scrolledtext, messagebox, filedialog
    from tkinter import font as tkfont
    HAS_TK = True
except Exception:  # pragma: no cover
    HAS_TK = False

try:  # PDF export needs fpdf2 (NOT the old "fpdf" package)
    import fpdf as _fpdf_mod
    from fpdf import FPDF
    _v = getattr(_fpdf_mod, "__version__", None) or getattr(_fpdf_mod, "FPDF_VERSION", "0")
    HAS_FPDF2 = int(str(_v).split(".")[0]) >= 2
except Exception:  # pragma: no cover
    FPDF = object
    HAS_FPDF2 = False

# --------------------------------------------------------------------------- #
# Global configuration
# --------------------------------------------------------------------------- #
APP_NAME = "SheetXpert"
APP_VERSION = "2.0.0"
PLATFORM_LABEL = "Linux"
ROOT_REQUIRED = "ROOT_REQUIRED"  # marker for values that need elevated rights

COLOR_MAIN = "#2196F3"
COLOR_MAIN_DARK = "#1976D2"
COLOR_BG = "#F5F7FA"
COLOR_TEXT = "#212121"
COLOR_MUTED = "#757575"

# Values that firmware vendors leave as placeholders
JUNK_VALUES = {
    "", "none", "n/a", "na", "unknown", "not specified", "not applicable", "default string",
    "to be filled by o.e.m.", "to be filled by oem", "system serial number", "system product name",
    "system manufacturer", "system version", "o.e.m.", "oem", "0", "00000000", "0123456789",
    "123456789", "1234567890", "default", "not available", "base board serial number",
    "chassis serial number", "type1productconfigid", "ffffffff-ffff-ffff-ffff-ffffffffffff",
}

UI = {
    "es": {
        "title": "Ficha Técnica del Equipo",
        "get_info": "Obtener ficha técnica",
        "lang_btn": "English",
        "export_pdf": "Exportar PDF",
        "export_txt": "Guardar TXT / JSON",
        "copy": "Copiar",
        "copied": "Ficha copiada al portapapeles.",
        "hint": "Pulsa «Obtener ficha técnica» para analizar este equipo.",
        "collecting": "Recopilando información…",
        "ready": "Listo: {} datos recopilados.",
        "save_pdf_as": "Guardar ficha como…",
        "save_txt_as": "Guardar ficha en texto o JSON…",
        "ok_title": "Éxito",
        "ok_msg": "Archivo guardado en:\n{}",
        "err_title": "Error",
        "err_msg": "No se pudo completar la operación.\n{}",
        "file_prefix": "Ficha",
        "root_required": "(requiere permisos de administrador)",
        "need_fpdf2": "Falta la librería fpdf2. Instálala con:  pip install fpdf2\n"
                      "(si tienes el paquete antiguo «fpdf», desinstálalo antes: pip uninstall fpdf)",
        "generated": "Generado el",
        "yes": "Sí", "no": "No",
        "fw_active": "Activo", "fw_inactive": "Inactivo", "fw_partial": "Parcial", "fw_unknown": "Desconocido",
        "bat_charging": "Cargando", "bat_discharging": "Descargando", "bat_full": "Completa",
        "bat_not_charging": "Conectada (sin cargar)", "bat_unknown": "Desconocido",
        "disk_line": "{mount}  -  {free} GB libres de {total} GB",
        "col_mount": "Montaje", "col_device": "Dispositivo", "col_fs": "Formato",
        "col_total": "Total", "col_free": "Libre", "col_used": "Usado",
        "sec_general": "General", "sec_hardware": "Hardware", "sec_storage": "Almacenamiento",
        "sec_network": "Red", "sec_battery": "Batería", "sec_security": "Seguridad",
        "no_gui": "No hay entorno gráfico disponible: mostrando la ficha en modo texto.",
    },
    "en": {
        "title": "System Technical Sheet",
        "get_info": "Get System Info",
        "lang_btn": "Español",
        "export_pdf": "Export PDF",
        "export_txt": "Save TXT / JSON",
        "copy": "Copy",
        "copied": "Sheet copied to the clipboard.",
        "hint": "Press “Get System Info” to scan this computer.",
        "collecting": "Collecting information…",
        "ready": "Done: {} items collected.",
        "save_pdf_as": "Save sheet as…",
        "save_txt_as": "Save sheet as text or JSON…",
        "ok_title": "Success",
        "ok_msg": "File saved at:\n{}",
        "err_title": "Error",
        "err_msg": "The operation could not be completed.\n{}",
        "file_prefix": "Sheet",
        "root_required": "(requires administrator rights)",
        "need_fpdf2": "The fpdf2 library is missing. Install it with:  pip install fpdf2\n"
                      "(if you have the old “fpdf” package, remove it first: pip uninstall fpdf)",
        "generated": "Generated on",
        "yes": "Yes", "no": "No",
        "fw_active": "Active", "fw_inactive": "Inactive", "fw_partial": "Partial", "fw_unknown": "Unknown",
        "bat_charging": "Charging", "bat_discharging": "Discharging", "bat_full": "Full",
        "bat_not_charging": "Plugged in (not charging)", "bat_unknown": "Unknown",
        "disk_line": "{mount}  -  {free} GB free of {total} GB",
        "col_mount": "Mount", "col_device": "Device", "col_fs": "Format",
        "col_total": "Total", "col_free": "Free", "col_used": "Used",
        "sec_general": "General", "sec_hardware": "Hardware", "sec_storage": "Storage",
        "sec_network": "Network", "sec_battery": "Battery", "sec_security": "Security",
        "no_gui": "No graphical environment available: showing the sheet in text mode.",
    },
}

LABELS = {
    "computer_name": {"es": "Nombre del equipo", "en": "Computer Name"},
    "os": {"es": "Sistema operativo", "en": "Operating System"},
    "os_version": {"es": "Versión del SO", "en": "OS Version"},
    "os_build": {"es": "Build del SO", "en": "OS Build"},
    "kernel": {"es": "Kernel", "en": "Kernel"},
    "arch": {"es": "Arquitectura", "en": "Architecture"},
    "install_date": {"es": "Fecha de instalación SO", "en": "OS Installation Date"},
    "last_boot": {"es": "Último arranque", "en": "Last Boot"},
    "uptime": {"es": "Tiempo encendido", "en": "Uptime"},
    "desktop": {"es": "Entorno de escritorio", "en": "Desktop Environment"},
    "session": {"es": "Tipo de sesión", "en": "Session Type"},
    "virtualization": {"es": "Virtualización", "en": "Virtualization"},
    "brand": {"es": "Marca", "en": "Brand"},
    "model": {"es": "Modelo", "en": "Model"},
    "board": {"es": "Placa base", "en": "Motherboard"},
    "system_serial": {"es": "Número de serie del sistema", "en": "System Serial Number"},
    "board_serial": {"es": "Serie de la placa base", "en": "Motherboard Serial"},
    "bios_serial": {"es": "Serial BIOS", "en": "BIOS Serial"},
    "bios_version": {"es": "Versión BIOS/UEFI", "en": "BIOS/UEFI Version"},
    "bios_date": {"es": "Fecha BIOS/UEFI", "en": "BIOS/UEFI Date"},
    "cpu": {"es": "Procesador (CPU)", "en": "Processor (CPU)"},
    "cpu_cores": {"es": "Núcleos / Hilos", "en": "Cores / Threads"},
    "ram_gb": {"es": "RAM total", "en": "Total RAM"},
    "gpu": {"es": "Gráficos (GPU)", "en": "Graphics (GPU)"},
    "screen": {"es": "Resolución de pantalla", "en": "Screen Resolution"},
    "disks": {"es": "Discos", "en": "Disks"},
    "ipv4": {"es": "IP(s) IPv4", "en": "IPv4 Address(es)"},
    "mac": {"es": "Dirección MAC", "en": "MAC Address"},
    "ssid": {"es": "Red Wi-Fi conectada (SSID)", "en": "Connected Wi-Fi (SSID)"},
    "battery_charge": {"es": "Batería - Carga", "en": "Battery - Charge"},
    "battery_status": {"es": "Batería - Estado", "en": "Battery - Status"},
    "battery_health": {"es": "Batería - Capacidad (salud)", "en": "Battery - Capacity (health)"},
    "battery_cycles": {"es": "Batería - Ciclos de carga", "en": "Battery - Charge Cycles"},
    "firewall": {"es": "Cortafuegos (firewall)", "en": "Firewall"},
}

SECTIONS = [
    ("general", ["computer_name", "os", "os_version", "os_build", "kernel", "arch", "install_date",
                 "last_boot", "uptime", "desktop", "session", "virtualization"]),
    ("hardware", ["brand", "model", "board", "system_serial", "board_serial", "bios_serial",
                  "bios_version", "bios_date", "cpu", "cpu_cores", "ram_gb", "gpu", "screen"]),
    ("storage", ["disks"]),
    ("network", ["ipv4", "mac", "ssid"]),
    ("battery", ["battery_charge", "battery_status", "battery_health", "battery_cycles"]),
    ("security", ["firewall"]),
]


# --------------------------------------------------------------------------- #
# Generic helpers
# --------------------------------------------------------------------------- #
def clean(value):
    """Return a stripped string, or '' for empty / placeholder values."""
    if value is None:
        return ""
    text = str(value).replace("\x00", "").strip()
    return "" if text.lower() in JUNK_VALUES else text


def gb(num_bytes):
    return f"{num_bytes / 1024 ** 3:.2f}"


def fmt_uptime(seconds):
    seconds = max(0, int(seconds))
    days, rest = divmod(seconds, 86400)
    hours, rest = divmod(rest, 3600)
    minutes = rest // 60
    parts = []
    if days:
        parts.append(f"{days}d")
    if days or hours:
        parts.append(f"{hours}h")
    parts.append(f"{minutes}m")
    return " ".join(parts)


def resource_path(relative_path):
    """Absolute path of a bundled resource (works from source and from PyInstaller)."""
    base = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relative_path)


def count_items(data):
    return len(data)


# --------------------------------------------------------------------------- #
# Report building (shared by GUI, text, TXT and PDF)
# --------------------------------------------------------------------------- #
def format_value(key, value, lang):
    ui = UI[lang]
    if value == ROOT_REQUIRED:
        return ui["root_required"]
    if key == "disks":
        lines = []
        for d in value:
            line = ui["disk_line"].format(mount=d["mount"], free=gb(d["free"]), total=gb(d["total"]))
            extra = ", ".join(x for x in (d.get("fs"), d.get("device")) if x)
            lines.append(f"{line}  [{extra}]" if extra else line)
        return lines
    if key == "battery_status":
        return ui.get("bat_" + str(value), str(value))
    if key == "firewall":
        state = ui.get("fw_" + value.get("state", "unknown"), ui["fw_unknown"])
        name = value.get("name")
        return f"{state} ({name})" if name else state
    if key == "battery_charge":
        return f"{value} %"
    if key == "ram_gb":
        return f"{value} GB"
    return str(value)


def build_sections(data, lang):
    """-> [(section_id, section_title, [(key, label, value_or_lines), ...]), ...]"""
    result = []
    for sec_id, keys in SECTIONS:
        rows = []
        for key in keys:
            if key in data and data[key] not in (None, "", []):
                rows.append((key, LABELS[key][lang], format_value(key, data[key], lang)))
        if rows:
            result.append((sec_id, UI[lang]["sec_" + sec_id], rows))
    return result


def text_report(data, lang):
    ui = UI[lang]
    sections = build_sections(data, lang)
    width = max([len(label) for _, _, rows in sections for _, label, _ in rows] or [10])
    title = f"{APP_NAME} - {ui['title']}"
    stamp = f"{data.get('computer_name', '')}  |  {ui['generated']} {datetime.now():%Y-%m-%d %H:%M}"
    lines = [title, "=" * max(len(title), len(stamp)), stamp, ""]
    for _, sec_title, rows in sections:
        lines.append(f"[{sec_title}]")
        for _, label, value in rows:
            if isinstance(value, list):
                lines.append(f"  {label}:")
                lines.extend(f"      {v}" for v in value)
            else:
                lines.append(f"  {label:<{width}} : {value}")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


# --------------------------------------------------------------------------- #
# PDF export (fpdf2) - full Unicode thanks to the bundled DejaVu fonts
# --------------------------------------------------------------------------- #
def find_fonts():
    folders = [
        resource_path(os.path.join("assets", "fonts")),
        "/usr/share/fonts/truetype/dejavu", "/usr/share/fonts/dejavu", "/usr/share/fonts/TTF",
        "/usr/share/fonts/dejavu-sans-fonts", "/usr/share/fonts/truetype/ttf-dejavu",
        r"C:\Windows\Fonts",
    ]
    for folder in folders:
        regular, bold = os.path.join(folder, "DejaVuSans.ttf"), os.path.join(folder, "DejaVuSans-Bold.ttf")
        if os.path.isfile(regular) and os.path.isfile(bold):
            return regular, bold
    return None, None


class ReportPDF(FPDF):
    family = "Helvetica"
    footer_text = ""

    def footer(self):
        self.set_y(-13)
        self.set_font(self.family, "", 8)
        self.set_text_color(130, 130, 130)
        self.cell(0, 6, f"{self.footer_text}   -   {self.page_no()}/{{nb}}", align="C")


def create_pdf_sheet(data, path, lang="es"):
    if not HAS_FPDF2:
        raise RuntimeError(UI[lang]["need_fpdf2"])
    ui = UI[lang]
    regular, bold = find_fonts()
    pdf = ReportPDF(format="A4")
    if regular and bold:
        pdf.add_font("Body", "", regular)
        pdf.add_font("Body", "B", bold)
        pdf.family = "Body"
        T = str
    else:  # core fonts only understand latin-1
        pdf.family = "Helvetica"
        T = lambda s: str(s).encode("latin-1", "replace").decode("latin-1")  # noqa: E731
    fam = pdf.family
    now = datetime.now()
    host = data.get("computer_name", "")
    pdf.footer_text = T(f"{APP_NAME} {APP_VERSION}  |  {host}  |  {now:%Y-%m-%d %H:%M}")
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(True, margin=18)
    pdf.set_title(T(f"{APP_NAME} - {ui['title']}"))
    pdf.set_author(APP_NAME)
    pdf.set_creator(f"{APP_NAME} {APP_VERSION}")
    pdf.add_page()

    # Header: logo + title
    logo = resource_path(os.path.join("assets", "sheetxpert_256.png"))
    if os.path.isfile(logo):
        pdf.image(logo, x=10, y=8, w=24)
    pdf.set_xy(40, 11)
    pdf.set_font(fam, "B", 20)
    pdf.set_text_color(25, 118, 210)
    pdf.cell(0, 10, T(ui["title"]), new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(40)
    pdf.set_font(fam, "", 10)
    pdf.set_text_color(110, 110, 110)
    pdf.cell(0, 6, T(f"{host}  |  {ui['generated']} {now:%Y-%m-%d %H:%M:%S}"), new_x="LMARGIN", new_y="NEXT")
    pdf.set_draw_color(25, 118, 210)
    pdf.set_line_width(0.8)
    pdf.line(10, 36, 200, 36)
    pdf.set_y(40)

    for _, sec_title, rows in build_sections(data, lang):
        if pdf.will_page_break(30):
            pdf.add_page()
        pdf.ln(3)
        pdf.set_font(fam, "B", 12)
        pdf.set_text_color(25, 118, 210)
        pdf.cell(0, 8, T(sec_title), new_x="LMARGIN", new_y="NEXT")
        pdf.set_draw_color(25, 118, 210)
        pdf.set_line_width(0.3)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(2)
        for key, label, value in rows:
            if key == "disks":
                _pdf_disk_table(pdf, fam, T, ui, data["disks"])
                continue
            pdf.set_text_color(33, 33, 33)
            pdf.set_font(fam, "B", 10)
            pdf.cell(62, 6.5, T(label + ":"), new_x="RIGHT", new_y="TOP")
            pdf.set_font(fam, "", 10)
            pdf.multi_cell(0, 6.5, T(value), new_x="LMARGIN", new_y="NEXT")
    pdf.output(path)


def _pdf_disk_table(pdf, fam, T, ui, disks):
    cols = [("col_mount", 30), ("col_device", 56), ("col_fs", 24), ("col_total", 26), ("col_free", 26), ("col_used", 28)]
    cut = lambda s, n: s if len(s) <= n else s[: n - 3] + "..."  # noqa: E731
    pdf.set_draw_color(190, 200, 215)
    pdf.set_line_width(0.2)
    pdf.set_fill_color(232, 240, 254)
    pdf.set_text_color(33, 33, 33)
    pdf.set_font(fam, "B", 9)
    for i, (name, w) in enumerate(cols):
        last = i == len(cols) - 1
        pdf.cell(w, 7, T(ui[name]), border=1, fill=True,
                 new_x="LMARGIN" if last else "RIGHT", new_y="NEXT" if last else "TOP")
    pdf.set_font(fam, "", 9)
    for d in disks:
        used = (d["total"] - d["free"]) / d["total"] * 100 if d["total"] else 0
        values = [cut(d["mount"], 14), cut(d.get("device", ""), 30), cut(d.get("fs", ""), 12),
                  f"{gb(d['total'])} GB", f"{gb(d['free'])} GB", f"{used:.0f} %"]
        for i, ((_, w), val) in enumerate(zip(cols, values)):
            last = i == len(cols) - 1
            pdf.cell(w, 6.5, T(val), border=1,
                     new_x="LMARGIN" if last else "RIGHT", new_y="NEXT" if last else "TOP")
    pdf.ln(1)


# --------------------------------------------------------------------------- #
# Linux data collection
# Works on any distribution: it reads /proc and /sys directly and only uses
# external tools (ip, nmcli, lspci, ...) when they happen to be installed.
# --------------------------------------------------------------------------- #
import glob

_EXTRA_PATH = ["/usr/local/sbin", "/usr/sbin", "/sbin", "/usr/local/bin", "/usr/bin", "/bin"]


def _env():
    env = dict(os.environ)
    env["LC_ALL"] = "C"
    env["LANG"] = "C"
    parts = [p for p in env.get("PATH", "").split(os.pathsep) if p]
    parts += [p for p in _EXTRA_PATH if p not in parts]
    env["PATH"] = os.pathsep.join(parts)
    return env


def _run(cmd, timeout=6):
    """Run an external command (no shell). Returns stdout or '' if unavailable/failing."""
    env = _env()
    exe = shutil.which(cmd[0], path=env["PATH"])
    if not exe:
        return ""
    try:
        res = subprocess.run([exe] + list(cmd[1:]), capture_output=True, text=True, errors="replace",
                             timeout=timeout, env=env, stdin=subprocess.DEVNULL)
        return res.stdout.strip()
    except Exception:
        return ""


def _read(path):
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            return fh.read().strip()
    except OSError:
        return ""


def _dmi(name):
    """Read /sys/class/dmi/id/<name>. Returns ROOT_REQUIRED when the file exists but is root-only."""
    path = "/sys/class/dmi/id/" + name
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            return clean(fh.read())
    except PermissionError:
        return ROOT_REQUIRED
    except OSError:
        return ""


def _os_release():
    info = {}
    for path in ("/etc/os-release", "/usr/lib/os-release"):
        text = _read(path)
        if text:
            for line in text.splitlines():
                if "=" in line and not line.lstrip().startswith("#"):
                    k, v = line.split("=", 1)
                    info[k.strip()] = v.strip().strip('"').strip("'")
            break
    return info


def _cpu_name():
    first = {}
    for line in _read("/proc/cpuinfo").splitlines():
        k, _, v = line.partition(":")
        first.setdefault(k.strip(), v.strip())
    name = first.get("model name", "")
    if not name:  # ARM, RISC-V, s390x, ...
        found = []
        for line in _run(["lscpu"]).splitlines():
            k, _, v = line.partition(":")
            if k.strip() == "Model name" and v.strip() and v.strip() not in found:
                found.append(v.strip())
        name = " + ".join(found)
    if not name:
        name = first.get("Hardware") or first.get("cpu model") or first.get("cpu") or platform.processor()
    return re.sub(r"\s+", " ", name or "").strip()


def _cpu_cores():
    threads = os.cpu_count() or 0
    seen = set()
    for topo in glob.glob("/sys/devices/system/cpu/cpu[0-9]*/topology"):
        core = _read(topo + "/core_id")
        if core != "":
            seen.add((_read(topo + "/physical_package_id"), core))
    if seen and threads:
        return f"{len(seen)} / {threads}"
    return str(threads) if threads else ""


def _ram_gb():
    for line in _read("/proc/meminfo").splitlines():
        if line.startswith("MemTotal:"):
            try:
                return f"{int(line.split()[1]) / 1024 ** 2:.2f}"
            except (ValueError, IndexError):
                break
    return ""


_REAL_FS = {"ext2", "ext3", "ext4", "xfs", "btrfs", "f2fs", "jfs", "reiserfs", "vfat", "exfat", "ntfs",
            "ntfs3", "fuseblk", "zfs", "bcachefs", "hfsplus", "ufs", "nilfs2", "erofs"}


def _disks():
    disks, seen = [], {}
    for line in _read("/proc/self/mounts").splitlines():
        parts = line.split()
        if len(parts) < 3:
            continue
        source, mount, fstype = parts[0], parts[1], parts[2]
        mount = re.sub(r"\\([0-7]{3})", lambda m: chr(int(m.group(1), 8)), mount)
        if fstype not in _REAL_FS or not (source.startswith("/dev/") or fstype == "zfs"):
            continue
        if mount.startswith(("/snap", "/boot/efi", "/var/lib/docker", "/var/snap")):
            continue
        try:
            st = os.statvfs(mount)
        except OSError:
            continue
        total, free = st.f_blocks * st.f_frsize, st.f_bavail * st.f_frsize
        if total <= 0:
            continue
        entry = {"mount": mount, "device": source, "fs": fstype, "total": total, "free": free}
        if source in seen:  # bind mounts / btrfs subvolumes: keep the shortest mount point
            idx = seen[source]
            if len(mount) < len(disks[idx]["mount"]):
                disks[idx] = entry
            continue
        seen[source] = len(disks)
        disks.append(entry)
    if not disks:
        try:
            usage = shutil.disk_usage("/")
            disks.append({"mount": "/", "device": "", "fs": "", "total": usage.total, "free": usage.free})
        except OSError:
            pass
    return disks


def _default_iface():
    best = None
    for line in _read("/proc/net/route").splitlines()[1:]:
        f = line.split()
        try:
            if len(f) >= 8 and f[1] == "00000000" and int(f[7], 16) == 0:
                metric = int(f[6])
                if best is None or metric < best[0]:
                    best = (metric, f[0])
        except ValueError:
            continue
    return best[1] if best else ""


def _ipv4():
    found = []
    for line in _run(["ip", "-4", "-o", "addr", "show", "scope", "global"]).splitlines():
        m = re.search(r"^\d+:\s+(\S+)\s+inet\s+(\d+\.\d+\.\d+\.\d+)/", line)
        if m and not m.group(2).startswith("169.254."):
            found.append(f"{m.group(2)} ({m.group(1).rstrip(':')})")
    if not found:
        for ip in _run(["hostname", "-I"]).split():
            if re.match(r"^\d+\.\d+\.\d+\.\d+$", ip) and not ip.startswith(("127.", "169.254.")):
                found.append(ip)
    if not found:  # pure-Python fallback: no packet is actually sent
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("10.255.255.255", 1))
            found.append(s.getsockname()[0])
            s.close()
        except OSError:
            pass
    return ", ".join(x for x in found if not x.startswith("0.0.0.0"))


def _ssid():
    for line in _run(["nmcli", "-t", "-f", "ACTIVE,SSID", "dev", "wifi"]).splitlines():
        if line.startswith("yes:"):
            return line[4:].replace("\\:", ":").strip()
    ssid = _run(["iwgetid", "-r"])
    if ssid:
        return ssid
    for line in _run(["iw", "dev"]).splitlines():
        line = line.strip()
        if line.startswith("ssid "):
            return line[5:].strip()
    return ""


def _mac():
    zero = "00:00:00:00:00:00"
    default = _default_iface()
    names = [default] if default else []
    try:
        names += sorted(n for n in os.listdir("/sys/class/net") if n not in names)
    except OSError:
        pass
    fallback = ""
    for name in names:
        if name == "lo":
            continue
        base = "/sys/class/net/" + name
        addr = _read(base + "/address")
        if not addr or addr == zero:
            continue
        if name == default or "/virtual/" not in os.path.realpath(base):
            return addr
        if not fallback and _read(base + "/operstate") == "up":
            fallback = addr
    return fallback


def _screen(tk_fallback):
    found = []
    for line in _run(["xrandr", "--current"]).splitlines():
        m = re.match(r"^\s+(\d+)x(\d+)\s.*\*", line)
        if m:
            found.append(f"{m.group(1)} x {m.group(2)}")
    if not found:
        for status in sorted(glob.glob("/sys/class/drm/card*-*/status")):
            if _read(status) == "connected":
                mode = _read(os.path.dirname(status) + "/modes").splitlines()[:1]
                m = re.match(r"^(\d+)x(\d+)", mode[0]) if mode else None
                if m:
                    found.append(f"{m.group(1)} x {m.group(2)}")
    if not found and tk_fallback:
        found.append(f"{tk_fallback[0]} x {tk_fallback[1]}")
    return ", ".join(found)


def _gpu():
    found = []
    for line in _run(["lspci"]).splitlines():
        if re.search(r"VGA compatible controller|3D controller|Display controller", line):
            found.append(line.split(": ", 1)[-1].strip())
    if not found:
        for card in sorted(glob.glob("/sys/class/drm/card[0-9]")):
            driver = os.path.basename(os.path.realpath(card + "/device/driver"))
            if driver and driver != "driver":
                found.append(f"driver: {driver}")
    return " | ".join(found)


def _battery():
    for ps in sorted(glob.glob("/sys/class/power_supply/*")):
        if _read(ps + "/type") != "Battery" or _read(ps + "/scope") == "Device":
            continue
        res = {}
        cap = _read(ps + "/capacity")
        if cap.isdigit():
            res["battery_charge"] = cap
        status = _read(ps + "/status").lower().replace(" ", "_")
        if status:
            res["battery_status"] = status if status in ("charging", "discharging", "full", "not_charging") else "unknown"
        for prefix in ("energy", "charge"):
            full, design = _read(f"{ps}/{prefix}_full"), _read(f"{ps}/{prefix}_full_design")
            if full.isdigit() and design.isdigit() and int(design) > 0:
                full, design = int(full), int(design)
                text = f"{full / design * 100:.1f} %"
                if prefix == "energy":
                    text += f" ({full / 1e6:.1f} / {design / 1e6:.1f} Wh)"
                else:
                    volt = _read(ps + "/voltage_min_design")
                    if volt.isdigit():
                        text += f" ({full * int(volt) / 1e12:.1f} / {design * int(volt) / 1e12:.1f} Wh)"
                res["battery_health"] = text
                break
        cycles = _read(ps + "/cycle_count")
        if cycles.isdigit() and int(cycles) > 0:
            res["battery_cycles"] = cycles
        if res:
            return res
    return {}


def _install_date():
    ts = 0
    out = _run(["stat", "-c", "%W", "/"])
    if out.isdigit() and int(out) > 0:
        ts = int(out)
    if not ts:
        for path in ("/var/log/installer", "/var/log/anaconda", "/etc/machine-id", "/var/lib/dbus/machine-id"):
            try:
                ts = int(os.stat(path).st_mtime)
                break
            except OSError:
                continue
    return ("~" + datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")) if ts else ""


def _firewall():
    def active(unit):
        return _run(["systemctl", "is-active", unit]) == "active"

    ufw_conf = _read("/etc/ufw/ufw.conf")
    ufw_known = bool(ufw_conf) or bool(shutil.which("ufw", path=_env()["PATH"]))
    fwd_known = bool(shutil.which("firewall-cmd", path=_env()["PATH"]))
    if fwd_known and active("firewalld"):
        return {"state": "active", "name": "firewalld"}
    if ufw_known and (re.search(r"^ENABLED=yes", ufw_conf, re.M) and active("ufw")):
        return {"state": "active", "name": "ufw"}
    if ufw_known or fwd_known:
        names = [n for n, known in (("ufw", ufw_known), ("firewalld", fwd_known)) if known]
        return {"state": "inactive", "name": "/".join(names)}
    if active("nftables"):
        return {"state": "active", "name": "nftables"}
    return {"state": "unknown"}


def collect_system_data(screen_fallback=None):
    data = {}

    def put(key, value):
        if value not in (None, "", []):
            data[key] = value

    osr = _os_release()
    put("computer_name", socket.gethostname())
    put("os", osr.get("PRETTY_NAME") or " ".join(x for x in (osr.get("NAME"), osr.get("VERSION")) if x) or platform.system())
    put("kernel", platform.release())
    put("arch", platform.machine())
    put("install_date", _install_date())

    btime = next((int(l.split()[1]) for l in _read("/proc/stat").splitlines() if l.startswith("btime ")), 0)
    uptime = _read("/proc/uptime").split()
    if btime:
        put("last_boot", datetime.fromtimestamp(btime).strftime("%Y-%m-%d %H:%M:%S"))
    if uptime:
        try:
            put("uptime", fmt_uptime(float(uptime[0])))
        except ValueError:
            pass
    put("desktop", os.environ.get("XDG_CURRENT_DESKTOP") or os.environ.get("DESKTOP_SESSION"))
    put("session", os.environ.get("XDG_SESSION_TYPE"))
    virt = _run(["systemd-detect-virt"])
    if not virt and re.search(r"\bhypervisor\b", _read("/proc/cpuinfo")):
        virt = "yes"
    if virt and virt != "none":
        put("virtualization", virt)

    # Hardware (DMI on x86/UEFI, device-tree on ARM boards such as Raspberry Pi)
    dt_model = clean(_read("/sys/firmware/devicetree/base/model"))
    vendor, product, version = _dmi("sys_vendor"), _dmi("product_name"), _dmi("product_version")
    if vendor.lower().startswith("lenovo") and version and version != ROOT_REQUIRED:
        product = f"{version} ({product})" if product else version
    put("brand", vendor)
    put("model", product or dt_model)
    put("board", " ".join(x for x in (_dmi("board_vendor"), _dmi("board_name")) if x and x != ROOT_REQUIRED))
    serial = _dmi("product_serial")
    if not serial or serial == ROOT_REQUIRED:
        alt = clean(_read("/sys/firmware/devicetree/base/serial-number"))
        if not alt:
            alt = next((clean(l.split(":", 1)[1]) for l in _read("/proc/cpuinfo").splitlines() if l.startswith("Serial")), "")
        serial = alt or serial
    put("system_serial", serial)
    put("board_serial", _dmi("board_serial"))
    bios_ver, bios_vendor = _dmi("bios_version"), _dmi("bios_vendor")
    put("bios_version", f"{bios_ver} ({bios_vendor})" if bios_ver and bios_vendor else bios_ver)
    put("bios_date", _dmi("bios_date"))
    put("cpu", _cpu_name())
    put("cpu_cores", _cpu_cores())
    put("ram_gb", _ram_gb())
    put("gpu", _gpu())
    put("screen", _screen(screen_fallback))

    put("disks", _disks())
    put("ipv4", _ipv4())
    put("mac", _mac())
    put("ssid", _ssid())
    for key, value in _battery().items():
        put(key, value)
    put("firewall", _firewall())
    return data


def detect_language():
    for var in ("LC_ALL", "LC_MESSAGES", "LANGUAGE", "LANG"):
        value = os.environ.get(var, "")
        if value:
            return "es" if value.lower().startswith("es") else "en"
    return "en"


def gui_available():
    return HAS_TK and bool(os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY"))


def gui_prepare():
    pass


def platform_window_setup(window):
    pass


# --------------------------------------------------------------------------- #
# Graphical interface (Tkinter)
# --------------------------------------------------------------------------- #
class SheetXpertApp:
    def __init__(self, lang):
        gui_prepare()
        self.lang = lang
        self.data = None
        self.results = queue.Queue()
        self.root = tk.Tk(className=APP_NAME)
        self.root.withdraw()
        self.root.title(APP_NAME)
        self._set_icon(self.root)
        self._pick_fonts()
        self._build_ui()
        self._apply_language()
        self.root.after(50, self._splash)

    # ---- setup ---------------------------------------------------------- #
    def _set_icon(self, window):
        try:
            self._icon = tk.PhotoImage(file=resource_path(os.path.join("assets", "sheetxpert_256.png")))
            window.iconphoto(True, self._icon)
        except Exception:
            pass
        platform_window_setup(window)

    def _pick_fonts(self):
        families = set(tkfont.families(self.root))

        def pick(candidates, named):
            for name in candidates:
                if name in families:
                    return name
            return tkfont.nametofont(named).actual("family")

        self.f_ui = pick(["Segoe UI", "Noto Sans", "Ubuntu", "Cantarell", "Open Sans", "DejaVu Sans",
                          "Liberation Sans", "Helvetica"], "TkDefaultFont")
        self.f_mono = pick(["Consolas", "DejaVu Sans Mono", "Liberation Mono", "Noto Sans Mono", "Ubuntu Mono",
                            "Courier New"], "TkFixedFont")

    def _build_ui(self):
        r = self.root
        r.configure(bg=COLOR_BG)
        r.minsize(640, 520)
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Primary.TButton", foreground="white", background=COLOR_MAIN,
                        font=(self.f_ui, 12, "bold"), padding=(14, 9), borderwidth=0)
        style.map("Primary.TButton", background=[("disabled", "#90A4AE"), ("active", COLOR_MAIN_DARK)],
                  foreground=[("disabled", "#ECEFF1")])
        style.configure("Secondary.TButton", foreground=COLOR_MAIN_DARK, background="#E3F2FD",
                        font=(self.f_ui, 10), padding=(10, 6), borderwidth=0)
        style.map("Secondary.TButton", background=[("disabled", "#ECEFF1"), ("active", "#BBDEFB")],
                  foreground=[("disabled", "#9E9E9E")])
        r.grid_columnconfigure(0, weight=1)
        r.grid_rowconfigure(3, weight=1)

        top = tk.Frame(r, bg=COLOR_BG)
        top.grid(row=0, column=0, sticky="ew", padx=20, pady=(16, 6))
        top.grid_columnconfigure(0, weight=1)
        self.header = tk.Label(top, bg=COLOR_BG, fg=COLOR_MAIN, font=(self.f_ui, 16, "bold"), anchor="w")
        self.header.grid(row=0, column=0, sticky="w")
        tk.Label(top, text=f"{APP_NAME} {APP_VERSION}  ·  {PLATFORM_LABEL}", bg=COLOR_BG, fg=COLOR_MUTED,
                 font=(self.f_ui, 9), anchor="w").grid(row=1, column=0, sticky="w")
        self.btn_lang = ttk.Button(top, style="Secondary.TButton", command=self._toggle_language)
        self.btn_lang.grid(row=0, column=1, rowspan=2, sticky="e")

        self.btn_get = ttk.Button(r, style="Primary.TButton", command=self._collect)
        self.btn_get.grid(row=1, column=0, sticky="ew", padx=20, pady=(8, 6))

        self.progress = ttk.Progressbar(r, mode="indeterminate")
        self.progress.grid(row=2, column=0, sticky="ew", padx=20, pady=2)
        self.progress.grid_remove()

        self.text = scrolledtext.ScrolledText(r, wrap="word", font=(self.f_ui, 10), bg="white", fg=COLOR_TEXT,
                                              bd=1, relief="solid", padx=14, pady=10, state="disabled", height=18,
                                              spacing1=2, spacing3=2)
        self.text.grid(row=3, column=0, sticky="nsew", padx=20, pady=(6, 8))
        self.text.tag_configure("title", font=(self.f_ui, 13, "bold"), foreground=COLOR_MAIN, spacing3=4)
        self.text.tag_configure("section", font=(self.f_ui, 11, "bold"), foreground=COLOR_MAIN_DARK,
                                spacing1=12, spacing3=3)
        self.text.tag_configure("label", font=(self.f_ui, 10, "bold"))
        self.text.tag_configure("mono", font=(self.f_mono, 9), lmargin1=24, lmargin2=24)
        self.text.tag_configure("hint", foreground=COLOR_MUTED, font=(self.f_ui, 10, "italic"))

        bar = tk.Frame(r, bg=COLOR_BG)
        bar.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 4))
        self.btn_pdf = ttk.Button(bar, style="Secondary.TButton", command=self._export_pdf, state="disabled")
        self.btn_txt = ttk.Button(bar, style="Secondary.TButton", command=self._export_text, state="disabled")
        self.btn_copy = ttk.Button(bar, style="Secondary.TButton", command=self._copy, state="disabled")
        for b in (self.btn_pdf, self.btn_txt, self.btn_copy):
            b.pack(side="left", padx=(0, 8))

        self.status = tk.StringVar()
        tk.Label(r, textvariable=self.status, bg=COLOR_BG, fg=COLOR_MUTED, font=(self.f_ui, 9), anchor="w") \
            .grid(row=5, column=0, sticky="ew", padx=20, pady=(0, 10))

    # ---- window helpers ------------------------------------------------- #
    def _center(self, win, width, height):
        x = max(0, (win.winfo_screenwidth() - width) // 2)
        y = max(0, (win.winfo_screenheight() - height) // 3)
        win.geometry(f"{width}x{height}+{x}+{y}")

    def _splash(self):
        path = resource_path(os.path.join("assets", "sheetxpert_splash.png"))
        try:
            splash = tk.Toplevel(self.root)
            splash.overrideredirect(True)
            splash.configure(bg="white")
            self._splash_img = tk.PhotoImage(file=path)
            tk.Label(splash, image=self._splash_img, bg="white", bd=0).pack()
            self._center(splash, self._splash_img.width(), self._splash_img.height())
            try:
                splash.attributes("-topmost", True)
            except tk.TclError:
                pass
            splash.after(1600, lambda: (splash.destroy(), self._show_main()))
        except Exception:
            self._show_main()

    def _show_main(self):
        self._center(self.root, 780, 680)
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    # ---- language ------------------------------------------------------- #
    def _toggle_language(self):
        self.lang = "en" if self.lang == "es" else "es"
        self._apply_language()

    def _apply_language(self):
        ui = UI[self.lang]
        self.header.config(text=ui["title"])
        self.btn_lang.config(text=ui["lang_btn"])
        self.btn_get.config(text=ui["get_info"])
        self.btn_pdf.config(text=ui["export_pdf"])
        self.btn_txt.config(text=ui["export_txt"])
        self.btn_copy.config(text=ui["copy"])
        if self.data:
            self._render()
            self.status.set(ui["ready"].format(count_items(self.data)))
        else:
            self._show_hint()

    # ---- text area ------------------------------------------------------ #
    def _show_hint(self):
        self.text.config(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("end", "\n" + UI[self.lang]["hint"], "hint")
        self.text.config(state="disabled")

    def _render(self):
        t = self.text
        t.config(state="normal")
        t.delete("1.0", "end")
        t.insert("end", f"{self.data.get('computer_name', '')}\n", "title")
        for _, sec_title, rows in build_sections(self.data, self.lang):
            t.insert("end", sec_title + "\n", "section")
            for _, label, value in rows:
                if isinstance(value, list):
                    t.insert("end", label + ":\n", "label")
                    for line in value:
                        t.insert("end", line + "\n", "mono")
                else:
                    t.insert("end", label + ": ", "label")
                    t.insert("end", value + "\n")
        t.config(state="disabled")
        t.yview_moveto(0)

    # ---- collection (worker thread) ------------------------------------- #
    def _collect(self):
        self.btn_get.config(state="disabled")
        for b in (self.btn_pdf, self.btn_txt, self.btn_copy):
            b.config(state="disabled")
        self.progress.grid()
        self.progress.start(12)
        self.status.set(UI[self.lang]["collecting"])
        fallback = (self.root.winfo_screenwidth(), self.root.winfo_screenheight())

        def work():
            try:
                self.results.put(("ok", collect_system_data(fallback)))
            except Exception as exc:  # noqa: BLE001
                self.results.put(("err", exc))

        threading.Thread(target=work, daemon=True).start()
        self.root.after(100, self._poll)

    def _poll(self):
        try:
            kind, payload = self.results.get_nowait()
        except queue.Empty:
            self.root.after(100, self._poll)
            return
        self.progress.stop()
        self.progress.grid_remove()
        self.btn_get.config(state="normal")
        ui = UI[self.lang]
        if kind == "err":
            self.status.set("")
            messagebox.showerror(ui["err_title"], ui["err_msg"].format(payload), parent=self.root)
            return
        self.data = payload
        self._render()
        for b in (self.btn_pdf, self.btn_txt, self.btn_copy):
            b.config(state="normal")
        self.status.set(ui["ready"].format(count_items(payload)))

    # ---- export --------------------------------------------------------- #
    def _suggested_name(self, ext):
        host = re.sub(r"[^\w.-]", "_", self.data.get("computer_name", "PC"))
        return f"{UI[self.lang]['file_prefix']}_{host}_{datetime.now():%Y%m%d_%H%M%S}{ext}"

    def _export_pdf(self):
        ui = UI[self.lang]
        path = filedialog.asksaveasfilename(parent=self.root, defaultextension=".pdf", title=ui["save_pdf_as"],
                                            filetypes=[("PDF", "*.pdf")], initialfile=self._suggested_name(".pdf"))
        if not path:
            return
        try:
            create_pdf_sheet(self.data, path, self.lang)
            messagebox.showinfo(ui["ok_title"], ui["ok_msg"].format(path), parent=self.root)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(ui["err_title"], ui["err_msg"].format(exc), parent=self.root)

    def _export_text(self):
        ui = UI[self.lang]
        path = filedialog.asksaveasfilename(parent=self.root, defaultextension=".txt", title=ui["save_txt_as"],
                                            filetypes=[("Text", "*.txt"), ("JSON", "*.json")],
                                            initialfile=self._suggested_name(".txt"))
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as fh:
                if path.lower().endswith(".json"):
                    json.dump(self.data, fh, ensure_ascii=False, indent=2)
                else:
                    fh.write(text_report(self.data, self.lang))
            messagebox.showinfo(ui["ok_title"], ui["ok_msg"].format(path), parent=self.root)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(ui["err_title"], ui["err_msg"].format(exc), parent=self.root)

    def _copy(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(text_report(self.data, self.lang))
        self.root.update()  # keep the clipboard after the window closes on most X11 setups
        self.status.set(UI[self.lang]["copied"])

    def run(self):
        self.root.mainloop()


# --------------------------------------------------------------------------- #
# Entry point (GUI by default, command line for servers / scripts)
# --------------------------------------------------------------------------- #
def main(argv=None):
    parser = argparse.ArgumentParser(prog=APP_NAME.lower(), description=f"{APP_NAME} - system technical sheet")
    parser.add_argument("--cli", action="store_true", help="print the sheet as text (no GUI)")
    parser.add_argument("--json", action="store_true", help="print the raw data as JSON (no GUI)")
    parser.add_argument("--pdf", metavar="FILE", help="write the sheet to a PDF file (no GUI)")
    parser.add_argument("--lang", choices=("es", "en"), help="language (default: system language)")
    parser.add_argument("--version", action="version", version=f"{APP_NAME} {APP_VERSION} ({PLATFORM_LABEL})")
    args = parser.parse_args(argv)
    lang = args.lang or detect_language()

    headless = args.cli or args.json or args.pdf
    if not headless and not gui_available():
        print(UI[lang]["no_gui"], file=sys.stderr)
        headless = True
    if headless:
        data = collect_system_data()
        if args.json:
            print(json.dumps(data, ensure_ascii=False, indent=2))
        if args.pdf:
            try:
                create_pdf_sheet(data, args.pdf, lang)
            except Exception as exc:  # noqa: BLE001
                print(UI[lang]["err_msg"].format(exc), file=sys.stderr)
                return 1
            print(UI[lang]["ok_msg"].format(args.pdf), file=sys.stderr if args.json else sys.stdout)
        if args.cli or not (args.json or args.pdf):
            print(text_report(data, lang), end="")
        return 0

    try:
        app = SheetXpertApp(lang)
    except tk.TclError as exc:  # e.g. DISPLAY set but unreachable
        print(f"{UI[lang]['no_gui']} ({exc})", file=sys.stderr)
        print(text_report(collect_system_data(), lang), end="")
        return 0
    app.run()
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except BrokenPipeError:  # e.g. "sheetxpert --cli | head"
        try:
            sys.stdout.close()
        except Exception:
            pass
        sys.exit(0)
