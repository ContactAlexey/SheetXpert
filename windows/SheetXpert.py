# Copyright (c) 2025 Alexey
# All rights reserved.
# Strictly personal and educational use.
# Modification, redistribution, or commercial use without express permission is prohibited.
# See LICENSE file for more information.
#
# SheetXpert 2.0.0 - Windows edition
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
PLATFORM_LABEL = "Windows"
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
# Windows data collection
# ONE PowerShell process collects everything and returns JSON (the previous
# version spawned ~30 processes, which was slow and flashed console windows).
# --------------------------------------------------------------------------- #
import base64
import ctypes

PS_SCRIPT = r"""
$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
function Fmt($d) { if ($d) { $d.ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture) } else { '' } }

$os    = Get-CimInstance Win32_OperatingSystem
$cs    = Get-CimInstance Win32_ComputerSystem
$bios  = Get-CimInstance Win32_BIOS
$csp   = Get-CimInstance Win32_ComputerSystemProduct
$bb    = Get-CimInstance Win32_BaseBoard
$cpus  = @(Get-CimInstance Win32_Processor)
$ram   = (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum
if (-not $ram) { $ram = $cs.TotalPhysicalMemory }

$disks = @(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' | ForEach-Object {
    [pscustomobject]@{ mount = $_.DeviceID; device = $_.VolumeName; fs = $_.FileSystem; total = [double]$_.Size; free = [double]$_.FreeSpace }
})
$ips = @(Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike '169.254.*' -and $_.IPAddress -ne '127.0.0.1' } | ForEach-Object { "$($_.IPAddress) ($($_.InterfaceAlias))" })
$nic = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' -and $_.HardwareInterface } | Select-Object -First 1

$ssid = ''
foreach ($line in @(netsh wlan show interfaces)) {
    if ($line -match '^\s*SSID\s*:\s*(.+)$') { $ssid = $Matches[1].Trim(); break }
}

$bat   = Get-CimInstance Win32_Battery | Select-Object -First 1
$bfull = (Get-CimInstance -Namespace root/WMI -ClassName BatteryFullChargedCapacity | Select-Object -First 1).FullChargedCapacity
$bdes  = (Get-CimInstance -Namespace root/WMI -ClassName BatteryStaticData | Select-Object -First 1).DesignedCapacity

$vc  = @(Get-CimInstance Win32_VideoController)
$gpu = @($vc | ForEach-Object { $_.Name })
$scr = $vc | Where-Object { $_.CurrentHorizontalResolution } | Select-Object -First 1
$fw  = @(Get-NetFirewallProfile | ForEach-Object { "$($_.Enabled)" })

$out = [ordered]@{
    computer_name = $env:COMPUTERNAME
    os            = $os.Caption
    os_version    = $os.Version
    os_build      = "$($os.BuildNumber)"
    arch          = $os.OSArchitecture
    install_date  = (Fmt $os.InstallDate)
    last_boot     = (Fmt $os.LastBootUpTime)
    brand         = $cs.Manufacturer
    model         = $cs.Model
    board         = (("$($bb.Manufacturer) $($bb.Product)").Trim())
    system_serial = $csp.IdentifyingNumber
    board_serial  = $bb.SerialNumber
    bios_serial   = $bios.SerialNumber
    bios_version  = $bios.SMBIOSBIOSVersion
    cpu           = (($cpus | ForEach-Object { "$($_.Name)".Trim() } | Select-Object -Unique) -join ' + ')
    cpu_cores     = ($cpus | Measure-Object -Property NumberOfCores -Sum).Sum
    cpu_threads   = ($cpus | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
    ram_bytes     = $ram
    gpu           = ($gpu -join ' | ')
    screen_w      = $scr.CurrentHorizontalResolution
    screen_h      = $scr.CurrentVerticalResolution
    disks         = $disks
    ipv4          = $ips
    mac           = $nic.MacAddress
    ssid          = $ssid
    battery_present = [bool]$bat
    battery_charge  = $bat.EstimatedChargeRemaining
    battery_status  = $bat.BatteryStatus
    battery_full_mwh   = $bfull
    battery_design_mwh = $bdes
    firewall      = $fw
}
$out | ConvertTo-Json -Depth 5 -Compress
"""


def _powershell_json(timeout=120):
    """Run the collection script and return a dict ({} on failure)."""
    encoded = base64.b64encode(PS_SCRIPT.encode("utf-16-le")).decode("ascii")
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    for exe in ("powershell", "pwsh"):
        try:
            res = subprocess.run(
                [exe, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded],
                capture_output=True, timeout=timeout, creationflags=flags, stdin=subprocess.DEVNULL)
        except (OSError, subprocess.TimeoutExpired):
            continue
        out = res.stdout.decode("utf-8-sig", errors="replace")
        start = out.find("{")
        if start >= 0:
            try:
                return json.loads(out[start:])
            except ValueError:
                continue
    return {}


def _as_list(value):
    if value is None:
        return []
    return value if isinstance(value, list) else [value]


def _num(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


_BATTERY_STATUS = {1: "discharging", 2: "not_charging", 3: "full", 4: "discharging", 5: "discharging",
                   6: "charging", 7: "charging", 8: "charging", 9: "charging", 10: "unknown", 11: "not_charging"}


def collect_system_data(screen_fallback=None):
    raw = _powershell_json()
    data = {}

    def put(key, value):
        if value not in (None, "", []):
            data[key] = value

    if not raw:  # PowerShell unavailable: minimal fallback so the tool still produces something
        put("computer_name", socket.gethostname())
        put("os", platform.platform())
        put("arch", platform.machine())
        put("cpu", platform.processor())
        if screen_fallback:
            put("screen", f"{screen_fallback[0]} x {screen_fallback[1]}")
        try:
            usage = shutil.disk_usage(os.environ.get("SystemDrive", "C:") + "\\")
            put("disks", [{"mount": os.environ.get("SystemDrive", "C:"), "device": "", "fs": "",
                           "total": usage.total, "free": usage.free}])
        except OSError:
            pass
        return data

    for key in ("computer_name", "os", "os_version", "os_build", "arch", "install_date", "last_boot",
                "brand", "model", "board", "bios_version", "cpu", "gpu", "mac", "ssid"):
        put(key, clean(raw.get(key)))
    for key in ("system_serial", "board_serial", "bios_serial"):
        put(key, clean(raw.get(key)))
    if data.get("bios_serial") and data.get("bios_serial") == data.get("system_serial"):
        del data["bios_serial"]  # avoid showing the same serial twice

    boot = data.get("last_boot")
    if boot:
        try:
            put("uptime", fmt_uptime((datetime.now() - datetime.strptime(boot, "%Y-%m-%d %H:%M:%S")).total_seconds()))
        except ValueError:
            pass

    cores, threads = int(_num(raw.get("cpu_cores"))), int(_num(raw.get("cpu_threads")))
    if cores and threads:
        put("cpu_cores", f"{cores} / {threads}")
    elif threads:
        put("cpu_cores", str(threads))
    if _num(raw.get("ram_bytes")) > 0:
        put("ram_gb", gb(_num(raw["ram_bytes"])))

    width, height = int(_num(raw.get("screen_w"))), int(_num(raw.get("screen_h")))
    if width and height:
        put("screen", f"{width} x {height}")
    elif screen_fallback:
        put("screen", f"{screen_fallback[0]} x {screen_fallback[1]}")

    disks = []
    for d in _as_list(raw.get("disks")):
        if isinstance(d, dict) and _num(d.get("total")) > 0:
            disks.append({"mount": d.get("mount", ""), "device": clean(d.get("device")), "fs": clean(d.get("fs")),
                          "total": _num(d["total"]), "free": _num(d.get("free"))})
    put("disks", disks)
    put("ipv4", ", ".join(str(x) for x in _as_list(raw.get("ipv4"))))

    if raw.get("battery_present"):
        if raw.get("battery_charge") is not None:
            put("battery_charge", str(int(_num(raw["battery_charge"]))))
        put("battery_status", _BATTERY_STATUS.get(int(_num(raw.get("battery_status"))), "unknown"))
        full, design = _num(raw.get("battery_full_mwh")), _num(raw.get("battery_design_mwh"))
        if full and design:
            put("battery_health", f"{full / design * 100:.1f} % ({full / 1000:.1f} / {design / 1000:.1f} Wh)")
        elif full:
            put("battery_health", f"{full / 1000:.1f} Wh")

    states = [str(x).lower() in ("true", "1") for x in _as_list(raw.get("firewall"))]
    if states:
        state = "active" if all(states) else "inactive" if not any(states) else "partial"
        put("firewall", {"state": state, "name": "Windows Defender Firewall"})
    else:
        put("firewall", {"state": "unknown"})
    return data


def detect_language():
    try:
        lang_id = ctypes.windll.kernel32.GetUserDefaultUILanguage()
        return "es" if (lang_id & 0x3FF) == 0x0A else "en"
    except Exception:
        value = (os.environ.get("LANG") or "").lower()
        return "es" if value.startswith("es") else "en"


def gui_available():
    return HAS_TK


def gui_prepare():
    """Make Tk sharp on high-DPI displays."""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def platform_window_setup(window):
    try:
        window.iconbitmap(resource_path(os.path.join("assets", "sheetxpert.ico")))
    except Exception:
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
