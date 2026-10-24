#!/usr/bin/env python3
"""Genera las terminales animadas (SVG) que se muestran en el README.

Uso:   python3 docs/tools/make_terminals.py

Cada SVG es una simulación autónoma (solo CSS, sin JavaScript), así que GitHub la anima
dentro de un <img>. Los textos y los valores son ilustrativos: edita los guiones de abajo
y vuelve a ejecutar el script para regenerar las animaciones.
"""
import html
import os
import re

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "terminals")

# --------------------------------------------------------------------------- #
# Geometría y colores
# --------------------------------------------------------------------------- #
WIDTH = 840
CW = 7.8          # ancho de carácter (se fuerza con textLength para que sea exacto)
LH = 20           # alto de línea
PAD_X = 20
BAR_H = 38        # barra de título
HOLD = 3.6        # pausa final antes de repetir

THEMES = {
    "linux": dict(bg="#161b22", bar="#21262d", border="#30363d", title="#8b949e", fg="#e6edf3"),
    "windows": dict(bg="#0c0c0c", bar="#202020", border="#3a3a3a", title="#d0d0d0", fg="#cccccc"),
}
COLORS = {
    "g": "#3fb950", "b": "#58a6ff", "y": "#e3b341", "r": "#ff7b72",
    "c": "#56d4dd", "d": "#7d8590", "w": "#ffffff", "m": "#d2a8ff",
}

# --------------------------------------------------------------------------- #
# Utilidades
# --------------------------------------------------------------------------- #
TAG = re.compile(r"\{(\w)\}(.*?)\{/\}")


def parse(markup):
    """'texto {g}verde{/} texto' -> [(texto, None), ('verde', 'g'), ...]"""
    parts, pos = [], 0
    for m in TAG.finditer(markup):
        if m.start() > pos:
            parts.append((markup[pos:m.start()], None))
        parts.append((m.group(2), m.group(1)))
        pos = m.end()
    if pos < len(markup):
        parts.append((markup[pos:], None))
    return parts


def plain(markup):
    return "".join(t for t, _ in parse(markup))


def esc(s):
    return html.escape(s, quote=False)


class Anim:
    """Acumula @keyframes y los reparte por porcentaje de la duración total."""

    def __init__(self):
        self.css = []
        self.n = 0
        self.total = 1.0

    def pct(self, t):
        return max(0.0, min(100.0, t / self.total * 100.0))

    def window(self, t0, t1=None):
        """Clase CSS que hace visible el elemento entre t0 y t1 (t1=None: hasta el final)."""
        self.n += 1
        name = f"w{self.n}"
        frames = f"0%{{opacity:0}}{self.pct(t0):.3f}%{{opacity:1}}"
        if t1 is not None:
            frames += f"{self.pct(t1):.3f}%{{opacity:0}}"
        self.css.append(f"@keyframes {name}{{{frames}}}.{name}{{opacity:0;animation:{name} {self.total:.2f}s step-end infinite}}")
        return name

    def slide(self, t0, t1, dist, steps):
        """Desplaza el elemento 'dist' px hacia la derecha entre t0 y t1 en 'steps' saltos."""
        self.n += 1
        name = f"s{self.n}"
        frames = (f"0%{{transform:translateX(0)}}"
                  f"{self.pct(t0):.3f}%{{transform:translateX(0);animation-timing-function:steps({steps},end)}}"
                  f"{self.pct(t1):.3f}%{{transform:translateX({dist:.1f}px)}}"
                  f"100%{{transform:translateX({dist:.1f}px)}}")
        self.css.append(f"@keyframes {name}{{{frames}}}.{name}{{animation:{name} {self.total:.2f}s linear infinite}}")
        return name

    def grow(self, t0, t1, steps):
        """Escala horizontal 0 -> 1 (barra de progreso)."""
        self.n += 1
        name = f"p{self.n}"
        frames = (f"0%{{transform:scaleX(0)}}"
                  f"{self.pct(t0):.3f}%{{transform:scaleX(0);animation-timing-function:steps({steps},end)}}"
                  f"{self.pct(t1):.3f}%{{transform:scaleX(1)}}"
                  f"100%{{transform:scaleX(1)}}")
        self.css.append(f"@keyframes {name}{{{frames}}}.{name}{{transform-box:fill-box;transform-origin:0 50%;"
                        f"animation:{name} {self.total:.2f}s linear infinite}}")
        return name


def text_el(x, y, markup, fg, cls=""):
    """<text> con tspans de color; textLength fija el ancho exacto de la cuadrícula."""
    parts = parse(markup)
    n = sum(len(t) for t, _ in parts)
    if n == 0:
        return ""
    spans = "".join(
        f'<tspan fill="{COLORS[c]}">{esc(t)}</tspan>' if c else esc(t) for t, c in parts
    )
    c = f' class="{cls}"' if cls else ""
    return (f'<text{c} x="{x}" y="{y}" fill="{fg}" textLength="{n * CW:.1f}" '
            f'lengthAdjust="spacing" xml:space="preserve">{spans}</text>')


# --------------------------------------------------------------------------- #
# Construcción de un SVG a partir de un guion
# --------------------------------------------------------------------------- #
def build(name, theme, title, prompts, script):
    th = THEMES[theme]

    # 1) Línea temporal: asignamos a cada elemento su momento de aparición.
    t = 0.7
    lines = []          # (tipo, datos...)
    prompt = prompts["start"]
    for it in script:
        kind = it[0]
        if kind == "prompt":                      # cambia el prompt activo
            prompt = prompts[it[1]]
        elif kind == "cmd":
            _, text, secs = it
            ts = t
            type0 = ts + 0.35
            type1 = type0 + secs
            lines.append(dict(kind="cmd", prompt=prompt, text=text, ts=ts, t0=type0, t1=type1))
            t = type1 + 0.7
        elif kind == "out":
            lines.append(dict(kind="out", text=it[1], ts=t))
            t += it[2] if len(it) > 2 else 0.13
        elif kind == "bar":
            _, label, secs, suffix = it
            lines.append(dict(kind="bar", label=label, suffix=suffix, ts=t, t1=t + secs))
            t += secs + 0.25
        elif kind == "wait":
            t += it[1]
        elif kind == "idle":                      # prompt final con cursor parpadeando
            lines.append(dict(kind="cmd", prompt=prompt, text="", ts=t, t0=t, t1=t, idle=True))
            t += 0.1
    total = t + HOLD

    # 2) Cursor: visible desde que aparece la línea hasta que aparece la siguiente.
    cmds = [i for i, l in enumerate(lines) if l["kind"] == "cmd"]
    for i in cmds:
        nxt = next((lines[j]["ts"] for j in range(i + 1, len(lines))), None)
        lines[i]["hide"] = None if lines[i].get("idle") else nxt

    height = BAR_H + 22 + len(lines) * LH + 18
    an = Anim()
    an.total = total

    body = []
    for i, l in enumerate(lines):
        y = BAR_H + 22 + i * LH + 13
        x = PAD_X
        cls = an.window(l["ts"])
        if l["kind"] == "out":
            body.append(f'<g class="{cls}">{text_el(x, y, l["text"], th["fg"])}</g>')

        elif l["kind"] == "cmd":
            pl = len(plain(l["prompt"]))
            cx = x + pl * CW
            n = len(l["text"])
            g = [text_el(x, y, l["prompt"], th["fg"])]
            if n:
                g.append(text_el(cx, y, "{w}" + l["text"] + "{/}", th["fg"]))
                # tapa del color del fondo que se desplaza y va "escribiendo" el comando
                slide = an.slide(l["t0"], l["t1"], n * CW, n)
                g.append(f'<g class="{slide}"><rect x="{cx:.1f}" y="{y - 13}" width="{n * CW + CW:.1f}" '
                         f'height="{LH}" fill="{th["bg"]}"/></g>')
            # cursor
            cur_win = an.window(l["ts"], l["hide"])
            cursor = f'<rect class="blink" x="{cx:.1f}" y="{y - 12}" width="{CW:.1f}" height="16" fill="{th["fg"]}"/>'
            if n:
                cursor = f'<g class="{an.slide(l["t0"], l["t1"], n * CW, n)}">{cursor}</g>'
            g.append(f'<g class="{cur_win}">{cursor}</g>')
            body.append(f'<g class="{cls}">{"".join(g)}</g>')

        elif l["kind"] == "bar":
            label = l["label"]
            bx = x + (len(plain(label)) + 1) * CW
            bw = 220
            steps = 20
            fill = an.grow(l["ts"] + 0.15, l["t1"], steps)
            g = [text_el(x, y, label, th["fg"]),
                 f'<rect x="{bx:.1f}" y="{y - 9}" width="{bw}" height="8" rx="2" fill="#30363d"/>',
                 f'<g class="{fill}"><rect x="{bx:.1f}" y="{y - 9}" width="{bw}" height="8" rx="2" fill="{COLORS["g"]}"/></g>']
            # porcentaje a saltos
            span = l["t1"] - (l["ts"] + 0.15)
            for k, p in enumerate((0, 20, 40, 60, 80, 100)):
                a = l["ts"] + 0.15 + span * k / 6
                b = l["ts"] + 0.15 + span * (k + 1) / 6 if p < 100 else None
                pc = an.window(a, b)
                g.append(f'<g class="{pc}">{text_el(bx + bw + CW * 2, y, f"{p}%", th["fg"])}</g>')
            if l["suffix"]:
                sc = an.window(l["t1"])
                g.append(f'<g class="{sc}">{text_el(bx + bw + CW * 8, y, "{d}" + l["suffix"] + "{/}", th["fg"])}</g>')
            body.append(f'<g class="{cls}">{"".join(g)}</g>')

    # 3) Barra de título
    if theme == "linux":
        chrome = (
            f'<circle cx="22" cy="19" r="6" fill="#ff5f56"/><circle cx="42" cy="19" r="6" fill="#ffbd2e"/>'
            f'<circle cx="62" cy="19" r="6" fill="#27c93f"/>'
            f'<text x="{WIDTH / 2}" y="23" text-anchor="middle" fill="{th["title"]}" class="ui">{esc(title)}</text>'
        )
    else:
        chrome = (
            f'<text x="16" y="23" fill="{th["title"]}" class="ui">&gt;_  {esc(title)}</text>'
            f'<path d="M{WIDTH - 126} 19h12" stroke="{th["title"]}" stroke-width="1.2"/>'
            f'<rect x="{WIDTH - 82}" y="13" width="11" height="11" fill="none" stroke="{th["title"]}" stroke-width="1.2"/>'
            f'<path d="M{WIDTH - 36} 13l11 11M{WIDTH - 25} 13l-11 11" stroke="{th["title"]}" stroke-width="1.2"/>'
        )

    css = (
        "text{font-family:'Cascadia Mono','DejaVu Sans Mono','Menlo',Consolas,'Liberation Mono',monospace;font-size:13px}"
        ".ui{font-family:'Segoe UI',system-ui,Helvetica,Arial,sans-serif;font-size:12px}"
        "@keyframes blink{0%,55%{opacity:1}56%,100%{opacity:.05}}"
        ".blink{animation:blink 1.1s step-end infinite}"
        + "".join(an.css)
    )
    svg = f'''<svg xmlns="http://www.w3.org/2000/svg" width="{WIDTH}" height="{height}" viewBox="0 0 {WIDTH} {height}" role="img" aria-label="{esc(title)}">
<title>{esc(title)}</title>
<style>{css}</style>
<rect x="0.5" y="0.5" width="{WIDTH - 1}" height="{height - 1}" rx="10" fill="{th["bg"]}" stroke="{th["border"]}"/>
<path d="M0.5 {BAR_H}V11a10.5 10.5 0 0 1 10.5-10.5h{WIDTH - 22}A10.5 10.5 0 0 1 {WIDTH - 0.5} 11V{BAR_H}z" fill="{th["bar"]}"/>
{chrome}
{"".join(body)}
</svg>
'''
    os.makedirs(OUT, exist_ok=True)
    path = os.path.join(OUT, name + ".svg")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(svg)
    return path, len(svg), total


# --------------------------------------------------------------------------- #
# Prompts
# --------------------------------------------------------------------------- #
LINUX = {
    "start": "{g}usuario@equipo{/}:{b}~/Descargas{/}$ ",
    "linux": "{g}usuario@equipo{/}:{b}~/Descargas/SheetXpert/linux{/}$ ",
    "home": "{g}usuario@equipo{/}:{b}~{/}$ ",
}
WIN = {
    "start": r"C:\Users\usuario\Downloads>",
    "win": r"C:\Users\usuario\Downloads\SheetXpert\windows>",
}

# --------------------------------------------------------------------------- #
# Guiones
# --------------------------------------------------------------------------- #
CLI_ES_LINUX = [
    ("out", "SheetXpert - Ficha Técnica del Equipo", 0.08),
    ("out", "=====================================", 0.08),
    ("out", "portatil  |  Generado el 2026-10-24 19:05", 0.08),
    ("out", "", 0.08),
    ("out", "{c}[General]{/}", 0.08),
    ("out", "  Nombre del equipo       : portatil", 0.08),
    ("out", "  Sistema operativo       : Ubuntu 24.04.4 LTS", 0.08),
    ("out", "  Kernel                  : 6.8.0-48-generic", 0.08),
    ("out", "  Arquitectura            : x86_64", 0.08),
    ("out", "  Tiempo encendido        : 3h 12m", 0.08),
    ("out", "", 0.08),
    ("out", "{c}[Hardware]{/}", 0.08),
    ("out", "  Procesador (CPU)        : Intel(R) Core(TM) i5-1135G7 @ 2.40GHz", 0.08),
    ("out", "  Núcleos / Hilos         : 4 / 8", 0.08),
    ("out", "  RAM total               : 15.35 GB", 0.08),
    ("out", "", 0.08),
    ("out", "{c}[Seguridad]{/}", 0.08),
    ("out", "  Cortafuegos (firewall)  : ufw (activo)", 0.08),
]

SCRIPTS = {
    # ------------------------------------------------------------------ Windows
    "windows-run": ("windows", "Símbolo del sistema", WIN, [
        ("cmd", r"cd SheetXpert\windows", 1.0),
        ("prompt", "win"),
        ("cmd", "python --version", 1.0),
        ("out", "Python 3.12.4", 0.4),
        ("cmd", "run_windows.bat", 1.0),
        ("out", "Installing dependencies - internet needed once...", 0.5),
        ("out", "Collecting fpdf2>=2.7.0", 0.3),
        ("bar", "  Downloading fpdf2-2.8.3-py3-none-any.whl", 1.5, "250.0/250.0 kB"),
        ("out", "Collecting defusedxml, Pillow, fonttools (dependencias de fpdf2)", 0.5),
        ("out", "Installing collected packages: defusedxml, Pillow, fonttools, fpdf2", 0.9),
        ("out", "{g}Successfully installed Pillow-11.0.0 defusedxml-0.7.1 fonttools-4.55.0 fpdf2-2.8.3{/}", 0.6),
        ("out", "{d}→ Se abre la ventana de SheetXpert (la consola queda libre){/}", 0.3),
        ("idle",),
    ]),
    "windows-build": ("windows", "Símbolo del sistema", WIN, [
        ("prompt", "win"),
        ("cmd", "build_windows.bat", 1.0),
        ("out", r"Requirement already satisfied: fpdf2>=2.7.0 in c:\users\usuario\…\site-packages (2.8.3)", 0.4),
        ("bar", "Collecting pyinstaller", 1.4, ""),
        ("out", "{g}Successfully installed pyinstaller-6.11.1 altgraph-0.17.4 pefile-2023.2.7{/}", 0.7),
        ("out", "{d}120 INFO:{/} PyInstaller: 6.11.1, contrib hooks: 2024.10", 0.25),
        ("out", "{d}121 INFO:{/} Python: 3.12.4", 0.25),
        ("out", "{d}340 INFO:{/} Appending 'datas' from .spec / --add-data assets;assets", 0.5),
        ("out", "{d}2815 INFO:{/} Building EXE from EXE-00.toc", 0.8),
        ("out", "{d}6420 INFO:{/} Building EXE from EXE-00.toc completed successfully.", 0.5),
        ("out", "", 0.1),
        ("out", r"{g}Done: C:\Users\usuario\Downloads\SheetXpert\windows\dist\SheetXpert.exe{/}", 0.4),
        ("out", "Presione una tecla para continuar . . .", 0.3),
        ("idle",),
    ]),
    "windows-cli": ("windows", "Símbolo del sistema", WIN, [
        ("prompt", "win"),
        ("cmd", "python SheetXpert.py --cli", 1.3),
        ("out", "SheetXpert - Ficha Técnica del Equipo", 0.08),
        ("out", "=====================================", 0.08),
        ("out", "PC-OFICINA  |  Generado el 2026-10-24 19:05", 0.08),
        ("out", "", 0.08),
        ("out", "{c}[General]{/}", 0.08),
        ("out", "  Nombre del equipo       : PC-OFICINA", 0.08),
        ("out", "  Sistema operativo       : Microsoft Windows 11 Pro", 0.08),
        ("out", "  Arquitectura            : x64", 0.08),
        ("out", "", 0.08),
        ("out", "{c}[Hardware]{/}", 0.08),
        ("out", "  Procesador (CPU)        : Intel(R) Core(TM) i7-1165G7 @ 2.80GHz", 0.08),
        ("out", "  RAM total               : 16.00 GB", 0.5),
        ("out", "{d}  […]{/}", 0.5),
        ("cmd", "python SheetXpert.py --pdf ficha.pdf --lang en", 1.9),
        ("out", "File saved at:", 0.15),
        ("out", r"{g}ficha.pdf{/}", 0.3),
        ("idle",),
    ]),
    # -------------------------------------------------------------------- Linux
    "linux-run": ("linux", "usuario@equipo: ~/Descargas", LINUX, [
        ("cmd", "unzip SheetXpert.zip", 1.1),
        ("out", "Archive:  SheetXpert.zip", 0.1),
        ("out", "   creating: SheetXpert/", 0.1),
        ("out", "  inflating: SheetXpert/linux/run.sh", 0.1),
        ("out", "{d}  […]{/}", 0.4),
        ("cmd", "cd SheetXpert/linux", 0.9),
        ("prompt", "linux"),
        ("cmd", "python3 -c 'import tkinter' && echo 'tkinter OK'", 2.0),
        ("out", "{g}tkinter OK{/}", 0.4),
        ("cmd", "./run.sh", 0.7),
        ("out", "First run: creating a private Python environment (internet needed once)...", 0.2),
        ("wait", 2.6),
        ("out", "{d}→ Se abre la ventana de SheetXpert{/}", 0.3),
        ("idle",),
    ]),
    "linux-install": ("linux", "usuario@equipo: ~/Descargas/SheetXpert/linux", LINUX, [
        ("prompt", "linux"),
        ("cmd", "./install.sh", 0.8),
        ("out", "==> Installing system packages (python3, venv, tkinter)...", 0.4),
        ("out", "[sudo] contraseña para usuario: ", 1.4),
        ("out", "Leyendo lista de paquetes... Hecho", 0.2),
        ("out", "Creando árbol de dependencias... Hecho", 0.2),
        ("out", "Se instalarán los siguientes paquetes NUEVOS:  python3-tk python3-venv", 0.5),
        ("out", "Configurando python3-tk (3.12.3-0ubuntu1) ...", 0.9),
        ("out", "==> Installing SheetXpert to /home/usuario/.local/share/sheetxpert", 0.4),
        ("out", "==> Creating virtual environment and installing fpdf2 (internet needed)...", 0.2),
        ("wait", 2.4),
        ("out", "", 0.1),
        ("out", "{g}Done!{/} Launch it from your applications menu, or run:  {w}sheetxpert{/}", 0.3),
        ("out", "Command line mode:  sheetxpert --cli | --json | --pdf report.pdf", 0.4),
        ("cmd", "sheetxpert", 0.9),
        ("out", "{d}→ Se abre la ventana de SheetXpert{/}", 0.3),
        ("idle",),
    ]),
    "linux-install-system": ("linux", "usuario@equipo: ~/Descargas/SheetXpert/linux", LINUX, [
        ("prompt", "linux"),
        ("cmd", "sudo ./install.sh --system", 1.6),
        ("out", "[sudo] contraseña para usuario: ", 1.4),
        ("out", "==> Installing SheetXpert to /opt/sheetxpert", 0.4),
        ("out", "==> Creating virtual environment and installing fpdf2 (internet needed)...", 0.2),
        ("wait", 2.2),
        ("out", "", 0.1),
        ("out", "{g}Done!{/} Launch it from your applications menu, or run:  {w}sheetxpert{/}", 0.3),
        ("out", "Command line mode:  sheetxpert --cli | --json | --pdf report.pdf", 0.4),
        ("cmd", "which sheetxpert", 0.9),
        ("out", "/usr/local/bin/sheetxpert", 0.3),
        ("idle",),
    ]),
    "linux-uninstall": ("linux", "usuario@equipo: ~/Descargas/SheetXpert/linux", LINUX, [
        ("prompt", "linux"),
        ("cmd", "./uninstall.sh", 0.9),
        ("out", "{g}SheetXpert removed.{/}", 0.5),
        ("cmd", "sudo ./uninstall.sh --system", 1.7),
        ("out", "[sudo] contraseña para usuario: ", 1.2),
        ("out", "{g}SheetXpert removed.{/}", 0.4),
        ("idle",),
    ]),
    "linux-build": ("linux", "usuario@equipo: ~/Descargas/SheetXpert/linux", LINUX, [
        ("prompt", "linux"),
        ("cmd", "./build_linux.sh", 0.9),
        ("wait", 1.6),
        ("out", "{d}112 INFO:{/} PyInstaller: 6.11.1, contrib hooks: 2024.10", 0.25),
        ("out", "{d}113 INFO:{/} Python: 3.12.3", 0.25),
        ("out", "{d}331 INFO:{/} Appending 'datas' from .spec / --add-data assets:assets", 0.5),
        ("out", "{d}2690 INFO:{/} Building EXE from EXE-00.toc", 0.8),
        ("out", "{d}5934 INFO:{/} Building EXE from EXE-00.toc completed successfully.", 0.5),
        ("out", "", 0.1),
        ("out", "{g}Done: /home/usuario/Descargas/SheetXpert/linux/dist/sheetxpert{/}", 0.4),
        ("cmd", "./dist/sheetxpert --version", 1.5),
        ("out", "SheetXpert 2.0.0 (Linux)", 0.3),
        ("idle",),
    ]),
    "linux-cli": ("linux", "usuario@equipo: ~", LINUX, [
        ("prompt", "home"),
        ("cmd", "sheetxpert --cli", 1.1),
        *CLI_ES_LINUX,
        ("wait", 0.5),
        ("cmd", "sheetxpert --pdf ficha.pdf", 1.6),
        ("out", "Archivo guardado en:", 0.15),
        ("out", "{g}ficha.pdf{/}", 0.3),
        ("idle",),
    ]),
}

if __name__ == "__main__":
    for name, (theme, title, prompts, script) in SCRIPTS.items():
        path, size, total = build(name, theme, title, prompts, script)
        print(f"{os.path.relpath(path)}  {size / 1024:.1f} KB  {total:.1f} s")
