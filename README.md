# SheetXpert 2.0

Ficha técnica del equipo en un clic: recopila información del sistema, la muestra en una interfaz
gráfica (español / inglés) y la exporta a **PDF**, TXT o JSON.
*One-click system technical sheet with a GUI (Spanish / English) and PDF / TXT / JSON export.*

```
SheetXpert/
├── windows/            ← versión para Windows 10/11
│   ├── SheetXpert.py
│   ├── run_windows.bat       (ejecutar desde el código)
│   ├── build_windows.bat     (genera dist\SheetXpert.exe)
│   ├── requirements.txt
│   └── assets/
├── linux/              ← versión para cualquier distribución Linux
│   ├── sheetxpert.py
│   ├── run.sh                (ejecutar sin instalar)
│   ├── install.sh / uninstall.sh
│   ├── build_linux.sh        (binario único con PyInstaller)
│   ├── requirements.txt
│   └── assets/
├── .github/workflows/build.yml   (compila .exe y binario Linux en GitHub)
├── LICENSE
└── README.md
```

---

## Novedades de la versión 2.0

- **Versión Linux completa**: lee `/proc` y `/sys` directamente, así que no depende de ninguna distro
  concreta; solo usa herramientas externas (`ip`, `nmcli`, `lspci`…) si están instaladas.
- **Windows mucho más rápido**: una sola llamada a PowerShell que devuelve JSON, en lugar de ~30 procesos
  (y sin ventanas de consola parpadeando).
- **PDF con Unicode real**: tildes, ñ y símbolos ya no salen como `?` (fuente DejaVu incluida).
  PDF rediseñado con logo, secciones, tabla de discos y pie con número de página.
- **Interfaz mejorada**: informe por secciones, botones *Exportar PDF*, *Guardar TXT/JSON* y *Copiar*,
  cambio de idioma en vivo (también re-traduce el informe ya generado) e idioma inicial según el sistema.
- **Modo consola** (`--cli`, `--json`, `--pdf`) para servidores sin entorno gráfico.
- **Datos nuevos**: kernel, arquitectura, uptime, GPU, núcleos/hilos, placa base, fecha de BIOS,
  salud y ciclos de batería, virtualización, escritorio y sesión (Linux), estado real del cortafuegos.
- **Correcciones**: el firewall de Windows siempre salía «Unknown»; la batería mostraba códigos numéricos;
  la resolución de pantalla salía vacía en muchos equipos; los recursos se buscaban en el directorio
  actual (fallaba al lanzar desde otra carpeta y con mayúsculas/minúsculas en Linux); ya no se usa
  `shell=True`; los datos se recogen en un hilo sin tocar Tk desde fuera del hilo principal.
- Se eliminan los emojis de los botones (en Linux sin fuente emoji se veían como cuadrados).

> El flujo cambia ligeramente: antes se abría el diálogo de guardado nada más recopilar; ahora primero
> ves el informe y luego eliges exportar.

---

## Windows

Requisitos: Windows 10/11 y [Python 3.8+](https://www.python.org/downloads/) (marca *Add python.exe to PATH*).

| Quiero… | Hacer |
|---|---|
| Ejecutarlo | doble clic en `windows\run_windows.bat` (instala `fpdf2` la primera vez) |
| Generar el `.exe` | doble clic en `windows\build_windows.bat` → `windows\dist\SheetXpert.exe` |
| Línea de comandos | `python SheetXpert.py --cli` · `--json` · `--pdf ficha.pdf` · `--lang en` |

El `.exe` se compila con `--windowed` (sin consola), por lo que las opciones `--cli/--json` no muestran
salida en él; úsalas con `python SheetXpert.py`.

## Linux

Requisitos: Python 3.8+ y, para la interfaz gráfica, **tkinter**. Para el PDF: `fpdf2` (lo instalan los scripts).

**Opción A — sin instalar nada:** `cd linux && ./run.sh`
(crea un entorno virtual local `.venv` la primera vez si falta `fpdf2`).

**Opción B — instalar con icono en el menú de aplicaciones:**
```bash
cd linux
./install.sh            # instalación de usuario (~/.local); pide sudo solo para paquetes del sistema
sudo ./install.sh --system   # instalación global (/opt + /usr/local/bin)
sheetxpert              # o búscalo en el menú: «SheetXpert»
./uninstall.sh          # desinstalar
```

`install.sh` detecta el gestor de paquetes e instala Python + tkinter si faltan:

| Distribución | Gestor | Paquetes |
|---|---|---|
| Debian, Ubuntu, Mint, Pop!_OS, Kali… | apt | `python3 python3-venv python3-tk` |
| Fedora, RHEL, Rocky, Alma, CentOS Stream | dnf / yum | `python3 python3-tkinter` |
| Arch, Manjaro, EndeavourOS | pacman | `python tk` |
| openSUSE | zypper | `python3 python3-pip python3-tk` |
| Alpine | apk | `python3 py3-pip python3-tkinter` |
| Void | xbps | `python3 python3-tkinter` |
| Gentoo y otras | — | instala Python con soporte Tk manualmente (`USE=tk`) |

**Opción C — binario único** (sin Python en el equipo de destino): `./build_linux.sh` → `dist/sheetxpert`.
Para máxima compatibilidad compílalo en una distro antigua (Ubuntu 22.04 / Debian 11): funcionará en
esa versión de glibc y en las posteriores. Alternativa: el workflow de GitHub incluido lo compila en
Ubuntu 22.04 y también el `.exe` de Windows (pestaña *Actions → Build binaries → Run workflow*).

Uso en consola (también en servidores sin pantalla, donde pasa a modo texto automáticamente):
```bash
sheetxpert --cli              # ficha en texto
sheetxpert --json             # datos en JSON
sheetxpert --pdf ficha.pdf    # PDF sin abrir ventana
sheetxpert --lang en          # forzar idioma
```

### Notas sobre Linux

- **Números de serie**: el kernel los restringe a root. Verás «(requiere permisos de administrador)».
  Para obtenerlos usa el modo consola: `sudo sheetxpert --cli` (la GUI como root suele fallar en Wayland).
- **Fecha de instalación**: Linux no la guarda; se estima (con `~`) a partir de la fecha de creación de `/`
  o de `machine-id`.
- **RAM**: es la que ve el kernel, algo menos que la instalada (memoria reservada).
- **SSID**: requiere NetworkManager (`nmcli`), `iwgetid` o `iw`; en equipos sin Wi-Fi no aparece.
- **GPU**: requiere `lspci` (paquete `pciutils`); si no existe se muestra el driver del kernel.
- **Cortafuegos**: detecta `firewalld`, `ufw` y `nftables` (sin root no se pueden leer reglas de iptables/nft).
- ARM / Raspberry Pi: se usa el device-tree para modelo y número de serie.

## Datos que recoge

| Sección | Windows | Linux |
|---|---|---|
| General | nombre, SO, versión/build, arquitectura, instalación, último arranque, uptime | nombre, distro, kernel, arquitectura, instalación (aprox.), arranque, uptime, escritorio, sesión, virtualización |
| Hardware | marca, modelo, placa, series, BIOS, CPU, núcleos/hilos, RAM, GPU, pantalla | ídem (DMI / device-tree) |
| Almacenamiento | discos locales | particiones reales montadas (sin tmpfs/snap/loop) |
| Red | IPv4 + interfaz, MAC, SSID | IPv4 + interfaz, MAC de la ruta por defecto, SSID |
| Batería | carga, estado, salud | carga, estado, salud, ciclos |
| Seguridad | Firewall de Windows (3 perfiles) | firewalld / ufw / nftables |

Todo se ejecuta en local; no se envía nada a terceros ni se recogen credenciales.

## Estado de las pruebas

- **Linux**: probado en Ubuntu 24.04 (contenedor): modo consola, JSON, PDF, GUI con pantalla virtual,
  instalador en un HOME limpio y binario PyInstaller. Los lectores de batería, Wi-Fi, IP, MAC, discos,
  cortafuegos y pantalla se probaron con datos simulados de `/sys`, `/proc` y de las herramientas.
  No se ha podido probar en otras distros reales.
- **Windows**: el código común (interfaz, PDF, informe) es el mismo que en Linux y se probó con datos
  simulados de PowerShell, pero **no se ha ejecutado en un Windows real**. Si algún dato sale vacío,
  ejecuta `python SheetXpert.py --json` y revisa qué campo falta.
- No se incluye el `SheetXpert.exe` antiguo (es la versión 1.x): genera uno nuevo con `build_windows.bat`.

## Licencia

Copyright (c) 2025 Alexey. Todos los derechos reservados. Uso personal y educativo; la modificación,
redistribución o uso comercial requiere autorización expresa. Ver [LICENSE](LICENSE).
La fuente DejaVu incluida en `assets/fonts` tiene su propia licencia libre (`LICENSE-DejaVu.txt`).

---

## English summary

**Windows:** `windows\run_windows.bat` (run from source) or `windows\build_windows.bat` (builds `SheetXpert.exe`).
**Linux (any distro):** `cd linux && ./run.sh` (portable) or `./install.sh` (menu entry, auto-installs Python + tkinter
via apt/dnf/yum/pacman/zypper/apk/xbps); `./build_linux.sh` makes a single binary.
**Headless:** `sheetxpert --cli | --json | --pdf file.pdf [--lang es|en]`.

Highlights: distro-independent data collection from `/proc` and `/sys`; single-call PowerShell collector on Windows;
Unicode PDF (DejaVu bundled); live language switch; TXT/JSON export; headless mode. Serial numbers on Linux need
root (`sudo sheetxpert --cli`). Windows code paths were tested with simulated PowerShell output only, not on a real
Windows machine.
