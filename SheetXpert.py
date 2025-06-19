# Copyright (c) 2025 Alexey
# All rights reserved.
# Strictly personal and educational use.
# Modification, redistribution, or commercial use without express permission is prohibited.
# See LICENSE file for more information.

# Import necessary libraries
import subprocess  # For running shell commands
import tkinter as tk  # For creating the GUI
from tkinter import scrolledtext, messagebox, filedialog, ttk  # Various tkinter components
from fpdf import FPDF  # For creating PDF files
import threading  # For running tasks in separate threads
from datetime import datetime  # For handling date and time
import os  # For interacting with the operating system
import sys  # For system-specific parameters and functions
from PIL import Image, ImageTk  # For image processing

# Global configuration
NOMBRE_MARCA = "SheetXpert"  # Brand name
ICONO_NOMBRE = f"{NOMBRE_MARCA}.ico"  # Icon file name
IMAGEN_SPLASH = f"{NOMBRE_MARCA}.png"  # Splash image file name
idioma_actual = "es"  # Current language set to Spanish
traducciones = {  # Translations for UI elements in Spanish and English
    "es": {
        "titulo": "Ficha Técnica del Equipo",
        "obtener_ficha": "📄 Obtener ficha técnica",
        "idioma_btn": "🌐 Inglés",
        "guardar_como": "Guardar ficha como...",
        "exito_titulo": "Éxito",
        "exito_msj": "PDF guardado en:\n{}",
        "error_titulo": "Error",
        "error_msj": "No se pudo guardar el PDF.\n{}",
        "nombre_archivo_prefijo": "Ficha"
    },
    "en": {
        "titulo": "System Technical Sheet",
        "obtener_ficha": "📄 Get System Info",
        "idioma_btn": "🌐 Español",
        "guardar_como": "Save Sheet as...",
        "exito_titulo": "Success",
        "exito_msj": "PDF saved at:\n{}",
        "error_titulo": "Error",
        "error_msj": "Could not save PDF.\n{}",
        "nombre_archivo_prefijo": "Sheet"
    }
}

# Executes a PowerShell command and returns the output
def run_powershell(cmd):
    try:
        result = subprocess.run(["powershell", "-Command", cmd], capture_output=True, text=True, shell=True)
        return result.stdout.strip()  # Return the output of the command
    except Exception as e:
        return f"Error: {e}"  # Return error message if command fails

# Formats WMI date to a readable format
def format_date_wmi(wmi_date):
    try:
        if wmi_date and len(wmi_date) >= 14:
            date = wmi_date[:14]  # Extract the date part
            dt = datetime.strptime(date, "%Y%m%d%H%M%S")  # Parse the date
            return dt.strftime("%Y-%m-%d %H:%M:%S")  # Return formatted date
    except Exception:
        pass
    return ""  # Return empty string if formatting fails

# Collects various data about the PC using PowerShell commands
def get_pc_data():
    data = {}
    data['Computer Name'] = run_powershell('hostname')  # Get computer name
    data['IPv4 Address(es)'] = run_powershell("(Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike '169.*' -and $_.IPAddress -ne '127.0.0.1' }).IPAddress -join ', '")  # Get IPv4 addresses
    data['Connected Network (SSID)'] = run_powershell('(netsh wlan show interfaces) -match "^\\s*SSID\\s*:\\s*(.+)$" | ForEach-Object { ($_ -split ":")[1].Trim() } | Select-Object -First 1')  # Get connected network SSID
    data['Operating System'] = run_powershell("(Get-CimInstance Win32_OperatingSystem).Caption")  # Get OS name
    ram_bytes = run_powershell("(Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum")  # Get total RAM in bytes
    try:
        ram_gb = round(int(ram_bytes) / (1024**3), 2)  # Convert bytes to GB
        data['Total RAM (GB)'] = str(ram_gb)  # Store total RAM
    except ValueError:
        data['Total RAM (GB)'] = "N/A"  # Handle conversion error
    disks_raw = run_powershell(
        'Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | '
        'Select-Object DeviceID, @{Name="SizeGB";Expression={[math]::Round($_.Size/1GB,2)}}, '
        '@{Name="FreeGB";Expression={[math]::Round($_.FreeSpace/1GB,2)}} | '
        'Format-Table -HideTableHeaders | Out-String')  # Get disk information
    disks = []
    for line in disks_raw.strip().splitlines():  # Process each line of disk info
        parts = line.split()
        if len(parts) == 3:
            disks.append((parts[0], parts[2], parts[1]))  # Append disk info as tuple
    if disks:
        data['Disks'] = disks  # Store disk information
    data['Brand'] = run_powershell("(Get-CimInstance Win32_ComputerSystem).Manufacturer")  # Get computer brand
    data['Model'] = run_powershell("(Get-CimInstance Win32_ComputerSystem).Model")  # Get computer model
    data['BIOS Serial'] = run_powershell("(Get-CimInstance Win32_BIOS).SerialNumber")  # Get BIOS serial number
    data['MAC'] = run_powershell("(Get-NetAdapter | Where-Object {$_.Status -eq 'Up'} | Select-Object -First 1).MacAddress")  # Get MAC address
    data['BIOS Version'] = run_powershell("(Get-CimInstance Win32_BIOS).SMBIOSBIOSVersion")  # Get BIOS version
    data['Processor (CPU)'] = run_powershell("(Get-CimInstance Win32_Processor).Name")  # Get CPU name
    data['System Serial Number'] = run_powershell("(Get-CimInstance Win32_ComputerSystemProduct).IdentifyingNumber")  # Get system serial number
    width = run_powershell("(Get-CimInstance Win32_DesktopMonitor | Select-Object -First 1).ScreenWidth")  # Get screen width
    height = run_powershell("(Get-CimInstance Win32_DesktopMonitor | Select-Object -First 1).ScreenHeight")  # Get screen height
    if width and height and width.isdigit() and height.isdigit():
        data['Screen Resolution'] = f"{width} x {height}"  # Store screen resolution
    data['Battery - Charge (%)'] = run_powershell("(Get-WmiObject -Class Win32_Battery | Select-Object -ExpandProperty EstimatedChargeRemaining)")  # Get battery charge percentage
    data['Battery - Status'] = run_powershell("(Get-WmiObject -Class Win32_Battery | Select-Object -ExpandProperty BatteryStatus)")  # Get battery status
    data['Battery - Capacity'] = run_powershell("(Get-WmiObject -Namespace root\\WMI -Class BatteryFullChargedCapacity | Select-Object -ExpandProperty FullChargedCapacity)")  # Get battery capacity
    data['OS Installation Date'] = format_date_wmi(run_powershell("(Get-CimInstance Win32_OperatingSystem).InstallDate"))  # Get OS installation date
    data['Last Boot'] = format_date_wmi(run_powershell("(Get-CimInstance Win32_OperatingSystem).LastBootUpTime"))  # Get last boot time
    data['Windows Version'] = run_powershell("(Get-CimInstance Win32_OperatingSystem).Version")  # Get Windows version
    data['Windows Build'] = run_powershell("(Get-CimInstance Win32_OperatingSystem).BuildNumber")  # Get Windows build number
    fw_status = run_powershell("(Get-NetFirewallProfile | Select-Object -First 1).Enabled")  # Check if firewall is enabled
    data['Firewall Enabled'] = "Yes" if fw_status == '1' else "No" if fw_status == '0' else "Unknown"  # Store firewall status
    return {k: v for k, v in data.items() if v and v != "N/A"}  # Return only valid data

# Formats the collected data into a readable sheet
def format_sheet(data):
    # Titles of keys in Spanish and English
    titles = {
        "Computer Name": {"es": "Nombre del equipo", "en": "Computer Name"},
        "IPv4 Address(es)": {"es": "IP(s) IPv4", "en": "IPv4 Address(es)"},
        "Connected Network (SSID)": {"es": "Red conectada (SSID)", "en": "Connected Network (SSID)"},
        "Operating System": {"es": "Sistema operativo", "en": "Operating System"},
        "Total RAM (GB)": {"es": "RAM total (GB)", "en": "Total RAM (GB)"},
        "Disks": {"es": "Discos", "en": "Disks"},
        "Brand": {"es": "Marca", "en": "Brand"},
        "Model": {"es": "Modelo", "en": "Model"},
        "BIOS Serial": {"es": "Serial BIOS", "en": "BIOS Serial"},
        "MAC": {"es": "MAC", "en": "MAC Address"},
        "BIOS Version": {"es": "Versión BIOS", "en": "BIOS Version"},
        "Processor (CPU)": {"es": "Procesador (CPU)", "en": "Processor (CPU)"},
        "System Serial Number": {"es": "Número de serie del sistema", "en": "System Serial Number"},
        "Screen Resolution": {"es": "Resolución pantalla", "en": "Screen Resolution"},
        "Battery - Charge (%)": {"es": "Batería - Carga (%)", "en": "Battery - Charge (%)"},
        "Battery - Status": {"es": "Batería - Estado", "en": "Battery - Status"},
        "Battery - Capacity": {"es": "Batería - Capacidad", "en": "Battery - Capacity"},
        "OS Installation Date": {"es": "Fecha de instalación SO", "en": "OS Installation Date"},
        "Last Boot": {"es": "Último arranque", "en": "Last Boot"},
        "Windows Version": {"es": "Versión Windows", "en": "Windows Version"},
        "Windows Build": {"es": "Build Windows", "en": "Windows Build"},
        "Firewall Enabled": {"es": "Firewall Activado", "en": "Firewall Enabled"},
    }

    lines = [traducciones[idioma_actual]["titulo"] + "\n" + "="*30 + "\n"]  # Add title and separator
    for key, value in data.items():  # Iterate through data
        if key == "Disks":
            lines.append(titles[key][idioma_actual] + ":")  # Add disks title
            for d in value:
                lines.append(f"  {d[0]}: {d[1]} GB free of {d[2]} GB")  # Add disk details
        else:
            lines.append(f"{titles[key][idioma_actual]}: {value}")  # Add other details
    return "\n".join(lines)  # Return formatted string

# Creates a PDF sheet from the collected data
def create_pdf_sheet(data, file):
    pdf = FPDF()  # Create a PDF object
    pdf.add_page()  # Add a new page
    pdf.set_auto_page_break(auto=True, margin=15)  # Set auto page break
    pdf.set_fill_color(25, 118, 210)  # Set fill color for title
    pdf.set_text_color(255, 255, 255)  # Set text color for title
    pdf.set_font("Arial", 'B', 18)  # Set font for title
    
    # Add title and handle encoding
    title = traducciones[idioma_actual]["titulo"].encode('latin-1', 'replace').decode('latin-1')  # Encode title
    pdf.cell(0, 15, title, 0, 1, 'C', fill=True)  # Add title to PDF
    
    pdf.ln(5)  # Add a line break
    pdf.set_text_color(33, 33, 33)  # Set text color for content
    pdf.set_font("Arial", 'B', 12)  # Set font for content
    pdf.set_draw_color(25, 118, 210)  # Set line color
    pdf.set_line_width(0.8)  # Set line width
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())  # Draw a line
    pdf.ln(5)  # Add a line break

    # Use translated titles when generating pdf
    titles = {
        "Computer Name": {"es": "Nombre del equipo", "en": "Computer Name"},
        "IPv4 Address(es)": {"es": "IP(s) IPv4", "en": "IPv4 Address(es)"},
        "Connected Network (SSID)": {"es": "Red conectada (SSID)", "en": "Connected Network (SSID)"},
        "Operating System": {"es": "Sistema operativo", "en": "Operating System"},
        "Total RAM (GB)": {"es": "RAM total (GB)", "en": "Total RAM (GB)"},
        "Disks": {"es": "Discos", "en": "Disks"},
        "Brand": {"es": "Marca", "en": "Brand"},
        "Model": {"es": "Modelo", "en": "Model"},
        "BIOS Serial": {"es": "Serial BIOS", "en": "BIOS Serial"},
        "MAC": {"es": "MAC", "en": "MAC Address"},
        "BIOS Version": {"es": "Versión BIOS", "en": "BIOS Version"},
        "Processor (CPU)": {"es": "Procesador (CPU)", "en": "Processor (CPU)"},
        "System Serial Number": {"es": "Número de serie del sistema", "en": "System Serial Number"},
        "Screen Resolution": {"es": "Resolución pantalla", "en": "Screen Resolution"},
        "Battery - Charge (%)": {"es": "Batería - Carga (%)", "en": "Battery - Charge (%)"},
        "Battery - Status": {"es": "Batería - Estado", "en": "Battery - Status"},
        "Battery - Capacity": {"es": "Batería - Capacidad", "en": "Battery - Capacity"},
        "OS Installation Date": {"es": "Fecha de instalación SO", "en": "OS Installation Date"},
        "Last Boot": {"es": "Último arranque", "en": "Last Boot"},
        "Windows Version": {"es": "Versión Windows", "en": "Windows Version"},
        "Windows Build": {"es": "Build Windows", "en": "Windows Build"},
        "Firewall Enabled": {"es": "Firewall Activado", "en": "Firewall Enabled"},
    }
    for key, value in data.items():  # Iterate through data
        if key == "Disks":
            pdf.set_font("Arial", 'B', 12)  # Set font for disks
            pdf.cell(0, 10, titles[key][idioma_actual] + ":", 0, 1)  # Add disks title
            pdf.set_font("Arial", '', 12)  # Set font for disk details
            for d in value:
                pdf.cell(0, 8, f"  {d[0]}: {d[1]} GB free of {d[2]} GB", 0, 1)  # Add disk details
        else:
            pdf.set_font("Arial", 'B', 12)  # Set font for other details
            pdf.cell(50, 10, f"{titles[key][idioma_actual]}:", 0, 0)  # Add title for the key
            pdf.set_font("Arial", '', 12)  # Set font for value
            # Handle encoding
            value = str(value).encode('latin-1', 'replace').decode('latin-1')  # Encode value
            pdf.multi_cell(0, 10, value)  # Add value to PDF

    pdf.output(file)  # Save the PDF to the specified file

# Returns the absolute path of a resource file
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # For bundled applications
    except Exception:
        base_path = os.path.abspath(".")  # For normal execution
    return os.path.join(base_path, relative_path)  # Return full path

# Centers the window on the screen
def center_window(win, width, height):
    screen_width = win.winfo_screenwidth()  # Get screen width
    screen_height = win.winfo_screenheight()  # Get screen height
    x = int((screen_width / 2) - (width / 2))  # Calculate x position
    y = int((screen_height / 2) - (height / 2))  # Calculate y position
    win.geometry(f"{width}x{height}+{x}+{y}")  # Set window size and position

# Displays the sheet and saves it as a PDF
def show_sheet_and_save(data, sheet, window, text_area):
    center_window(window, 700, 600)  # Center the window
    text_area.grid()  # Show the text area
    text_area.delete(1.0, tk.END)  # Clear previous content
    text_area.insert(tk.END, sheet)  # Insert the new sheet content
    computer_name = data.get('Computer Name', 'Computer')  # Get computer name
    date_time = datetime.now().strftime("%Y%m%d_%H%M%S")  # Get current date and time
    prefix_name = traducciones[idioma_actual].get("nombre_archivo_prefijo", "Sheet" if idioma_actual == "en" else "Ficha")  # Get file prefix
    suggested_name = f"{prefix_name}_{computer_name}_{date_time}.pdf"  # Create suggested file name
    file = filedialog.asksaveasfilename (defaultextension=".pdf",
                                           filetypes=[("PDF files", "*.pdf")],
                                           title=traducciones[idioma_actual]["guardar_como"],
                                           initialfile=suggested_name)  # Open save dialog for PDF
    if file:  # If a file was selected
        try:
            print(f"Attempting to save PDF at: {file}")  # Debug message
            create_pdf_sheet(data, file)  # Create the PDF sheet
            messagebox.showinfo(traducciones[idioma_actual]["exito_titulo"], traducciones[idioma_actual]["exito_msj"].format(file))  # Show success message
        except Exception as e:
            print(f"Error saving PDF: {e}")  # Debug message
            messagebox.showerror(traducciones[idioma_actual]["error_titulo"], traducciones[idioma_actual]["error_msj"].format(e))  # Show error message

# Changes the language of the UI
def change_language():
    global idioma_actual, header, btn_obtain, btn_language
    idioma_actual = "en" if idioma_actual == "es" else "es"  # Toggle language
    header.config(text=traducciones[idioma_actual]["titulo"])  # Update header text
    btn_obtain.config(text=traducciones[idioma_actual]["obtener_ficha"])  # Update button text
    btn_language.config(text=traducciones[idioma_actual]["idioma_btn"])  # Update language button text

# Creates a splash window that appears on startup
def splash_window():
    splash = tk.Tk()  # Create a new Tkinter window
    splash.overrideredirect(True)  # Remove window decorations
    splash.configure(bg="white")  # Set background color
    splash_width, splash_height = 200, 200  # Set splash window size
    screen_width = splash.winfo_screenwidth()  # Get screen width
    screen_height = splash.winfo_screenheight()  # Get screen height
    x = (screen_width // 2) - (splash_width // 2)  # Calculate x position
    y = (screen_height // 2) - (splash_height // 2)  # Calculate y position
    splash.geometry(f"{splash_width}x{splash_height}+{x}+{y}")  # Set window size and position
    icon_path = resource_path(ICONO_NOMBRE)  # Get icon path
    try:
        splash.iconbitmap(icon_path)  # Set window icon
    except Exception:
        pass  # Ignore errors if icon cannot be set
    try:
        image_path = resource_path(IMAGEN_SPLASH)  # Get splash image path
        image = Image.open(image_path)  # Open image
        image = image.resize((splash_width, splash_height), Image.LANCZOS)  # Resize image
        photo = ImageTk.PhotoImage(image)  # Create PhotoImage object
        label = tk.Label(splash, image=photo, bg="white")  # Create label for image
        label.image = photo  # Keep a reference to avoid garbage collection
        label.pack(expand=True)  # Pack label into the window
    except:
        label = tk.Label(splash, text=NOMBRE_MARCA, font=("Segoe UI", 16), bg="white", fg="#2196F3")  # Fallback label
        label.pack(expand=True)  # Pack fallback label

    def close_splash():
        splash.destroy()  # Close splash window
        start_main_window()  # Start the main application window

    splash.after(2000, close_splash)  # Close splash after 2 seconds
    splash.mainloop()  # Run the splash window event loop

# Starts the main application window
def start_main_window():
    global header, btn_obtain, btn_language
    window = tk.Tk()  # Create a new Tkinter window
    window.resizable(False, False)  # Disable resizing
    window.title(NOMBRE_MARCA)  # Set window title
    center_window(window, 400, 200)  # Center the window
    try:
        window.iconbitmap(resource_path(ICONO_NOMBRE))  # Set window icon
    except:
        pass  # Ignore errors if icon cannot be set
    main_color = "#2196F3"  # Main color for the UI
    background_color = "#F5F7FA"  # Background color for the UI
    text_color = "#212121"  # Text color for the UI
    area_color = "#FFFFFF"  # Color for text area
    window.configure(bg=background_color)  # Set window background color
    general_font = ("Segoe UI", 10)  # General font for the UI
    title_font = ("Segoe UI", 14, "bold")  # Font for the title

    header = tk.Label(window, text=traducciones[idioma_actual]["titulo"], bg=background_color,
                      fg=main_color, font=title_font, pady=15)  # Create header label
    header.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")  # Place header in grid

    btn_language = tk.Button(window, text=traducciones[idioma_actual]["idioma_btn"], command=change_language,
                             bg=background_color, fg=main_color, font=general_font, bd=0, cursor="hand2")  # Create language button
    btn_language.place(x=10, y=10)  # Place language button

    style = ttk.Style()  # Create a new style
    style.theme_use("clam")  # Use clam theme
    style.configure("Custom.TButton", foreground="white", background=main_color,
                    font=("Segoe UI", 12, "bold"), padding=10, borderwidth=0)  # Configure button style
    style.map("Custom.TButton", background=[("active", "#1976D2"), ("disabled", "#90A4AE")])  # Map button styles

    text_area = scrolledtext.ScrolledText(window, width=80, height=30,
                                          font=general_font, bg=area_color, fg=text_color,
                                          bd=1, relief="solid", wrap=tk.WORD, padx=10, pady=10)  # Create text area
    text_area.grid(row=2, column=0, padx=20, pady=(5, 10), sticky="nsew")  # Place text area in grid
    text_area.grid_remove()  # Initially hide text area

    progress = ttk.Progressbar(window, mode='indeterminate')  # Create progress bar
    progress.grid(row=3, column=0, padx=20, pady=5, sticky='ew')  # Place progress bar in grid
    progress.grid_remove()  # Initially hide progress bar

    def task_get_sheet():
        try:
            data = get_pc_data()  # Get PC data
            sheet = format_sheet(data)  # Format the data into a sheet
            window.after(0, show_sheet_and_save, data, sheet, window, text_area)  # Show and save the sheet
        finally:
            window.after(0, lambda: btn_obtain.config(state=tk.NORMAL))  # Enable button after task
            window.after(0, lambda: progress.stop())  # Stop progress bar
            window.after(0, lambda: progress.grid_remove())  # Hide progress bar

    def button_get():
        btn_obtain.config(state=tk.DISABLED)  # Disable button during task
        text_area.grid_remove()  # Hide text area
        text_area.delete(1.0, tk.END)  # Clear text area
        progress.grid()  # Show progress bar
        progress.start()  # Start progress bar animation
        threading.Thread(target=task_get_sheet, daemon=True).start()  # Start task in a new thread

    btn_obtain = ttk.Button(window, text=traducciones[idioma_actual]["obtener_ficha"],
                             command=button_get, style="Custom.TButton")  # Create button to get sheet
    btn_obtain.grid(row=1, column=0, padx=20, pady=15, sticky='ew')  # Place button in grid

    window.grid_columnconfigure(0, weight=1)  # Configure column weight
    window.grid_rowconfigure(2, weight=1)  # Configure row weight
    window.mainloop()  # Run the main window event loop

# Run the app
splash_window()  # Start the application with the splash window