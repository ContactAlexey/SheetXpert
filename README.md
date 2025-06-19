<body>
  <h1>Technical README – SheetXpert</h1>

  <h2>Project Summary</h2>
  <p>
    SheetXpert is a Python-based tool that gathers technical information from a Windows system using PowerShell commands, formats it into a structured report, and allows export to a PDF file. It features a graphical user interface (GUI) built with Tkinter and supports language switching between Spanish and English.
  </p>

  <h2>Main Features</h2>
  <ul>
    <li>System data collection using PowerShell.</li>
    <li>Intuitive graphical interface with multilingual support.</li>
    <li>Export to PDF using FPDF.</li>
    <li>Compatible with Windows systems.</li>
    <li>Visual organization of disk, network, BIOS, battery, processor info, and more.</li>
  </ul>

  <h2>Functionality Overview</h2>

  <h3>1. Data Collection</h3>
  <p>The <code>get_pc_data()</code> function runs multiple PowerShell commands to collect system information:</p>
  <ul>
    <li>Computer name, IP addresses, connected network, manufacturer, model, etc.</li>
    <li>Total RAM, hard drives, processor, BIOS version.</li>
    <li>Battery status, screen resolution, installation and last boot dates.</li>
    <li>Windows version and build, firewall status.</li>
  </ul>
  <p>The results are stored in a structured dictionary.</p>

  <h3>2. Date Formatting</h3>
  <p>The <code>format_date_wmi()</code> function converts WMI-formatted dates into a readable format (YYYY-MM-DD HH:MM:SS).</p>

  <h3>3. Report Display</h3>
  <p>
    The <code>format_sheet(data)</code> function takes the collected data and transforms it into readable plain text, using translations based on the selected language. The structure includes categorized sections (disks, network, battery, etc.).
  </p>

  <h3>4. PDF Export</h3>
  <p>
    <code>create_pdf_sheet()</code> generates a cleanly formatted PDF document: centered title, section headers, and character encoding support.
  </p>

  <h3>5. Graphical Interface</h3>
  <ul>
    <li>Built with Tkinter and components like <code>ScrolledText</code>, <code>ttk.Button</code>, and <code>messagebox</code>.</li>
    <li>Includes splash image and custom window icon.</li>
    <li>Supports real-time language switching using <code>change_language()</code>.</li>
    <li>Window centering via <code>center_window()</code>.</li>
  </ul>

  <h3>6. Report Saving</h3>
  <p>
    <code>show_sheet_and_save()</code> displays the report in the interface and opens a save dialog for the PDF. A default file name is suggested, including the computer name and current date.
  </p>

  <h2>Technical Considerations</h2>
  <ul>
    <li>The script is intended for local use only, with no data sent to third parties.</li>
    <li>It does not include sensitive information or user credentials.</li>
    <li>The code is modular and can be extended with new commands or formats.</li>
  </ul>

  <h2>License</h2>
  <p>
    This software is protected by copyright and is for personal and educational use only.
    Modification, redistribution, or commercial use is prohibited without explicit authorization.
  </p>
</body>
