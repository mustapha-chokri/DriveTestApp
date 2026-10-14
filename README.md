# DriveTestApp 📡

**DriveTestApp** is an automated desktop application designed for Telecom & Drive Test engineers. It streamlines the Radio Network Planning (RNP) data preparation pipeline, merges cell parameters across 2G, 3G, and 4G technologies into standardized ADT Excel templates, executes validation macros, and exports `.NBF` (Nemo / TEMS) configuration files.

---

## 🌟 Key Features

- **Automated Data Migration**: Extracts cellular parameters from source RNP workbooks (`2G RNP`, `3G RNP`, `4G RNP`, `Input Data`) and injects them into the ADT template.
- **Excel Macro COM Automation**: Triggers Excel VBA macros (`PasteValues2G3G4G`) silently in the background without freezing or opening Excel UI.
- **NBF Generation**: Automatically extracts configuration cell coordinates (`O1:O13`) for each technology and produces ready-to-use `.NBF` files for Drive Test tools.
- **Modern User Interface**: Built with `ttkbootstrap` featuring an interactive dark theme (`superhero`), background multi-threading, live logging, and one-click folder navigation.
- **Modular Architecture**: Designed for maintainability, unit testing, and extensibility.

---

## 🏗️ Project Architecture

```text
DriveTestApp/
├── run.py                          # Primary application entry point
├── main.py                         # Backward-compatible entry point
├── requirements.txt                # Python package dependencies
├── config.example.json             # Sample configuration template
├── .gitignore                      # Git ignore rules for logs, caches, and local paths
├── src/
│   ├── config.py                   # Configuration manager
│   ├── core/                       # Business logic
│   │   ├── excel_processor.py      # openpyxl RNP-to-ADT data pipeline
│   │   ├── macro_runner.py         # win32com COM automation & NBF exporter
│   │   └── exceptions.py           # Domain-specific custom exceptions
│   ├── ui/                         # User interface
│   │   ├── app.py                  # Main window & thread handling
│   │   └── components.py           # Custom UI widgets and viewers
│   └── utils/                      # Utilities
│       ├── logger.py               # Unified logger (File + UI streaming)
│       └── shortcut.py             # Desktop shortcut generator
└── scripts/
    └── create_desktop_shortcut.ps1 # PowerShell shortcut script with pythonw support
```

---

## 🚀 Installation & Setup

### Prerequisites
- **Operating System**: Windows 10 / 11
- **Python**: Python 3.9+
- **Microsoft Excel**: Desktop version installed (required for COM automation and VBA macros)

### 1. Clone the repository
```bash
git clone https://github.com/mustapha-chokri/DriveTestApp.git
cd DriveTestApp
```

### 2. Create and activate a Virtual Environment (Recommended)
```bash
python -m venv .venv
.venv\Scripts\activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

---

## 💻 Usage

### Launch the Application
Run the main script:
```bash
python run.py
```
*(Or `python main.py`)*

### Create a Desktop Shortcut
To launch the application directly from your desktop without a terminal console:
- Using Python:
  ```bash
  python src/utils/shortcut.py
  ```
- Or using PowerShell:
  ```powershell
  powershell -ExecutionPolicy Bypass -File scripts/create_desktop_shortcut.ps1
  ```

---

## ⚙️ Configuration

On first run, the app generates a local `config.json` file saving your last used paths:
- **RNP Source File**: Path to the raw Radio Network Planning workbook (`.xlsx` or `.xlsm`).
- **Template ADT**: Path to the master macro template (`applic_DT.xlsm`).
- **Output Directory**: Target directory for the generated workbook and the `NBF_Output` folder.
- **Target City**: The operational city code (e.g. `TANGER`).

---

## 🛡️ License & Author

Developed by **Mustapha Chokri**.  
Engineered for Telecom Network Automation.
