# Excel Audit

**Excel Audit** is a Python-based GUI tool built with PySide6 designed to streamline the merging and auditing of "Buyer" and "Seller" Excel Sheets(.xlsx). It ensures data integrity by enforcing strict header validation and allowing users to define custom header rows and merge keys for each file.

## Quick Start (For Users)

If you just want to use the tool without setting up Python, download the standalone executable for your operating system:

  * **Windows (.exe):** [Download Latest Windows Binary](https://github.com/Crisp79/excel_audit/releases/download/v1.0.0/ExcelAudit_Windows.exe)
  * **Linux (Binary):** [Download Latest Linux Binary](https://github.com/Crisp79/excel_audit/releases/download/v1.0.0/ExcelAudit_Linux)

> **Note:** On Windows, if you get a "SmartScreen" warning, click "More info" and "Run anyway." This is common for unsigned community-built binaries.

-----

## Developer Setup (For Contributors)

If you want to modify the code or run it from the source, follow these steps. This project uses **uv** for fast, reliable dependency management.

### 1\. Prerequisites

Ensure you have [uv](https://github.com/astral-sh/uv) installed on your system.

### 2\. Clone and Initialize

```powershell
# Clone the repository
git clone https://github.com/yourusername/excel-audit.git
cd excel-audit

# Create environment and install dependencies
uv sync
```

### 3\. Running the Application

```powershell
uv run main.py
```

### 4\. Building your own Binaries

To package the app yourself using PyInstaller:

```powershell
uv run pyinstaller --noconsole --onefile --name "ExcelAudit" main.py
```

-----

## Usage Instructions

### 1\. File Preparation

  * **Header Consistency:** All sheets in a single file must have the same header row index.
  * **Formatting:** Headers must have exact spaces and dots (e.g., `Serial. No` vs `Serial No`).
  * **Permissions:** Ensure the files are closed in Excel before running the audit.

### 2\. GUI Configuration

1.  **File Paths:** Select your "Buyer" and "Seller" Excel files.
2.  **Mapping:** Specify the **Header Row Index** (usually 0) and the **Column Name** to use as the merge key.
3.  **Output:** Set the destination folder and base filename.

-----

## Project Structure

```text
excel-audit/
├── main.py          # PySide6 GUI Entry Point
├── logic.py         # Backend Excel processing logic
├── utils.py         # Data import and cleaning logic
├── formatter.py     # Row Formatting logic
├── pyproject.toml   # Project metadata and dependencies
├── uv.lock          # Locked dependency versions
└── LICENSE          # MIT License
```

-----

## License

Copyright (c) 2026 **Akshith Vuppala**

This project is licensed under the **MIT License**. See the [LICENSE](./LICENSE) file for the full text.
