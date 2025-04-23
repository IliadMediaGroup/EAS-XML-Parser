# EAS XML Parser

A Windows desktop application for parsing Emergency Alert System (EAS) XML files and generating Excel reports for Required Weekly Tests (RWT) and Required Monthly Tests (RMT). The app updates a single Excel file per month with weekly data, creating new files for new months. Built with PyQt5 and packaged as an MSI installer.

## Features

- User-friendly GUI to select XML files and output directory.
- Parses EAS XML files to extract RWT and RMT data.
- Generates Excel reports with formatted tables, highlighting unparsed data in red.
- Updates monthly Excel files with weekly data, creating new files for new months.
- Logs processing status and errors in the GUI.
- Deployable as an MSI installer for Windows.

## Prerequisites

- Windows 10 or 11 (64-bit).
- Python 3.8+ (for development or running from source).
- Inno Setup (for building the MSI installer).

## Installation

### Option 1: Install via MSI

1. Download the latest `EASParserSetup.msi` from the Releases page.
2. Run the installer and follow the prompts to install the app.
3. Launch “EAS XML Parser” from the Start menu or desktop shortcut.

### Option 2: Run from Source

1. Clone the repository:

   ```bash
   git clone https://github.com/IliadMediaGroup/eas-xml-parser.git
   cd eas-xml-parser
   ```
2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```
3. Run the app:

   ```bash
   python eas_parser_app.py
   ```

## Usage

1. Launch the app.
2. Click “Select XML Files” to choose one or more EAS XML files.
3. Click “Select Output Directory” to specify where Excel reports will be saved.
4. Click “Parse and Generate Excel Reports” to process the files.
5. View the log in the app for processing status and errors.
6. Check the output directory for the Excel file (e.g., `EAS_2025-04.xlsx`), updated weekly for the current month.

## Building the MSI Installer

1. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```
2. Build the executable with PyInstaller:

   ```bash
   pyinstaller --name EASParser --icon=icon.ico --onefile eas_parser_app.py
   ```
3. Install Inno Setup.
4. Open `eas_parser_setup.iss` in Inno Setup Compiler.
5. Ensure `dist/EASParser.exe` and `icon.ico` are in the same directory as `eas_parser_setup.iss`.
6. Compile to generate `Output/EASParserSetup.msi`.

## Project Structure

- `eas_parser_app.py`: Main application script with GUI and parsing logic.
- `eas_parser_setup.iss`: Inno Setup script for MSI packaging.
- `icon.ico`: Application icon (optional).
- `requirements.txt`: Python dependencies.
- `LICENSE`: MIT License for open-source usage.

## Contributing

Contributions are welcome! Please:

1. Fork the repository.
2. Create a feature branch (`git checkout -b feature/YourFeature`).
3. Commit changes (`git commit -m "Add YourFeature"`).
4. Push to the branch (`git push origin feature/YourFeature`).
5. Open a Pull Request.

## License

This project is licensed under the MIT License. See the LICENSE file for details.

## Contact

For questions or support, open an issue on GitHub or contact \[your email or preferred contact method\].

---

*Built for broadcasters to simplify EAS compliance reporting.*
