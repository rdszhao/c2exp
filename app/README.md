# Concept2 Logbook Exporter

A standalone Windows application to export Concept2 workout data to Excel.

## For Users

### Quick Start
1. Download `Concept2Exporter.exe`
2. Run the application
3. Click "Login to Concept2" and authorize in your browser
4. Click "Fetch Workouts" to download your data
5. Click "Export to Excel" to save the file

### Features
- Exports all workout data from your Concept2 Logbook
- Creates per-person sheets (based on workout notes)
- Shows LEFT and RIGHT sides side-by-side (for paddle erg)
- Includes stroke-by-stroke power data
- Highlights peak power strokes in yellow
- Shows average and peak power summaries

### Notes Format
For paddle erg workouts, add notes like "Name L" or "Name R" to indicate:
- The person's name
- Which side (L for Left, R for Right)

Example: "John L" or "Jane R"

---

## For Developers

### Building the Windows Executable

#### Prerequisites
- Windows 10/11
- Python 3.10 or higher
- Git (optional)

#### Build Steps

1. **Copy the app folder to a Windows machine**

2. **Run the build script:**
   ```cmd
   cd app
   build_windows.bat
   ```

3. **Find the executable at:** `dist\Concept2Exporter.exe`

#### Manual Build (Alternative)

```cmd
# Create virtual environment
python -m venv venv
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Build executable
pyinstaller concept2_export.spec --clean
```

### Project Structure
```
app/
├── concept2_export.py    # Main application code
├── concept2_export.spec  # PyInstaller configuration
├── requirements.txt      # Python dependencies
├── build_windows.bat     # Windows build script
└── README.md            # This file
```

### Distribution
The single `Concept2Exporter.exe` file contains everything needed:
- No Python installation required on user machines
- No additional DLLs or files needed
- Users just download and run

### OAuth Credentials
The Concept2 API credentials are embedded in the application.
Each user authenticates with their own Concept2 account.
