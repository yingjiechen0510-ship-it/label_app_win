# build_win.ps1 - Build Windows .exe with PyInstaller (GUI mode)

python -m pip install --upgrade pip
if (Test-Path requirements.txt) {
    pip install -r requirements.txt
}
pip install pyinstaller

# Add data: put templates at bundle root so relative open("KMART模板.xlsx") works
$addDataArgs = @()
if (Test-Path "KMART模板.xlsx") { $addDataArgs += @("--add-data", "KMART模板.xlsx;.") }
if (Test-Path "TARGET模板.xlsx") { $addDataArgs += @("--add-data", "TARGET模板.xlsx;.") }

# Optional icon
$iconArg = @()
if (Test-Path "icon.ico") { $iconArg += @("--icon", "icon.ico") }

# Use --windowed for Tkinter GUI; remove it if you want a console app
$pyiArgs = @(
  "--onefile",
  "--windowed",
  "--name", "LabelApp",
  "--hidden-import", "label_app",
  "--paths", "src"
) + $iconArg + $addDataArgs + @("src/launcher.py")

pyinstaller @pyiArgs
