# build_win.ps1 - 用 PyInstaller 打包 Windows EXE（GUI）

python -m pip install --upgrade pip
if (Test-Path requirements.txt) { pip install -r requirements.txt }
pip install pyinstaller

# 把模板打进可执行根目录（注意分号 ; 是 Windows 的写法）
$addDataArgs = @()
if (Test-Path "KMART模板.xlsx")  { $addDataArgs += @("--add-data", "KMART模板.xlsx;.") }
if (Test-Path "TARGET模板.xlsx") { $addDataArgs += @("--add-data", "TARGET模板.xlsx;.") }

# 可选图标
$iconArg = @()
if (Test-Path "icon.ico") { $iconArg += @("--icon", "icon.ico") }

# 直接以 src/label_app.py 作为入口（不再使用 launcher.py）
$pyiArgs = @(
  "--onefile",
  "--windowed",
  "--name", "LabelApp",
  "--paths", "src"            # 让分析期能找到 src 下其它模块
) + $iconArg + $addDataArgs + @("src/label_app.py")

pyinstaller @pyiArgs
