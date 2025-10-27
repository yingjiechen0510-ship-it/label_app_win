python -m pip install --upgrade pip
if (Test-Path requirements.txt) { pip install -r requirements.txt }
pip install pyinstaller

# 把模板打进可执行文件根目录（注意分号 ; 是 Windows 的写法）
$addDataArgs = @()
if (Test-Path "KMART模板.xlsx")  { $addDataArgs += @("--add-data", "KMART模板.xlsx;.") }
if (Test-Path "TARGET模板.xlsx") { $addDataArgs += @("--add-data", "TARGET模板.xlsx;.") }

# 可选图标
$iconArg = @()
if (Test-Path "icon.ico") { $iconArg += @("--icon", "icon.ico") }

# 直接以 src/label_app.py 作为入口
$pyiArgs = @(
  "--onefile",
  "--windowed",
  "--name", "LabelApp",
  "--paths", "src"          # 让分析器能找到 src 下的模块
) + $iconArg + $addDataArgs + @("src/label_app.py")

pyinstaller @pyiArgs
