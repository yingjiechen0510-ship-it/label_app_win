param(
  [ValidateSet("x86","x64")]
  [string]$Arch = "x64"
)

$ErrorActionPreference = "Stop"

python -m pip install -U pip
pip install -r requirements.txt

# 清理旧产物
if (Test-Path build) { Remove-Item build -Recurse -Force }
if (Test-Path dist)  { Remove-Item dist  -Recurse -Force }

# 你的入口如果是 src\launcher.py 就保持它；否则换成你的主脚本
pyinstaller --onefile --noconsole `
  --name "LabelApp-$Arch" `
  --add-data "KMART模板.xlsx;." `
  --add-data "TARGET模板.xlsx;." `
  "src\launcher.py"

Write-Host "Built: dist\LabelApp-$Arch.exe"
