# LabelApp（Windows .exe 构建项目）

此项目已集成你的 `label_app.py` 与模板：`KMART模板.xlsx`、`TARGET模板.xlsx`。  
通过 GitHub Actions 在 Windows 上打包生成 `LabelApp.exe`，使用者无需安装 Python。

## 使用方法
1. 将本项目推到 GitHub（建议私有仓库）。
2. 在 GitHub → Actions → 运行 **Build Windows EXE**。
3. 运行完成后，在该运行记录的 **Artifacts** 下载 `LabelApp-windows`（内含 `LabelApp.exe`）。
4. 发给 Windows 用户，双击即可运行。

## 说明
- 打包入口：`src/launcher.py`（会把工作目录切换到 PyInstaller 的 `_MEIPASS`，然后执行 `label_app`）。
- 模板文件被打包到可执行文件目录（bundle root），因此 `open("KMART模板.xlsx")` 这类相对路径语句可直接使用。
- 若你的应用是命令行工具，请在 `build_win.ps1` 中删除 `--windowed`。

## 本地验证（可选，Mac 上）
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python src/label_app.py
```

## 依赖
- `requirements.txt` 已根据导入自动写入：`pandas`、`openpyxl`。若还需其它第三方包，请自行追加。
