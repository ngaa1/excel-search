Write-Host "安装依赖..."
pip install -r requirements.txt

Write-Host "安装PyInstaller..."
pip install pyinstaller

Write-Host "开始打包..."
pyinstaller --onefile --name excel-search --hidden-import=openpyxl.cell._writer main.py

Write-Host "打包完成！"
Write-Host "可执行文件位于 dist 目录中"

Read-Host "按任意键继续..."