# 确保使用UTF-8编码
$PSDefaultParameterValues['*:Encoding'] = 'utf8'
[console]::InputEncoding = [System.Text.Encoding]::UTF8
[console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 安装依赖
Write-Output "安装依赖..."
pip install -r requirements.txt

# 安装PyInstaller
Write-Output "安装PyInstaller..."
pip install pyinstaller

# 开始打包
Write-Output "开始打包..."
pyinstaller --onefile --name excel-search --hidden-import=openpyxl.cell._writer main.py

# 完成
Write-Output "打包完成！"
Write-Output "可执行文件位于 dist 目录中"

Read-Host "按任意键继续..."