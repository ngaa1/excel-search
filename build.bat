@echo off

echo 安装依赖...
pip install -r requirements.txt

echo 安装PyInstaller...
pip install pyinstaller

echo 开始打包...
pyinstaller --onefile --name excel-search main.py

echo 打包完成！
echo 可执行文件位于 dist 目录中

pause