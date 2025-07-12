@echo off
echo 正在创建虚拟环境...
python -m venv venv
echo 正在激活虚拟环境...
call venv\Scripts\activate.bat

echo 正在安装依赖...
pip install -r requirements.txt
pip install pyinstaller

echo 正在使用PyInstaller打包程序...
pyinstaller --onefile --windowed --icon=weight-scale.ico --add-data "weight-scale.ico;." "E550串口测试V63.py"

echo 打包完成！可执行文件位于dist文件夹中
echo 按任意键退出
pause > nul
