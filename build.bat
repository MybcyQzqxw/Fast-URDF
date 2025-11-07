@echo off
:: 创建虚拟环境
python -m venv myenv

:: 激活虚拟环境
call myenv\Scripts\activate

:: 升级 pip 到最新版本
python.exe -m pip install --upgrade pip

:: 安装依赖
pip install -r requirements.txt

:: 安装 pyinstaller
pip install pyinstaller

:: 使用 pyinstaller 打包程序
pyinstaller --noconsole --onefile --icon="favicon.ico" --add-data "favicon.ico;." "Fast URDF.py"