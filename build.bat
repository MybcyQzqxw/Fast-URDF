@echo off
:: 创建虚拟环境
python -m venv myenv

:: 激活虚拟环境
call myenv\Scripts\activate

:: 安装依赖
pip install -r requirements.txt

:: 使用 pyinstaller 打包程序
pyinstaller --noconsole --onefile --icon="favicon.ico" --add-data "favicon.ico;." "Fast URDF.py"