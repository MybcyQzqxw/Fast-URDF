@echo off
:: Clean old build files and directories
echo Cleaning old build files...

if exist "build" (
    echo Deleting build directory...
    rmdir /s /q "build"
)

if exist "dist" (
    echo Deleting dist directory...
    rmdir /s /q "dist"
)

if exist "myenv" (
    echo Deleting myenv directory...
    rmdir /s /q "myenv"
)

if exist "Fast URDF.spec" (
    echo Deleting Fast URDF.spec file...
    del /f /q "Fast URDF.spec"
)

echo Cleanup complete!
echo.

:: Create virtual environment
echo Creating virtual environment...
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