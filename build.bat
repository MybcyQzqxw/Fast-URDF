@echo off
:: 清理旧的构建文件和目录
echo 正在清理旧的构建文件...

if exist "build" (
    echo 删除 build 目录...
    rmdir /s /q "build"
)

if exist "dist" (
    echo 删除 dist 目录...
    rmdir /s /q "dist"
)

if exist "myenv" (
    echo 删除 myenv 目录...
    rmdir /s /q "myenv"
)

if exist "Fast URDF.spec" (
    echo 删除 Fast URDF.spec 文件...
    del /f /q "Fast URDF.spec"
)

echo 清理完成！
echo.

:: 创建虚拟环境
echo 创建虚拟环境...
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