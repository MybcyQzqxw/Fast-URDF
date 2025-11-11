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

:: Activate virtual environment
call myenv\Scripts\activate

:: Upgrade pip to the latest version
python.exe -m pip install --upgrade pip

:: Install dependencies
pip install -r requirements.txt

:: Install pyinstaller
pip install pyinstaller

:: Use pyinstaller to package the program
pyinstaller --noconsole --onefile --icon="favicon.ico" --add-data "favicon.ico;." "Fast URDF.py"