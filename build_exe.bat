@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Creating the build environment...
  python -m venv .venv
  if errorlevel 1 goto fail
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 goto fail
)
".venv\Scripts\python.exe" -m pip install nuitka ordered-set zstandard
if errorlevel 1 goto fail
".venv\Scripts\python.exe" -m nuitka --standalone --assume-yes-for-downloads --zig --enable-plugin=pyside6 --windows-console-mode=disable --windows-icon-from-ico=assets\app.ico --company-name=EditablePPTX --product-name=EditablePPTX --file-version=1.0.0.0 --product-version=1.0.0.0 --file-description="Convert images into editable PowerPoint files" --copyright=EditablePPTX --output-dir=build --output-filename=EditablePPTX.exe --include-package=winrt --include-package-data=pptx --nofollow-import-to=editable_pptx.selftest --nofollow-import-to=fastapi --nofollow-import-to=uvicorn --nofollow-import-to=cv2 --nofollow-import-to=tkinter launch.py
if errorlevel 1 goto fail
echo.
echo Built: build\launch.dist\EditablePPTX.exe
exit /b 0
:fail
echo Build failed.
exit /b 1
