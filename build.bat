@echo off
echo Installing requirements...
pip install -r requirements.txt

echo Building executable via Nuitka...
python -m nuitka ^
    --standalone ^
    --onefile ^
    --plugin-enable=tk-inter ^
    --include-package-data=customtkinter ^
    --include-package=aiohttp ^
    --include-package=PIL ^
    --include-package=win32com ^
    --include-module=pythoncom ^
    --include-module=win32com.client ^
    --include-data-files=app.ico=app.ico ^
    --include-data-files=Metropolis-ExtraBold.otf=Metropolis-ExtraBold.otf ^
    --nofollow-import-to=win32com.test ^
    --windows-console-mode=disable ^
    --output-dir=dist ^
    --output-filename="Roblox SCC.exe" ^
    --windows-icon-from-ico=app.ico ^
    main.py

echo Build completed! Look into 'dist' folder.
pause