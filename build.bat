@echo off
echo Installing requirements...
pip install -r requirements.txt

echo Building executable via Nuitka...
python -m nuitka ^
    --onefile ^
    --plugin-enable=tk-inter ^
    --include-package-data=customtkinter ^
    --include-module=aiohttp ^
    --include-module=PIL.Image ^
    --include-module=PIL.ImageTk ^
    --include-module=PIL.ImageFont ^
    --include-module=pythoncom ^
    --include-data-files=app.ico=app.ico ^
    --include-data-files=Metropolis-ExtraBold.otf=Metropolis-ExtraBold.otf ^
    --nofollow-import-to=win32com.test ^
    --nofollow-import-to=unittest ^
    --nofollow-import-to=pydoc ^
    --windows-console-mode=disable ^
    --output-dir=dist ^
    --output-filename="Roblox SCC.exe" ^
    --windows-icon-from-ico=app.ico ^
    main.py

echo Build completed! Look into 'dist' folder.
pause