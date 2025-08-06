pyinstaller --clean --onefile -F -n pasa-rpa-digitacao-peg ^
    --add-data "%APPDATA%\..\Local\ms-playwright\chromium-1105;./playwright/driver/package/.local-browsers/chromium-1105" ^
    --hidden-import=pandas ^
    --icon=mitrius.ico ^
    --version-file=version.txt ^
    --add-data "%APPDATA%\..\Local\.certifi\cacert.pem;certifi" ^
    main.py

copy config.yml dist /y
copy logo.png dist /y
copy email.html dist /y

cd dist

copy pasa-rpa-digitacao-peg.exe .. /y

cd ..