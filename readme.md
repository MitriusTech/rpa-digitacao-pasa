# PASA - Digitação de PEG

Digitação de PEG no Benner GSP a partir de protocolos de reembolso do MobileSaúde.

## Run Locally

Clone the project

```bash
  git clone https://github.com/XXXXXXXXXXXXX/rpa-template.git
```

Go to the project directory

```bash
  cd pasa-rpa-digitacao-peg
```

Create virtual environment

```bash
  python -m venv .venv
  .venv\Scripts\activate
```

Install dependencies

```bash
  python.exe -m pip install --upgrade pip

  pip install -r requirements.txt

  playwright install chromium
  
  pip install pyinstaller
  
```

Start the automation

```bash
  python main.py
```

## Acknowledgements

 - [Readme on line editor](https://readme.so/editor)


##  Build

Specify Chromium directory and version (inside build.bat)

```bash

--add-data "%APPDATA%\..\Local\ms-playwright\{chromium_folder};./playwright/driver/package/.local-browsers/{chromium_version}"

EX: --add-data "%APPDATA%\..\Local\ms-playwright\chromium-1097;./playwright/driver/package/.local-browsers/1097" ^

```

Generate .exe

```bash
    ./build.bat
```

## Requirements

Create requirements.txt file (activate .venv before these codes)

```bash
  python.exe -m pip install --upgrade pip

  pip install pipreqs

  pipreqs --force --encoding=utf8 --ignore bin,etc,include,lib,lib64,.venv

  pip freeze > requirements.txt

```
## Security

Encrypt and decrypt config.yml

Use current bot user password on Windows

```bash
    easy-vault encrypt config.yml

    easy-vault decrypt config.yml
```