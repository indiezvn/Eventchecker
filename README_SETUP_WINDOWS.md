# EventInspector Setup on Windows

File nay danh rieng cho may Windows.


## 1. Yeu cau toi thieu

Can co:

- Git for Windows
- Python 3.11 hoac 3.12

Neu test Android:

- Android SDK
- `adb`

Neu build installer:

- Inno Setup


## 2. Kiem tra cong cu san co

PowerShell:

```powershell
git --version
py --version
```

Neu co Android:

```powershell
adb devices
```


## 3. Clone repo

PowerShell:

```powershell
git clone https://github.com/indiezvn/Eventchecker.git
cd Eventchecker
```


## 4. Tao virtualenv

PowerShell:

```powershell
py -3 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

CMD:

```bat
py -3 -m venv .venv
.venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
```


## 5. Chay local

```powershell
python desktop_app.py
```

Neu app khong tim thay `adb`, co the set:

```powershell
$env:ADB_PATH="C:\Users\<user>\AppData\Local\Android\Sdk\platform-tools\adb.exe"
python desktop_app.py
```


## 6. Build Windows portable

CMD hoac PowerShell:

```bat
build\windows\build_portable.bat
```

Output:

- `dist\EventInspector\`


## 7. Build Windows installer

Doc them:

- `build/windows/README_INSTALLER_WINDOWS.md`


## 8. Thu muc local state tren Windows

State nam o:

```text
%LOCALAPPDATA%\EventInspector
```

Mo nhanh bang Run hoac CMD:

```bat
explorer "%LOCALAPPDATA%\EventInspector"
```


## 9. Neu update bi ket

CMD:

```bat
del /f /q "%LOCALAPPDATA%\EventInspector\update_state_v250.json"
del /f /q "%LOCALAPPDATA%\EventInspector\remote_update_config_v250.json"
rmdir /s /q "%LOCALAPPDATA%\EventInspector\updates_v250"
rmdir /s /q "%LOCALAPPDATA%\EventInspector\updates_v250_tmp"
```


## 10. Quy trinh dev tren Windows

```powershell
git checkout main
git pull origin main
git checkout -b feature/<ten-task>
```

Sau khi sua:

```powershell
python desktop_app.py

# neu can test package
# build\windows\build_portable.bat

# khi moi thu OK thi moi push
git add <file>
git commit -m "Mo ta thay doi"
git push origin feature/<ten-task>
```

Flow khuyen nghi tren Windows:

1. Keo code moi nhat
2. Tao branch rieng
3. Sua code local
4. Chay `python desktop_app.py`
5. Neu can test packaging thi build local
6. Neu test OK moi commit/push
