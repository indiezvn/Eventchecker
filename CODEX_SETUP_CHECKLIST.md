# Codex Setup Checklist

File nay danh cho Codex hoac dev moi tren may khac.

Muc tieu:

- Setup repo nay tu dau tren may moi
- Chay duoc app local
- Co the tiep tuc code, commit, push
- Hieu ro phan nao la source chung, phan nao la local-only


## 1. Thu tu doc file trong repo

Khi Codex mo repo nay tren may moi, hay doc theo thu tu:

1. `README_SETUP_TEAM.md`
2. `README_SETUP_WINDOWS.md` hoac `README_SETUP_MAC.md`
3. `README_APP.md`
4. `requirements.txt`
5. `desktop_app.py`
6. `Log_checker.py`

Neu can build:

7. `build/windows/README_PORTABLE_WINDOWS.md`
8. `build/windows/README_INSTALLER_WINDOWS.md`
9. `build/macos/README_INSTALLER_MACOS.md`


## 2. Nhiem vu setup toi thieu

Codex can dam bao cac buoc sau:

1. Xac nhan may da co `git`
2. Xac nhan may da co Python 3.11 hoac 3.12
3. Clone repo
4. Tao `.venv`
5. Cai `requirements.txt`
6. Chay `python desktop_app.py`
7. Kiem tra app co mo local server duoc hay khong


## 3. Lenh setup nhanh

macOS / Linux:

```bash
git clone https://github.com/indiezvn/Eventchecker.git
cd Eventchecker
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python desktop_app.py
```

Windows PowerShell:

```powershell
git clone https://github.com/indiezvn/Eventchecker.git
cd Eventchecker
py -3 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python desktop_app.py
```

Windows CMD:

```bat
git clone https://github.com/indiezvn/Eventchecker.git
cd Eventchecker
py -3 -m venv .venv
.venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
python desktop_app.py
```


## 4. Kiem tra sau setup

Sau khi chay app, Codex can kiem tra:

- app co mo duoc khong
- co loi import module khong
- local server dang nghe o port nao
- log co ghi vao thu muc user data khong

Neu chay `desktop_app.py` ma loi `ModuleNotFoundError`, quay lai kiem tra:

- `.venv` da activate chua
- `pip install -r requirements.txt` da thanh cong chua


## 5. User data / local state

App dung thu muc local state khac nhau theo OS:

Windows:

- `%LOCALAPPDATA%\EventInspector`

macOS:

- `~/Library/Application Support/EventInspector`

Trong do co the co:

- `app.log`
- `update_state_v250.json`
- `remote_update_config_v250.json`
- `updates_v250/`

Day la state local, khong phai source code.


## 6. Nguon code nao la source chung

Nhung file Codex nen uu tien sua:

- `Log_checker.py`
- `desktop_app.py`
- `remote_update.py`
- `Updates_2_5/remote_manifest.json`
- file trong `tools/`
- file trong `build/`

Khong sua / khong commit:

- `__pycache__/`
- `.venv/`
- `dist/`
- `.worktrees/`
- state trong thu muc user data


## 7. Quy tac lam viec Git

Mac dinh:

- `main` = source chung

Neu lam task moi:

```bash
git checkout main
git pull origin main
git checkout -b feature/<ten-task>
```

Sau khi sua:

```bash
git add <file-can-thiet>
git commit -m "Mo ta thay doi"
git push origin feature/<ten-task>
```


## 7A. Workflow ma Codex phai theo

Tren may dev moi, Codex khong duoc mac dinh sua xong la push ngay.

Workflow dung:

1. Pull code moi nhat
2. Tao branch rieng
3. Sua code local
4. Chay app local
5. Neu can thi build local
6. Neu local test OK moi commit/push

Trinh tu khuyen nghi:

```bash
git checkout main
git pull origin main
git checkout -b feature/<ten-task>

# sua code
python desktop_app.py

# neu can build:
# macOS -> bash build/macos/build_macos.sh
# Windows -> build\windows\build_portable.bat

# sau khi test OK moi push
git add <file-can-thiet>
git commit -m "Mo ta thay doi"
git push origin feature/<ten-task>
```

Codex can uu tien:

- test local truoc
- push sau
- khong commit file local-only


## 7B. Neu co 2 may cung dang lam

Neu co nhieu may cung dev, Codex phai lam theo quy tac:

1. Moi may = 1 branch rieng
2. Push thuong xuyen, khong giu code local qua lau
3. Truoc khi tiep tuc lam, phai `git fetch` hoac `git pull`
4. Truoc khi merge, phai dua `origin/main` vao branch hien tai

Lenh mau:

```bash
git fetch origin
git checkout main
git pull origin main
git checkout <branch-cua-minh>
git merge origin/main
```

Neu Codex gap conflict:

- khong tu y force push
- khong reset manh
- bao lai tinh trang conflict
- chi resolve khi conflict ro rang va nho


## 7C. Thao tac Git co the giao cho Codex

Codex co the tu dong lam:

- `git status`
- `git checkout`
- `git fetch`
- `git pull`
- `git add`
- `git commit`
- `git push`
- merge `origin/main` vao branch hien tai

Can nguoi xac nhan:

- merge vao `main`
- force push
- reset / revert co rui ro
- xoa branch


## 8. Khi nao phai sua payload update

Neu thay doi chi de dev local:

- co the chi sua `Log_checker.py` / `desktop_app.py`

Neu thay doi can app client nhan qua updater:

- sua payload va manifest trong `Updates_2_5/`
- cap nhat `sha256`
- tang `version` trong manifest
- commit va push


## 9. Khi nao can build

Khong phai thay doi nao cung can build.

Chi can build khi:

- can tao app moi cho user
- can tao `.dmg`
- can tao `.exe`
- can test packaging

Con neu chi dang sua source va chay local:

- `python desktop_app.py` la du

Neu may moi co vai tro giong may chinh hien tai, thi duoc phep:

- sua source code
- build local
- chay app local
- test xong moi push len Git


## 10. Cach bao cao sau khi setup xong

Sau khi Codex setup xong tren may moi, nen bao cao ngan gon:

- da clone repo chua
- da tao `.venv` chua
- da cai dependency chua
- da chay duoc `desktop_app.py` chua
- co thieu `adb` hay tool iOS khong
- co can quyen them de build/push khong
