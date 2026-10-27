# EventInspector Team Setup Guide

Muc tieu cua file nay:
- Dua source code sang may moi de tiep tuc phat trien.
- Cho 2-3 nguoi co the cung lam viec tren cung 1 repo.
- Dam bao Codex tren may khac mo repo len la co the setup va tiep tuc lam chung.


## 1. Tong quan repo

Repo GitHub hien tai:

- `https://github.com/indiezvn/Eventchecker.git`

Nhanh chinh:

- `main`: source code chung, on dinh de team cung phat trien.

Thanh phan chinh trong repo:

- `desktop_app.py`: entry point cho desktop app.
- `Log_checker.py`: backend + frontend HTML/JS chinh cua tool.
- `remote_update.py`: logic updater/cache update.
- `requirements.txt`: dependency Python.
- `build/macos/`: script build macOS.
- `build/windows/`: script build Windows.
- `tools/`: script ho tro reset/update local.
- `Updates_2_5/remote_manifest.json`: manifest update duy nhat cho kenh chinh.


## 2. Yeu cau tren may moi

Can co san:

- Git
- Python 3.11 hoac 3.12
- Internet de clone repo va cai dependency

Neu can test Android:

- Android SDK + `adb`

Neu can test iOS:

- macOS
- `libimobiledevice` / `idevicesyslog` hoac bo tool iOS ma team dang dung

Neu can build Windows app:

- Windows
- Python
- Neu build installer thi can Inno Setup

Neu can build macOS app:

- macOS
- Python
- Xcode Command Line Tools


## 3. Clone repo ve may moi

```bash
git clone https://github.com/indiezvn/Eventchecker.git
cd Eventchecker
```

Kiem tra nhanh:

```bash
git remote -v
git branch
```


## 4. Tao moi truong Python

macOS / Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

Windows PowerShell:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Windows CMD:

```bat
python -m venv .venv
.venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
```


## 5. Chay app local

Chay desktop app:

```bash
python desktop_app.py
```

Mac/Linux:
- App thuong mo local server tai `http://127.0.0.1:5001`
- Neu port `5001` dang ban, app se tu chon port khac

Neu muon chay backend truc tiep de debug:

```bash
python Log_checker.py
```

Neu may khong nhan Android device:

- kiem tra `adb devices`
- neu can, set bien moi truong `ADB_PATH`

Vi du macOS:

```bash
export ADB_PATH="$HOME/Library/Android/sdk/platform-tools/adb"
python desktop_app.py
```


## 6. Kiem tra dependency thiet bi

Android:

```bash
adb devices
```

iOS tren macOS:

- Neu team dung `idevicesyslog`, kiem tra tool do truoc
- Neu team dung `tidevice`, cai rieng tren may can test

Luu y:
- Repo nay khong tu dong cai bo tool iOS system-level
- Moi dev tu setup phan iOS test tren may cua minh


## 7. Cach lam viec nhieu nguoi

Khuyen nghi workflow:

- Khong code truc tiep tren `main`
- Moi task tao 1 branch rieng
- Xong thi push branch va merge lai
- Moi may dev deu duoc code local, build local, chay thu local
- Chi push len Git khi da test local thay on

Buoc lam viec moi ngay:

```bash
git checkout main
git pull origin main
git checkout -b feature/ten-task
```

Sau khi sua:

```bash
git add .
git commit -m "Mo ta thay doi"
git push origin feature/ten-task
```

Sau do tao Pull Request tren GitHub.

Neu team chua dung PR nghiem tuc, van nen giu toi thieu quy tac:

- Moi nguoi 1 branch
- Merge xong moi quay lai `main`
- Luon `git pull origin main` truoc khi tao branch moi


## 7A. Flow lam viec chuan cua team

Day la flow khuyen nghi cho moi may dev, bao gom may dang dung Codex:

1. Keo code moi nhat tu GitHub
2. Tao branch rieng de sua
3. Sua source code local tren may do
4. Chay app local de test nhanh
5. Neu can, build local de test packaging
6. Neu thay on, moi commit va push len GitHub
7. Sau do moi merge vao nhanh chung

Flow ngan gon:

```bash
git checkout main
git pull origin main
git checkout -b feature/<ten-task>

# sua code local
# chay local
python desktop_app.py

# neu can build local
# macOS: bash build/macos/build_macos.sh
# Windows: build\windows\build_portable.bat

# test xong, neu OK moi push
git add <file-can-thiet>
git commit -m "Mo ta thay doi"
git push origin feature/<ten-task>
```

Nguyen tac quan trong:

- Local build va local test truoc
- Khong push code chua test
- Khong commit file build, cache, state local
- Neu la thay doi cho updater/client, phai cap nhat payload va manifest


## 7B. Neu 2 may cung dang code thi dong bo the nao

Neu may A va may B cung dang sua app, muc tieu la:

- khong miss code cua nhau
- khong de 1 may giu code local qua lau
- khong push de len `main`

Quy tac:

1. Moi may lam tren branch rieng
2. Push thuong xuyen len GitHub
3. Truoc khi tiep tuc lam, phai `fetch` / `pull` code moi nhat
4. Truoc khi merge vao nhanh chung, phai cap nhat `main` vao branch hien tai

Flow de nghi:

May A:

```bash
git checkout main
git pull origin main
git checkout -b feature/may-a-task
```

May B:

```bash
git checkout main
git pull origin main
git checkout -b feature/may-b-task
```

Trong qua trinh lam:

```bash
git add <file-can-thiet>
git commit -m "WIP: phan viec da xong"
git push origin <branch-cua-minh>
```

Truoc khi merge:

```bash
git fetch origin
git checkout <branch-cua-minh>
git merge origin/main
```

Hoac:

```bash
git fetch origin
git checkout <branch-cua-minh>
git rebase origin/main
```

Nguyen tac de khong miss code:

- Khong sua truc tiep tren `main`
- Khong giu code local qua lau ma khong push
- Push commit nho, deu
- Truoc khi tiep tuc lam task cu, luon `fetch` code moi nhat
- Neu 2 may cung sua 1 file lon nhu `Log_checker.py`, nen bao nhau truoc


## 7C. Thao tac Git nao co the giao cho Codex

Codex co the lam an toan:

- `git status`
- `git checkout`
- `git branch`
- `git fetch`
- `git pull`
- `git add`
- `git commit`
- `git push`
- `git merge origin/main` vao branch hien tai
- giai thich conflict co ban

Nhung viec nen co nguoi xac nhan truoc:

- merge thang vao `main`
- force push
- `git reset --hard`
- xoa branch
- revert commit co anh huong den nguoi khac
- resolve conflict phuc tap

Rule de giao Git cho Codex:

- test local OK roi moi commit/push
- khong commit file build/cache/state local
- neu gap conflict, Codex nen dung lai va bao
- khong push thang `main` neu chua duoc yeu cau ro rang


## 8. Quy tac file nen va khong nen commit

Nen commit:

- `Log_checker.py`
- `desktop_app.py`
- `remote_update.py`
- `requirements.txt`
- file trong `build/`
- file trong `tools/`
- `Updates_2_5/remote_manifest.json`
- cac file README/huong dan

Khong nen commit:

- `__pycache__/`
- `*.pyc`
- file build da dong goi
- `dist/EventInspector.dmg`
- `dist/EventInspector.app`
- `dist/EventInspector/`
- `.venv/`
- `.worktrees/`

Repo da co `.gitignore` cho phan lon cac file local/build.


## 9. Build app

### 9.1 Build macOS

```bash
bash build/macos/build_macos.sh
```

Output:

- `dist/EventInspector.app`
- `dist/EventInspector.dmg`


### 9.2 Build Windows portable

Chay tren Windows:

```bat
build\windows\build_portable.bat
```

Output:

- `dist\EventInspector\`


### 9.3 Build Windows installer

Doc them:

- `build/windows/README_INSTALLER_WINDOWS.md`


## 9A. Khi nao can build local

Moi may dev co the build local tren chinh may do.

Can build local khi:

- muon test ban app dong goi
- muon test startup cua `.app`, `.dmg`, `.exe`
- muon kiem tra logic updater trong ban packaged

Khong can build local moi lan neu:

- chi sua logic Python/HTML va co the test bang `python desktop_app.py`

Quy tac thuc te:

- Sua nho: chay local truoc
- Sua lien quan package/update/restart: nen build local de test them
- Chi push khi local test da on


## 10. Remote update / payload update

App hien co co che update payload thong qua:

- `remote_update.py`
- `Updates_2_5/remote_manifest.json`

Khi sua mot tinh nang ma app can nhan qua updater:

1. Sua source chinh can thiet
2. Dong bo `sha256`, version va `payload_ref` trong `Updates_2_5/remote_manifest.json`; `payload_ref` phai la commit SHA 40 ky tu chua cac payload do
3. Tang version/badge neu can
4. Commit va push tren `main` sau khi test

Neu chi sua source local de dev, khong nhat thiet phai publish updater ngay.


## 11. Cach tiep tuc bang Codex tren may khac

Neu mo repo nay bang Codex o may khac, lam nhu sau:

1. Clone repo
2. Cai Python + dependency
3. Mo Codex tai root repo nay
4. Bao Codex doc cac file sau truoc khi sua:
   - `README_SETUP_TEAM.md`
   - `README_APP.md`
   - `requirements.txt`
   - `desktop_app.py`
   - `Log_checker.py`

Neu Codex can biet cach van hanh team:

- Code tren branch rieng
- Khong commit file build/cache local
- Neu can publish update, sua dung manifest/payload

Luu y:

- Trong may hien tai co file AGENTS tham chieu toi file local:
  - `/Users/truc.bui/.codex/RTK.md`
- Tren may khac co the khong co duong dan nay
- Neu Codex o may moi bao thieu file huong dan local, chi can tiep tuc dua tren repo nay va file `README_SETUP_TEAM.md`


## 12. Neu dev moi muon lay ban moi nhat

```bash
git checkout main
git pull origin main
```

Neu dang o branch rieng:

```bash
git fetch origin
git merge origin/main
```

Hoac:

```bash
git rebase origin/main
```


## 13. Xu ly xung dot co ban

Neu Git bao conflict:

1. Mo file bi conflict
2. Tim cac doan:

```text
<<<<<<<
=======
>>>>>>>
```

3. Chon noi dung dung
4. Xoa marker conflict
5. Sau do:

```bash
git add <file>
git commit
```


## 14. Checkpoint / revert

Repo nay co the dung tag checkpoint de quay lai moc on dinh.

Xem tag:

```bash
git tag
```

Xem commit hien tai:

```bash
git rev-parse --short HEAD
```

Checkout tam 1 checkpoint:

```bash
git checkout <tag-name>
```

Tao branch moi tu checkpoint:

```bash
git checkout -b hotfix/from-checkpoint <tag-name>
```


## 15. Checklist setup nhanh cho may moi

1. Clone repo
2. Tao `.venv`
3. `pip install -r requirements.txt`
4. Kiem tra `adb` neu test Android
5. `python desktop_app.py`
6. Tao branch rieng truoc khi sua
7. Commit va push branch len GitHub


## 16. Neu co van de

Neu app khong chay:

- kiem tra da activate `.venv` chua
- kiem tra `pip install -r requirements.txt` da xong chua
- kiem tra port local co dang bi chiem khong
- kiem tra `adb` / tool iOS co san khong

Neu updater khong dung:

- kiem tra `remote_manifest.json`
- kiem tra `sha256`
- kiem tra `payload_ref` co ton tai tren remote va cac URL payload deu tro vao cung commit SHA


## 17. Khuyen nghi cho team

De lam viec on dinh voi 2-3 nguoi:

- dung GitHub private repo
- moi task la 1 branch
- merge qua PR
- giu `main` sach va chay duoc
- ghi ro trong commit message la sua source local hay publish update payload


## 18. Harness engineering

Muc tieu cua harness:

- bat regression truoc khi push
- giu cac bug lap lai khong chui qua release
- khong can mo UI, khong can device, khong can build app

File chay harness:

```bash
python3 tools/eventchecker_harness.py --profile local
```

`--profile local` la check cho code dang phat trien: manifest duoc doi chieu
voi blob bat bien cua `payload_ref`, con cac gate doi worktree phai trung
manifest se khong lam fail local test chi vi dang co thay doi chua commit.

Truoc khi commit/push mot release, dong bo manifest tu dung commit se phat
hanh roi chay day du release gate:

```bash
python3 tools/sync_update_manifest.py --payload-ref <commit-sha> --version <release-version>
python3 tools/eventchecker_harness.py --profile release
```

Lenh `sync_update_manifest.py` chi cap nhat `Updates_2_5/remote_manifest.json`;
no tu loai cac preset/mapping khoi payload va khong build, commit hay push.

Harness dang check nhung phan de vo nhat:

- shape cua `remote_manifest.json`
- `sha256` trong manifest phai khop payload that
- URL payload trong manifest phai tro vao cung mot commit SHA bat bien
- mapping package -> game code
- state cua installation id khi doi package
- log contract exact-match cho Crashlytics Android va iOS
- release marker trong `Log_checker.py`

Neu muon chi chay tung nhom test:

```bash
python3 tools/eventchecker_harness.py --only test_installation_id_state_machine test_package_code_mapping
```

Nguyen tac su dung:

- chay harness truoc khi commit
- neu harness fail thi sua xong moi push
- neu sua `Updates_2_5/remote_manifest.json` hoac flow update thi harness la bat buoc
- harness khong thay cho local run app, ma la them 1 lop chan loi

Quy tac bat buoc cho du lieu lay tu Git:

- Test upload/download preset phai tao hoac sua mot file da duoc commit trong
  Git, roi cho client tai qua HTTP; khong duoc chi mock `requests.get` hoac
  chi kiem tra chuoi trong source.
- Test phai dung cung mot client/process cho hai lan tai: lan dau xac nhan gia
  tri commit cu, sua sang commit moi, bam Reload, roi xac nhan gia tri moi va
  `sha256`/`preset_revision` da thay doi.
- Neu test chi kiem tra cau truc ma khong chung minh gia tri moi den duoc
  client thi xem nhu chua dat, khong duoc push release.
- Test bat buoc: `python3 tools/eventchecker_harness.py --only
  test_services_checker_git_value_reload_from_real_commit
  test_sdk_preset_git_value_reload_from_real_commit`.
