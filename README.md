# 🔬 lab_chore — CSNL 실험 행정 자동화

서울대학교 뇌인지과학과 CSNL 실험참여자비 양식 작성 및 업로드 자동화 도구입니다.

---

## 📦 기능

| 탭 | 기능 |
|----|------|
| **📝 참가자 양식 작성** | 이름·주민번호·계좌·참여시간·전자서명 입력 → `실험참여자비 양식_이름.xlsx` 자동 생성 |
| **📊 업로드 양식 생성** | 폴더 내 모든 참가자 양식을 읽어 `일회성경비지급자_업로드양식_작성.xlsx` 한 파일로 통합 |

---

## 🚀 설치 (최초 1회만)

### 방법 A — 다운로드 (권장, git 불필요)

1. 이 페이지 오른쪽 상단 **`<> Code` → `Download ZIP`** 클릭
2. 압축 해제 후 폴더 안으로 이동
3. 터미널에서 실행:

```bash
bash setup.sh
```

### 방법 B — git clone

```bash
git clone https://github.com/Joonoh991119/lab_chore.git
cd lab_chore
bash setup.sh
```

> **Python이 없는 경우** `setup.sh` 실행 전 한 번만:
> ```bash
> brew install python
> ```
> Homebrew가 없으면 → [brew.sh](https://brew.sh) 에서 설치

`setup.sh` 가 하는 일:
- Python 탐색 (pyenv / brew / system 순서)
- `streamlit`, `openpyxl`, `streamlit-drawable-canvas`, `Pillow` 자동 설치
- `.app` 실행 권한 설정 및 macOS Gatekeeper 우회 처리

---

## 🖥️ 사용법

### 1. .app을 템플릿 폴더에 복사

```
📂 실험 데이터 폴더/
├── 실험참여자비 양식(중견).xlsx          ← 템플릿 (필수)
├── template_일회성경비지급자 업로드양식.xlsx
├── 실험참여자비GUI.app                  ← 여기에 복사
└── ...
```

### 2. 더블클릭으로 실행

`실험참여자비GUI.app` 더블클릭 → 브라우저에서 GUI 자동 오픈

> 처음 실행 시 패키지 자동 설치(약 30~60초), 이후부터는 즉시 실행됩니다.

### 3. 탭 1 — 참가자 양식 작성

- 상단 **📂 찾기** 버튼으로 폴더 선택 (또는 경로 직접 입력)
- 참가자 정보 입력 → 캔버스에 전자서명 → **저장** 클릭
- `실험참여자비 양식_이름.xlsx` 가 같은 폴더에 생성됩니다

### 4. 탭 2 — 업로드 양식 생성

- 폴더 내 `실험참여자비 양식_*.xlsx` 파일을 자동 스캔
- 신규/기등록 여부를 미리보기 테이블로 확인
- **업로드 양식 생성** 클릭 → `일회성경비지급자_업로드양식_작성.xlsx` 생성
- 기존 행의 폰트·테두리·정렬이 그대로 유지됩니다

---

## ⚙️ 고급 — 터미널에서 직접 실행

```bash
# GUI (streamlit)
cd /path/to/templates/folder
streamlit run /path/to/lab_chore/실험참여자비GUI.app/Contents/Resources/app.py

# 업로드 양식 CLI (일괄 처리)
LAB_CHORE_DIR=/path/to/templates/folder \
  python3 실험참여자비GUI.app/Contents/Resources/upload_updater.py --all
```

---

## 🔧 필요 환경

| 항목 | 버전 |
|------|------|
| macOS | 12 Monterey 이상 |
| Python | 3.9 이상 (brew / pyenv 모두 지원) |
| streamlit | ≥ 1.32 |
| openpyxl | ≥ 3.1 |
| streamlit-drawable-canvas | ≥ 0.9.3 |
| Pillow | ≥ 10.0 |

---

## ⚠️ 개인정보 보호

`실험참여자비 양식_*.xlsx`, `일회성경비지급자_업로드양식_작성.xlsx` 등 개인정보 파일은 `.gitignore`에 의해 **자동으로 git 추적에서 제외**됩니다.

---

*CSNL · Dept. of Brain and Cognitive Sciences · Seoul National University*
