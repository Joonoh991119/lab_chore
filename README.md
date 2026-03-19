# 🔬 lab_chore — CSNL 실험 행정 자동화

서울대학교 뇌인지과학과 CSNL 실험참여자비 양식 작성 및 업로드 자동화 도구입니다.

---

## 기능

| 도구 | 설명 |
|------|------|
| **실험참여자비GUI.app** | 참가자 정보 입력 → `실험참여자비 양식_이름.xlsx` 자동 생성 |
| **upload_updater.py** | 작성된 양식들을 읽어 업로드 양식에 행 자동 추가 |

---

## 설치 (최초 1회)

```bash
git clone https://github.com/joonop99/lab_chore.git
cd lab_chore
bash setup.sh
```

> Python 없는 경우: `brew install python` 먼저 실행

---

## GUI 사용법

1. `실험참여자비GUI.app`을 **Excel 템플릿 파일들과 같은 폴더**에 복사
2. 더블클릭 → 브라우저에서 GUI 자동 실행
3. 정보 입력 후 저장 → `실험참여자비 양식_이름.xlsx` 생성

```
📂 실험 데이터 폴더/
├── 실험참여자비 양식(중견).xlsx   ← 템플릿 (필수)
├── template_일회성경비지급자 업로드양식.xlsx
├── 실험참여자비GUI.app            ← 여기에 복사
└── 실험참여자비 양식_홍길동.xlsx   ← GUI가 생성
```

---

## 업로드 양식 자동화

```bash
# 폴더 내 모든 실험참여자비 양식_*.xlsx 일괄 처리
python3 실험참여자비GUI.app/Contents/Resources/upload_updater.py --all

# 특정 파일만
python3 실험참여자비GUI.app/Contents/Resources/upload_updater.py "실험참여자비 양식_홍길동.xlsx"
```

결과: `일회성경비지급자_업로드양식_작성.xlsx` (중복 방지, END 마커 자동 관리)

---

## 셀 매핑 (Sheet `17`)

| 셀 | 내용 |
|----|------|
| B16 | 이름 | D16 | 소속 | E16 | 주민등록번호 |
| F16 | 이메일 | G16 | 은행명 | I16 | 계좌번호 |
| L16 | 예금주 | D19 | 지급액 | C10 | 활용일자 |

---

## 개인정보 보호

`실험참여자비 양식_*.xlsx` 등 개인정보가 포함된 파일은 `.gitignore`에 의해 자동으로 git 추적에서 제외됩니다.

---

*CSNL · Dept. of Brain and Cognitive Sciences · Seoul National University*
