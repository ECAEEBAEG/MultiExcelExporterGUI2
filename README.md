# 📊 MultiExcelExporterGUI2 (PowerShell Excel 일괄 처리 도구)

## 📌 개요
이 PowerShell 스크립트는 Excel 파일을 일괄 처리하기 위한 GUI 기반 자동화 도구입니다.

주요 기능:
- 입력/출력 폴더를 GUI로 선택
- 폴더 내 모든 Excel(.xlsx) 파일 처리
- 각 파일의 모든 시트를 새로운 파일로 복사
- 진행 상태를 실시간 ProgressBar로 표시
- 결과 파일을 `PSC_` 접두어로 저장

---

## ⚙️ 주요 기능

### 🗂 1. 폴더 선택 (GUI)
- Windows 폼 기반 폴더 선택 창 제공
- 원본 Excel 파일이 있는 폴더 선택
- 결과 파일 저장 폴더 선택

---

### 📁 2. Excel 파일 처리
- 선택한 폴더의 `.xlsx` 파일 전체 조회
- Microsoft Excel COM 객체를 이용하여 파일 열기
- 모든 시트를 새 워크북으로 복사
- 결과 파일 저장

저장 형식: PSC_<원본파일명>.xlsx


---

### 📊 3. 진행 상황 UI (ProgressBar)
Windows Forms UI를 통해 다음 정보를 실시간 표시:
- 현재 처리 중인 파일 이름
- 전체 진행률 (%)
- ProgressBar 상태

---

## 🧠 처리 로직 (핵심 동작)

각 Excel 파일에 대해 다음 순서로 처리됩니다:

1. Excel 파일 열기 (COM)
2. 새 워크북 생성
3. 기본 시트 삭제
4. 원본 모든 시트 복사
5. 새 파일로 저장
6. 파일 닫기

---

## 🖥 실행 환경

### ✔️ 필수 환경
- Windows 운영체제
- Microsoft Excel 설치 필수
- PowerShell 5.1 이상

### ✔️ 필요 구성
- Excel COM Object (`Excel.Application`)
- .NET WinForms 라이브러리
  - `System.Windows.Forms`
  - `System.Drawing`

---

## 🚀 실행 방법

```powershell
powershell.exe -STA -File "SingleFileExporter.ps1"

⚠️ 중요:
GUI 사용을 위해 반드시 -STA 모드로 실행해야 합니다.

## 전체 흐름도

스크립트 시작
   ↓
STA 모드 확인
   ↓
원본 폴더 선택
   ↓
저장 폴더 선택
   ↓
Excel 파일 목록 로드
   ↓
각 파일 반복 처리
   ├─ Excel 열기
   ├─ 시트 복사
   ├─ 새 파일 저장 (PSC_ 접두어)
   ├─ ProgressBar 업데이트
   ↓
Excel 종료
   ↓
완료 메시지 출력
