# Word to HWP 변환 시스템

워드 파일을 업로드하여 HWP 파일로 변환하는 웹 애플리케이션입니다.

## 시스템 요구사항

- Windows OS (한컴오피스 한글 2014 이상 설치 필요)
- Python 3.8 이상

## 빠른 시작

### 1. 의존성 설치 (최초 1회만)
```bash
install.bat
```

### 2. 서버 실행 (추천)
```bash
start.bat
```
메뉴에서 "2. 서버 시작 + 브라우저 열기" 선택

또는 간단하게:
```bash
start_server.bat
```

### 3. 브라우저 접속
http://localhost:8000

## 배치 파일 목록

| 파일 | 설명 |
|------|------|
| `start.bat` | 통합 메뉴 (서버 시작/종료, 설치 등) |
| `start_server.bat` | 서버만 시작 |
| `stop_server.bat` | 실행 중인 서버 종료 |
| `install.bat` | Python 의존성 설치 |
| `open_browser.bat` | 브라우저에서 앱 열기 |

## 프로젝트 구조

```
.
├── backend/           # Python FastAPI 백엔드
│   ├── main.py       # FastAPI 메인 애플리케이션
│   ├── hwp_converter.py  # 한컴오피스 COM 자동화
│   └── requirements.txt
├── frontend/         # 프론트엔드 (HTML/CSS/JS)
│   ├── index.html    # 메인 페이지
│   ├── style.css     # 다크 테마 스타일
│   └── app.js        # JavaScript
├── templates/        # 저장된 템플릿
├── uploads/          # 업로드된 워드 파일
├── outputs/          # 생성된 HWP 파일
├── start.bat         # 통합 실행 메뉴
├── start_server.bat  # 서버 시작
└── install.bat       # 의존성 설치
```

## 수동 설치 및 실행

### 1. Python 의존성 설치

```bash
cd backend
pip install -r requirements.txt
```

### 2. 서버 실행

```bash
cd backend
python main.py
```

### 3. 브라우저 접속
http://localhost:8000

## 기능

1. **워드 파일 업로드**: DOCX 파일을 업로드
2. **템플릿 저장**: 자주 사용하는 양식을 템플릿으로 저장
3. **HWP 변환**: 업로드된 워드 파일을 HWP로 변환
4. **템플릿 관리**: 저장된 템플릿 조회 및 사용

## 호환성

- 한글 2014, 2018, 2020, 2022 버전 지원
