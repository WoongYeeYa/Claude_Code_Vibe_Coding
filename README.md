# DOCX to HWP 변환 시스템

워드 파일을 업로드하여 HWP 파일로 변환하는 웹 애플리케이션입니다.

## 🚀 빠른 시작 (NEW! Java 버전)

### 한 번에 설치 및 실행 (권장)

```bash
빠른_시작.bat
```

이 스크립트가 자동으로:
- ✅ Java 17 설치
- ✅ Maven 설치
- ✅ 프로젝트 빌드
- ✅ 서버 시작
- ✅ 브라우저 열기

**끝!** 더 이상 Hancom Office 설치가 필요 없습니다!

### 장점
- ✅ **Windows 독립적**: Hancom Office 불필요
- ✅ **크로스 플랫폼**: Windows, Linux, macOS 모두 지원
- ✅ **안정적인 셀 병합**: TableCellMerger로 완벽한 표 처리
- ✅ **빠른 변환**: COM API 없이 직접 파일 조작

자세한 내용: `설치_가이드.md` 참고

---

## 📝 시스템 요구사항

### Java 버전 (권장)
- ✅ Java 17 이상
- ✅ Maven 3.6 이상
- ❌ Hancom Office 불필요

### Python 버전 (레거시)
- Windows OS
- Hancom Office 한글 2014 이상
- Python 3.8 이상

---

## 🎯 빠른 시작 (상세)

### 방법 1: Java 버전 (권장)

```bash
# 1. 자동 설치 및 실행
빠른_시작.bat

# 또는 수동으로:
cd backend-java
build.bat      # 빌드 (최초 1회)
run.bat        # 실행
```

브라우저 접속: http://localhost:8000

### 방법 2: Python 버전 (레거시)

```bash
# 1. 의존성 설치 (최초 1회)
install.bat

# 2. 서버 시작
start.bat      # 메뉴에서 선택
# 또는
start_server.bat

# 3. 브라우저 접속
http://localhost:8000
```

## 📂 프로젝트 구조

```
.
├── backend-java/          # ⭐ Java 백엔드 (권장)
│   ├── src/main/java/com/converter/
│   │   ├── Application.java          # Spring Boot 메인
│   │   ├── DocxToHwpConverter.java   # 변환 로직 (hwplib)
│   │   └── ConverterController.java  # REST API
│   ├── pom.xml           # Maven 의존성
│   ├── build.bat         # 빌드 스크립트
│   └── run.bat           # 실행 스크립트
│
├── backend/              # Python 백엔드 (레거시)
│   ├── main.py          # FastAPI
│   └── hwp_converter.py # COM API
│
├── frontend/            # 프론트엔드 (공통)
│   ├── index.html
│   ├── style.css
│   └── app.js
│
├── hwplib/              # hwplib 라이브러리 소스
│
├── legacy-python/       # 📦 Python 레거시 버전 (백업)
│   ├── backend/        # Python 백엔드 (복사본)
│   ├── frontend/       # 프론트엔드 (복사본)
│   ├── *.bat           # Python 실행 스크립트들
│   ├── test_*.py       # 테스트 파일들
│   └── README.md       # 레거시 버전 설명
│
├── uploads/             # 업로드 파일
├── outputs/             # 변환 결과
├── templates/           # 템플릿
│
├── 빠른_시작.bat        # ⭐ 자동 설치 및 실행
├── 설치_가이드.md       # 설치 가이드
├── JAVA_MIGRATION_GUIDE.md  # 마이그레이션 가이드
└── 현재_기능_정리.md    # 기능 정리
```

## 📜 배치 파일 목록

### Java 버전
| 파일 | 설명 |
|------|------|
| `빠른_시작.bat` | ⭐ 자동 설치 및 실행 (권장) |
| `backend-java/build.bat` | 프로젝트 빌드 |
| `backend-java/run.bat` | 서버 실행 |

### Python 버전 (레거시)
| 파일 | 설명 |
|------|------|
| `start.bat` | 통합 메뉴 |
| `start_server.bat` | 서버 시작 |
| `stop_server.bat` | 서버 종료 |
| `install.bat` | 의존성 설치 |

## 🎨 기능

### Java 버전 (hwplib)
1. ✅ **DOCX 파일 업로드**: 워드 파일 업로드
2. ✅ **HWP 변환**: DOCX → HWP 변환
3. ✅ **표 변환**: 행/열 자동 변환
4. ✅ **셀 병합**: 가로/세로/블록 병합 지원 ⭐
5. ✅ **텍스트 변환**: 문단 및 텍스트 변환
6. 🚧 **텍스트 서식**: 색상, 폰트, 정렬 (추후 추가)
7. 🚧 **셀 배경색** (추후 추가)

### Python 버전 (레거시)
1. ✅ **워드 파일 업로드**: DOCX 파일 업로드
2. ✅ **템플릿 저장**: 자주 사용하는 양식 저장
3. ✅ **HWP 변환**: COM API 기반 변환
4. ✅ **템플릿 관리**: 템플릿 조회 및 사용
5. ❌ **셀 병합**: 지원 안 됨
6. ❌ **셀 배경색**: 지원 안 됨

---

## 🔧 트러블슈팅

문제가 발생하면:
1. `설치_가이드.md` 참고
2. `JAVA_MIGRATION_GUIDE.md` 참고
3. 로그 확인

---

## 📚 문서

- `README.md` (현재 파일) - 프로젝트 개요
- `설치_가이드.md` - 설치 및 실행 가이드
- `JAVA_MIGRATION_GUIDE.md` - Python → Java 마이그레이션 가이드
- `현재_기능_정리.md` - 완성된 기능 정리
- `backend-java/README.md` - Java 백엔드 상세 문서

---

## 💡 추천 사용 방법

### 신규 사용자
```bash
빠른_시작.bat  # 자동으로 모든 것을 설치하고 실행
```

### 개발자
```bash
cd backend-java
build.bat      # 빌드
run.bat        # 실행
```

---

## 🆚 버전 비교

| 기능 | Python (레거시) | Java (권장) |
|------|----------------|-------------|
| 셀 병합 | ❌ | ✅ |
| 셀 배경색 | ❌ | 🚧 |
| Hancom Office 필요 | ✅ | ❌ |
| 크로스 플랫폼 | ❌ | ✅ |
| 속도 | 느림 | 빠름 |
| 안정성 | 낮음 | 높음 |

---

## 📄 라이선스

- hwplib: Apache 2.0
- Apache POI: Apache 2.0
- Spring Boot: Apache 2.0

---

## 📦 Python 레거시 버전

Python + COM API 기반 구버전이 `legacy-python/` 폴더에 보관되어 있습니다.

### 레거시 버전 사용 (권장하지 않음)

```bash
cd legacy-python
install.bat          # 의존성 설치
start_server.bat     # 서버 시작
```

**주의:**
- ❌ 셀 병합 지원 안 됨
- ❌ 셀 배경색 지원 안 됨
- ⚠️ Hancom Office 설치 필수
- ⚠️ Windows 전용

**권장:** 대신 새 Java 버전 사용 (`빠른_시작.bat`)

자세한 내용: `legacy-python/README.md` 참고
