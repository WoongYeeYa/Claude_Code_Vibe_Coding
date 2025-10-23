# DOCX to HWP 변환기

Apache POI와 hwplib을 사용하여 DOCX 파일을 HWP 파일로 변환하는 웹 애플리케이션입니다.

## 🚀 빠른 시작

### 1️⃣ 최초 설치 (한 번만)
```
install.bat 더블클릭
```
hwplib과 백엔드 서버를 자동으로 빌드합니다.

### 2️⃣ 서버 시작
```
start_server.bat 더블클릭
```

### 3️⃣ 브라우저에서 접속
```
http://localhost:8000
```

### 4️⃣ 서버 종료
서버 창에서 `Ctrl + C` 또는 `stop_server.bat` 실행

---

## 📁 실행 파일

| 파일 | 설명 |
|------|------|
| **install.bat** | 최초 설치 (hwplib + 백엔드 빌드) |
| **build.bat** | 백엔드만 다시 빌드 |
| **start_server.bat** | 서버 시작 |
| **stop_server.bat** | 서버 종료 |

---

## 🔧 기술 스택

### Backend
- **Java 17**
- **Spring Boot 3.2.0**
- **Apache POI 5.2.5** - DOCX 파싱
- **hwplib 1.1.10** - HWP 생성
- **Maven** - 빌드 도구

### Frontend
- **HTML5 / CSS3 / JavaScript**

---

## ✅ 현재 지원 기능

- ✅ 기본 텍스트 변환 (한글 완벽 지원)
- ✅ 테이블 구조 변환 (행, 열, 셀)
- ✅ 셀 내용 추출 및 렌더링
- ✅ 파일 업로드/다운로드
- ✅ REST API
- ✅ 상세 로그

---

## 🔜 향후 개선 예정

- [ ] **텍스트 스타일** - 굵기, 이탤릭, 색상, 크기
- [ ] **테이블 스타일** - 테두리, 배경색, 셀 병합
- [ ] **이미지** - 이미지 추출 및 삽입
- [ ] **복잡한 서식** - 머리글, 바닥글, 페이지 번호
- [ ] **글꼴 매핑** - DOCX 폰트 → HWP 폰트 자동 변환

---

## 🌐 API 엔드포인트

| Method | Endpoint | 설명 |
|--------|----------|------|
| GET | `/api/health` | 서버 상태 확인 |
| POST | `/api/upload` | DOCX 파일 업로드 |
| POST | `/api/convert` | DOCX → HWP 변환 |
| GET | `/api/download/{fileId}` | HWP 파일 다운로드 |

### API 사용 예시

```bash
# 1. 서버 상태 확인
curl http://localhost:8000/api/health

# 2. 파일 업로드
curl -X POST -F "file=@document.docx" http://localhost:8000/api/upload

# 3. 변환 (file_id는 업로드 응답에서 받음)
curl -X POST -d "file_id=abc-123" http://localhost:8000/api/convert

# 4. 다운로드
curl -o output.hwp http://localhost:8000/api/download/abc-123
```

---

## 📂 프로젝트 구조

```
Claude_Code_Vibe_Coding/
├── backend-java/              # Spring Boot 백엔드
│   ├── src/main/java/com/converter/
│   │   ├── Application.java           # 메인 애플리케이션
│   │   ├── controller/
│   │   │   └── ConverterController.java   # REST API
│   │   ├── service/
│   │   │   └── DocxToHwpConverter.java    # 변환 로직
│   │   └── config/
│   │       └── WebConfig.java             # 웹 설정
│   ├── src/main/resources/
│   │   └── application.properties     # 서버 설정
│   └── pom.xml                        # Maven 의존성
│
├── frontend/                  # 프론트엔드
│   ├── index.html            # 메인 페이지
│   ├── app.js                # JavaScript
│   └── style.css             # 스타일
│
├── hwplib-main/              # HWP 라이브러리
│
├── install.bat               # 최초 설치
├── build.bat                 # 빌드
├── start_server.bat          # 서버 시작
├── stop_server.bat           # 서버 종료
└── README.md                 # 이 파일
```

---

## ⚙️ 수동 실행 방법

### Maven 빌드
```bash
cd backend-java
"C:\Program Files\apache-maven-3.9.11\bin\mvn.cmd" clean package -DskipTests
```

### 서버 실행
```bash
cd backend-java
"C:\Program Files\Java\jdk-17\bin\java.exe" -jar target\docx-to-hwp-converter-1.0.0.jar
```

---

## 📝 로그 확인

- **위치**: `backend-java/logs/converter.log`
- **콘솔 로그**: 서버 실행 창에서 실시간 확인

---

## 🐛 문제 해결

### ❌ install.bat 실패
- Java 17이 설치되어 있는지 확인
- Maven이 설치되어 있는지 확인
- 경로가 올바른지 확인

### ❌ 서버 시작 실패
- JAR 파일이 빌드되었는지 확인: `backend-java\target\docx-to-hwp-converter-1.0.0.jar`
- 포트 8000이 사용 중인지 확인: `netstat -ano | findstr :8000`

### ❌ 변환 실패
- 업로드한 파일이 DOCX 형식인지 확인
- 서버 로그 확인: `backend-java\logs\converter.log`
- 파일 크기 제한: 최대 50MB

---

## 💡 팁

- **서버 자동 재시작**: 코드 수정 후 `build.bat` 실행 → 서버 재시작
- **포트 변경**: `backend-java/src/main/resources/application.properties`에서 `server.port` 수정
- **로그 레벨 변경**: 같은 파일에서 `logging.level` 수정

---

## 📄 라이센스

이 프로젝트는 개발 중입니다.

---

## 🤝 기여

버그 리포트나 기능 제안은 환영합니다!
