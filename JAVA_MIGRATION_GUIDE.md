# Java 백엔드 마이그레이션 가이드

## 개요

Python COM API 기반 변환기를 Java hwplib 기반으로 재작성했습니다.

## 왜 Java로 변경했나?

### Python COM API의 한계
- ❌ Windows 전용 (Hancom Office 설치 필수)
- ❌ 셀 병합 실패 (`HTableCellBlock` 오류)
- ❌ 셀 배경색 설정 실패 (서버 예외 오류)
- ❌ 불안정한 백그라운드 프로세스
- ❌ 느린 변환 속도

### Java hwplib의 장점
- ✅ 크로스 플랫폼 (Windows, Linux, macOS)
- ✅ Hancom Office 설치 불필요
- ✅ 완벽한 셀 병합 (`TableCellMerger.mergeCell`)
- ✅ 직접 파일 조작 (빠르고 안정적)
- ✅ 오픈소스 라이브러리 (Apache 2.0)

## 프로젝트 구조

```
Claude_Code_Vibe_Coding/
├── backend/                 # 기존 Python 백엔드 (폐기 예정)
├── backend-java/            # 새로운 Java 백엔드 ⭐
│   ├── src/main/java/com/converter/
│   │   ├── Application.java
│   │   ├── DocxToHwpConverter.java
│   │   ├── ConverterController.java
│   │   └── WebConfig.java
│   ├── src/main/resources/
│   │   └── application.properties
│   ├── pom.xml
│   ├── build.bat
│   ├── run.bat
│   └── README.md
├── frontend/                # 프론트엔드 (변경 없음)
│   ├── index.html
│   ├── app.js
│   └── style.css
└── hwplib/                  # hwplib 소스코드 (참고용)
```

## 설치 및 실행

### 1. 사전 요구사항

#### Java 17 설치
1. [Oracle JDK 17](https://www.oracle.com/java/technologies/javase/jdk17-archive-downloads.html) 또는 [OpenJDK 17](https://adoptium.net/) 다운로드
2. 설치 후 환경 변수 설정:
   ```
   JAVA_HOME=C:\Program Files\Java\jdk-17
   Path에 %JAVA_HOME%\bin 추가
   ```
3. 확인:
   ```bash
   java -version
   # java version "17.0.x" 출력되어야 함
   ```

#### Maven 3.6+ 설치
1. [Apache Maven](https://maven.apache.org/download.cgi) 다운로드 (Binary zip archive)
2. 압축 해제 (예: `C:\Program Files\Apache\maven`)
3. 환경 변수 설정:
   ```
   MAVEN_HOME=C:\Program Files\Apache\maven
   Path에 %MAVEN_HOME%\bin 추가
   ```
4. 확인:
   ```bash
   mvn -version
   # Apache Maven 3.x.x 출력되어야 함
   ```

### 2. 빌드 및 실행

#### 방법 1: 배치 파일 사용 (권장)

```bash
# 빌드
cd backend-java
build.bat

# 실행
run.bat
```

#### 방법 2: Maven 명령어 직접 사용

```bash
cd backend-java

# 빌드
mvn clean package

# 실행
mvn spring-boot:run
```

### 3. 서버 접속

서버가 시작되면:
- 프론트엔드: http://localhost:8000
- API 문서: http://localhost:8000/api/health

## API 엔드포인트 (기존과 동일)

| Method | Endpoint | 설명 |
|--------|----------|------|
| GET | `/api/health` | 헬스 체크 |
| POST | `/api/upload` | DOCX 파일 업로드 |
| POST | `/api/convert` | HWP로 변환 |
| GET | `/api/download/{file_id}` | HWP 다운로드 |
| GET | `/` | 프론트엔드 |

### 응답 형식 예시

#### Health Check
```json
{
  "status": "ok",
  "converter_type": "hwplib (Java)",
  "hwp_installed": "not required",
  "hwp_version": "hwplib 1.1.10"
}
```

#### Upload
```json
{
  "success": true,
  "file_id": "123e4567-e89b-12d3-a456-426614174000",
  "original_filename": "test.docx",
  "saved_filename": "123e4567-e89b-12d3-a456-426614174000.docx",
  "message": "파일이 성공적으로 업로드되었습니다"
}
```

## 지원 기능

### ✅ 현재 지원 (v1.0.0)
- 텍스트 추출 및 변환
- 문단 단위 변환
- 표 변환 (행/열)
- **셀 병합 (가로/세로/블록)** ⭐ NEW
- 기본 테두리 및 서식

### 🚧 추후 추가 예정
- 텍스트 색상 (빨강, 파랑, 녹색 등)
- 폰트 크기 및 스타일
- 문단 정렬 (좌/중/우/양쪽)
- 들여쓰기 및 내어쓰기
- 셀 배경색
- 이미지 삽입

## 테스트

### 간단한 표 테스트

```bash
# 테스트 파일 준비 (이미 존재)
test_simple_table.docx

# 서버 실행 후 프론트엔드에서 업로드 및 변환
```

### 복잡한 표 테스트 (셀 병합)

```bash
# 테스트 파일 준비 (이미 존재)
test_complex_table.docx

# 서버 실행 후 프론트엔드에서 업로드 및 변환
# 결과: 셀 병합이 정상적으로 처리됨 ✅
```

## 트러블슈팅

### Maven 빌드 실패
```bash
# 의존성 강제 업데이트
mvn clean install -U
```

### Java 버전 오류
```bash
# Java 버전 확인
java -version

# 17 미만이면 JDK 17 설치 필요
```

### Port 8000 already in use
```bash
# 기존 Python 서버 종료
stop_server.bat

# 또는 다른 포트 사용 (application.properties에서 변경)
server.port=8080
```

## 성능 비교

| 항목 | Python COM API | Java hwplib |
|------|----------------|-------------|
| 변환 속도 | 느림 (백그라운드 프로세스) | 빠름 (직접 파일 조작) |
| 셀 병합 | ❌ 실패 | ✅ 성공 |
| 셀 배경색 | ❌ 실패 | 🚧 구현 예정 |
| 메모리 사용 | 높음 (HWP 프로세스) | 낮음 |
| 안정성 | 낮음 | 높음 |
| 크로스 플랫폼 | ❌ Windows 전용 | ✅ 모든 OS |

## 다음 단계

1. ✅ Java 백엔드 완성
2. 🔄 프론트엔드 연동 테스트
3. 📋 불필요한 Python 코드 정리
4. 🎨 텍스트 서식 기능 추가
5. 🖼️ 이미지 변환 기능 추가

## 참고 자료

- [hwplib GitHub](https://github.com/neolord0/hwplib)
- [Apache POI Documentation](https://poi.apache.org/)
- [Spring Boot Documentation](https://spring.io/projects/spring-boot)
