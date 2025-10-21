# DOCX to HWP Converter (Java + hwplib)

Java 기반 DOCX to HWP 변환기 (hwplib 사용)

## 특징

- ✅ **Windows 독립적**: Hancom Office 설치 불필요
- ✅ **크로스 플랫폼**: Windows, Linux, macOS 모두 지원
- ✅ **안정적인 셀 병합**: `TableCellMerger`를 통한 완벽한 표 셀 병합
- ✅ **직접 파일 조작**: COM API 없이 HWP 파일 직접 생성
- ✅ **동일한 API**: 기존 Python 백엔드와 동일한 REST API 제공

## 요구사항

- Java 17 이상
- Maven 3.6 이상

## 빌드 및 실행

### 방법 1: Maven 사용

```bash
# 빌드
mvn clean package

# 실행
java -jar target/docx-to-hwp-converter-1.0.0.jar
```

### 방법 2: Maven 직접 실행

```bash
mvn spring-boot:run
```

## API 엔드포인트

기존 Python FastAPI와 동일한 엔드포인트 제공:

- `GET /api/health` - 헬스 체크
- `POST /api/upload` - DOCX 파일 업로드
- `POST /api/convert` - HWP로 변환
- `GET /api/download/{file_id}` - 변환된 HWP 다운로드
- `GET /` - 프론트엔드 index.html

## 의존성

- Spring Boot 3.2.0 - REST API 프레임워크
- hwplib 1.1.10 - HWP 파일 조작
- Apache POI 5.2.5 - DOCX 파일 읽기

## 프로젝트 구조

```
backend-java/
├── src/main/java/com/converter/
│   ├── Application.java           # Spring Boot 메인 클래스
│   ├── DocxToHwpConverter.java    # 핵심 변환 로직
│   ├── ConverterController.java   # REST API 컨트롤러
│   └── WebConfig.java             # 정적 파일 서빙 설정
├── src/main/resources/
│   └── application.properties     # Spring Boot 설정
└── pom.xml                        # Maven 의존성 설정
```

## 지원 기능

### ✅ 현재 지원
- 텍스트 변환
- 문단 단위 변환
- 표 변환 (행/열)
- 셀 병합 (가로/세로/2x2 블록)
- 기본 테두리 및 서식

### 🚧 추후 추가 예정
- 텍스트 색상 및 폰트
- 문단 정렬 (좌/중/우/양쪽)
- 들여쓰기 및 내어쓰기
- 셀 배경색
- 이미지 삽입
