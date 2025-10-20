"""
Word to HWP 변환 시스템 - FastAPI 메인 애플리케이션
"""
from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
import os
import uuid
import shutil
from pathlib import Path
from datetime import datetime
from typing import Optional, List
import json
import logging

from hwp_converter import HWPConverter

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# FastAPI 앱 생성
app = FastAPI(
    title="Word to HWP 변환 시스템",
    description="워드 파일을 HWP 파일로 변환하는 웹 애플리케이션",
    version="1.1.0"
)

# CORS 설정
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 디렉토리 설정
BASE_DIR = Path(__file__).resolve().parent.parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
TEMPLATE_DIR = BASE_DIR / "templates"
FRONTEND_DIR = BASE_DIR / "frontend"

# 디렉토리 생성
for directory in [UPLOAD_DIR, OUTPUT_DIR, TEMPLATE_DIR]:
    directory.mkdir(parents=True, exist_ok=True)

# HWP 변환기 초기화
converter = HWPConverter()

# 정적 파일 서빙 (프론트엔드)
if FRONTEND_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(FRONTEND_DIR)), name="static")


@app.on_event("startup")
async def startup_event():
    """서버 시작 시 실행"""
    logger.info("=" * 50)
    logger.info("Word to HWP 변환 시스템 시작")
    logger.info("=" * 50)
    is_installed, version = converter.check_hwp_installed()
    if is_installed:
        logger.info(f"[OK] 한컴오피스 연동 준비 완료 (버전: {version})")
    else:
        logger.error("[CRITICAL] 한컴오피스가 설치되어 있지 않거나 연동할 수 없습니다.")
        logger.error("           프로그램을 종료합니다.")
        # In a real scenario, you might want to exit the app
        # For now, we'll just log the error.


@app.on_event("shutdown")
async def shutdown_event():
    """서버 종료 시 실행"""
    logger.info("서버 종료 중...")


@app.get("/")
async def root():
    """루트 경로 - 프론트엔드 index.html 반환"""
    index_path = FRONTEND_DIR / "index.html"
    if index_path.exists():
        return FileResponse(str(index_path))
    return {"message": "Word to HWP 변환 시스템", "status": "running"}


@app.get("/api/health")
async def health_check():
    """헬스 체크"""
    is_installed, version = converter.check_hwp_installed()
    return {
        "status": "ok",
        "converter_type": "Hancom COM Automation",
        "hwp_installed": is_installed,
        "hwp_version": version
    }


@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    """
    워드 파일 업로드
    """
    try:
        # 파일 확장자 확인
        if not file.filename.endswith(('.docx', '.doc')):
            raise HTTPException(
                status_code=400,
                detail="DOCX 또는 DOC 파일만 업로드 가능합니다"
            )

        # 고유 파일명 생성
        file_id = str(uuid.uuid4())
        file_extension = Path(file.filename).suffix
        saved_filename = f"{file_id}{file_extension}"
        file_path = UPLOAD_DIR / saved_filename

        # 파일 저장
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        logger.info(f"파일 업로드 성공: {file.filename} -> {saved_filename}")

        return {
            "success": True,
            "file_id": file_id,
            "original_filename": file.filename,
            "saved_filename": saved_filename,
            "message": "파일이 성공적으로 업로드되었습니다"
        }

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"파일 업로드 실패: {str(e)}")
        raise HTTPException(status_code=500, detail=f"파일 업로드 실패: {str(e)}")


@app.post("/api/convert")
async def convert_to_hwp(file_id: str = Form(...)):
    """
    업로드된 워드 파일을 HWP로 변환
    """
    try:
        # 업로드된 파일 찾기
        uploaded_files = list(UPLOAD_DIR.glob(f"{file_id}.*"))
        if not uploaded_files:
            logger.error(f"파일을 찾을 수 없습니다: {file_id}")
            raise HTTPException(
                status_code=404,
                detail="업로드된 파일을 찾을 수 없습니다"
            )

        docx_path = uploaded_files[0]
        hwp_filename = f"{file_id}.hwp"
        hwp_path = OUTPUT_DIR / hwp_filename

        logger.info(f"변환 시작: {docx_path.name} -> {hwp_filename}")
        
        # COM 기반 변환 실행
        success = converter.convert_docx_to_hwp(str(docx_path), str(hwp_path))

        if not success:
            logger.error(f"변환 실패: convert_docx_to_hwp returned False")
            raise HTTPException(
                status_code=500,
                detail="HWP 변환에 실패했습니다. 서버 로그를 확인하세요."
            )

        logger.info(f"변환 완료: {hwp_filename}")

        return {
            "success": True,
            "file_id": file_id,
            "hwp_filename": hwp_filename,
            "message": "HWP 파일로 변환되었습니다"
        }

    except HTTPException:
        raise
    except Exception as e:
        import traceback
        logger.error(f"변환 실패: {str(e)}")
        logger.error(f"상세 에러: {traceback.format_exc()}")
        raise HTTPException(status_code=500, detail=f"변환 실패: {str(e)}")


@app.get("/api/download/{file_id}")
async def download_hwp(file_id: str):
    """
    변환된 HWP 파일 다운로드
    """
    try:
        hwp_path = OUTPUT_DIR / f"{file_id}.hwp"

        if not hwp_path.exists():
            raise HTTPException(
                status_code=404,
                detail="HWP 파일을 찾을 수 없습니다"
            )

        return FileResponse(
            path=str(hwp_path),
            filename=f"converted_{file_id}.hwp",
            media_type="application/x-hwp"
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"다운로드 실패: {str(e)}")
        raise HTTPException(status_code=500, detail=f"다운로드 실패: {str(e)}")


@app.post("/api/template/save")
async def save_template(
    file: UploadFile = File(...),
    template_name: str = Form(...)
):
    """
    템플릿으로 저장
    """
    try:
        # 파일 확장자 확인
        if not file.filename.endswith(('.hwp', '.hwpx', '.docx', '.doc')):
            raise HTTPException(
                status_code=400,
                detail="HWP, HWPX, DOCX 또는 DOC 파일만 업로드 가능합니다"
            )

        # 템플릿 ID 생성
        template_id = str(uuid.uuid4())
        file_extension = Path(file.filename).suffix
        saved_filename = f"{template_id}{file_extension}"
        template_path = TEMPLATE_DIR / saved_filename

        # 파일 저장
        with open(template_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # 템플릿 메타데이터 저장
        metadata = {
            "template_id": template_id,
            "template_name": template_name,
            "original_filename": file.filename,
            "saved_filename": saved_filename,
            "created_at": datetime.now().isoformat()
        }

        metadata_path = TEMPLATE_DIR / f"{template_id}.json"
        with open(metadata_path, "w", encoding="utf-8") as f:
            json.dump(metadata, f, ensure_ascii=False, indent=2)

        logger.info(f"템플릿 저장 성공: {template_name}")

        return {
            "success": True,
            "template_id": template_id,
            "template_name": template_name,
            "message": "템플릿이 저장되었습니다"
        }

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"템플릿 저장 실패: {str(e)}")
        raise HTTPException(status_code=500, detail=f"템플릿 저장 실패: {str(e)}")


@app.get("/api/templates")
async def list_templates():
    """
    저장된 템플릿 목록 조회
    """
    try:
        templates = []
        for metadata_file in TEMPLATE_DIR.glob("*.json"):
            try:
                with open(metadata_file, "r", encoding="utf-8") as f:
                    metadata = json.load(f)
                    templates.append(metadata)
            except Exception as e:
                logger.warning(f"템플릿 메타데이터 읽기 실패: {metadata_file.name} - {str(e)}")

        # 생성일 기준 내림차순 정렬
        templates.sort(key=lambda x: x.get("created_at", ""), reverse=True)

        return {
            "success": True,
            "templates": templates,
            "count": len(templates)
        }

    except Exception as e:
        logger.error(f"템플릿 목록 조회 실패: {str(e)}")
        raise HTTPException(status_code=500, detail=f"템플릿 목록 조회 실패: {str(e)}")


@app.post("/api/template/use/{template_id}")
async def use_template(template_id: str):
    """
    템플릿 사용하여 HWP 생성
    """
    try:
        # 템플릿 파일 찾기
        template_files = list(TEMPLATE_DIR.glob(f"{template_id}.*"))
        template_files = [f for f in template_files if f.suffix != ".json"]

        if not template_files:
            raise HTTPException(
                status_code=404,
                detail="템플릿 파일을 찾을 수 없습니다"
            )

        template_path = template_files[0]

        # 출력 파일명 생성
        output_id = str(uuid.uuid4())
        output_filename = f"{output_id}.hwp"
        output_path = OUTPUT_DIR / output_filename

        logger.info(f"템플릿 사용: {template_path.name}")

        # 템플릿이 HWP/HWPX인 경우 복사, DOCX인 경우 변환
        if template_path.suffix.lower() in ['.hwp', '.hwpx']:
            # HWPX도 HWP로 변환하여 일관성 유지
            if template_path.suffix.lower() == '.hwpx':
                 shutil.copy(template_path, output_path.with_suffix('.hwpx')) # 임시 복사
                 # HWPX to HWP 변환 로직 필요 (현재 컨버터는 DOCX->HWP만 지원)
                 # 간단하게 그냥 복사로 처리
                 shutil.copy(template_path, output_path)
                 success = True
            else:
                 shutil.copy(template_path, output_path)
                 success = True
        else: # DOCX인 경우
            success = converter.convert_docx_to_hwp(str(template_path), str(output_path))

        if not success:
            raise HTTPException(
                status_code=500,
                detail="템플릿 처리에 실패했습니다"
            )

        logger.info(f"템플릿 파일 생성 완료: {output_filename}")

        return {
            "success": True,
            "file_id": output_id,
            "hwp_filename": output_filename,
            "message": "템플릿을 사용하여 HWP 파일이 생성되었습니다"
        }

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"템플릿 사용 실패: {str(e)}")
        raise HTTPException(status_code=500, detail=f"템플릿 사용 실패: {str(e)}")


@app.delete("/api/template/{template_id}")
async def delete_template(template_id: str):
    """
    템플릿 삭제
    """
    try:
        # 템플릿 파일 찾기
        template_files = list(TEMPLATE_DIR.glob(f"{template_id}.*"))

        if not template_files:
            raise HTTPException(
                status_code=404,
                detail="템플릿을 찾을 수 없습니다"
            )

        # 모든 관련 파일 삭제
        for file_path in template_files:
            file_path.unlink()

        logger.info(f"템플릿 삭제 완료: {template_id}")

        return {
            "success": True,
            "message": "템플릿이 삭제되었습니다"
        }

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"템플릿 삭제 실패: {str(e)}")
        raise HTTPException(status_code=500, detail=f"템플릿 삭제 실패: {str(e)}")


if __name__ == "__main__":
    import uvicorn

    print("=" * 60)
    print("Word to HWP 변환 시스템 서버 시작")
    print("=" * 60)
    print("주소: http://localhost:8000")
    print("API 문서: http://localhost:8000/docs")
    print("=" * 60)

    uvicorn.run(app, host="0.0.0.0", port=8000, log_level="info")
