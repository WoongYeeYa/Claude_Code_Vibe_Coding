"""
DOCX to HWP 변환 테스트 스크립트
"""
import sys
import os

# 백엔드 모듈 경로 추가
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'backend'))

from hwp_converter import HWPConverter

def test_conversion():
    """테스트 파일을 HWP로 변환"""
    docx_path = "test_formatting.docx"
    hwp_path = "test_formatting.hwp"

    if not os.path.exists(docx_path):
        print(f"[ERROR] 테스트 파일을 찾을 수 없습니다: {docx_path}")
        return False

    print("=" * 60)
    print("DOCX to HWP 변환 테스트 시작")
    print("=" * 60)
    print(f"입력 파일: {docx_path}")
    print(f"출력 파일: {hwp_path}")
    print()

    converter = HWPConverter()

    # 한컴오피스 설치 확인
    is_installed, version = converter.check_hwp_installed()
    if not is_installed:
        print("[ERROR] 한컴오피스가 설치되지 않았거나 연동할 수 없습니다.")
        return False

    print(f"[OK] 한컴오피스 버전: {version}")
    print()

    # 변환 실행
    print("변환 중...")
    success = converter.convert_docx_to_hwp(docx_path, hwp_path)

    if success:
        if os.path.exists(hwp_path):
            file_size = os.path.getsize(hwp_path)
            print()
            print("=" * 60)
            print("[OK] 변환 성공!")
            print("=" * 60)
            print(f"생성된 파일: {hwp_path}")
            print(f"파일 크기: {file_size:,} bytes")
            print()
            print("테스트 항목:")
            print("  1. 폰트 색깔 (빨강, 초록, 파랑, 보라, 주황)")
            print("  2. 글 정렬 (왼쪽, 가운데, 오른쪽, 양쪽정렬)")
            print("  3. 들여쓰기 (첫줄, 왼쪽, 오른쪽, 복합)")
            print("  4. 내어쓰기 (Hanging Indent)")
            print("  5. 복합 서식 (색깔+정렬+들여쓰기)")
            print()
            print("생성된 HWP 파일을 열어서 서식이 제대로 적용되었는지 확인하세요.")
            return True
        else:
            print("[ERROR] 변환은 성공했지만 파일이 생성되지 않았습니다.")
            return False
    else:
        print("[ERROR] 변환 실패")
        return False

if __name__ == "__main__":
    test_conversion()
