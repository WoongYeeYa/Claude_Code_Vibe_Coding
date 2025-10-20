"""
DOCX 파일의 문단 서식 정보를 추출하는 스크립트
"""
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

docx_path = "test_formatting.docx"

doc = Document(docx_path)

print("=" * 80)
print("DOCX 파일 문단 서식 분석")
print("=" * 80)

for idx, para in enumerate(doc.paragraphs):
    if not para.text.strip():
        continue

    print(f"\n[문단 {idx+1}] {para.text[:50]}...")

    # 정렬
    alignment_names = {
        None: "None",
        WD_ALIGN_PARAGRAPH.LEFT: "LEFT",
        WD_ALIGN_PARAGRAPH.CENTER: "CENTER",
        WD_ALIGN_PARAGRAPH.RIGHT: "RIGHT",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "JUSTIFY"
    }
    print(f"  정렬: {alignment_names.get(para.alignment, 'UNKNOWN')}")

    # 들여쓰기
    fmt = para.paragraph_format

    if fmt.left_indent:
        print(f"  왼쪽 여백: {fmt.left_indent.pt} pt")

    if fmt.right_indent:
        print(f"  오른쪽 여백: {fmt.right_indent.pt} pt")

    if fmt.first_line_indent:
        print(f"  첫줄 들여쓰기: {fmt.first_line_indent.pt} pt")

    if fmt.space_before:
        print(f"  문단 위 간격: {fmt.space_before.pt} pt")

    if fmt.space_after:
        print(f"  문단 아래 간격: {fmt.space_after.pt} pt")

print("\n" + "=" * 80)
