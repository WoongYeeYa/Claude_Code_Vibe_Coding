"""
한컴오피스 한글 COM 자동화 모듈
Word 파일을 HWP 파일로 변환
"""
import os
import win32com.client
import pythoncom
import logging
from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class HWPConverter:
    """한컴오피스 한글 자동화 클래스"""

    def __init__(self):
        self.hwp = None
        self.is_initialized = False

    def initialize(self):
        """한글 프로그램 초기화 (백그라운드 모드, 보안 경고 없음)"""
        try:
            # 기존 인스턴스 완전히 정리
            if self.hwp:
                try:
                    self.hwp.Quit()
                except:
                    pass
                finally:
                    self.hwp = None

            # 잠시 대기 (프로세스 종료 대기)
            import time
            time.sleep(0.5)

            pythoncom.CoInitialize()
            self.hwp = win32com.client.Dispatch("HWPFrame.HwpObject")

            # 한글 창 숨기기
            try:
                self.hwp.XHwpWindows.Active_XHwpWindow.Visible = False
            except:
                pass

            try:
                self.hwp.Visible = False
            except:
                pass

            # 보안 경고 자동 허용
            self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

            try:
                self.hwp.HParameterSet.HFileOpenSave.AutoAction = True
            except:
                pass

            self.is_initialized = True
            logger.info("한컴오피스 한글 초기화 성공 (백그라운드 모드)")
            return True
        except Exception as e:
            logger.error(f"한컴오피스 한글 초기화 실패: {str(e)}")
            import traceback
            logger.error(f"상세 에러: {traceback.format_exc()}")
            self.is_initialized = False
            return False

    def close(self):
        """한글 프로그램 종료"""
        try:
            if self.hwp:
                try:
                    self.hwp.Quit()
                except:
                    pass
                finally:
                    self.hwp = None

            try:
                pythoncom.CoUninitialize()
            except:
                pass

            self.is_initialized = False
            logger.info("한컴오피스 한글 종료")

            # 프로세스 완전 정리를 위한 대기
            import time
            time.sleep(0.3)

        except Exception as e:
            logger.error(f"한컴오피스 한글 종료 실패: {str(e)}")

    def get_hwp_version(self):
        """설치된 한글 버전 확인"""
        try:
            if not self.is_initialized:
                self.initialize()

            version = self.hwp.Version
            logger.info(f"한컴오피스 한글 버전: {version}")
            return version
        except Exception as e:
            logger.error(f"버전 확인 실패: {str(e)}")
            return None

    def _apply_paragraph_format(self, paragraph):
        """문단 서식 적용 (정렬, 들여쓰기 등)"""
        try:
            # 문단 모양 가져오기
            self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

            # 정렬 설정
            alignment_map = {
                WD_PARAGRAPH_ALIGNMENT.LEFT: 0,      # 왼쪽
                WD_PARAGRAPH_ALIGNMENT.CENTER: 1,     # 가운데
                WD_PARAGRAPH_ALIGNMENT.RIGHT: 2,      # 오른쪽
                WD_PARAGRAPH_ALIGNMENT.JUSTIFY: 3,    # 양쪽
                None: 0  # 기본값: 왼쪽
            }

            alignment = alignment_map.get(paragraph.alignment, 0)
            self.hwp.HParameterSet.HParaShape.Align = alignment

            # 들여쓰기 설정 (twips 단위를 hwp 단위로 변환)
            if paragraph.paragraph_format.left_indent:
                try:
                    left_indent = int(paragraph.paragraph_format.left_indent.twips * 0.05)
                    self.hwp.HParameterSet.HParaShape.IndentLeft = left_indent
                except:
                    pass

            if paragraph.paragraph_format.first_line_indent:
                try:
                    first_indent = int(paragraph.paragraph_format.first_line_indent.twips * 0.05)
                    self.hwp.HParameterSet.HParaShape.IndentFirst = first_indent
                except:
                    pass

            # 문단 간격 설정
            if paragraph.paragraph_format.space_before:
                try:
                    space_before = int(paragraph.paragraph_format.space_before.twips * 0.05)
                    self.hwp.HParameterSet.HParaShape.SpaceAbove = space_before
                except:
                    pass

            if paragraph.paragraph_format.space_after:
                try:
                    space_after = int(paragraph.paragraph_format.space_after.twips * 0.05)
                    self.hwp.HParameterSet.HParaShape.SpaceBelow = space_after
                except:
                    pass

            # 문단 서식 적용
            self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

        except Exception as e:
            logger.debug(f"문단 서식 적용 실패 (무시): {str(e)}")

    def _convert_run(self, run):
        """Run 단위로 텍스트와 서식 변환"""
        if not run.text or run.text == "":
            return

        # 서식 적용
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

        try:
            # 굵기
            if run.bold:
                self.hwp.HParameterSet.HCharShape.Bold = 1

            # 기울임
            if run.italic:
                self.hwp.HParameterSet.HCharShape.Italic = 1

            # 밑줄
            if run.underline:
                try:
                    self.hwp.HParameterSet.HCharShape.UnderlineType = 1
                except:
                    pass

            # 폰트
            if run.font.name:
                try:
                    self.hwp.HParameterSet.HCharShape.FaceNameHangul = run.font.name
                    self.hwp.HParameterSet.HCharShape.FaceNameLatin = run.font.name
                    self.hwp.HParameterSet.HCharShape.FaceNameHanja = run.font.name
                    self.hwp.HParameterSet.HCharShape.FaceNameJapanese = run.font.name
                    self.hwp.HParameterSet.HCharShape.FaceNameOther = run.font.name
                    self.hwp.HParameterSet.HCharShape.FaceNameSymbol = run.font.name
                    self.hwp.HParameterSet.HCharShape.FaceNameUser = run.font.name
                except:
                    pass

            # 폰트 크기
            if run.font.size:
                try:
                    font_size_pt = run.font.size.pt if hasattr(run.font.size, 'pt') else 10
                    self.hwp.HParameterSet.HCharShape.Height = int(font_size_pt * 100)
                except:
                    pass

            # 글자 색상
            if run.font.color and run.font.color.rgb:
                try:
                    rgb = run.font.color.rgb
                    # RGB를 HWP 색상 값으로 변환
                    color_value = (rgb[0]) + (rgb[1] << 8) + (rgb[2] << 16)
                    self.hwp.HParameterSet.HCharShape.TextColor = color_value
                except:
                    pass

            # 서식 적용
            self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        except Exception as e:
            logger.warning(f"서식 적용 중 일부 실패: {str(e)}")

        # 텍스트 입력
        self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
        self.hwp.HParameterSet.HInsertText.Text = run.text
        self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)

    def _convert_paragraph(self, paragraph):
        """문단 변환 (서식 포함)"""
        # 문단 서식 먼저 적용
        self._apply_paragraph_format(paragraph)

        # Run별로 텍스트와 서식 적용
        has_text = False
        if paragraph.runs:
            for run in paragraph.runs:
                if run.text:  # 빈 텍스트가 아닌 경우만
                    self._convert_run(run)
                    has_text = True

        # runs가 없거나 runs에 텍스트가 없는 경우
        if not has_text and paragraph.text.strip():
            self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
            self.hwp.HParameterSet.HInsertText.Text = paragraph.text
            self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)

    def _get_merged_cells_info(self, table):
        """병합된 셀 정보 추출"""
        merged_cells = []

        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                # 셀의 grid span 확인
                tc = cell._element
                tcPr = tc.tcPr

                if tcPr is not None:
                    # 수직 병합 확인
                    vMerge = tcPr.vMerge
                    if vMerge is not None:
                        # vMerge가 있으면 병합된 셀
                        val = vMerge.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                        if val == 'restart' or val is None:
                            # 병합 시작 셀
                            # 아래로 몇 개 병합되는지 계산
                            merge_count = 1
                            for check_row_idx in range(row_idx + 1, len(table.rows)):
                                check_cell = table.rows[check_row_idx].cells[col_idx]
                                check_tc = check_cell._element
                                check_tcPr = check_tc.tcPr
                                if check_tcPr is not None:
                                    check_vMerge = check_tcPr.vMerge
                                    if check_vMerge is not None:
                                        check_val = check_vMerge.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                                        if check_val == 'continue' or (check_val is None and check_vMerge is not None):
                                            merge_count += 1
                                        else:
                                            break
                                    else:
                                        break
                                else:
                                    break

                            if merge_count > 1:
                                merged_cells.append({
                                    'row': row_idx,
                                    'col': col_idx,
                                    'rowspan': merge_count,
                                    'colspan': 1
                                })

                    # 수평 병합 확인
                    gridSpan = tcPr.gridSpan
                    if gridSpan is not None:
                        span = int(gridSpan.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 1))
                        if span > 1:
                            # 이미 vertical merge로 추가된 셀이면 colspan 업데이트
                            found = False
                            for mc in merged_cells:
                                if mc['row'] == row_idx and mc['col'] == col_idx:
                                    mc['colspan'] = span
                                    found = True
                                    break
                            if not found:
                                merged_cells.append({
                                    'row': row_idx,
                                    'col': col_idx,
                                    'rowspan': 1,
                                    'colspan': span
                                })

        return merged_cells

    def _convert_table(self, table):
        """표 변환 (셀 병합, 테두리 포함)"""
        try:
            rows = len(table.rows)
            cols = len(table.columns) if rows > 0 else 0

            if rows == 0 or cols == 0:
                return

            logger.info(f"표 변환 중: {rows}행 x {cols}열")

            # 병합 정보 추출
            merged_cells = self._get_merged_cells_info(table)
            logger.info(f"병합된 셀: {len(merged_cells)}개")

            # HWP에 표 삽입
            self.hwp.HAction.GetDefault("TableCreate", self.hwp.HParameterSet.HTableCreation.HSet)
            self.hwp.HParameterSet.HTableCreation.Rows = rows
            self.hwp.HParameterSet.HTableCreation.Cols = cols
            self.hwp.HParameterSet.HTableCreation.WidthType = 2  # 문서 너비에 맞춤
            self.hwp.HParameterSet.HTableCreation.HeightType = 0  # 자동
            self.hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", cols)

            # 각 열의 너비를 균등하게 설정
            col_width = int(60000 / cols)
            for i in range(cols):
                self.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(i, col_width)

            # 표 생성
            result = self.hwp.HAction.Execute("TableCreate", self.hwp.HParameterSet.HTableCreation.HSet)

            if not result:
                logger.warning("표 생성 실패")
                return

            # 표의 첫 번째 셀로 이동
            self.hwp.Run("TableCellBlock")
            self.hwp.Run("Cancel")

            # 각 셀에 내용 채우기
            for row_idx, row in enumerate(table.rows):
                for col_idx, cell in enumerate(row.cells):
                    # 셀 내용 변환
                    cell_text = cell.text.strip()

                    # 셀의 정렬 확인
                    if cell.paragraphs:
                        first_para = cell.paragraphs[0]
                        alignment_map = {
                            WD_PARAGRAPH_ALIGNMENT.LEFT: 0,
                            WD_PARAGRAPH_ALIGNMENT.CENTER: 1,
                            WD_PARAGRAPH_ALIGNMENT.RIGHT: 2,
                            WD_PARAGRAPH_ALIGNMENT.JUSTIFY: 3,
                            None: 0
                        }
                        alignment = alignment_map.get(first_para.alignment, 0)

                        # 셀 정렬 설정
                        try:
                            self.hwp.HAction.GetDefault("CellBorderFill", self.hwp.HParameterSet.HCellBorderFill.HSet)
                            self.hwp.HParameterSet.HCellBorderFill.Align = alignment
                            self.hwp.HAction.Execute("CellBorderFill", self.hwp.HParameterSet.HCellBorderFill.HSet)
                        except:
                            pass

                    if cell_text:
                        # 텍스트 입력
                        self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
                        self.hwp.HParameterSet.HInsertText.Text = cell_text
                        self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)

                    # 다음 셀로 이동 (마지막 셀이 아닌 경우)
                    if not (row_idx == rows - 1 and col_idx == cols - 1):
                        self.hwp.Run("TableRightCell")

            # 셀 병합 처리
            if merged_cells:
                logger.info(f"셀 병합 처리 시작: {len(merged_cells)}개")
                for merge_info in merged_cells:
                    try:
                        # 병합 시작 셀로 이동
                        # 표의 첫 셀로 이동
                        self.hwp.Run("TableCellBlock")
                        self.hwp.Run("Cancel")

                        # 목표 셀까지 이동
                        for r in range(merge_info['row']):
                            self.hwp.Run("TableDownCell")
                        for c in range(merge_info['col']):
                            self.hwp.Run("TableRightCell")

                        # 병합할 범위 선택
                        self.hwp.Run("TableCellBlockExtend")

                        # 아래로 확장
                        for r in range(merge_info['rowspan'] - 1):
                            self.hwp.Run("TableLowerCellAppend")

                        # 오른쪽으로 확장
                        for c in range(merge_info['colspan'] - 1):
                            self.hwp.Run("TableRightCellAppend")

                        # 셀 병합 실행
                        self.hwp.Run("TableMergeCell")

                        logger.debug(f"셀 병합 완료: ({merge_info['row']},{merge_info['col']}) {merge_info['rowspan']}x{merge_info['colspan']}")
                    except Exception as e:
                        logger.warning(f"셀 병합 실패: {str(e)}")

            # 표 밖으로 나가기
            self.hwp.HAction.Run("TableOut")

            logger.info("표 변환 완료")

        except Exception as e:
            logger.error(f"표 변환 실패: {str(e)}")
            import traceback
            logger.error(f"상세 에러: {traceback.format_exc()}")

    def convert_docx_to_hwp(self, docx_path: str, hwp_path: str) -> bool:
        """
        DOCX 파일을 HWP 파일로 변환
        python-docx로 파싱 후 HWP로 재구성

        Args:
            docx_path: 변환할 DOCX 파일 경로
            hwp_path: 저장할 HWP 파일 경로

        Returns:
            bool: 변환 성공 여부
        """
        try:
            # 매번 새로 초기화
            logger.info("한글 프로그램 초기화 중...")
            if not self.initialize():
                logger.error("한글 프로그램 초기화 실패")
                return False

            # 경로를 절대 경로로 변환
            docx_path = os.path.abspath(docx_path)
            hwp_path = os.path.abspath(hwp_path)

            # 파일 존재 확인
            if not os.path.exists(docx_path):
                logger.error(f"DOCX 파일을 찾을 수 없습니다: {docx_path}")
                return False

            logger.info(f"변환 시작: {docx_path} -> {hwp_path}")

            # DOCX 파일 파싱
            try:
                doc = Document(docx_path)
                logger.info(f"DOCX 파일 파싱 완료: {len(doc.paragraphs)}개 문단")
            except Exception as e:
                logger.error(f"DOCX 파일 파싱 실패: {str(e)}")
                return False

            # 새 HWP 문서 생성
            self.hwp.HAction.Run("FileNew")
            logger.info("새 HWP 문서 생성")

            # 문서 요소별로 변환
            for element in doc.element.body:
                if isinstance(element, CT_P):  # 문단
                    # Document 객체에서 해당 문단 찾기
                    for para in doc.paragraphs:
                        if para._element == element:
                            self._convert_paragraph(para)
                            # 문단 끝에 Enter
                            self.hwp.HAction.Run("BreakPara")
                            break

                elif isinstance(element, CT_Tbl):  # 표
                    # Document 객체에서 해당 표 찾기
                    for table in doc.tables:
                        if table._element == element:
                            self._convert_table(table)
                            # 표 다음에 Enter
                            self.hwp.HAction.Run("BreakPara")
                            break

            # HWPX 파일로 저장
            logger.info(f"HWPX 파일로 저장: {hwp_path}")
            save_result = self.hwp.SaveAs(hwp_path, "HWPX", "")

            if save_result:
                logger.info(f"SaveAs 성공")
            else:
                logger.warning(f"SaveAs 반환값: {save_result}")

            # 문서 닫기
            self.hwp.Clear(1)

            # 변환 결과 확인
            if os.path.exists(hwp_path):
                file_size = os.path.getsize(hwp_path)
                logger.info(f"변환 성공: {hwp_path} (크기: {file_size} bytes)")

                # 한글 프로그램 종료
                try:
                    self.hwp.Quit()
                    self.hwp = None
                    self.is_initialized = False
                    logger.info("한글 프로그램 종료")
                except:
                    pass

                return True
            else:
                logger.error("HWP 파일이 생성되지 않았습니다")

                # 한글 프로그램 종료
                try:
                    self.hwp.Quit()
                    self.hwp = None
                    self.is_initialized = False
                except:
                    pass

                return False

        except Exception as e:
            import traceback
            logger.error(f"변환 실패: {str(e)}")
            logger.error(f"상세 에러: {traceback.format_exc()}")
            try:
                self.hwp.Clear(1)
            except:
                pass

            # 한글 프로그램 종료
            try:
                if self.hwp:
                    self.hwp.Quit()
                    self.hwp = None
                    self.is_initialized = False
            except:
                pass

            return False

    def check_hwp_installed(self) -> bool:
        """한컴오피스 한글 설치 여부 확인"""
        try:
            pythoncom.CoInitialize()
            hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
            hwp.Quit()
            pythoncom.CoUninitialize()
            return True
        except:
            return False
