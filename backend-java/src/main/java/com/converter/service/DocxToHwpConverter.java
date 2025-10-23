package com.converter.service;

import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.object.bodytext.Section;
import kr.dogfoot.hwplib.object.bodytext.paragraph.Paragraph;
import kr.dogfoot.hwplib.object.bodytext.paragraph.charshape.ParaCharShape;
import kr.dogfoot.hwplib.object.bodytext.paragraph.header.ParaHeader;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.ParaText;
import kr.dogfoot.hwplib.object.docinfo.CharShape;
import kr.dogfoot.hwplib.object.docinfo.FaceName;
import kr.dogfoot.hwplib.object.bodytext.control.ControlTable;
import kr.dogfoot.hwplib.object.bodytext.control.ControlType;
import kr.dogfoot.hwplib.object.bodytext.control.table.*;
import kr.dogfoot.hwplib.object.bodytext.control.gso.textbox.TextVerticalAlignment;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.sectiondefine.TextDirection;
import kr.dogfoot.hwplib.object.bodytext.control.gso.textbox.LineChange;
import kr.dogfoot.hwplib.tool.blankfilemaker.BlankFileMaker;
import kr.dogfoot.hwplib.writer.HWPWriter;
import org.apache.poi.xwpf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.UnsupportedEncodingException;
import java.util.List;

/**
 * DOCX를 HWP로 변환하는 서비스
 *
 * <p>Apache POI를 사용하여 DOCX 파일을 파싱하고,
 * hwplib을 사용하여 HWP 파일을 생성합니다.</p>
 */
@Service
public class DocxToHwpConverter {

    private static final Logger logger = LoggerFactory.getLogger(DocxToHwpConverter.class);

    private HWPFile hwpFile;
    private Section section;

    /**
     * DOCX 파일을 HWP 파일로 변환
     */
    public boolean convert(File docxFile, File hwpOutputFile) {
        logger.info("=".repeat(80));
        logger.info("DOCX → HWP 변환 시작");
        logger.info("입력: {}", docxFile.getAbsolutePath());
        logger.info("출력: {}", hwpOutputFile.getAbsolutePath());
        logger.info("=".repeat(80));

        try {
            // 1. DOCX 파일 열기
            logger.info("[1/4] DOCX 파일 읽기 중...");
            FileInputStream fis = new FileInputStream(docxFile);
            XWPFDocument document = new XWPFDocument(fis);
            logger.info("✓ DOCX 파일 로드 완료");

            // 2. HWP 파일 초기화
            logger.info("[2/4] HWP 파일 초기화 중...");
            initializeHWPFile();
            logger.info("✓ HWP 파일 초기화 완료");

            // 3. DOCX 내용을 HWP로 변환
            logger.info("[3/4] DOCX 내용 변환 중...");
            convertContent(document);
            logger.info("✓ 내용 변환 완료");

            // 4. HWP 파일 저장
            logger.info("[4/4] HWP 파일 저장 중...");
            HWPWriter.toFile(this.hwpFile, hwpOutputFile.getAbsolutePath());
            logger.info("✓ HWP 파일 저장 완료: {}", hwpOutputFile.getAbsolutePath());

            // 리소스 정리
            document.close();
            fis.close();

            logger.info("=".repeat(80));
            logger.info("✓ 변환 성공!");
            logger.info("=".repeat(80));

            return true;

        } catch (Exception e) {
            logger.error("=".repeat(80));
            logger.error("✗ 변환 실패!", e);
            logger.error("=".repeat(80));
            return false;
        }
    }

    /**
     * HWP 파일 초기화
     */
    private void initializeHWPFile() throws Exception {
        hwpFile = BlankFileMaker.make();
        section = hwpFile.getBodyText().addNewSection();
        logger.debug("HWP 파일 및 섹션 초기화 완료");
    }

    /**
     * DOCX 내용을 HWP로 변환
     */
    private void convertContent(XWPFDocument document) {
        int paragraphCount = 0;
        int tableCount = 0;

        for (IBodyElement element : document.getBodyElements()) {
            if (element instanceof XWPFParagraph) {
                XWPFParagraph para = (XWPFParagraph) element;
                convertParagraph(para);
                paragraphCount++;
            } else if (element instanceof XWPFTable) {
                XWPFTable table = (XWPFTable) element;
                convertTable(table);
                tableCount++;
            }
        }

        logger.info("문단 변환: {} 개", paragraphCount);
        logger.info("테이블 변환: {} 개", tableCount);
    }

    /**
     * 문단 변환
     */
    private void convertParagraph(XWPFParagraph para) {
        try {
            Paragraph hwpPara = section.addNewParagraph();

            // ParaHeader 설정
            ParaHeader ph = hwpPara.getHeader();
            ph.setLastInList(true);
            ph.setParaShapeId(1);
            ph.setStyleId((short) 1);

            // 텍스트 추출
            String text = para.getText();
            if (text == null || text.trim().isEmpty()) {
                logger.debug("빈 문단");
                return;
            }

            logger.debug("문단 텍스트: \"{}\"", text);

            // 텍스트 추가
            hwpPara.createText();
            ParaText pt = hwpPara.getText();

            try {
                pt.addString(text);
            } catch (UnsupportedEncodingException e) {
                logger.error("텍스트 인코딩 실패", e);
            }

            // CharShape 추가
            hwpPara.createCharShape();
            ParaCharShape pcs = hwpPara.getCharShape();
            pcs.addParaCharShape(0, 0); // position 0, charShapeId 0 (기본)

        } catch (Exception e) {
            logger.error("문단 변환 실패", e);
        }
    }

    /**
     * 테이블 변환
     */
    private void convertTable(XWPFTable table) {
        logger.debug("=".repeat(60));
        logger.debug("테이블 변환 시작");

        try {
            int rows = table.getNumberOfRows();
            if (rows == 0) {
                logger.warn("테이블 행이 없습니다");
                return;
            }

            int cols = table.getRow(0).getTableCells().size();
            logger.debug("테이블 크기: {} 행 x {} 열", rows, cols);

            // 테이블 컨트롤 생성
            Paragraph hwpPara = section.addNewParagraph();
            hwpPara.createText(); // Must create text before addExtendCharForTable
            hwpPara.getText().addExtendCharForTable();
            ControlTable ctrlTable = (ControlTable) hwpPara.addNewControl(ControlType.Table);

            // 테이블 설정
            Table hwpTable = ctrlTable.getTable();
            hwpTable.setRowCount(rows);
            hwpTable.setColumnCount(cols);

            // 셀 개수 설정
            for (int r = 0; r < rows; r++) {
                hwpTable.getCellCountOfRowList().add(cols);
            }

            // 행과 셀 추가
            for (int r = 0; r < rows; r++) {
                XWPFTableRow docxRow = table.getRow(r);
                Row hwpRow = ctrlTable.addNewRow();

                for (int c = 0; c < cols && c < docxRow.getTableCells().size(); c++) {
                    Cell hwpCell = hwpRow.addNewCell();

                    // ListHeader 설정
                    ListHeaderForCell lh = hwpCell.getListHeader();
                    lh.setParaCount(1);
                    lh.getProperty().setTextDirection(TextDirection.Horizontal);
                    lh.getProperty().setLineChange(LineChange.Normal);
                    lh.getProperty().setTextVerticalAlignment(TextVerticalAlignment.Center);
                    lh.setColIndex(c);
                    lh.setRowIndex(r);
                    lh.setColSpan(1);
                    lh.setRowSpan(1);

                    // 셀 내용 추가 - 모든 셀에 문단 필수!
                    XWPFTableCell docxCell = docxRow.getCell(c);
                    String cellText = "";
                    if (docxCell != null) {
                        cellText = docxCell.getText();
                        if (cellText == null) {
                            cellText = "";
                        }
                    }

                    // 빈 셀이라도 반드시 문단 추가
                    Paragraph cellPara = hwpCell.getParagraphList().addNewParagraph();

                    // ParaHeader 설정
                    ParaHeader ph = cellPara.getHeader();
                    ph.setLastInList(true);
                    ph.setParaShapeId(1);
                    ph.setStyleId((short) 1);

                    // 텍스트 추가 - 빈 셀도 반드시 빈 문자열 추가!
                    cellPara.createText();
                    try {
                        if (cellText.isEmpty()) {
                            // 빈 셀에는 공백 하나라도 추가
                            cellPara.getText().addString(" ");
                        } else {
                            cellPara.getText().addString(cellText);
                            logger.debug("셀 [{}][{}]: \"{}\"", r, c, cellText);
                        }
                    } catch (UnsupportedEncodingException e) {
                        logger.error("셀 텍스트 인코딩 실패", e);
                    }

                    // CharShape 추가
                    cellPara.createCharShape();
                    cellPara.getCharShape().addParaCharShape(0, 0);
                }
            }

            logger.debug("✓ 테이블 변환 완료");

        } catch (Exception e) {
            logger.error("테이블 변환 실패", e);
        }
    }
}
