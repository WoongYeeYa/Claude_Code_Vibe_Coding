package com.converter;

import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.object.bodytext.Section;
import kr.dogfoot.hwplib.object.bodytext.control.ControlTable;
import kr.dogfoot.hwplib.object.bodytext.control.ControlType;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.CtrlHeaderGso;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.gso.*;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.sectiondefine.TextDirection;
import kr.dogfoot.hwplib.object.bodytext.control.gso.textbox.LineChange;
import kr.dogfoot.hwplib.object.bodytext.control.gso.textbox.TextVerticalAlignment;
import kr.dogfoot.hwplib.object.bodytext.control.table.*;
import kr.dogfoot.hwplib.object.bodytext.paragraph.Paragraph;
import kr.dogfoot.hwplib.object.bodytext.paragraph.charshape.ParaCharShape;
import kr.dogfoot.hwplib.object.bodytext.paragraph.header.ParaHeader;
import kr.dogfoot.hwplib.object.bodytext.paragraph.lineseg.LineSegItem;
import kr.dogfoot.hwplib.object.bodytext.paragraph.lineseg.ParaLineSeg;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.ParaText;
import kr.dogfoot.hwplib.object.docinfo.BorderFill;
import kr.dogfoot.hwplib.object.docinfo.CharShape;
import kr.dogfoot.hwplib.object.docinfo.ParaShape;
import kr.dogfoot.hwplib.object.docinfo.borderfill.*;
import kr.dogfoot.hwplib.object.docinfo.borderfill.fillinfo.PatternFill;
import kr.dogfoot.hwplib.object.docinfo.borderfill.fillinfo.PatternType;
import kr.dogfoot.hwplib.reader.HWPReader;
import kr.dogfoot.hwplib.tool.TableCellMerger;
import kr.dogfoot.hwplib.tool.blankfilemaker.BlankFileMaker;
import kr.dogfoot.hwplib.writer.HWPWriter;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.math.BigInteger;
import java.util.*;

/**
 * DOCX to HWP converter using hwplib
 */
@Service
public class DocxToHwpConverter {

    private static final Logger logger = LoggerFactory.getLogger(DocxToHwpConverter.class);

    private HWPFile hwpFile;
    private int zOrder = 0;
    private int borderFillIDForCell;
    private Map<String, Integer> charShapeCache = new HashMap<>();

    /**
     * Convert DOCX file to HWP file
     *
     * @param docxPath Input DOCX file path
     * @param hwpPath Output HWP file path
     * @return true if conversion successful, false otherwise
     */
    public boolean convert(String docxPath, String hwpPath) {
        try {
            logger.info("Starting conversion: {} -> {}", docxPath, hwpPath);

            XWPFDocument docx = new XWPFDocument(new FileInputStream(docxPath));
            hwpFile = BlankFileMaker.make();
            if (hwpFile == null) {
                throw new Exception("Failed to create blank HWP file");
            }

            Section section = hwpFile.getBodyText().getSectionList().get(0);
            List<IBodyElement> bodyElements = docx.getBodyElements();

            if (bodyElements.isEmpty()) {
                logger.info("Input DOCX is empty, writing a blank HWP file.");
            } else {
                // Reuse the first paragraph for the first element
                IBodyElement firstElement = bodyElements.get(0);
                Paragraph firstPara = section.getParagraph(0);
                    // Clear the default paragraph content
                firstPara.createText();
                firstPara.createCharShape();
                firstPara.createLineSeg();

                if (firstElement instanceof XWPFParagraph) {
                    populateParagraph(firstPara, (XWPFParagraph) firstElement);
                } else if (firstElement instanceof XWPFTable) {
                    processTable((XWPFTable) firstElement, firstPara);
                }

                // Process remaining elements
                for (int i = 1; i < bodyElements.size(); i++) {
                    IBodyElement element = bodyElements.get(i);
                    if (element instanceof XWPFParagraph) {
                        populateParagraph(section.addNewParagraph(), (XWPFParagraph) element);
                    } else if (element instanceof XWPFTable) {
                        processTable((XWPFTable) element, section.addNewParagraph());
                    }
                }
            }

            HWPWriter.toFile(hwpFile, hwpPath);

            logger.info("Conversion completed successfully");
            docx.close();
            return true;

        } catch (Exception e) {
            logger.error("Conversion failed: {}", e.getMessage(), e);
            return false;
        }
    }

    private int getCharShapeIdForRun(XWPFRun run) {
        // 1. Build a unique key for the style
        String fontName = run.getFontFamily() == null ? "Batang" : run.getFontFamily();
        int fontSize = run.getFontSize() == -1 ? 10 : run.getFontSize();
        String color = run.getColor() == null ? "000000" : run.getColor();
        boolean isBold = run.isBold();
        String styleKey = String.format("font:%s;size:%d;color:%s;bold:%b", fontName, fontSize, color, isBold);

        // 2. Return from cache if style already exists
        if (charShapeCache.containsKey(styleKey)) {
            return charShapeCache.get(styleKey);
        }

        // 3. Create new CharShape if not in cache
        logger.info("Creating new CharShape for style: {}", styleKey);
        CharShape cs = hwpFile.getDocInfo().addNewCharShape();

        // Font Face (Simplified: using default Batang font)
        cs.getFaceNameIds().setForAll(1);

        // Font Size
        cs.setBaseSize(ptToLineHeight(fontSize));
        cs.getRatios().setForAll((short) 100);

        // Bold
        cs.getProperty().setBold(isBold);

        // Color
        long bgr = Long.parseLong(color.substring(4, 6) + color.substring(2, 4) + color.substring(0, 2), 16);
        cs.getCharColor().setValue(bgr);

        // 4. Add to cache and return ID
        int newId = hwpFile.getDocInfo().getCharShapeList().size() - 1;
        charShapeCache.put(styleKey, newId);
        return newId;
    }

    private void populateParagraph(Paragraph hwpPara, XWPFParagraph docxPara) throws UnsupportedEncodingException {
        // Set paragraph alignment and header
        int paraShapeId = getParaShapeIdForAlignment(docxPara);
        setParaHeader(hwpPara, paraShapeId);

        // Initialize text and charshape objects
        hwpPara.createText();
        ParaText paraText = hwpPara.getText();
        hwpPara.createCharShape();
        ParaCharShape paraCharShape = hwpPara.getCharShape();
        
        int currentPos = 0;

        // Process each run to apply styles
        for (XWPFRun run : docxPara.getRuns()) {
            String text = run.getText(0);
            if (text == null || text.isEmpty()) {
                continue;
            }

            // Get style for this run
            int charShapeId = getCharShapeIdForRun(run);
            
            // Add text and apply shape
            paraText.addString(text);
            paraCharShape.addParaCharShape(currentPos, charShapeId);
            
            currentPos += text.length();
        }

        // If paragraph is empty, add a single empty string to avoid corruption
        if (currentPos == 0) {
            paraText.addString("");
            paraCharShape.addParaCharShape(0, 1); // Default shape
        }

        // Set paragraph line segment (currently disabled to prevent corruption)
        // setParaLineSeg(hwpPara, currentPos);
    }

    /**
     * Process a DOCX table and add to HWP section
     */
    private void processTable(XWPFTable docxTable, Paragraph hwpPara) throws UnsupportedEncodingException {
        logger.info("Processing table with {} rows", docxTable.getRows().size());
        logger.info("Creating new HWP table control.");

        // Create table control in the given paragraph

        // Initialize ParaText if null
        if (hwpPara.getText() == null) {
            hwpPara.createText();
        }

        hwpPara.getText().addExtendCharForTable();

        ControlTable table = (ControlTable) hwpPara.addNewControl(ControlType.Table);

        // Set table properties
        setTableHeader(table, docxTable);
        setTableRecord(table, docxTable);

        // Add table cells
        addTableCells(table, docxTable);

        // Process cell merges
        processCellMerges(table, docxTable);
    }

    /**
     * Set table header (position, size, etc.)
     */
    private void setTableHeader(ControlTable table, XWPFTable docxTable) {
        CtrlHeaderGso ctrlHeader = table.getHeader();
        ctrlHeader.getProperty().setLikeWord(false);
        ctrlHeader.getProperty().setApplyLineSpace(false);
        ctrlHeader.getProperty().setVertRelTo(VertRelTo.Para);
        ctrlHeader.getProperty().setVertRelativeArrange(RelativeArrange.TopOrLeft);
        ctrlHeader.getProperty().setHorzRelTo(HorzRelTo.Para);
        ctrlHeader.getProperty().setHorzRelativeArrange(RelativeArrange.TopOrLeft);
        ctrlHeader.getProperty().setVertRelToParaLimit(false);
        ctrlHeader.getProperty().setAllowOverlap(false);
        ctrlHeader.getProperty().setWidthCriterion(WidthCriterion.Absolute);
        ctrlHeader.getProperty().setHeightCriterion(HeightCriterion.Absolute);
        ctrlHeader.getProperty().setProtectSize(false);
        ctrlHeader.getProperty().setTextFlowMethod(TextFlowMethod.FitWithText);
        ctrlHeader.getProperty().setTextHorzArrange(TextHorzArrange.BothSides);
        ctrlHeader.getProperty().setObjectNumberSort(ObjectNumberSort.Table);

        // Set default size
        ctrlHeader.setxOffset(mmToHwp(20.0));
        ctrlHeader.setyOffset(mmToHwp(20.0));
        ctrlHeader.setWidth(mmToHwp(150.0));
        ctrlHeader.setHeight(mmToHwp(50.0) * docxTable.getRows().size());
        ctrlHeader.setzOrder(zOrder++);
        ctrlHeader.setOutterMarginLeft(0);
        ctrlHeader.setOutterMarginRight(0);
        ctrlHeader.setOutterMarginTop(0);
        ctrlHeader.setOutterMarginBottom(0);
    }

    /**
     * Set table record (row/column count, spacing, etc.)
     */
    private void setTableRecord(ControlTable table, XWPFTable docxTable) {
        Table tableRecord = table.getTable();
        tableRecord.getProperty().setDivideAtPageBoundary(DivideAtPageBoundary.DivideByCell);
        tableRecord.getProperty().setAutoRepeatTitleRow(false);

        int rowCount = docxTable.getRows().size();
        int colCount = docxTable.getRow(0).getTableCells().size();

        tableRecord.setRowCount(rowCount);
        tableRecord.setColumnCount(colCount);
        tableRecord.setCellSpacing(0);
        tableRecord.setLeftInnerMargin(0);
        tableRecord.setRightInnerMargin(0);
        tableRecord.setTopInnerMargin(0);
        tableRecord.setBottomInnerMargin(0);
        tableRecord.setBorderFillId(getBorderFillIDForTableOuterLine());

        // Set cell count for each row
        for (int i = 0; i < rowCount; i++) {
            tableRecord.getCellCountOfRowList().add(colCount);
        }
    }

    /**
     * Add cells to table
     */
    private void addTableCells(ControlTable table, XWPFTable docxTable) throws UnsupportedEncodingException {
        try {
            int rowIndex = 0;
            int colCount = docxTable.getRow(0).getTableCells().size();
            logger.info("Adding table cells: {} rows, {} cols", docxTable.getRows().size(), colCount);

            for (XWPFTableRow docxRow : docxTable.getRows()) {
                Row hwpRow = table.addNewRow();
                logger.debug("Processing row {}", rowIndex);

                for (int colIndex = 0; colIndex < colCount; colIndex++) {
                    try {
                        Cell hwpCell = hwpRow.addNewCell();

                        // Set cell header
                        setListHeaderForCell(hwpCell, colIndex, rowIndex, colCount);

                        // Set cell content
                        if (colIndex < docxRow.getTableCells().size()) {
                            XWPFTableCell docxCell = docxRow.getCell(colIndex);
                            setParagraphForCell(hwpCell, docxCell);
                        } else {
                            setParagraphForCell(hwpCell, null);
                        }
                    } catch (Exception e) {
                        logger.error("Error processing cell[{},{}]: {}", rowIndex, colIndex, e.getMessage(), e);
                        throw e;
                    }
                }
                rowIndex++;
            }
        } catch (Exception e) {
            logger.error("Error in addTableCells: {}", e.getMessage(), e);
            throw new RuntimeException("Failed to add table cells", e);
        }
    }

    /**
     * Process cell merges from DOCX table
     */
    private void processCellMerges(ControlTable table, XWPFTable docxTable) {
        try {
            Map<String, CellMergeInfo> mergeMap = analyzeCellMerges(docxTable);

            for (CellMergeInfo mergeInfo : mergeMap.values()) {
                logger.info("Merging cells: row={} col={} rowSpan={} colSpan={}",
                    mergeInfo.startRow, mergeInfo.startCol, mergeInfo.rowSpan, mergeInfo.colSpan);

                // Calculate end positions
                int endRow = mergeInfo.startRow + mergeInfo.rowSpan - 1;
                int endCol = mergeInfo.startCol + mergeInfo.colSpan - 1;

                // Merge cells using hwplib
                TableCellMerger.mergeCell(table,
                    mergeInfo.startRow, mergeInfo.startCol,
                    endRow, endCol);
            }
        } catch (Exception e) {
            logger.error("Failed to process cell merges: {}", e.getMessage(), e);
        }
    }

    /**
     * Analyze cell merges in DOCX table
     */
    private Map<String, CellMergeInfo> analyzeCellMerges(XWPFTable docxTable) {
        Map<String, CellMergeInfo> mergeMap = new HashMap<>();
        Map<Integer, Integer> vMergeTracker = new HashMap<>();

        int rowIndex = 0;
        for (XWPFTableRow row : docxTable.getRows()) {
            int colIndex = 0;
            for (XWPFTableCell cell : row.getTableCells()) {
                CTTcPr tcPr = cell.getCTTc().getTcPr();

                if (tcPr != null) {
                    // Horizontal merge (gridSpan)
                    int hSpan = 1;
                    if (tcPr.isSetGridSpan()) {
                        hSpan = tcPr.getGridSpan().getVal().intValue();
                    }

                    // Vertical merge
                    int vSpan = 1;
                    boolean isVMergeStart = false;

                    if (tcPr.isSetVMerge()) {
                        CTVMerge vMerge = tcPr.getVMerge();
                        if (vMerge.getVal() == null || vMerge.getVal().equals(STMerge.RESTART)) {
                            // Start of vertical merge
                            vMergeTracker.put(colIndex, rowIndex);
                            isVMergeStart = true;
                        } else if (vMerge.getVal().equals(STMerge.CONTINUE)) {
                            // Continuation of vertical merge
                            if (vMergeTracker.containsKey(colIndex)) {
                                int startRow = vMergeTracker.get(colIndex);
                                String key = startRow + "," + colIndex;

                                if (mergeMap.containsKey(key)) {
                                    mergeMap.get(key).rowSpan = rowIndex - startRow + 1;
                                }
                            }
                        }
                    }

                    // Record merge if span > 1
                    if (hSpan > 1 || isVMergeStart) {
                        String key = rowIndex + "," + colIndex;
                        CellMergeInfo info = new CellMergeInfo();
                        info.startRow = rowIndex;
                        info.startCol = colIndex;
                        info.rowSpan = vSpan;
                        info.colSpan = hSpan;
                        mergeMap.put(key, info);
                    }
                }

                colIndex++;
            }
            rowIndex++;
        }

        return mergeMap;
    }

    /**
     * Set cell header (list header for cell)
     */
    private void setListHeaderForCell(Cell cell, int colIndex, int rowIndex, int totalCols) {
        ListHeaderForCell lh = cell.getListHeader();
        lh.setParaCount(1);
        lh.getProperty().setTextDirection(TextDirection.Horizontal);
        lh.getProperty().setLineChange(LineChange.Normal);
        lh.getProperty().setTextVerticalAlignment(TextVerticalAlignment.Center);
        lh.getProperty().setProtectCell(false);
        lh.getProperty().setEditableAtFormMode(false);
        lh.setColIndex(colIndex);
        lh.setRowIndex(rowIndex);
        lh.setColSpan(1);
        lh.setRowSpan(1);

        // Calculate width based on total columns
        double cellWidth = 150.0 / totalCols;
        lh.setWidth(mmToHwp(cellWidth));
        lh.setHeight(mmToHwp(10.0));
        lh.setLeftMargin(0);
        lh.setRightMargin(0);
        lh.setTopMargin(0);
        lh.setBottomMargin(0);
        lh.setBorderFillId(getBorderFillIDForCell());
        lh.setTextWidth(mmToHwp(cellWidth));
        lh.setFieldName("");
    }

    /**
     * Set paragraph for cell
     */
    private void setParagraphForCell(Cell cell, XWPFTableCell docxCell) throws UnsupportedEncodingException {
        try {
            if (cell == null) {
                logger.error("Cell is null in setParagraphForCell");
                return;
            }

            if (docxCell != null && !docxCell.getParagraphs().isEmpty()) {
                for (XWPFParagraph docxCellPara : docxCell.getParagraphs()) {
                    Paragraph targetHwpPara = cell.getParagraphList().addNewParagraph();
                    populateParagraph(targetHwpPara, docxCellPara);
                }
            } else {
                // If DOCX cell is empty, ensure at least one empty HWP paragraph is added
                Paragraph targetHwpPara = cell.getParagraphList().addNewParagraph();
                populateParagraph(targetHwpPara, new XWPFDocument().createParagraph());
            }
        } catch (Exception e) {
            logger.error("Error in setParagraphForCell: {}", e.getMessage(), e);
            throw new RuntimeException("Failed to set paragraph for cell", e);
        }
    }

    /**
     * Set paragraph header
     */
    private void setParaHeader(Paragraph p, int paraShapeId) {
        ParaHeader ph = p.getHeader();
        ph.setLastInList(true);
        ph.setParaShapeId(paraShapeId);
        ph.setStyleId((short) 1);
        ph.getDivideSort().setDivideSection(false);
        ph.getDivideSort().setDivideMultiColumn(false);
        ph.getDivideSort().setDividePage(false);
        ph.getDivideSort().setDivideColumn(false);
        ph.setCharShapeCount(1);
        ph.setRangeTagCount(0);
        ph.setLineAlignCount(1);
        ph.setInstanceID(0);
        ph.setIsMergedByTrack(0);
    }

    /**
     * Set paragraph text
     */
    private void setParaText(Paragraph p, String text) throws UnsupportedEncodingException {
        p.createText();
        ParaText pt = p.getText();

        // hwplib requires at least one character in CharList
        // Even for empty text, we need to add at least empty string
        if (text != null && !text.isEmpty()) {
            pt.addString(text);
        } else {
            // Add empty string to prevent IndexOutOfBoundsException
            pt.addString("");
        }
    }

    /**
     * Set paragraph character shape based on DOCX formatting
     */
    private void setParaCharShape(Paragraph p, XWPFParagraph docxPara) {
        p.createCharShape();
        ParaCharShape pcs = p.getCharShape();
        pcs.addParaCharShape(0, 1);
    }

    /**
     * Set paragraph line segment
     */
    private void setParaLineSeg(Paragraph p, int textLength) {
        p.createLineSeg();
        ParaLineSeg pls = p.getLineSeg();
        LineSegItem lsi = pls.addNewLineSegItem();

        lsi.setTextStartPosition(0);
        lsi.setLineVerticalPosition(0);
        lsi.setLineHeight(ptToLineHeight(10.0));
        lsi.setTextPartHeight(ptToLineHeight(10.0));
        lsi.setDistanceBaseLineToLineVerticalPosition(ptToLineHeight(10.0 * 0.85));
        lsi.setLineSpace(ptToLineHeight(3.0));
        lsi.setStartPositionFromColumn(0);
        lsi.setSegmentWidth((int) mmToHwp(50.0));
        lsi.getTag().setFirstSegmentAtLine(true);
        lsi.getTag().setLastSegmentAtLine(true);
    }

    /**
     * Set paragraph alignment based on DOCX alignment
     */
    private int getParaShapeIdForAlignment(XWPFParagraph docxPara) {
        ParaShape ps = hwpFile.getDocInfo().addNewParaShape();
        
        // Set properties using setter methods
        ps.setLeftMargin(0);
        ps.setRightMargin(0);
        ps.setIndent(0);

        // Get alignment from DOCX
        ParagraphAlignment alignment = docxPara.getAlignment();
        logger.info("Processing paragraph alignment: {}", alignment);

        switch (alignment) {
            case LEFT:
                ps.getProperty1().setAlignment(kr.dogfoot.hwplib.object.docinfo.parashape.Alignment.Left);
                break;
            case CENTER:
                ps.getProperty1().setAlignment(kr.dogfoot.hwplib.object.docinfo.parashape.Alignment.Center);
                break;
            case RIGHT:
                ps.getProperty1().setAlignment(kr.dogfoot.hwplib.object.docinfo.parashape.Alignment.Right);
                break;
            case BOTH:
                ps.getProperty1().setAlignment(kr.dogfoot.hwplib.object.docinfo.parashape.Alignment.Justify);
                break;
            default:
                ps.getProperty1().setAlignment(kr.dogfoot.hwplib.object.docinfo.parashape.Alignment.Justify);
                break;
        }

        return hwpFile.getDocInfo().getParaShapeList().size() - 1;
    }

    /**
     * Get border fill ID for table outer line
     */
    private int getBorderFillIDForTableOuterLine() {
        if (borderFillIDForCell > 0) {
            return borderFillIDForCell;
        }

        BorderFill bf = hwpFile.getDocInfo().addNewBorderFill();
        bf.getProperty().set3DEffect(false);
        bf.getProperty().setShadowEffect(false);
        bf.getProperty().setSlashDiagonalShape(SlashDiagonalShape.None);
        bf.getProperty().setBackSlashDiagonalShape(BackSlashDiagonalShape.None);

        // Set all borders to none for outer line
        setBorder(bf.getLeftBorder(), BorderType.None);
        setBorder(bf.getRightBorder(), BorderType.None);
        setBorder(bf.getTopBorder(), BorderType.None);
        setBorder(bf.getBottomBorder(), BorderType.None);
        setBorder(bf.getDiagonalBorder(), BorderType.None);

        // Set fill
        bf.getFillInfo().getType().setPatternFill(true);
        bf.getFillInfo().createPatternFill();
        PatternFill pf = bf.getFillInfo().getPatternFill();
        pf.setPatternType(PatternType.None);
        pf.getBackColor().setValue(-1);
        pf.getPatternColor().setValue(0);

        return hwpFile.getDocInfo().getBorderFillList().size();
    }

    /**
     * Get border fill ID for table cell
     */
    private int getBorderFillIDForCell() {
        if (borderFillIDForCell > 0) {
            return borderFillIDForCell;
        }

        BorderFill bf = hwpFile.getDocInfo().addNewBorderFill();
        bf.getProperty().set3DEffect(false);
        bf.getProperty().setShadowEffect(false);
        bf.getProperty().setSlashDiagonalShape(SlashDiagonalShape.None);
        bf.getProperty().setBackSlashDiagonalShape(BackSlashDiagonalShape.None);

        // Set all borders to solid
        setBorder(bf.getLeftBorder(), BorderType.Solid);
        setBorder(bf.getRightBorder(), BorderType.Solid);
        setBorder(bf.getTopBorder(), BorderType.Solid);
        setBorder(bf.getBottomBorder(), BorderType.Solid);
        setBorder(bf.getDiagonalBorder(), BorderType.None);

        // Set fill
        bf.getFillInfo().getType().setPatternFill(true);
        bf.getFillInfo().createPatternFill();
        PatternFill pf = bf.getFillInfo().getPatternFill();
        pf.setPatternType(PatternType.None);
        pf.getBackColor().setValue(-1);
        pf.getPatternColor().setValue(0);

        borderFillIDForCell = hwpFile.getDocInfo().getBorderFillList().size();
        return borderFillIDForCell;
    }

    /**
     * Set border properties
     */
    private void setBorder(EachBorder border, BorderType type) {
        border.setType(type);
        border.setThickness(BorderThickness.MM0_5);
        border.getColor().setValue(0x0);
    }

    /**
     * Convert millimeters to HWP units
     */
    private long mmToHwp(double mm) {
        return (long) (mm * 72000.0f / 254.0f + 0.5f);
    }

    /**
     * Convert points to line height
     */
    private int ptToLineHeight(double pt) {
        return (int) (pt * 100.0f);
    }

    /**
     * Inner class to store cell merge information
     */
    private static class CellMergeInfo {
        int startRow;
        int startCol;
        int rowSpan = 1;
        int colSpan = 1;
    }
}
