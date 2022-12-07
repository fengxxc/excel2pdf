package com.github.fengxxc;

import com.github.fengxxc.util.ExcelUtil;
import com.github.fengxxc.util.ITextUtil;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.properties.UnitValue;
import com.itextpdf.layout.properties.VerticalAlignment;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.function.Consumer;

/**
 * @author fengxxc
 * @date 2022-12-05
 */
public class Excel2PDF {
    /**
     * Excel 转 PDF
     * @param is Excel文件 输入流
     * @param os PDF文件 输出流
     * @throws IOException
     */
    public static void process(InputStream is, OutputStream os) throws IOException {
        process(is, os, null, null);
    }

    /**
     * Excel 转 PDF
     * @param is Excel文件 输入流
     * @param os PDF文件 输出流
     * @param documentCallback document建立后的回调函数
     * @throws IOException
     */
    public static void process(InputStream is, OutputStream os, Consumer<Document> documentCallback) throws IOException {
        process(is, os, null, documentCallback);
    }

    /**
     * Excel 转 PDF
     * @param is Excel文件 输入流
     * @param os PDF文件 输出流
     * @param columnWidths PDF表格列宽
     * @param documentCallback document建立后的回调函数
     * @throws IOException
     */
    public static void process(InputStream is, OutputStream os, UnitValue[] columnWidths, Consumer<Document> documentCallback) throws IOException {
        final Workbook workbook = WorkbookFactory.create(is);
        // final XSSFWorkbook workbook = new XSSFWorkbook(is);
        final Sheet sheet0 = workbook.getSheetAt(0); // TODO all sheet

        // init pdf document
        final PdfWriter pdfWriter = new PdfWriter(os);
        final PdfDocument pdfDocument = new PdfDocument(pdfWriter);
        final Document document = new Document(pdfDocument);
        if (documentCallback != null) {
            documentCallback.accept(document);
        }

        if (columnWidths == null) {
            columnWidths = ExcelUtil.findMaxRowColWidths(sheet0);
        }
        final Table table = new Table(columnWidths, false);
        table.setWidth(new UnitValue(UnitValue.PERCENT, 100)).setVerticalAlignment(VerticalAlignment.MIDDLE);

        final List<CellRangeAddress> combineCellList = ExcelUtil.getCombineCellList(sheet0);
        for (Row row : sheet0) {
            int colIdx = 0;
            for (Cell cell : row) {
                // type
                final CellType cellType = cell.getCellType();
                // value
                final String value = ExcelUtil.getCellValue(workbook, cell);
                // style
                final CellStyle cellStyle = cell.getCellStyle();
                // cell width
                final int columnWidth = sheet0.getColumnWidth(cell.getColumnIndex());
                // is combine cell?
                final CellRangeAddress cellRangeAddress = ExcelUtil.getCombineCellRangeAddress(combineCellList, cell);
                final boolean isCombineCell = cellRangeAddress != null ? true : false;
                final boolean isFirstInCombineCell = ExcelUtil.isFirstInCombineCell(cellRangeAddress, cell);

                // log(cell.getRowIndex() + "," + cell.getColumnIndex() + ": " + value + ", type: " + cellType + ", width: " + columnWidth + ", combine: " + isCombineCell);

                // to pdf
                // 补充空单元格
                for (int i = 0; i < cell.getColumnIndex() - colIdx; i++) {
                    table.addCell(ITextUtil.cellNull());
                }
                int rowSpan = 1;
                int colSpan = 1;
                if (isFirstInCombineCell) {
                    rowSpan = cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow() + 1;
                    colSpan = cellRangeAddress.getLastColumn() - cellRangeAddress.getFirstColumn() + 1;
                }
                final com.itextpdf.layout.element.Cell pdfCell = ITextUtil.cell(rowSpan, colSpan, value);
                setPdfTableCellStyle(workbook, cellStyle, pdfCell);
                if (!isCombineCell || isFirstInCombineCell) {
                    table.addCell(pdfCell);
                }
                colIdx = cell.getColumnIndex() + 1;
            }
            table.startNewRow();
        }
        document.add(table);

        document.close();
    }

    private static void setPdfTableCellStyle(Workbook workbook, CellStyle cellStyle, com.itextpdf.layout.element.Cell pdfCell) {
        // border style
        pdfCell.setBorderTop(ITextUtil.POI_TO_ITEXT_BORDER.get(cellStyle.getBorderTop().getCode()));
        pdfCell.setBorderRight(ITextUtil.POI_TO_ITEXT_BORDER.get(cellStyle.getBorderRight().getCode()));
        pdfCell.setBorderBottom(ITextUtil.POI_TO_ITEXT_BORDER.get(cellStyle.getBorderBottom().getCode()));
        pdfCell.setBorderLeft(ITextUtil.POI_TO_ITEXT_BORDER.get(cellStyle.getBorderLeft().getCode()));

        // background color
        final Color fillForegroundColorColor = cellStyle.getFillForegroundColorColor();
        final int[] colorARGB = ExcelUtil.getColorARGB(fillForegroundColorColor);
        if (colorARGB != null) {
            pdfCell.setBackgroundColor(new DeviceRgb(colorARGB[1], colorARGB[2], colorARGB[3]));
        }

        // does font bold
        final int fontIndex = cellStyle.getFontIndexAsInt();
        final Font font = workbook.getFontAt(fontIndex);
        final boolean bold = font.getBold();
        if (bold) {
            pdfCell.setBold();
        }

        // font color
        final int[] fontColorRGB = ExcelUtil.getFontColorARGB(font, workbook);
        if (fontColorRGB != null) {
            pdfCell.setFontColor(new DeviceRgb(fontColorRGB[1], fontColorRGB[2], fontColorRGB[3]));
        }


        // font size
        final short fontHeightInPoints = font.getFontHeightInPoints();
        pdfCell.setFontSize(fontHeightInPoints);

        // text align
        final HorizontalAlignment alignment = cellStyle.getAlignment();
        switch (alignment) {
            case FILL:
            case JUSTIFY:
                pdfCell.setTextAlignment(TextAlignment.JUSTIFIED);
                break;
            case RIGHT:
                pdfCell.setTextAlignment(TextAlignment.RIGHT);
                break;
            case CENTER:
            case CENTER_SELECTION:
                pdfCell.setTextAlignment(TextAlignment.CENTER);
                break;
            case LEFT:
            case GENERAL:
                pdfCell.setTextAlignment(TextAlignment.LEFT);
                break;
            case DISTRIBUTED:
                pdfCell.setTextAlignment(TextAlignment.JUSTIFIED_ALL);
                break;
            default:
        }

        // vertical align
        final org.apache.poi.ss.usermodel.VerticalAlignment verticalAlignment = cellStyle.getVerticalAlignment();
        switch (verticalAlignment) {
            case DISTRIBUTED:
            case CENTER:
            case JUSTIFY:
                pdfCell.setVerticalAlignment(VerticalAlignment.MIDDLE);
                break;
            case TOP:
                pdfCell.setVerticalAlignment(VerticalAlignment.TOP);
                break;
            case BOTTOM:
                pdfCell.setVerticalAlignment(VerticalAlignment.BOTTOM);
                break;
            default:
        }
    }

    private static void log(String text) {
        System.out.println(text);
    }
}
