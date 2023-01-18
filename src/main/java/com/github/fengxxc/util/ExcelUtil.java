package com.github.fengxxc.util;

import com.itextpdf.layout.properties.UnitValue;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

/**
 * @author fengxxc
 * @date 2022-12-05
 */
public class ExcelUtil {
    public static Cell getAboveCell(Cell cell) {
        final int rowIndex = cell.getRowIndex();
        final int columnIndex = cell.getColumnIndex();
        final Sheet sheet = cell.getSheet();
        return getAboveCell(rowIndex, columnIndex, sheet);
    }

    public static Cell getAboveCell(int rowIndex, int columnIndex, Sheet sheet) {
        if (rowIndex <= 0) {
            return null;
        }
        final Row aboveRow = sheet.getRow(rowIndex - 1);
        if (aboveRow == null) {
            return null;
        }
        final Cell aboveRowCell = aboveRow.getCell(columnIndex);
        return aboveRowCell;
    }

    public static Cell getLeftCell(Cell cell) {
        final int columnIndex = cell.getColumnIndex();
        if (columnIndex <= 0) {
            return null;
        }
        final Cell leftCell = cell.getRow().getCell(columnIndex - 1);
        return leftCell;
    }

    /**
     * 找到最大行，并获取各列宽百分比值
     * @param sheet
     * @return
     */
    public static UnitValue[] findMaxRowColWidths(Sheet sheet) {
        Row maxRow = null;
        int max = 0;
        for (Row row : sheet) {
            final int count = row.getPhysicalNumberOfCells();
            if (count > max) {
                maxRow = row;
                max = count;
            }
        }
        int sum = 0;
        int[] colWidthCache = new int[max];
        for (int i = 0; i < max; i++) {
            final int width = sheet.getColumnWidth(i);
            sum += width;
            colWidthCache[i] = width;
        }
        final UnitValue[] unitValues = new UnitValue[max];
        for (int i = 0; i < colWidthCache.length; i++) {
            final float widthPercent = (float) colWidthCache[i] / sum * 100;
            unitValues[i] = new UnitValue(UnitValue.PERCENT, widthPercent);
        }
        return unitValues;
    }

    /**
     * ARGB 十六进制形式 转 数值形式
     * "FFD0CECE" -> [255, 208, 206, 206]
     * @param argbHex
     * @return
     */
    public static int[] argbHex2argb(String argbHex) {
        String aHex = argbHex.substring(0, 2);
        String rHex = argbHex.substring(2, 4);
        String gHex = argbHex.substring(4, 6);
        String bHex = argbHex.substring(6, 8);
        int a2 = Integer.parseInt(aHex, 16);
        int r2 = Integer.parseInt(rHex, 16);
        int g2 = Integer.parseInt(gHex, 16);
        int b2 = Integer.parseInt(bHex, 16);
        return new int[]{a2, r2, g2, b2};
    }

    public static int[] getColorARGB(Color color) {
        if (color == null) {
            return null;
        }
        int[] result = new int[4];
        if (color instanceof XSSFColor) {
            XSSFColor xc = (XSSFColor) color;
            // final byte[] argb = xc.getARGB();
            final String argbHex = xc.getARGBHex();
            if (argbHex != null) {
                final int[] argb = argbHex2argb(argbHex);
                for (int i = 0; i < argb.length; i++) {
                    if (argb[i] == -1) {
                        result[i] = 255;
                    } else {
                        result[i] = argb[i];
                    }
                }
                // result = new int[]{argb[0], argb[1], argb[2], argb[3]};
            }
        } else if (color instanceof HSSFColor) {
            HSSFColor hc = (HSSFColor) color;
            short[] triplet = hc.getTriplet();
            if (triplet != null) {
                if (triplet[0] == 0 && triplet[1] == 0 && triplet[2] == 0) {
                    // 没有颜色会返回 0, 0, 0，这代表黑色，得重置为 255, 255, 255
                    result = new int[]{0xff, 0xff, 0xff, 0xff};
                } else {
                    result = new int[]{0xff, triplet[0], triplet[1], triplet[2]};
                }
            }
        }
        return result;
    }

    public static int[] getFontColorARGB(Font font, Workbook workbook) {
        if (font == null || workbook == null) {
            return null;
        }
        int[] result = new int[4];
        if (font instanceof XSSFFont) {
            final XSSFColor xssfColor = ((XSSFFont) font).getXSSFColor();
            if (xssfColor == null) {
                return null;
            }
            // byte[] fontColorRGB = xssfColor.getRGB();
            final String argbHex = xssfColor.getARGBHex();
            if (argbHex != null) {
                final int[] argb = argbHex2argb(argbHex);
                for (int i = 0; i < argb.length; i++) {
                    if (argb[i] == -1) {
                        result[i] = 255;
                    } else {
                        result[i] = argb[i];
                    }
                }
            }
        } else if (font instanceof HSSFFont) {
            final HSSFColor hssfColor = ((HSSFFont) font).getHSSFColor(((HSSFWorkbook) workbook));
            if (hssfColor == null) {
                return null;
            }
            short[] triplet = hssfColor.getTriplet();
            if (triplet != null) {
                result = new int[]{0xff, triplet[0], triplet[1], triplet[2]};
            }
        }
        return result;
    }

    /**
     * 获取合并单元格集合
     * @param sheet
     * @return
     */
    public static List<CellRangeAddress> getCombineCellList(Sheet sheet) {
        List<CellRangeAddress> list = new ArrayList<>();
        //获得一个 sheet 中合并单元格的数量
        int sheetMergerCount = sheet.getNumMergedRegions();
        //遍历所有的合并单元格
        for(int i = 0; i < sheetMergerCount; i++) {
            //获得合并单元格保存进list中
            CellRangeAddress ca = sheet.getMergedRegion(i);
            list.add(ca);
        }
        return list;
    }

    /**
     * 获取cell的合并单元格信息
     * 若cell不是合并单元格，返回null
     * @param combineCellList
     * @param cell
     * @return
     */
    public static CellRangeAddress getCombineCellRangeAddress(List<CellRangeAddress> combineCellList, Cell cell) {
        for(CellRangeAddress ca : combineCellList) {
            if(ca.isInRange(cell)) {
                return ca;
            }
        }
        return null;
    }

    /**
     * 是否是合并单元格中的第一个单元格
     * @param cellRangeAddress
     * @param cell
     * @return
     */
    public static boolean isFirstInCombineCell(CellRangeAddress cellRangeAddress, Cell cell) {
        if (cellRangeAddress == null) {
            return false;
        }
        if (cell.getRowIndex() == cellRangeAddress.getFirstRow() && cell.getColumnIndex() == cellRangeAddress.getFirstColumn()) {
            return true;
        }
        return false;
    }

    public static String getCellValue(Workbook workbook, Cell cell) {
        String cellValue = "";
        if (cell == null) {
            return cellValue;
        }
        switch (cell.getCellType()) {
            case NUMERIC:
                cellValue = getDateValue(cell.getCellStyle().getDataFormat(), cell.getCellStyle().getDataFormatString(),
                        cell.getNumericCellValue());
                if (cellValue == null) {
                    cellValue = String.valueOf(cell.getNumericCellValue());
                }
                break;
            case STRING:
                cellValue = String.valueOf(cell.getStringCellValue());
                break;
            case BOOLEAN:
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA:
                /**
                 * 格式化单元格
                 */
                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                cellValue = getCellValue(evaluator.evaluate(cell));
                break;
            case BLANK:
                cellValue = "";
                break;
            case ERROR:
                cellValue = String.valueOf(cell.getErrorCellValue());
                break;
            case _NONE:
                cellValue = "";
                break;
            default:
                cellValue = "";
                break;
        }
        return cellValue;
    }

    public static String getCellValue(CellValue formulaValue) {
        String cellValue = "";
        if (formulaValue == null) {
            return cellValue;
        }

        switch (formulaValue.getCellType()) {
            case NUMERIC:
                cellValue = String.valueOf(formulaValue.getNumberValue());
                break;
            case STRING:
                cellValue = String.valueOf(formulaValue.getStringValue());
                break;
            case BOOLEAN:
                cellValue = String.valueOf(formulaValue.getBooleanValue());
                break;
            case BLANK:
                cellValue = "";
                break;
            case ERROR:
                cellValue = String.valueOf(formulaValue.getErrorValue());
                break;
            case _NONE:
                cellValue = "";
                break;
            default:
                cellValue = "";
                break;
        }
        return cellValue;

    }

    public static String getDateValue(Short dataFormat, String dataFormatString, double value) {
        if (!DateUtil.isValidExcelDate(value)) {
            return null;
        }

        Date date = DateUtil.getJavaDate(value);

        /**
         * 年月日时分秒
         */
        if (Constants.EXCEL_FORMAT_INDEX_DATE_NYRSFM_STRING.contains(dataFormatString)) {
            return Constants.COMMON_DATE_FORMAT.format(date);
        }

        /**
         * 年月日
         */
        if (Constants.EXCEL_FORMAT_INDEX_DATE_NYR_STRING.contains(dataFormatString)) {
            return Constants.COMMON_DATE_FORMAT_NYR.format(date);
        }

        /**
         * 年月 中文
         */
        if (Constants.EXCEL_FORMAT_INDEX_DATE_NY_STRING_CN.equals(dataFormatString)) {
            return Constants.COMMON_DATE_FORMAT_NY_CN.format(date);
        }

        /**
         * 年月
         */
        if (Constants.EXCEL_FORMAT_INDEX_DATE_NY_STRING.contains(dataFormatString) || Constants.EXCEL_FORMAT_INDEX_DATA_EXACT_NY.equals(dataFormat)) {
            return Constants.COMMON_DATE_FORMAT_NY.format(date);
        }

        /**
         * 月日
         */
        if (Constants.EXCEL_FORMAT_INDEX_DATE_YR_STRING.contains(dataFormatString) || Constants.EXCEL_FORMAT_INDEX_DATA_EXACT_YR.equals(dataFormat)) {
            return Constants.COMMON_DATE_FORMAT_YR.format(date);

        }

        /**
         * 月
         */
        if (Constants.EXCEL_FORMAT_INDEX_DATE_Y_STRING.contains(dataFormatString)) {
            return Constants.COMMON_DATE_FORMAT_Y.format(date);
        }

        /**
         * 时间格式
         */
        if (Constants.EXCEL_FORMAT_INDEX_TIME_STRING.contains(dataFormatString) || Constants.EXCEL_FORMAT_INDEX_TIME_EXACT.contains(dataFormat)) {
            return Constants.COMMON_TIME_FORMAT.format(DateUtil.getJavaDate(value));
        }

        return null;
    }

    public interface Constants {

        /**
         * 年月日时分秒 默认格式
         */
        SimpleDateFormat COMMON_DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        /**
         * 时间 默认格式
         */
        SimpleDateFormat COMMON_TIME_FORMAT = new SimpleDateFormat("HH:mm:ss");
        /**
         * 年月日 默认格式
         */
        SimpleDateFormat COMMON_DATE_FORMAT_NYR = new SimpleDateFormat("yyyy-MM-dd");
        /**
         * 年月 默认格式
         */
        SimpleDateFormat COMMON_DATE_FORMAT_NY = new SimpleDateFormat("yyyy-MM");
        /**
         * 年月 中文格式
         */
        SimpleDateFormat COMMON_DATE_FORMAT_NY_CN = new SimpleDateFormat("yyyy年MM月");
        /**
         * 月日 默认格式
         */
        SimpleDateFormat COMMON_DATE_FORMAT_YR = new SimpleDateFormat("MM-dd");
        /**
         * 月 默认格式
         */
        SimpleDateFormat COMMON_DATE_FORMAT_Y = new SimpleDateFormat("MM");

        /**
         * date-年月日时分秒
         */
        List<String> EXCEL_FORMAT_INDEX_DATE_NYRSFM_STRING = Arrays.asList(
                "yyyy/m/d\\ h:mm;@", "m/d/yy h:mm", "yyyy/m/d\\ h:mm\\ AM/PM",
                "[$-409]yyyy/m/d\\ h:mm\\ AM/PM;@", "yyyy/mm/dd\\ hh:mm:dd", "yyyy/mm/dd\\ hh:mm", "yyyy/m/d\\ h:m", "yyyy/m/d\\ h:m:s",
                "yyyy/m/d\\ h:mm", "m/d/yy h:mm;@", "yyyy/m/d\\ h:mm\\ AM/PM;@"
        );
        /**
         * date-年月日
         */
        List<String> EXCEL_FORMAT_INDEX_DATE_NYR_STRING = Arrays.asList(
                "m/d/yy", "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy",
                "[DBNum1][$-804]yyyy\"年\"m\"月\"d\"日\";@", "yyyy\"年\"m\"月\"d\"日\";@", "yyyy/m/d;@", "yy/m/d;@", "m/d/yy;@",
                "[$-409]d/mmm/yy", "[$-409]dd/mmm/yy;@", "reserved-0x1F", "reserved-0x1E", "mm/dd/yy;@", "yyyy/mm/dd", "d-mmm-yy",
                "[$-409]d\\-mmm\\-yy;@", "[$-409]d\\-mmm\\-yy", "[$-409]dd\\-mmm\\-yy;@", "[$-409]dd\\-mmm\\-yy",
                "[DBNum1][$-804]yyyy\"年\"m\"月\"d\"日\"", "yy/m/d", "mm/dd/yy", "dd\\-mmm\\-yy"
        );

        String EXCEL_FORMAT_INDEX_DATE_NY_STRING_CN = "yyyy\"年\"m\"月\";@";
        /**
         * date-年月
         */
        List<String> EXCEL_FORMAT_INDEX_DATE_NY_STRING = Arrays.asList(
                "[DBNum1][$-804]yyyy\"年\"m\"月\";@", "[DBNum1][$-804]yyyy\"年\"m\"月\"",
                "yyyy\"年\"m\"月\"", "[$-409]mmm\\-yy;@", "[$-409]mmm\\-yy",
                "[$-409]mmm/yy;@", "[$-409]mmm/yy", "[$-409]mmmm/yy;@","[$-409]mmmm/yy",
                "[$-409]mmmmm/yy;@", "[$-409]mmmmm/yy", "mmm-yy", "yyyy/mm", "mmm/yyyy",
                "[$-409]mmmm\\-yy;@", "[$-409]mmmmm\\-yy;@", "mmmm\\-yy", "mmmmm\\-yy"
        );
        /**
         * date-月日
         */
        List<String> EXCEL_FORMAT_INDEX_DATE_YR_STRING = Arrays.asList(
                "[DBNum1][$-804]m\"月\"d\"日\";@", "[DBNum1][$-804]m\"月\"d\"日\"",
                "m\"月\"d\"日\";@", "m\"月\"d\"日\"", "[$-409]d/mmm;@", "[$-409]d/mmm",
                "m/d;@", "m/d", "d-mmm", "d-mmm;@", "mm/dd", "mm/dd;@", "[$-409]d\\-mmm;@", "[$-409]d\\-mmm"
        );
        /**
         * date-月X
         */
        List<String> EXCEL_FORMAT_INDEX_DATE_Y_STRING = Arrays.asList("[$-409]mmmmm;@","mmmmm","[$-409]mmmmm");
        /**
         * time - 时间
         */
        List<String> EXCEL_FORMAT_INDEX_TIME_STRING = Arrays.asList(
                "mm:ss.0", "h:mm", "h:mm\\ AM/PM", "h:mm:ss", "h:mm:ss\\ AM/PM",
                "reserved-0x20", "reserved-0x21", "[DBNum1]h\"时\"mm\"分\"", "[DBNum1]上午/下午h\"时\"mm\"分\"", "mm:ss",
                "[h]:mm:ss", "h:mm:ss;@", "[$-409]h:mm:ss\\ AM/PM;@", "h:mm;@", "[$-409]h:mm\\ AM/PM;@",
                "h\"时\"mm\"分\";@", "h\"时\"mm\"分\"\\ AM/PM;@", "h\"时\"mm\"分\"ss\"秒\";@", "h\"时\"mm\"分\"ss\"秒\"_ AM/PM;@", "上午/下午h\"时\"mm\"分\";@",
                "上午/下午h\"时\"mm\"分\"ss\"秒\";@", "[DBNum1][$-804]h\"时\"mm\"分\";@", "[DBNum1][$-804]上午/下午h\"时\"mm\"分\";@", "h:mm AM/PM","h:mm:ss AM/PM",
                "[$-F400]h:mm:ss\\ AM/PM"
        );
        /**
         * date-当formatString为空的时候-年月
         */
        Short EXCEL_FORMAT_INDEX_DATA_EXACT_NY = 57;
        /**
         * date-当formatString为空的时候-月日
         */
        Short EXCEL_FORMAT_INDEX_DATA_EXACT_YR = 58;
        /**
         * time-当formatString为空的时候-时间
         */
        List<Short> EXCEL_FORMAT_INDEX_TIME_EXACT = Arrays.asList(new Short[]{55, 56});

    }
}
