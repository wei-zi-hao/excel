package com.ek.common.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.UnsupportedEncodingException;
import java.text.NumberFormat;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ValideDataUti {
    private static final String DOUBLE_PATTERN1 = "^[-\\+]?\\d+(\\.\\d*)?|\\.\\d+$";
    private static final String DOUBLE_PATTERN2 = "[0]";
    private static NumberFormat numberFormat = NumberFormat.getInstance();
    private static final String DECIMAL_PATTERN = "([0])((\\.\\d{0,N})|(\\.))?";
    private static final String INTEGER_PATTERN = "[\\d&&[^0]](\\d{0,5})";
    private static final String LONGITIDU_PATTERN = "/^[\\-\\+]?(0?\\d{1,2}\\.\\d{1,5}|1[0-7]?\\d{1}\\.\\d{1,5}|180\\.0{1,5})$/";

    private ValideDataUti() {
    }

    public static boolean isSheetEmpty(Sheet sheet) {
        return null == sheet || sheet.getPhysicalNumberOfRows() <= 1 || sheet.getLastRowNum() <= 0;
    }

    public static String cleanString(String str) {
        String replace = "";
        if (str != null) {
            Pattern pattern = Pattern.compile("\\s*|\t|\r|\n");
            Matcher matcher = pattern.matcher(str);
            replace = matcher.replaceAll("");
        }

        return replace;
    }

    public static String getCellValue(Cell cell) {
        String cellValue = null;
        if (cell == null) {
            return "";
        } else {
            if (cell.getCellType() == CellType.NUMERIC) {
                cell.setCellType(CellType.STRING);
            }

            switch (cell.getCellType()) {
                case NUMERIC:
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                case STRING:
                    cellValue = String.valueOf(cell.getStringCellValue());
                    break;
                case BOOLEAN:
                    cellValue = String.valueOf(cell.getBooleanCellValue());
                    break;
                case BLANK:
                    cellValue = "";
                    break;
                case FORMULA:
                    if (cell.getCellStyle().getDataFormatString().indexOf("%") == -1 && !cell.getRow().getCell(1).getStringCellValue().equals("100G容量升级软件包")) {
                        cellValue = "未知类型";
                    } else if (cell.getRow().getCell(0).getStringCellValue() != "") {
                        cellValue = "未知类型";
                    } else {
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                case ERROR:
                    cellValue = "非法字符";
                    break;
                default:
                    cellValue = "未知类型";
            }

            String trim;
            if (cellValue.contains("\n")) {
                trim = cellValue.replace('\n', ' ');
                cellValue = trim;
            }

            if (cellValue.contains(" ") || cellValue.contains("　")) {
                trim = cellValue.replace('　', ' ').trim();
                cellValue = trim;
            }

            return cellValue;
        }
    }

    public static boolean isDouble(String str) {
        return str.matches("^[-\\+]?\\d+(\\.\\d*)?|\\.\\d+$") || str.matches("[0]");
    }

    public static boolean isPositiveInteger(String str) {
        return str.matches("[\\d&&[^0]](\\d{0,5})");
    }

    public static String getUTF8String(String content) {
        String utf8String = "";

        try {
            utf8String = new String(content.getBytes(), "UTF-8");
        } catch (UnsupportedEncodingException var3) {
        }

        return utf8String;
    }
}
