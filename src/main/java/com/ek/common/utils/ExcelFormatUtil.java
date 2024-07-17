package com.ek.common.utils;


import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.UUID;

public class ExcelFormatUtil {
    private static final String DOUBLE_PATTERN1 = "^[-\\+]?\\d+(\\.\\d*)?|\\.\\d+$";
    private static final String DOUBLE_PATTERN2 = "[0]";
    private static ExcelFormatUtil instance;

    public ExcelFormatUtil() {
    }

    public static ExcelFormatUtil getInstance() {
        if (instance == null) {
            instance = new ExcelFormatUtil();
        }

        return instance;
    }

    public void setFreezeAndFilter(Sheet sheet) {
        sheet.createFreezePane(0, 1, 0, 1);
        int count = sheet.getRow(0).getPhysicalNumberOfCells();
        if (count > 0) {
            CellRangeAddress cellAddresses = new CellRangeAddress(0, 0, 0, count - 1);
            sheet.setAutoFilter(cellAddresses);
        }

    }

    public void setFreezeAndFilterAdvanced(Sheet sheet, int rowIndex) {
        sheet.createFreezePane(0, 0, 0, 0);
        sheet.createFreezePane(0, rowIndex + 1, 0, rowIndex + 1);
        int count = sheet.getRow(0).getPhysicalNumberOfCells();
        if (count > 0) {
            CellRangeAddress cellAddresses = new CellRangeAddress(rowIndex, rowIndex, 0, count - 1);
            sheet.setAutoFilter(cellAddresses);
        }

    }

    public void setMergedAllRegion(Sheet sheet) {
        int lastrow = sheet.getPhysicalNumberOfRows();
        int lastcolumn = sheet.getRow(1).getPhysicalNumberOfCells();

        for(int i = 0; i < lastcolumn; ++i) {
            int begin = 0;
            int count = 0;

            for(int j = 0; j < lastrow - 1; ++j) {
                Cell cell1 = sheet.getRow(j).getCell(i);
                Cell cell2 = sheet.getRow(j + 1).getCell(i);
                if (cell1.getStringCellValue().equals(cell2.getStringCellValue())) {
                    ++count;
                    if (count == 1) {
                        begin = j;
                    }
                } else if (count > 0) {
                    this.startMerge(sheet, begin, begin + count, i, i);
                    count = 0;
                }

                if (j == lastrow - 2 && count > 0) {
                    this.startMerge(sheet, begin, begin + count, i, i);
                    begin = begin + count + 1;
                    count = 0;
                }
            }
        }

    }

    private void startMerge(Sheet sheet, int firstrow, int lastrown, int firstcolu, int lastcolu) {
        sheet.addMergedRegion(new CellRangeAddress(firstrow, lastrown, firstcolu, lastcolu));
    }

    public void setMergedIntervalColumnRegion(Sheet sheet, int start, int stop) {
        int endrow = sheet.getPhysicalNumberOfRows();

        for(int i = start; i < stop; ++i) {
            int begin = 0;
            int count = 0;

            for(int j = 0; j < endrow - 1; ++j) {
                Cell cell1 = sheet.getRow(j).getCell(i);
                Cell cell2 = sheet.getRow(j + 1).getCell(i);
                if (cell1.getStringCellValue().equals(cell2.getStringCellValue())) {
                    ++count;
                    if (count == 1) {
                        begin = j;
                    }
                } else if (count > 0) {
                    this.startMerge(sheet, begin, begin + count, i, i);
                    count = 0;
                }

                if (j == endrow - 2 && count > 0) {
                    this.startMerge(sheet, begin, begin + count, i, i);
                    begin = begin + count + 1;
                    count = 0;
                }
            }
        }

    }

    public void setMergedSomeColumnRegion(Sheet sheet, int[] numbers) {
        int lastrow = sheet.getPhysicalNumberOfRows();
        int[] var4 = numbers;
        int var5 = numbers.length;

        for(int var6 = 0; var6 < var5; ++var6) {
            int i = var4[var6];
            int begin = 0;
            int count = 0;

            for(int j = 0; j < lastrow - 1; ++j) {
                Cell cell1 = sheet.getRow(j).getCell(i);
                Cell cell2 = sheet.getRow(j + 1).getCell(i);
                if (cell1.getStringCellValue().equals(cell2.getStringCellValue())) {
                    ++count;
                    if (count == 1) {
                        begin = j;
                    }
                } else if (count > 0) {
                    this.startMerge(sheet, begin, begin + count, i, i);
                    count = 0;
                }

                if (j == lastrow - 2 && count > 0) {
                    this.startMerge(sheet, begin, begin + count, i, i);
                    begin = begin + count + 1;
                    count = 0;
                }
            }
        }

    }

    public void setMergedSomeRegion(Sheet sheet, int startcolu, int stopcolu, int startrow, int stoprow) {
        int lastrow = sheet.getPhysicalNumberOfRows();
        if (stoprow > lastrow) {
            stoprow = lastrow;
        }

        for(int i = startcolu; i < stopcolu; ++i) {
            int begin = startrow;
            int count = 0;

            for(int j = startrow; j < stoprow - 1; ++j) {
                Cell cell1 = sheet.getRow(j).getCell(i);
                Cell cell2 = sheet.getRow(j + 1).getCell(i);
                if (cell1.getStringCellValue().equals(cell2.getStringCellValue())) {
                    ++count;
                    if (count == 1) {
                        begin = j;
                    }
                } else if (count > 0) {
                    this.startMerge(sheet, begin, begin + count, i, i);
                    count = 0;
                }

                if (j == lastrow - 2 && count > 0) {
                    this.startMerge(sheet, begin, begin + count, i, i);
                    begin = begin + count + 1;
                    count = 0;
                }
            }
        }

    }

    public void setMergedRegionByOneColumn(Sheet sheet, int flag, int start, int stop) {
        int endrow = sheet.getPhysicalNumberOfRows();
        int begin = 0;
        int count = 0;

        for(int j = 0; j < endrow - 1; ++j) {
            Cell cell1 = sheet.getRow(j).getCell(flag);
            Cell cell2 = sheet.getRow(j + 1).getCell(flag);
            if (cell1.getStringCellValue().equals(cell2.getStringCellValue())) {
                ++count;
                if (count == 1) {
                    begin = j;
                }
            } else if (count > 0) {
                this.setMergedSomeRegions(sheet, start, stop, begin, begin + count + 1);
                count = 0;
            }

            if (j == endrow - 2 && count > 0) {
                this.setMergedSomeRegions(sheet, start, stop, begin, begin + count + 1);
                begin = begin + count + 1;
                count = 0;
            }
        }

    }

    private void setMergedSomeRegions(Sheet sheet, int startcolu, int stopcolu, int startrow, int stoprow) {
        for(int i = startcolu; i < stopcolu; ++i) {
            int begin = startrow;
            int count = 0;

            for(int j = startrow; j < stoprow - 1; ++j) {
                Cell cell1 = sheet.getRow(j).getCell(i);
                Cell cell2 = sheet.getRow(j + 1).getCell(i);
                int delete;
                if (cell1.getStringCellValue().equals(cell2.getStringCellValue())) {
                    ++count;
                    if (count == 1) {
                        begin = j;
                    }
                } else if (count > 0) {
                    if (i == 0) {
                        for(delete = 1; delete <= count; ++delete) {
                            sheet.getRow(begin + delete).getCell(i).setCellValue("");
                        }
                    }

                    this.startMerge(sheet, begin, begin + count, i, i);
                    count = 0;
                }

                if (j == stoprow - 2 && count > 0) {
                    if (i == 0) {
                        for(delete = 1; delete <= count; ++delete) {
                            sheet.getRow(begin + delete).getCell(i).setCellValue("");
                        }
                    }

                    this.startMerge(sheet, begin, begin + count, i, i);
                    begin = begin + count + 1;
                    count = 0;
                }
            }
        }

    }

    public CellStyle setCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        Font font = workbook.createFont();
        font.setFontName("黑体");
        font.setFontHeightInPoints((short)13);
        cellStyle.setFont(font);
        return cellStyle;
    }

    public CellStyle setContentCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setWrapText(true);
        Font font = workbook.createFont();
        font.setFontName("Arial");
        cellStyle.setFont(font);
        return cellStyle;
    }

    public void exportExcelInLocal(Workbook workbook, String filename) {
        int count = filename.lastIndexOf("/");
        String path = filename.substring(0, count + 1);
        String nameleft = filename.substring(0, filename.lastIndexOf("."));
        String nameright = filename.substring(filename.lastIndexOf("."));
        FileOutputStream fileOutputStream = null;
        if ((nameright.equals(".xls") || nameright.equals(".xlsx")) && count != -1) {
            try {
                File file = new File(path);
                if (!file.exists()) {
                    try {
                        file.mkdirs();
                    } catch (Exception var19) {
                    }
                } else {
                    int flag = 0;
                    file = new File(filename);
                    if (file.exists()) {
                        while(flag < 5) {
                            file = new File(filename);
                            if (!file.exists()) {
                                break;
                            }

                            ++flag;
                            filename = nameleft + "(" + flag + ")" + nameright;
                        }
                    }
                }

                fileOutputStream = new FileOutputStream(filename);
                workbook.write(fileOutputStream);
            } catch (IOException var20) {
            } finally {
                try {
                    workbook.close();
                    fileOutputStream.close();
                } catch (IOException var18) {
                }

            }
        }

    }

    public void exportExcelOverwriteLocal(Workbook workbook, String filename) {
        int count = filename.lastIndexOf("/");
        String path = filename.substring(0, count + 1);
        String nameright = filename.substring(filename.lastIndexOf("."));
        FileOutputStream fileOutputStream = null;
        if (nameright.equals(".xls") || nameright.equals(".xlsx")) {
            try {
                File file = new File(path);
                if (!file.exists()) {
                    try {
                        file.mkdirs();
                    } catch (Exception var18) {
                    }
                }

                fileOutputStream = new FileOutputStream(filename);
                workbook.write(fileOutputStream);
            } catch (IOException var19) {
            } finally {
                try {
                    if (fileOutputStream != null) {
                        workbook.close();
                        fileOutputStream.close();
                    }
                } catch (IOException var17) {
                }

            }
        }

    }

    public void exportExcelNotClose(Workbook workbook, String filename) {
        int count = filename.lastIndexOf("/");
        String path = filename.substring(0, count + 1);
        String nameright = filename.substring(filename.lastIndexOf("."));
        FileOutputStream fileOutputStream = null;
        if (nameright.equals(".xls") || nameright.equals(".xlsx")) {
            try {
                File file = new File(path);
                if (!file.exists()) {
                    try {
                        file.mkdirs();
                    } catch (Exception var18) {
                    }
                }

                fileOutputStream = new FileOutputStream(filename);
                workbook.write(fileOutputStream);
            } catch (IOException var19) {
            } finally {
                try {
                    if (fileOutputStream != null) {
                        fileOutputStream.close();
                    }
                } catch (IOException var17) {
                }

            }
        }

    }

    public void setRowHeight(Sheet sheet, int rowSize) {
        int rowNum = sheet.getPhysicalNumberOfRows();

        for(int i = 0; i < rowNum; ++i) {
            sheet.getRow(i).setHeightInPoints((float)rowSize);
        }

    }

    public void setOtherRowHeight(Sheet sheet, int rowSize) {
        int rowNum = sheet.getPhysicalNumberOfRows();

        for(int i = 1; i < rowNum; ++i) {
            sheet.getRow(i).setHeightInPoints((float)rowSize);
        }

    }

    public void getImgHWsetOtherRowHeight(Sheet sheet, int rowSize) {
        int rowNum = sheet.getPhysicalNumberOfRows();

        for(int i = 1; i < rowNum; ++i) {
            sheet.getRow(i).setHeightInPoints((float)rowSize);
        }

    }

    public void setOneRowHeight(Sheet sheet, int rowNum, int rowSize) {
        sheet.getRow(rowNum).setHeightInPoints((float)rowSize);
    }

    public void setAutoRowHeight(Sheet sheet, int[] colWidth) {
        int rowCount = sheet.getPhysicalNumberOfRows();
        sheet.getRow(0).setHeightInPoints(25.0F);

        for(int rowNum = 1; rowNum < rowCount; ++rowNum) {
            double rowSize = 0.0;
            int columnCount = sheet.getRow(rowNum).getPhysicalNumberOfCells();

            for(int colNum = 0; colNum < columnCount; ++colNum) {
                String value = sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
                if (value != null && value.length() != 0) {
                    int valuelength = value.getBytes().length;
                    double size = (double)valuelength / (2.1 * (double)colWidth[colNum]);
                    if (size > rowSize) {
                        rowSize = size;
                    }
                }
            }

            if (rowSize < 1.0) {
                rowSize = 1.0;
            }

            sheet.getRow(rowNum).setHeightInPoints((float)(rowSize * 25.0));
        }

    }

    public void setAutoColumnWidth(Sheet sheet) {
        int columns;
        for(columns = 0; columns < sheet.getPhysicalNumberOfRows(); ++columns) {
            sheet.getRow(columns).setHeightInPoints(25.0F);
        }

        columns = sheet.getRow(0).getPhysicalNumberOfCells();
        if (sheet instanceof SXSSFSheet) {
            ((SXSSFSheet)sheet).trackAllColumnsForAutoSizing();
        }

        for(int i = 0; i < columns; ++i) {
            sheet.autoSizeColumn(i, true);
            sheet.setColumnWidth(i, sheet.getColumnWidth(i) * 12 / 10);
        }

        this.setFreezeAndFilter(sheet);
    }

    public void setColumnWidth(Sheet sheet, int columnsize) {
        int columns = sheet.getRow(0).getPhysicalNumberOfCells();

        for(int i = 0; i < columns; ++i) {
            sheet.setColumnWidth(i, columnsize * 256);
        }

    }

    public void setEachColumnWidth(Sheet sheet, int[] columns) {
        int columncount = sheet.getRow(0).getPhysicalNumberOfCells();
        int col = 0;
        if (columns.length <= columncount) {
            int[] var5 = columns;
            int var6 = columns.length;

            for(int var7 = 0; var7 < var6; ++var7) {
                int colsize = var5[var7];
                sheet.setColumnWidth(col, colsize * 256);
                ++col;
            }
        }

        this.setFreezeAndFilter(sheet);
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
                case FORMULA:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        cellValue = String.valueOf(cell.getDateCellValue());
                    } else {
                        cellValue = String.valueOf(cell.getCellFormula());
                    }
                    break;
                case BLANK:
                    cellValue = "";
                    break;
                case ERROR:
                    cellValue = "非法字符";
                    break;
                default:
                    cellValue = "未知类型";
            }

            return cellValue;
        }
    }

    public static int getIntegerNumber(Cell cell) {
        String cellValue = getCellValue(cell);
        return !cellValue.isEmpty() && isNumeric(cellValue) ? Integer.parseInt(cellValue) : 0;
    }

    public static long getLongNumber(Cell cell) {
        String cellValue = getCellValue(cell);
        return !cellValue.isEmpty() && isNumeric(cellValue) ? Long.parseLong(cellValue) : 0L;
    }

    public static boolean isNumeric(String str) {
        int i = str.length();

        char chr;
        do {
            --i;
            if (i < 0) {
                return true;
            }

            chr = str.charAt(i);
        } while(chr >= '0' && chr <= '9');

        return false;
    }

    public static double getDoubleNumber(Cell cell) {
        String cellValue = getCellValue(cell);
        return !cellValue.isEmpty() && isDouble(cellValue) ? Double.parseDouble(cellValue) : 0.0;
    }

    public static double getDoubleValue(String cellValue) {
        return !cellValue.isEmpty() && isDouble(cellValue) ? Double.parseDouble(cellValue) : 0.0;
    }

    public static boolean isDouble(String str) {
        return str.matches("^[-\\+]?\\d+(\\.\\d*)?|\\.\\d+$") || str.matches("[0]");
    }

    public boolean isFixed() {
        return true;
    }

    public void addDropList(Workbook workbook, Sheet sheet, String[] strings, int col) {
        String hiddenSheetName = "a" + UUID.randomUUID().toString().replace("-", "").substring(1, 31);
        String formulaId = "form" + UUID.randomUUID().toString().replace("-", "");
        HSSFWorkbook workbook1 = (HSSFWorkbook)workbook;
        HSSFSheet hiddenSheet = workbook1.createSheet(hiddenSheetName);
        HSSFRow row = null;
        HSSFCell cell = null;

        for(int i = 0; i < strings.length; ++i) {
            row = hiddenSheet.createRow(i);
            cell = row.createCell(col);
            cell.setCellValue(strings[i]);
        }

        HSSFName namedCell = workbook1.createName();
        namedCell.setNameName(formulaId);
        if (strings.length != 0) {
            namedCell.setRefersToFormula(hiddenSheetName + "!A$1:A$" + strings.length);
        }

        HSSFSheet sheet1 = (HSSFSheet)sheet;
        HSSFDataValidationHelper dvHelper = new HSSFDataValidationHelper(sheet1);
        DataValidationConstraint dvConstraint = dvHelper.createFormulaListConstraint(formulaId);
        CellRangeAddressList addressList = new CellRangeAddressList(1, Integer.MAX_VALUE, col, col);
        DataValidation validation = dvHelper.createValidation(dvConstraint, addressList);
        sheet.addValidationData(validation);
        workbook1.setSheetHidden(workbook1.getSheetIndex(hiddenSheetName), true);
    }
}
