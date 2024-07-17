package com.ek.project.service;

import com.ek.common.utils.DateUtils;
import com.ek.common.utils.ExcelFormatUtil;
import com.ek.common.utils.ValideDataUti;
import com.ek.common.utils.file.FileUtils;
import com.ek.framework.web.damain.AjaxResult;
import com.ek.project.common.GlobalHashMapHolder;
import com.ek.project.common.dto.ExcelDetail;
import com.ek.project.common.dto.RawData;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Service
public class ExcelService {

    @Autowired
    private GlobalHashMapHolder globalHashMapHolder;

    private static final Logger log = LoggerFactory.getLogger(ExcelService.class);

    public AjaxResult getExcel(List<MultipartFile> multipartFiles) throws IOException {
        Workbook workbook_old = getInput(multipartFiles.get(0));
        Workbook workbook_new = getInput(multipartFiles.get(1));
        String workbookOldKey = DateUtils.dateTimeNow() + "old";
        String workbookNewKey = DateUtils.dateTimeNow() + "new";
        globalHashMapHolder.putIntoGlobalHashMap(workbookOldKey, workbook_old);
        globalHashMapHolder.putIntoGlobalHashMap(workbookNewKey, workbook_new);
        ArrayList<String> excelList = new ArrayList<>();
        for (int i = 0; i < workbook_new.getNumberOfSheets(); ++i) {
            Sheet sheet_new = workbook_new.getSheetAt(i);
            String sheetName = sheet_new.getSheetName();
            Sheet sheet_old = workbook_old.getSheet(sheetName.trim());
            if (sheet_old != null && (sheet_new.getRow(5).getCell(12) != null && sheet_new.getRow(5).getCell(12).getStringCellValue().equals("物料编码") || sheet_new.getRow(6).getCell(12) != null && sheet_new.getRow(6).getCell(12).getStringCellValue().equals("物料编码"))) {
                excelList.add(sheet_new.getSheetName());
            }
        }
        AjaxResult ajax = AjaxResult.success();
        ajax.put("excelList", excelList);
        ajax.put("workbookOldKey", workbookOldKey);
        ajax.put("workbookNewKey", workbookNewKey);
        return ajax;
    }

    public void downloadExcel(String fileName, HttpServletResponse response) throws IOException {
        response.setContentType(MediaType.APPLICATION_OCTET_STREAM_VALUE);
        FileUtils.setAttachmentResponseHeader(response, fileName);
        FileUtils.writeBytes((byte[]) globalHashMapHolder.getGlobalHashMap().get(fileName), response.getOutputStream());
        globalHashMapHolder.getGlobalHashMap().clear();
        log.info("清除缓存成功");

    }

    public AjaxResult handleExcel(ExcelDetail excelDetail) throws IOException {

        Workbook workbook_new = (Workbook) globalHashMapHolder.getGlobalHashMap().get(excelDetail.getWorkbookNewKey());
        Workbook workbook_old = (Workbook) globalHashMapHolder.getGlobalHashMap().get(excelDetail.getWorkbookOldKey());

        ExcelFormatUtil formatUtil = ExcelFormatUtil.getInstance();
        CellStyle colorCellStyleContentR = formatUtil.setContentCellStyle(workbook_new);
        CellStyle colorCellStyleContentY = formatUtil.setContentCellStyle(workbook_new);
        CellStyle colorCellStyleContentG = formatUtil.setContentCellStyle(workbook_new);
        CellStyle colorCellStyleContentGr = formatUtil.setContentCellStyle(workbook_old);
        colorCellStyleContentR.setFillForegroundColor(IndexedColors.PINK.index);
        colorCellStyleContentY.setFillForegroundColor(IndexedColors.GREEN.index);
        colorCellStyleContentG.setFillForegroundColor(IndexedColors.YELLOW.index);
        colorCellStyleContentGr.setFillForegroundColor(IndexedColors.BLUE.index);
        colorCellStyleContentR.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        colorCellStyleContentY.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        colorCellStyleContentG.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        colorCellStyleContentGr.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        List<String> selectedExcel = excelDetail.getSelectSheet();

        for (String selectSheet : excelDetail.getSelectSheet()) {
            log.info("需要标注的Excel页: {}", selectSheet);
        }


        Iterator var17 = selectedExcel.iterator();

        while (var17.hasNext()) {
            String excelName = (String) var17.next();
            log.info("开始处理" + excelName);
            Sheet sheet_new = workbook_new.getSheet(excelName);
            Sheet sheet_old = workbook_old.getSheet(excelName);
            List<String[]> excelContentsList_new = readExcelContent(sheet_new);
            List<String[]> excelContentsList_old = readExcelContent(sheet_old);
            List<Integer> grayRawIndex = new ArrayList();
            List<Integer> greenRawIndex = new ArrayList();
            List<Integer> yellowRawIndex = new ArrayList();
            HashMap<Integer, List<Integer>> redRawIndex = new HashMap();
            List<String> commonBoardId = new ArrayList();
            Map<String, RawData> new_BoardIds = new HashMap();
            Map<String, RawData> old_BoardIds = new HashMap();

            int f;
            String[] col_old;
            RawData rawData_old;
            Double realPrice;
            Double countDown;
            Double realPrice_2;
            for (f = 7; f < excelContentsList_new.size(); ++f) {
                col_old = (String[]) excelContentsList_new.get(f);
                if (col_old.length >= 13 && col_old[12] != "") {
                    if (col_old[12].equals("新板卡")) {
                        greenRawIndex.add(f);
                    } else {
                        rawData_old = new RawData();
                        rawData_old.setIndex(col_old[0]);
                        rawData_old.setName(col_old[1]);
                        rawData_old.setBoard_code(col_old[2]);
                        rawData_old.setUnit(col_old[3]);
                        realPrice = 0.0;
                        if (col_old[4] != "未知类型") {
                            rawData_old.setPrice(col_old[4]);
                            rawData_old.setDiscode(col_old[5]);
                            realPrice = Double.valueOf(col_old[4]) * (1.0 - Double.valueOf(col_old[5]));
                            rawData_old.setRealPrice(realPrice.toString());
                        } else {
                            rawData_old.setRealPrice(col_old[6]);
                            rawData_old.setDiscode(col_old[5]);
                            realPrice = Double.valueOf(col_old[6]) / (1.0 - Double.valueOf(col_old[5]));
                            rawData_old.setPrice(realPrice.toString());
                        }

                        rawData_old.setCountDownPercent(col_old[7]);
                        countDown = Double.valueOf(col_old[7]) * realPrice;
                        rawData_old.setCountDown(countDown.toString());
                        realPrice_2 = realPrice + countDown;
                        rawData_old.setRealPrice_2(realPrice_2.toString());
                        rawData_old.setDesc(col_old[10]);
                        rawData_old.setWayToinstall(col_old[11]);
                        rawData_old.setBoardId(col_old[12]);
                        rawData_old.setOthers(col_old[13]);
                        rawData_old.setRawNum(f);
                        commonBoardId.add(rawData_old.getBoardId());
                        new_BoardIds.put(rawData_old.getBoardId(), rawData_old);
                    }
                }
            }

            if (((String[]) excelContentsList_old.get(5))[12].equals("物料编码")) {
                for (f = 7; f < excelContentsList_old.size(); ++f) {
                    col_old = (String[]) excelContentsList_old.get(f);
                    if (col_old.length >= 12 && col_old[12] != "") {
                        rawData_old = new RawData();
                        rawData_old.setIndex(col_old[0]);
                        rawData_old.setName(col_old[1]);
                        rawData_old.setBoard_code(col_old[2]);
                        rawData_old.setUnit(col_old[3]);
                        realPrice = 0.0;
                        if (col_old[4] != "未知类型") {
                            rawData_old.setPrice(col_old[4]);
                            rawData_old.setDiscode(col_old[5]);
                            realPrice = Double.valueOf(col_old[4]) * (1.0 - Double.valueOf(col_old[5]));
                            rawData_old.setRealPrice(realPrice.toString());
                        } else {
                            rawData_old.setRealPrice(col_old[6]);
                            rawData_old.setDiscode(col_old[5]);
                            realPrice = Double.valueOf(col_old[6]) / (1.0 - Double.valueOf(col_old[5]));
                            rawData_old.setPrice(realPrice.toString());
                        }

                        rawData_old.setCountDownPercent(col_old[7]);
                        countDown = Double.valueOf(col_old[7]) * realPrice;
                        rawData_old.setCountDown(countDown.toString());
                        realPrice_2 = realPrice + countDown;
                        rawData_old.setRealPrice_2(realPrice_2.toString());
                        rawData_old.setDesc(col_old[10]);
                        rawData_old.setWayToinstall(col_old[11]);
                        rawData_old.setBoardId(col_old[12]);
                        rawData_old.setOthers(col_old[13]);
                        rawData_old.setRawNum(f);
                        old_BoardIds.put(rawData_old.getBoardId(), rawData_old);
                    }
                }
            } else if (excelContentsList_old.get(5).length >= 15 && ((String[]) excelContentsList_old.get(5))[15].equals("物料编码")) {
                for (f = 7; f < excelContentsList_old.size(); ++f) {
                    col_old = excelContentsList_old.get(f);
                    if (col_old.length >= 15 && col_old[15] != "") {
                        rawData_old = new RawData();
                        rawData_old.setIndex(col_old[0]);
                        rawData_old.setName(col_old[1]);
                        rawData_old.setBoard_code(col_old[2]);
                        rawData_old.setUnit(col_old[3]);
                        realPrice = 0.0;
                        if (col_old[4] != "未知类型") {
                            rawData_old.setPrice(col_old[4]);
                            rawData_old.setDiscode(col_old[5]);
                            realPrice = Double.valueOf(col_old[4]) * (1.0 - Double.valueOf(col_old[5]));
                            rawData_old.setRealPrice(realPrice.toString());
                        } else {
                            rawData_old.setRealPrice(col_old[6]);
                            rawData_old.setDiscode(col_old[5]);
                            realPrice = Double.valueOf(col_old[6]) / (1.0 - Double.valueOf(col_old[5]));
                            rawData_old.setPrice(realPrice.toString());
                        }

                        rawData_old.setCountDownPercent(col_old[7]);
                        countDown = Double.valueOf(col_old[7]) * realPrice;
                        rawData_old.setCountDown(countDown.toString());
                        realPrice_2 = realPrice + countDown;
                        rawData_old.setRealPrice_2(realPrice_2.toString());
                        rawData_old.setDesc(col_old[13]);
                        rawData_old.setWayToinstall(col_old[14]);
                        rawData_old.setBoardId(col_old[15]);
                        rawData_old.setOthers(col_old[16]);
                        rawData_old.setRawNum(f);
                        old_BoardIds.put(rawData_old.getBoardId(), rawData_old);
                    }
                }
            } else {
                for (f = 7; f < excelContentsList_old.size(); ++f) {
                    col_old = (String[]) excelContentsList_old.get(f);
                    if (col_old.length >= 16 && col_old[16] != "") {
                        rawData_old = new RawData();
                        rawData_old.setIndex(col_old[0]);
                        rawData_old.setName(col_old[1]);
                        rawData_old.setBoard_code(col_old[2]);
                        rawData_old.setUnit(col_old[3]);
                        realPrice = 0.0;
                        if (col_old[4] != "未知类型") {
                            rawData_old.setPrice(col_old[4]);
                            rawData_old.setDiscode(col_old[5]);
                            realPrice = Double.valueOf(col_old[4]) * (1.0 - Double.valueOf(col_old[5]));
                            rawData_old.setRealPrice(realPrice.toString());
                        } else {
                            rawData_old.setRealPrice(col_old[6]);
                            rawData_old.setDiscode(col_old[5]);
                            realPrice = Double.valueOf(col_old[6]) / (1.0 - Double.valueOf(col_old[5]));
                            rawData_old.setPrice(realPrice.toString());
                        }

                        rawData_old.setCountDownPercent(col_old[7]);
                        countDown = Double.valueOf(col_old[7]) * realPrice;
                        rawData_old.setCountDown(countDown.toString());
                        realPrice_2 = realPrice + countDown;
                        rawData_old.setRealPrice_2(realPrice_2.toString());
                        rawData_old.setDesc(col_old[14]);
                        rawData_old.setWayToinstall(col_old[15]);
                        rawData_old.setBoardId(col_old[16]);
                        rawData_old.setOthers(col_old[17]);
                        rawData_old.setRawNum(f);
                        old_BoardIds.put(rawData_old.getBoardId(), rawData_old);
                    }
                }
            }

            commonBoardId.retainAll(old_BoardIds.keySet());
            Iterator var49 = new_BoardIds.keySet().iterator();

            while (true) {
                String boardId;
                while (var49.hasNext()) {
                    boardId = (String) var49.next();
                    if (!commonBoardId.contains(boardId)) {
                        rawData_old = (RawData) new_BoardIds.get(boardId);
                        greenRawIndex.add(rawData_old.getRawNum());
                    } else {
                        rawData_old = (RawData) old_BoardIds.get(boardId);
                        RawData new_rawData = (RawData) new_BoardIds.get(boardId);
                        List<Integer> colNum = checkRawData(rawData_old, new_rawData);
                        if (colNum.isEmpty()) {
                            yellowRawIndex.add(new_rawData.getRawNum());
                        } else {
                            Iterator var56 = colNum.iterator();

                            while (var56.hasNext()) {
                                Integer num = (Integer) var56.next();
                                if (redRawIndex.containsKey(new_rawData.getRawNum())) {
                                    ((List) redRawIndex.get(new_rawData.getRawNum())).add(num);
                                } else {
                                    List<Integer> cellNum = new ArrayList();
                                    cellNum.add(num);
                                    redRawIndex.put(new_rawData.getRawNum(), cellNum);
                                }
                            }
                        }
                    }
                }

                var49 = old_BoardIds.keySet().iterator();

                while (var49.hasNext()) {
                    boardId = (String) var49.next();
                    if (!commonBoardId.contains(boardId)) {
                        rawData_old = (RawData) old_BoardIds.get(boardId);
                        grayRawIndex.add(rawData_old.getRawNum());
                    }
                }

                log.info("解析完成，开始给" + excelName + "上色");
                var49 = greenRawIndex.iterator();

                Integer red;
                Row row;
                int i;
                while (var49.hasNext()) {
                    red = (Integer) var49.next();
                    row = sheet_new.getRow(red + 1);

                    for (i = 0; i < 13; ++i) {
                        if (row.getCell(i) != null) {
                            row.getCell(i).setCellStyle(colorCellStyleContentG);
                        }
                    }
                }

                var49 = yellowRawIndex.iterator();

                while (var49.hasNext()) {
                    red = (Integer) var49.next();
                    row = sheet_new.getRow(red + 1);

                    for (i = 0; i < 13; ++i) {
                        if (row.getCell(i) != null) {
                            row.getCell(i).setCellStyle(colorCellStyleContentY);
                        }
                    }
                }

                var49 = grayRawIndex.iterator();

                while (var49.hasNext()) {
                    red = (Integer) var49.next();
                    row = sheet_old.getRow(red + 1);

                    for (i = 0; i < 17; ++i) {
                        if (row.getCell(i) != null) {
                            row.getCell(i).setCellStyle(colorCellStyleContentGr);
                        }
                    }
                }

                var49 = redRawIndex.keySet().iterator();

                while (var49.hasNext()) {
                    red = (Integer) var49.next();
                    row = sheet_new.getRow(red + 1);
                    Iterator var59 = ((List) redRawIndex.get(red)).iterator();

                    while (var59.hasNext()) {
                        Integer i2 = (Integer) var59.next();
                        if (i2 != 6 && i2 != 8) {
                            row.getCell(i2).setCellStyle(colorCellStyleContentR);
                        }
                    }
                }

                for (f = 0; f < excelContentsList_old.size(); ++f) {
                    Row row2 = sheet_old.getRow(f);

                    for (int del = row2.getLastCellNum() - 1; del > 30; --del) {
                        row2.removeCell(row2.getCell(del));
                    }
                }
                break;
            }
        }

        log.info("所有表单上色完毕，正在导出文件");

        // 假设已经生成了两个Excel文件的数据字节数组
        byte[] excelFile1Bytes = getByteArrayFromWorkbook(workbook_new);
        log.info("导出新表单完成");
        byte[] excelFile2Bytes = getByteArrayFromWorkbook(workbook_old);
        log.info("导出旧表单完成");
        // 创建一个ZipOutputStream用于打包
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ZipOutputStream zos = new ZipOutputStream(baos);

        // 添加每个Excel文件到zip中
        ZipEntry entry1 = new ZipEntry(DateUtils.dateTimeNow()+"_标注后_" + excelDetail.getNewExcelName());
        zos.putNextEntry(entry1);
        zos.write(excelFile1Bytes);
        zos.closeEntry();

        ZipEntry entry2 = new ZipEntry(DateUtils.dateTimeNow()+"_标注后_" + excelDetail.getOldExcelName());
        zos.putNextEntry(entry2);
        zos.write(excelFile2Bytes);
        zos.closeEntry();

        // 关闭ZipOutputStream
        zos.close();

        String fileName = "excel标色_" + DateUtils.dateTimeNow() + ".zip";
        globalHashMapHolder.getGlobalHashMap().put(fileName, baos.toByteArray());

        // 移除缓存
        globalHashMapHolder.getGlobalHashMap().remove(excelDetail.getWorkbookNewKey());
        globalHashMapHolder.getGlobalHashMap().remove(excelDetail.getWorkbookOldKey());

        AjaxResult ajax = AjaxResult.success();
        ajax.put("fileName", fileName);
        return ajax;
    }

    public byte[] getByteArrayFromWorkbook(Workbook workbook) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try {
            // 对于XLSX格式的文件
            if (workbook instanceof XSSFWorkbook) {
                workbook.write(baos);
            } else if (workbook instanceof HSSFWorkbook) {
                // 对于老版本的HSSF（.xls格式），也适用类似的过程
                workbook.write(baos);
            }
        } finally {
            workbook.close(); // 记得关闭Workbook以释放资源
        }

        return baos.toByteArray();
    }

    public static List<String[]> readExcelContent(Sheet sheet) {
        List<String[]> list = new ArrayList();
        Row firstRow = sheet.getRow(getFirstRowNum(sheet));
        int firstCellNum = firstRow.getFirstCellNum();
        int lastCellNum = firstRow.getPhysicalNumberOfCells();

        for (int rowNum = getFirstRowNum(sheet) + 1; rowNum <= getLastRowNum(sheet); ++rowNum) {
            Row row = sheet.getRow(rowNum);
            if (null == row) {
                break;
            }

            String[] cells = new String[lastCellNum];
            boolean isEmptyRow = true;

            for (int cellNum = firstCellNum; cellNum < lastCellNum; ++cellNum) {
                if (cellNum < 0) {
                    ++cellNum;
                } else {
                    Cell cell = row.getCell(cellNum);
                    cells[cellNum] = ValideDataUti.getCellValue(cell);
                    if (cells[cellNum] != null) {
                        isEmptyRow = false;
                    }
                }
            }

            if (isEmptyRow) {
                break;
            }

            list.add(cells);
        }

        return list;
    }

    private static int getFirstRowNum(Sheet sheet) {
        return sheet.getFirstRowNum();
    }

    private static int getLastRowNum(Sheet sheet) {
        return sheet.getLastRowNum();
    }

    private Workbook getInput(MultipartFile file) throws IOException {
        Workbook workbook = null;
        String fileName = file.getOriginalFilename();
        if (fileName.endsWith(".xls")) {
            workbook = new HSSFWorkbook(file.getInputStream());
        } else if (fileName.toLowerCase().endsWith(".xlsx")) {
            workbook = new XSSFWorkbook(file.getInputStream());
        } else {
            log.error("路径错误，请在输入中填写合法的输入文件路径及名称，如：D:\\输入文件.xls");
        }

        log.info("读取"+fileName+"完成");

        return (Workbook) workbook;
    }

    private static String getOutputPath(String path) {
        String edName = path;
        String temp1 = edName.substring(edName.indexOf(".xls"));
        String temp2 = edName.substring(0, edName.indexOf(".xls"));
        Date date = new Date();
        DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
        temp2 = temp2 + "_标注后_" + dateFormat.format(date);
        String f = temp2 + temp1;
        return f;
    }

    private static List<Integer> checkRawData(RawData rawData1, RawData rawData2) {
        List<Integer> difColIndex = new ArrayList();
        if (!rawData1.getIndex().trim().equals(rawData2.getIndex().trim())) {
            difColIndex.add(0);
        }

        if (!rawData1.getName().trim().equals(rawData2.getName().trim())) {
            difColIndex.add(1);
        }

        if (!rawData1.getBoard_code().trim().equals(rawData2.getBoard_code().trim())) {
            difColIndex.add(2);
        }

        if (!rawData1.getUnit().trim().equals(rawData2.getUnit().trim())) {
            difColIndex.add(3);
        }

        if (!rawData1.getPrice().trim().equals(rawData2.getPrice().trim())) {
            difColIndex.add(4);
        }

        if (!rawData1.getDiscode().trim().equals(rawData2.getDiscode().trim())) {
            difColIndex.add(5);
        }

        if (!rawData1.getRealPrice().trim().equals(rawData2.getRealPrice().trim())) {
            difColIndex.add(6);
        }

        if (!rawData1.getCountDownPercent().trim().equals(rawData2.getCountDownPercent().trim())) {
            difColIndex.add(7);
        }

        if (!rawData1.getCountDown().trim().equals(rawData2.getCountDown().trim())) {
            difColIndex.add(8);
        }

        if (!rawData1.getRealPrice_2().trim().equals(rawData2.getRealPrice_2().trim())) {
            difColIndex.add(9);
        }

        if (!rawData1.getDesc().trim().equals(rawData2.getDesc().trim())) {
            difColIndex.add(10);
        }

        if (!rawData1.getWayToinstall().trim().equals(rawData2.getWayToinstall().trim())) {
            difColIndex.add(11);
        }

        return difColIndex;
    }
}
