package com.ek.project.common.dto;

import java.util.List;

public class ExcelDetail {
    private String workbookOldKey;
    private String workbookNewKey;
    private String oldExcelName;
    private String newExcelName;
    private List<String> selectSheet;

    public String getWorkbookOldKey() {
        return workbookOldKey;
    }

    public void setWorkbookOldKey(String workbookOldKey) {
        this.workbookOldKey = workbookOldKey;
    }

    public String getWorkbookNewKey() {
        return workbookNewKey;
    }

    public void setWorkbookNewKey(String workbookNewKey) {
        this.workbookNewKey = workbookNewKey;
    }

    public List<String> getSelectSheet() {
        return selectSheet;
    }

    public void setSelectSheet(List<String> selectSheet) {
        this.selectSheet = selectSheet;
    }

    public String getOldExcelName() {
        return oldExcelName;
    }

    public void setOldExcelName(String oldExcelName) {
        this.oldExcelName = oldExcelName;
    }

    public String getNewExcelName() {
        return newExcelName;
    }

    public void setNewExcelName(String newExcelName) {
        this.newExcelName = newExcelName;
    }
}
