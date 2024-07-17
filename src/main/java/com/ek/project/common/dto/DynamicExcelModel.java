package com.ek.project.common.dto;

import java.util.Map;

public class DynamicExcelModel {
    private Map<String, String> rowData;

    public Map<String, String> getRowData() {
        return rowData;
    }

    public void setRowData(Map<String, String> rowData) {
        this.rowData = rowData;
    }
}
