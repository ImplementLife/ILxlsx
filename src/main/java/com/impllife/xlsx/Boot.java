package com.impllife.xlsx;

import com.impllife.xlsx.service.ExcelService;

public class Boot {
    private static final ExcelService excelService = new ExcelService();
    public static void main(String[] args) {
        excelService.createByTemplateStat();
    }
}
