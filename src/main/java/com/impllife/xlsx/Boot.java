package com.impllife.xlsx;

import com.impllife.xlsx.service.ExcelService;
import com.impllife.xlsx.service.ExcelServiceImpl;

public class Boot {
    private static final ExcelService excelService = new ExcelServiceImpl();
    public static void main(String[] args) {
        excelService.createStat();
    }
}
