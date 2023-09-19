package com.impllife.xlsx;

import com.impllife.xlsx.data.Transaction;
import com.impllife.xlsx.service.ExcelService;

import java.io.File;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

public class Boot {
    public static void main(String[] args) {
        ExcelService excelService = new ExcelService();
        String fileName = "data/stat.xlsx";

        File resFile = new File(fileName);
        if (resFile.exists()) resFile.delete();

        File dataFolder = new File("data");
        Set<Transaction> set = new HashSet<>();
        for (File file : dataFolder.listFiles()) {
            set.addAll(excelService.readExcelData(file.getAbsolutePath()));
        }

        excelService.createSheetTrn(fileName, new ArrayList<>(set));
    }
}
