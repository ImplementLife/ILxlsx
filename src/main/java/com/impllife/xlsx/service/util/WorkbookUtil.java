package com.impllife.xlsx.service.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

public final class WorkbookUtil {
    private WorkbookUtil() {}

    public static Workbook read(String filePath) {
        try (FileInputStream fis = new FileInputStream(filePath)) {
            return WorkbookFactory.create(fis);
        } catch (Exception e) {
            throw new IllegalStateException(e);
        }
    }

    public static Workbook create(String fileName) {
        if (fileName.toLowerCase().endsWith("xlsx")) {
            return new XSSFWorkbook();
        } else if (fileName.toLowerCase().endsWith("xls")) {
            return new HSSFWorkbook();
        } else {
            throw new IllegalArgumentException("File extension does not support.");
        }
    }

    public static Workbook readOrCreate(String fileName) {
        File file = new File(fileName);
        if (file.exists()) {
            return read(fileName);
        } else {
            return create(fileName);
        }
    }

    public static void processWorkbookFile(String fileName, Consumer<Workbook> con) {
        Workbook workbook = readOrCreate(fileName);
        con.accept(workbook);
        write(fileName, workbook);
    }

    public static void write(String fileName, Workbook workbook) {
        try (FileOutputStream fos = new FileOutputStream(fileName)) {
            workbook.write(fos);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void removeSheet(Workbook workbook, String sheetName) {
        int trnSheetIndex = workbook.getSheetIndex(sheetName);
        if (trnSheetIndex != -1) workbook.removeSheetAt(trnSheetIndex);
    }

    public static void removeSheet(String fileName, String sheetName) {
        Workbook workbook = read(fileName);

        int trnSheetIndex = workbook.getSheetIndex(sheetName);
        if (trnSheetIndex != -1) workbook.removeSheetAt(trnSheetIndex);

        write(fileName, workbook);
    }

    public static Map<String, Sheet> getSheets(Workbook workbook) {
        Map<String, Sheet> result = new HashMap<>();
        int numberOfSheets = workbook.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            result.put(sheet.getSheetName(), sheet);
        }
        return result;
    }

    public static boolean isMergedCell(List<CellRangeAddress> mergedRegions, Cell cell) {
        for (CellRangeAddress mergedRegion : mergedRegions) {
            if (mergedRegion.isInRange(cell)) {
                return true;
            }
        }
        return false;
    }

    public static Sheet cloneSheet(Workbook workbook, String origName, String newName) {
        Sheet cloneSheet = workbook.cloneSheet(workbook.getSheetIndex(origName));
        workbook.setSheetName(workbook.getSheetIndex(cloneSheet), newName);
        return cloneSheet;
    }
}
