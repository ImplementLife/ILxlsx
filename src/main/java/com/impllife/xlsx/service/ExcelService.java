package com.impllife.xlsx.service;

import com.impllife.xlsx.data.Transaction;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelService {
    private Map<String, Sheet> getSheets(Workbook workbook) {
        Map<String, Sheet> result = new HashMap<>();
        int numberOfSheets = workbook.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            result.put(sheet.getSheetName(), sheet);
        }
        return result;
    }

    private Workbook getWorkbook(String fileName) {
        return getWorkbook(fileName, null);
    }

    private Workbook getWorkbook(String fileName, FileInputStream fis) {
        try {
            if (fileName.toLowerCase().endsWith("xlsx")) {
                if (fis != null) {
                    return new XSSFWorkbook(fis);
                } else {
                    return new XSSFWorkbook();
                }
            } else if (fileName.toLowerCase().endsWith("xls")) {
                if (fis != null) {
                    return new HSSFWorkbook(fis);
                } else {
                    return new HSSFWorkbook();
                }
            } else {
                throw new IllegalArgumentException("File extension not support.");
            }
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }
    }

    private boolean isMergedCell(List<CellRangeAddress> mergedRegions, Cell cell) {
        for (CellRangeAddress mergedRegion : mergedRegions) {
            if (mergedRegion.isInRange(cell)) {
                return true;
            }
        }
        return false;
    }

    private Map<String, Integer> predefinedHeaders = new HashMap<>();
    {
        predefinedHeaders.put("Дата", 0);
        predefinedHeaders.put("Час", 1);
        predefinedHeaders.put("Категорія", 3);
        predefinedHeaders.put("Опис операції", 4);
        predefinedHeaders.put("Сума в валюті картки", 5);
    }

    public List<Transaction> readExcelData(String fileName) {
        List<Transaction> result = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(fileName)) {
            Workbook workbook = getWorkbook(fileName, fis);

            Sheet sheet = workbook.getSheetAt(0);
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            trnFill: for (Row row : sheet) {
                Transaction transaction = new Transaction();
                for (Cell cell : row) {
                    if (isMergedCell(mergedRegions, cell)) {
                        continue trnFill;
                    }

                    int columnIndex = cell.getColumnIndex();
                    if (columnIndex == 0) {
                        transaction.setDate(cell.toString());
                    }
                    if (columnIndex == 1) {
                        transaction.setTime(cell.toString());
                    }
                    if (columnIndex == 2) {
                        transaction.setCategory(cell.toString());
                    }
                    if (columnIndex == 4) {
                        transaction.setDscr(cell.toString());
                    }
                    if (columnIndex == 5) {
                        transaction.setSum(cell.toString());
                    }
                }
                boolean valid = true;
                try {
                    new BigDecimal(transaction.getSum());
                } catch (Throwable e) {
                    valid = false;
                }

                if (valid) {
                    result.add(transaction);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return result;
    }

    public void removeSheet(String fileName, String sheetName) {
        Workbook workbook;
        try (FileInputStream fis = new FileInputStream(fileName)) {
            workbook = getWorkbook(fileName, fis);
            int trnSheetIndex = workbook.getSheetIndex("Trn");
            if (trnSheetIndex != -1) workbook.removeSheetAt(trnSheetIndex);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        write(fileName, workbook);
    }

    public void createSheetTrn(String fileName, List<Transaction> transactions) {
        Workbook workbook = getWorkbook(fileName);
        putData(workbook, transactions);
        write(fileName, workbook);
    }

    public void addSheetTrn(String fileName, List<Transaction> transactions) {
        Workbook workbook;
        try (FileInputStream fis = new FileInputStream(fileName)) {
            workbook = getWorkbook(fileName, fis);
            int trnSheetIndex = workbook.getSheetIndex("Trn");
            if (trnSheetIndex != -1) workbook.removeSheetAt(trnSheetIndex);
            putData(workbook, transactions);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        write(fileName, workbook);
    }

    private void putData(Workbook workbook, List<Transaction> transactions) {
        Sheet sheet = workbook.createSheet("Trn");


        CellStyle fullDateCellStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        short format = createHelper.createDataFormat().getFormat("dd.MM.yyyy HH:mm");
        fullDateCellStyle.setDataFormat(format);
        fullDateCellStyle.setAlignment(HorizontalAlignment.CENTER);


        int rowIndex = 0;
        int colIndex = 0;
        Row row = sheet.createRow(rowIndex++);
        row.createCell(colIndex++).setCellValue("Повна дата");
        row.createCell(colIndex++).setCellValue("Дата");
        row.createCell(colIndex++).setCellValue("Час");
        row.createCell(colIndex++).setCellValue("Категорія");
        row.createCell(colIndex++).setCellValue("Опис");
        row.createCell(colIndex++).setCellValue("Сума");
        for (Transaction transaction : transactions) {
            row = sheet.createRow(rowIndex++);
            colIndex = 0;

            Cell dateCell = row.createCell(colIndex++);
            row.createCell(colIndex++).setCellValue(transaction.getDate());
            row.createCell(colIndex++).setCellValue(transaction.getTime());
            row.createCell(colIndex++).setCellValue(transaction.getCategory());
            row.createCell(colIndex++).setCellValue(transaction.getDscr());
            row.createCell(colIndex++).setCellValue(transaction.getSum());
            dateCell.setCellFormula("B" + rowIndex + "+C" + rowIndex);
            dateCell.setCellStyle(fullDateCellStyle);

        }

        sheet.setColumnWidth(0, (16+1)*256);
        sheet.setAutoFilter(new CellRangeAddress(0, rowIndex-1,0, colIndex-1));
    }

    private void write(String fileName, Workbook workbook) {
        try (FileOutputStream fos = new FileOutputStream(fileName)) {
            workbook.write(fos);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
