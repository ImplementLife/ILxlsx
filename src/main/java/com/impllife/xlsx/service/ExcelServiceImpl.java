package com.impllife.xlsx.service;

import com.impllife.xlsx.data.Stat;
import com.impllife.xlsx.data.Transaction;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.MathContext;
import java.math.RoundingMode;
import java.util.*;
import java.util.function.BiConsumer;

import static com.impllife.xlsx.service.Util.concatDateAndTime;

public class ExcelServiceImpl implements ExcelService {
    private enum ColumnDefinition {
        DATE            (0,"Date",          (c, t) -> t.setDate(Util.parseDateByPattern(c.getStringCellValue(), "dd.MM.yyyy"))),
        TIME            (1,"Time",          (c, t) -> t.setTime(Util.parseDateByPattern(c.getStringCellValue(), "HH:mm"))),
        CATEGORY        (2,"Category",      (c, t) -> t.setCategory(c.getStringCellValue())),
        DESCRIPTION     (4,"Description",   (c, t) -> t.setDscr(c.getStringCellValue())),
        SUM             (5,"Sum",           (c, t) -> t.setSum(BigDecimal.valueOf(c.getNumericCellValue()).setScale(2, RoundingMode.CEILING))),
        ;
        private final Integer index;
        private final String name;
        private final BiConsumer<Cell, Transaction> consumer;

        public void fillValue(Cell cell, Transaction transaction) {
            consumer.accept(cell, transaction);
        }

        ColumnDefinition(int index, String name, BiConsumer<Cell, Transaction> consumer) {
            this.index = index;
            this.name = name;
            this.consumer = consumer;
        }
        private static ColumnDefinition getInstance(String name) {
            for (ColumnDefinition value : values()) {
                if (value.name.equals(name)) return value;
            }
            return null;
        }
        private static ColumnDefinition getInstance(Integer index) {
            for (ColumnDefinition value : values()) {
                if (value.index.equals(index)) return value;
            }
            return null;
        }
    }

    @Override
    public List<Transaction> readData(String fileName) {
        List<Transaction> result = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(fileName)) {
            Workbook workbook = getWorkbook(fileName, fis);

            Sheet sheet = workbook.getSheetAt(0);
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            trnFill: for (Row row : sheet) {
                try {
                    Transaction transaction = new Transaction();
                    for (Cell cell : row) {
                        if (isMergedCell(mergedRegions, cell)) {
                            continue trnFill;
                        }

                        ColumnDefinition definition = ColumnDefinition.getInstance(cell.getColumnIndex());
                        if (definition != null) {
                            definition.fillValue(cell, transaction);
                        }
                    }
                    transaction.setFullDate(concatDateAndTime(transaction.getDate(), transaction.getTime()));
                    result.add(transaction);
                } catch (Throwable t) {
                    //not valid row
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return result;
    }

    @Override
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

    @Override
    public void createSheetTrn(String fileName, List<Transaction> transactions) {
        Workbook workbook = getWorkbook(fileName);
        putData(workbook, transactions);
        write(fileName, workbook);
    }

    @Override
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

    public void createSheetStat(String fileName, List<Stat> stats) {
        Workbook workbook = getWorkbook(fileName);
        Sheet sheet = workbook.createSheet("Stat");

        CellStyle fullDateCellStyle = getTimeFormat(workbook, "dd.MM.yyyy");
        fullDateCellStyle.setAlignment(HorizontalAlignment.CENTER);

        int rowIndex = 0;
        int colIndex = 0;
        Row row = sheet.createRow(rowIndex++);
        row.createCell(colIndex++).setCellValue("Date");
        row.createCell(colIndex++).setCellValue("Sum");
        for (Stat stat : stats) {
            row = sheet.createRow(rowIndex++);
            colIndex = 0;

            Cell dateCell = row.createCell(colIndex++);
            dateCell.setCellValue(stat.getDate());
            dateCell.setCellStyle(fullDateCellStyle);
            row.createCell(colIndex++).setCellValue(stat.getSum().doubleValue());
        }


        sheet.setColumnWidth(0, (10+1)*256);

        write(fileName, workbook);
    }


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

    private void putData(Workbook workbook, List<Transaction> transactions) {
        Sheet sheet = workbook.createSheet("Trn");

        CellStyle fullDateCellStyle = getTimeFormat(workbook, "dd.MM.yyyy HH:mm");
        fullDateCellStyle.setAlignment(HorizontalAlignment.CENTER);


        int rowIndex = 0;
        int colIndex = 0;
        Row row = sheet.createRow(rowIndex++);
        row.createCell(colIndex++).setCellValue("Повна дата");
//        row.createCell(colIndex++).setCellValue("Дата");
//        row.createCell(colIndex++).setCellValue("Час");
        row.createCell(colIndex++).setCellValue("Категорія");
        row.createCell(colIndex++).setCellValue("Опис");
        row.createCell(colIndex++).setCellValue("Сума");

        for (Transaction transaction : transactions) {
            row = sheet.createRow(rowIndex++);
            colIndex = 0;

            Cell dateCell = row.createCell(colIndex++);
            dateCell.setCellValue(transaction.getFullDate());
            dateCell.setCellStyle(fullDateCellStyle);
//            row.createCell(colIndex++).setCellValue(transaction.getDate());
//            row.createCell(colIndex++).setCellValue(transaction.getTime());
            row.createCell(colIndex++).setCellValue(transaction.getCategory());
            row.createCell(colIndex++).setCellValue(transaction.getDscr());
            row.createCell(colIndex++).setCellValue(transaction.getSum().doubleValue());
        }

        sheet.setColumnWidth(0, (16+1)*256);
        sheet.setColumnWidth(1, (30+1)*256);
        sheet.setColumnWidth(2, (40+1)*256);
        sheet.setColumnWidth(3, (10+1)*256);
        sheet.setAutoFilter(new CellRangeAddress(0, rowIndex-1,0, colIndex-1));
    }

    private void putDataForAllDays(Workbook workbook, List<Transaction> transactions) {
        Sheet sheet = workbook.createSheet("Trn All Days");

        CellStyle fullDateCellStyle = getTimeFormat(workbook, "dd.MM.yyyy HH:mm");
        fullDateCellStyle.setAlignment(HorizontalAlignment.CENTER);

        int rowIndex = 0;
        int colIndex = 0;
        Row row = sheet.createRow(rowIndex++);
        row.createCell(colIndex++).setCellValue("Повна дата");
//        row.createCell(colIndex++).setCellValue("Дата");
//        row.createCell(colIndex++).setCellValue("Час");
        row.createCell(colIndex++).setCellValue("Категорія");
        row.createCell(colIndex++).setCellValue("Опис");
        row.createCell(colIndex++).setCellValue("Сума");

        for (Transaction transaction : transactions) {
            row = sheet.createRow(rowIndex++);
            colIndex = 0;

            Cell dateCell = row.createCell(colIndex++);
            dateCell.setCellValue(transaction.getFullDate());
            dateCell.setCellStyle(fullDateCellStyle);
//            row.createCell(colIndex++).setCellValue(transaction.getDate());
//            row.createCell(colIndex++).setCellValue(transaction.getTime());
            row.createCell(colIndex++).setCellValue(transaction.getCategory());
            row.createCell(colIndex++).setCellValue(transaction.getDscr());
            row.createCell(colIndex++).setCellValue(transaction.getSum().doubleValue());
        }

        sheet.setColumnWidth(0, (16+1)*256);
        sheet.setColumnWidth(1, (30+1)*256);
        sheet.setColumnWidth(2, (40+1)*256);
        sheet.setColumnWidth(3, (10+1)*256);
        sheet.setAutoFilter(new CellRangeAddress(0, rowIndex-1,0, colIndex-1));
    }

    private CellStyle getTimeFormat(Workbook workbook, String pattern) {
        CellStyle fullDateCellStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        short format = createHelper.createDataFormat().getFormat(pattern);
        fullDateCellStyle.setDataFormat(format);
        return fullDateCellStyle;
    }

    private void write(String fileName, Workbook workbook) {
        try (FileOutputStream fos = new FileOutputStream(fileName)) {
            workbook.write(fos);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
