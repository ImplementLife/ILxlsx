package com.impllife.xlsx.service;

import com.impllife.xlsx.Const;
import com.impllife.xlsx.data.Stat;
import com.impllife.xlsx.data.StatSrt;
import com.impllife.xlsx.data.Transaction;
import com.impllife.xlsx.service.util.DateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.function.BiConsumer;
import java.util.stream.Collectors;

import static com.impllife.xlsx.service.util.DateUtil.*;
import static com.impllife.xlsx.service.util.WorkbookUtil.*;
import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK;

public class ExcelServiceImpl implements ExcelService {
    private enum ColumnDefinition {
        DATE            (0,"Date",          (c, t) -> t.setDate(DateUtil.parseDateByPattern(c.getStringCellValue(), "dd.MM.yyyy"))),
        TIME            (1,"Time",          (c, t) -> t.setTime(DateUtil.parseDateByPattern(c.getStringCellValue(), "HH:mm"))),
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

    private final static List<String> ignore = new ArrayList<>();
    static {
        ignore.add("Зі своєї картки 51**22");
//        ignore.add("Чудновська Вікторія Леонідівна");
    }

    @Override
    public List<Transaction> readData(String fileName) {
        List<Transaction> result = new ArrayList<>();
        Workbook workbook = read(fileName);
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
        return result;
    }

    @Override
    public void createByTemplateStat() {
        Workbook template = read(Const.TEMPLATES_DIR + "template.xlsx");
        String fileName = Const.WORK_DIR + "tmpl_test.xlsx";
        File resFile = new File(fileName);
        if (resFile.exists()) resFile.delete();

        List<Transaction> transactions = readAndSortFilesTransactions();
        Map<String, List<Transaction>> groupByMonth = groupByMonth(fillEmptyDays(transactions));

//        Sheet monthStat = cloneSheet(template, "T1", "Months");
//        putStat(monthStat, monthStat);
//        Sheet allStat = cloneSheet(template, "T3", "Months");

        SimpleDateFormat dateFormat = new SimpleDateFormat("MM.yyyy");
        List<Map.Entry<String, List<Transaction>>> collect = groupByMonth.entrySet().stream().sorted((e1, e2) -> {
            try {
                Date date1 = dateFormat.parse(e1.getKey());
                Date date2 = dateFormat.parse(e2.getKey());
                return date2.compareTo(date1);
            } catch (ParseException ex) {
                ex.printStackTrace();
                return 0;
            }
        }).collect(Collectors.toList());

        for (Map.Entry<String, List<Transaction>> entry : collect) {
            Sheet monthStat = cloneSheet(template, "T2", entry.getKey());
            List<Transaction> transactionList = entry.getValue();
            List<StatSrt> monthStatistic = transactionList.stream()
                .collect(Collectors.groupingBy(Transaction::getDscr))
                .entrySet().stream()
                .map(e -> {
                    StatSrt stat = new StatSrt();
                    stat.setStr(e.getKey());
                    BigDecimal sum = e.getValue().stream().map(Transaction::getSum).reduce(BigDecimal.ZERO, BigDecimal::add);
                    stat.setSum(sum);
                    return stat;
                })
                .sorted(Comparator.comparing(StatSrt::getSum))
                .collect(Collectors.toList());

            putMonthStat(monthStat, transactionList, monthStatistic);
        }


        removeSheet(template, "T1");
        removeSheet(template, "T2");
        removeSheet(template, "T3");
        write(fileName, template);
    }

    private void putMonthStat(Sheet sheet, List<Transaction> trnList, List<StatSrt> monthStatistic) {
        int rowIndex = 0;
        int colIndex = -1;
        for (Transaction trn : trnList) {
            Row row = sheet.getRow(++rowIndex);
            colIndex = -1;

            row.getCell(++colIndex, CREATE_NULL_AS_BLANK).setCellValue(trn.getFullDate());
            row.getCell(++colIndex, CREATE_NULL_AS_BLANK).setCellValue(trn.getCategory());
            row.getCell(++colIndex, CREATE_NULL_AS_BLANK).setCellValue(trn.getDscr());
            row.getCell(++colIndex, CREATE_NULL_AS_BLANK).setCellValue(trn.getSum().doubleValue());
        }

        rowIndex = 0;
        colIndex += 2;
        for (StatSrt statSrt : monthStatistic) {
            Row row = sheet.getRow(++rowIndex);
            int colIndexStat = colIndex;

            row.getCell(++colIndexStat, CREATE_NULL_AS_BLANK).setCellValue(statSrt.getStr());
            row.getCell(++colIndexStat, CREATE_NULL_AS_BLANK).setCellValue(statSrt.getSum().doubleValue());
        }
    }

    private Map<String, List<Transaction>> groupByMonth(List<Transaction> transactions) {
        return transactions.stream()
            .collect(Collectors.groupingBy(transaction -> {
                Calendar calendar = Calendar.getInstance();
                calendar.setTime(transaction.getFullDate());
                int month = calendar.get(Calendar.MONTH) + 1;
                int year = calendar.get(Calendar.YEAR);
                return String.format("%02d.%d", month, year);
            }));
    }

    @Override
    public void createStat() {
        String fileName = Const.WORK_DIR + "stat.xlsx";
        {
            File resFile = new File(fileName);
            if (resFile.exists()) resFile.delete();
        }

        List<Transaction> transactions = readAndSortFilesTransactions();

        List<Stat> stat = getStat(fillEmptyDays(transactions));

        List<Stat> monthStat = getMonthStat(stat);
        addSheetMonthStat(fileName, monthStat);
        addSheetDaysStat(fileName, stat);
        addSheetTrn(fileName, transactions);
    }

    private List<Transaction> readAndSortFilesTransactions() {
        File dataFolder = new File(Const.WORK_DIR);
        Set<Transaction> set = new HashSet<>();
        for (File file : dataFolder.listFiles()) {
            if (file.isFile()) {
                set.addAll(readData(file.getAbsolutePath()));
            }
        }
        return set.stream().sorted(Comparator.comparing(Transaction::getFullDate)).collect(Collectors.toList());
    }

    private void addSheetMonthStat(String fileName, List<Stat> stats) {
        Workbook workbook = readOrCreate(fileName);
        putStat(workbook, "Статистика по місяцях", stats);
        write(fileName, workbook);
    }

    private void addSheetDaysStat(String fileName, List<Stat> stats) {
        Workbook workbook = readOrCreate(fileName);
        putStat(workbook, "Статистика по дням", stats);
        write(fileName, workbook);
    }

    private void addSheetTrn(String fileName, List<Transaction> transactions) {
        Workbook workbook = readOrCreate(fileName);

        removeSheet(workbook, "Виписка");
        putTrnData(workbook, transactions);

        write(fileName, workbook);
    }

    private void putStat(Sheet sheet, List<Stat> stats) {
        int rowIndex = -1;
        int colIndex = -1;
        Row row = sheet.getRow(++rowIndex);
        row.getCell(++colIndex).setCellValue("Date");
        row.getCell(++colIndex).setCellValue("Sum");
        for (Stat stat : stats) {
            row = sheet.getRow(++rowIndex);
            colIndex = -1;

            Cell dateCell = row.getCell(++colIndex);
            dateCell.setCellValue(stat.getDate());
            row.getCell(++colIndex).setCellValue(stat.getSum().doubleValue());
        }
    }

    private void putStat(Workbook workbook, String sheetName, List<Stat> stats) {
        Sheet sheet = workbook.createSheet(sheetName);

        CellStyle fullDateCellStyle = getTimeFormat(workbook, "dd.MM");
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
    }

    private void putTrnData(Workbook workbook, List<Transaction> transactions) {
        Sheet sheet = workbook.createSheet("Виписка");

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
            row.createCell(colIndex++).setCellFormula("MONTH(A" + rowIndex + ")");
        }

        sheet.setColumnHidden(4, true);
        sheet.setColumnWidth(0, (16+1)*256);
        sheet.setColumnWidth(1, (30+1)*256);
        sheet.setColumnWidth(2, (40+1)*256);
        sheet.setColumnWidth(3, (10+1)*256);
        sheet.setAutoFilter(new CellRangeAddress(0, rowIndex-1,0, colIndex-1));
    }

    private List<Stat> getMonthStat(List<Stat> stat) {
        List<Stat> result = new ArrayList<>();
        BigDecimal monthSum = BigDecimal.ZERO;
        Date month = stat.get(0).getDate();
        for (Stat s : stat) {
            if (!isSameMonth(month, s.getDate())) {
                result.add(new Stat(month, monthSum));
                monthSum = BigDecimal.ZERO;
                month = s.getDate();
            }
            monthSum = monthSum.add(s.getSum());
        }
        result.add(new Stat(month, monthSum));
        return result;
    }

    private List<Stat> getStat(List<Transaction> transactions) {
        List<Stat> result = new ArrayList<>();

        Map<Date, List<Transaction>> groupsByDays = transactions.stream().collect(Collectors.groupingBy(Transaction::getDate));
        for (Date date : groupsByDays.keySet()) {
            List<Transaction> day = groupsByDays.get(date);
            BigDecimal sum = BigDecimal.ZERO;
            for (Transaction trn : day) {
                if (!ignore.contains(trn.getDscr())) {
                    sum = sum.add(trn.getSum());
                }
            }
            Stat stat = new Stat();
            stat.setDate(date);
            stat.setSum(sum);
            result.add(stat);
        }
        return result.stream().sorted(Comparator.comparing(Stat::getDate)).collect(Collectors.toList());
    }

    private List<Transaction> fillEmptyDays(List<Transaction> transactions) {
        Set<Date> dates = transactions.stream().map(Transaction::getDate).collect(Collectors.toSet());
        Date firstTrn = transactions.get(0).getDate();
        Date lastTrn = transactions.get(transactions.size()-1).getDate();
        Date tempDate = firstTrn;
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(firstTrn);
        int emptyDays = 0;
        while (tempDate.before(lastTrn)) {
            calendar.add(Calendar.DAY_OF_MONTH, 1);
            tempDate = calendar.getTime();
            if (!dates.contains(tempDate)) {
                emptyDays++;
                Transaction emptyDayTrn = new Transaction();
                emptyDayTrn.setFullDate(tempDate);
                emptyDayTrn.setDate(tempDate);
                emptyDayTrn.setSum(BigDecimal.ZERO);
                emptyDayTrn.setDscr("Free day");
                emptyDayTrn.setCategory("Free day");
                transactions.add(emptyDayTrn);
            }
        }
        return transactions.stream().sorted(Comparator.comparing(Transaction::getFullDate)).collect(Collectors.toList());
    }

    private CellStyle getTimeFormat(Workbook workbook, String pattern) {
        CellStyle fullDateCellStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        short format = createHelper.createDataFormat().getFormat(pattern);
        fullDateCellStyle.setDataFormat(format);
        return fullDateCellStyle;
    }

}
