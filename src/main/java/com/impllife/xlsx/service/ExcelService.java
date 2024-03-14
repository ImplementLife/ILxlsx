package com.impllife.xlsx.service;

import com.impllife.xlsx.Const;
import com.impllife.xlsx.data.StatByString;
import com.impllife.xlsx.data.Transaction;
import com.impllife.xlsx.data.map.ColumnDefinition;
import com.impllife.xlsx.service.map.JSONLoader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

import static com.impllife.xlsx.service.util.DateUtil.concatDateAndTime;
import static com.impllife.xlsx.service.util.DateUtil.getCurrentDateTime;
import static com.impllife.xlsx.service.util.WorkbookUtil.*;
import static org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK;

public class ExcelService {
    private final Map<String, List<ColumnDefinition>> columnDefinitionsMap = new HashMap<>();

    public ExcelService() {
        readMappings();
    }

    public void readMappings() {
        columnDefinitionsMap.clear();
        columnDefinitionsMap.put("mono", JSONLoader.loadColumnDefinitions(Const.MAP_DIR + "mono_input_excel_mappings.json"));
        columnDefinitionsMap.put("p24", JSONLoader.loadColumnDefinitions(Const.MAP_DIR + "p24_input_excel_mappings.json"));
    }

    public List<Transaction> readData(String fileName, String bank, String tags) {
        List<Transaction> result = new ArrayList<>();
        Workbook workbook = read(fileName);
        Sheet sheet = workbook.getSheetAt(0);
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        List<ColumnDefinition> columnDefinitions = columnDefinitionsMap.get(bank);


        trnFill: for (Row row : sheet) {
            try {
                Transaction transaction = new Transaction();
                for (ColumnDefinition definition : columnDefinitions) {
                    Cell cell = row.getCell(definition.getIndex());
                    if (isMergedCell(mergedRegions, cell)) continue trnFill;

                    String setter = definition.getSetter();
                    if (setter.equals("setDate"))          transaction.setDate(definition.convert(cell));
                    else if (setter.equals("setTime"))     transaction.setTime(definition.convert(cell));
                    else if (setter.equals("setDateTime")) transaction.setFullDate(definition.convert(cell));
                    else if (setter.equals("setCategory")) transaction.setCategory(definition.convert(cell));
                    else if (setter.equals("setDscr"))     transaction.setDscr(definition.convert(cell));
                    else if (setter.equals("setSum"))      transaction.setSum(definition.convert(cell));
                }
                if (transaction.getFullDate() == null) {
                    transaction.setFullDate(concatDateAndTime(transaction.getDate(), transaction.getTime()));
                }
                transaction.setTags(tags);
                result.add(transaction);
            } catch (Throwable t) { /*not valid row*/ }
        }
        return result;
    }

    public void createByTemplateStat() {
        Workbook template = read(Const.TEMPLATES_DIR + "template.xlsx");
        String resultFileName = Const.RESULT_DIR + "stat_" + getCurrentDateTime() + ".xlsx";
        File resFile = new File(resultFileName);
        if (resFile.exists()) resFile.delete();

        List<Transaction> transactions = readAndSortFilesTransactions();
        Map<String, List<Transaction>> groupByMonth = groupByMonth(fillEmptyDays(transactions));

        SimpleDateFormat dateFormat = new SimpleDateFormat("MM.yyyy");
        List<Map.Entry<String, List<Transaction>>> sortedGroupByMonth = groupByMonth.entrySet().stream().sorted((e1, e2) -> {
            try {
                Date date1 = dateFormat.parse(e1.getKey());
                Date date2 = dateFormat.parse(e2.getKey());
                return date2.compareTo(date1);
            } catch (ParseException ex) {
                ex.printStackTrace();
                return 0;
            }
        }).toList();

        for (Map.Entry<String, List<Transaction>> entry : sortedGroupByMonth) {
            Sheet monthStat = cloneSheet(template, "T2", entry.getKey());
            List<Transaction> transactionList = entry.getValue();
            List<StatByString> monthStatistic = transactionList.stream()
                .collect(Collectors.groupingBy(Transaction::getDscr)).entrySet().stream()
                .map(e -> {
                    StatByString stat = new StatByString();
                    stat.setStr(e.getKey());
                    BigDecimal sum = e.getValue().stream()
                        .map(Transaction::getSum)
                        .reduce(BigDecimal.ZERO, BigDecimal::add);
                    stat.setSum(sum);
                    return stat;
                })
                .sorted(Comparator.comparing(StatByString::getSum))
                .toList();

            putMonthStat(monthStat, transactionList, monthStatistic);
        }


        removeSheet(template, "T1");
        removeSheet(template, "T2");
        removeSheet(template, "T3");
        write(resultFileName, template);
    }

    private void putMonthStat(Sheet sheet, List<Transaction> trnList, List<StatByString> monthStatistic) {
        int rowIndex = 0;
        int colIndex = -1;
        for (Transaction trn : trnList) {
            Row row = sheet.getRow(++rowIndex);
            colIndex = -1;

            row.getCell(++colIndex, CREATE_NULL_AS_BLANK).setCellValue(trn.getFullDate());
            row.getCell(++colIndex, CREATE_NULL_AS_BLANK).setCellValue(trn.getCategory());
            row.getCell(++colIndex, CREATE_NULL_AS_BLANK).setCellValue(trn.getDscr());
            row.getCell(++colIndex, CREATE_NULL_AS_BLANK).setCellValue(trn.getSum().doubleValue());
            row.getCell(++colIndex, CREATE_NULL_AS_BLANK).setCellValue("");
            row.getCell(++colIndex, CREATE_NULL_AS_BLANK).setCellValue(trn.getTags());
        }

        rowIndex = 0;
        colIndex += 1;
        for (StatByString statSrt : monthStatistic) {
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

    private List<Transaction> readAndSortFilesTransactions() {
        File dataFolder = new File(Const.INPUT_DATA_DIR);
        Set<Transaction> set = new HashSet<>();
        readAndSortFilesTransactions(dataFolder, set);
        return set.stream().sorted(Comparator.comparing(Transaction::getFullDate)).collect(Collectors.toList());
    }

    private void readAndSortFilesTransactions(File dataFolder, Set<Transaction> set) {
        for (File file : dataFolder.listFiles()) {
            if (file.isFile()) {
                File personFold = file.getParentFile();
                File bankFold = personFold.getParentFile();

                // Get the directory names as strings
                String personFoldName = personFold != null ? personFold.getName() : "Unknown";
                String bankFoldName = bankFold != null ? bankFold.getName() : "Unknown";
                String tags = String.format("#%s #%s", personFoldName, bankFoldName);

                set.addAll(readData(file.getAbsolutePath(), bankFoldName, tags));
            } else if (file.isDirectory()) {
                readAndSortFilesTransactions(file, set);
            }
        }
    }

    private List<Transaction> fillEmptyDays(List<Transaction> transactions) {
        Set<Date> dates = transactions.stream().map(Transaction::getDate).collect(Collectors.toSet());
        Date firstTrn = transactions.get(0).getDate();
        Date lastTrn = transactions.get(transactions.size()-1).getDate();
        Date tempDate = firstTrn;
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(firstTrn);
        while (tempDate.before(lastTrn)) {
            calendar.add(Calendar.DAY_OF_MONTH, 1);
            tempDate = calendar.getTime();
            if (!dates.contains(tempDate)) {
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
}
