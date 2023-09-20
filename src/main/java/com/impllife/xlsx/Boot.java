package com.impllife.xlsx;

import com.impllife.xlsx.data.Stat;
import com.impllife.xlsx.data.Transaction;
import com.impllife.xlsx.service.ExcelService;
import com.impllife.xlsx.service.ExcelServiceImpl;

import java.io.File;
import java.math.BigDecimal;
import java.util.*;
import java.util.stream.Collectors;

public class Boot {
    private static final ExcelService excelService = new ExcelServiceImpl();
    public static void main(String[] args) {
        createStat();
    }

    private static void createStat() {
        String fileName = "data/stat.xlsx";
        File resFile = new File(fileName);
        if (resFile.exists()) resFile.delete();

        File dataFolder = new File("data");
        Set<Transaction> set = new HashSet<>();
        for (File file : dataFolder.listFiles()) {
            set.addAll(excelService.readData(file.getAbsolutePath()));
        }
        List<Transaction> transactions = set.stream().sorted(Comparator.comparing(Transaction::getFullDate)).collect(Collectors.toList());

        List<Stat> stats = getStat(fillEmptyDays(transactions));
        excelService.createSheetStat(fileName, stats);
        excelService.addSheetTrn(fileName, transactions);
    }

    private final static List<String> ignore = new ArrayList<>();
    static {
        ignore.add("Зі своєї картки 51**22");
    }
    public static List<Stat> getStat(List<Transaction> transactions) {
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

    private static List<Transaction> fillEmptyDays(List<Transaction> transactions) {
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
        transactions = transactions.stream().sorted(Comparator.comparing(Transaction::getFullDate)).collect(Collectors.toList());
        return transactions;
    }
}
