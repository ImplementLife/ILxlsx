package com.impllife.xlsx.service;

import com.impllife.xlsx.data.Stat;
import com.impllife.xlsx.data.Transaction;

import java.util.List;

public interface ExcelService {
    List<Transaction> readData(String fileName);
    void removeSheet(String fileName, String sheetName);
    void createSheetTrn(String fileName, List<Transaction> transactions);
    void addSheetTrn(String fileName, List<Transaction> transactions);
    void addSheetStat(String fileName, List<Stat> stats);
    void createStat();
}
