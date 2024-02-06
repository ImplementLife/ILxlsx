package com.impllife.xlsx.service;

import com.impllife.xlsx.data.Transaction;

import java.util.List;

public interface ExcelService {
    List<Transaction> readData(String fileName);

    void createByTemplateStat();

    void createStat();
}
