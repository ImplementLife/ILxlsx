package com.impllife.xlsx.data.map;

import org.apache.poi.ss.usermodel.Cell;

public class StringConvert implements Convert<String> {
    @Override
    public String convert(Cell cell) {
        return cell.getStringCellValue();
    }
}
