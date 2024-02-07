package com.impllife.xlsx.data.map;

import com.impllife.xlsx.service.util.DateUtil;
import org.apache.poi.ss.usermodel.Cell;

import java.util.Date;

public class DateConvert implements Convert<Date> {
    private final String pattern;
    public DateConvert(String pattern) {
        this.pattern = pattern;
    }

    @Override
    public Date convert(Cell cell) {
        return DateUtil.parseDateByPattern(cell.getStringCellValue(), pattern);
    }
}
