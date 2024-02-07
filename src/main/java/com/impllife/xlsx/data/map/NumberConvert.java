package com.impllife.xlsx.data.map;

import org.apache.poi.ss.usermodel.Cell;

import java.math.BigDecimal;
import java.math.RoundingMode;

public class NumberConvert implements Convert<BigDecimal> {
    private int scale = 2;
    public NumberConvert(int scale) {
        this.scale = scale;
    }

    @Override
    public BigDecimal convert(Cell cell) {
        return BigDecimal.valueOf(cell.getNumericCellValue()).setScale(scale, RoundingMode.CEILING);
    }
}
