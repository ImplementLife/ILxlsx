package com.impllife.xlsx.data.map;

import org.apache.poi.ss.usermodel.Cell;

public interface Convert<T> {
    <T> T convert(Cell cell);
}
