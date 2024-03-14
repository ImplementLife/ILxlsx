package com.impllife.xlsx.data;

import java.math.BigDecimal;

public class StatByString {
    private String str;
    private BigDecimal sum;

    public StatByString() {
    }

    public StatByString(String date, BigDecimal sum) {
        this.str = date;
        this.sum = sum;
    }

    public String getStr() {
        return str;
    }
    public void setStr(String str) {
        this.str = str;
    }

    public BigDecimal getSum() {
        return sum;
    }
    public void setSum(BigDecimal sum) {
        this.sum = sum;
    }
}
