package com.impllife.xlsx.data;

import java.math.BigDecimal;
import java.util.Date;

public class Stat {
    private Date date;
    private BigDecimal sum;

    public Stat() {
    }

    public Stat(Date date, BigDecimal sum) {
        this.date = date;
        this.sum = sum;
    }

    public Date getDate() {
        return date;
    }
    public void setDate(Date date) {
        this.date = date;
    }

    public BigDecimal getSum() {
        return sum;
    }
    public void setSum(BigDecimal sum) {
        this.sum = sum;
    }
}
