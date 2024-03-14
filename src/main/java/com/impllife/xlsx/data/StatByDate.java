package com.impllife.xlsx.data;

import java.math.BigDecimal;
import java.util.Date;

public class StatByDate {
    private Date date;
    private BigDecimal sum;

    public StatByDate() {
    }

    public StatByDate(Date date, BigDecimal sum) {
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
