package com.impllife.xlsx.data;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.StringJoiner;

public class Transaction {
    private Date fullDate;
    private Date date;
    private Date time;
    private String category;
    private String dscr;
    private BigDecimal sum;

    public Date getFullDate() {
        return fullDate;
    }
    public void setFullDate(Date fullDate) {
        this.fullDate = fullDate;
    }

    public Date getDate() {
        return date;
    }
    public void setDate(Date date) {
        this.date = date;
    }

    public Date getTime() {
        return time;
    }
    public void setTime(Date time) {
        this.time = time;
    }

    public String getCategory() {
        return category;
    }
    public void setCategory(String category) {
        this.category = category;
    }

    public String getDscr() {
        return dscr;
    }
    public void setDscr(String dscr) {
        this.dscr = dscr;
    }

    public BigDecimal getSum() {
        return sum;
    }
    public void setSum(BigDecimal sum) {
        this.sum = sum;
    }

    @Override
    public String toString() {
        return new StringJoiner("|", Transaction.class.getSimpleName() + "[", "]")
            .add("date='" + new SimpleDateFormat("dd.MM.yyyy").format(date) + "'")
            .add("time='" + new SimpleDateFormat("HH.mm").format(time) + "'")
            .add("category='" + category + "'")
            .add("dscr='" + dscr + "'")
            .add("sum='" + sum + "'")
            .toString();
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        Transaction that = (Transaction) o;

        if (getFullDate() != null ? !getFullDate().equals(that.getFullDate()) : that.getFullDate() != null)
            return false;
        if (getDate() != null ? !getDate().equals(that.getDate()) : that.getDate() != null) return false;
        if (getTime() != null ? !getTime().equals(that.getTime()) : that.getTime() != null) return false;
        if (getCategory() != null ? !getCategory().equals(that.getCategory()) : that.getCategory() != null)
            return false;
        if (getDscr() != null ? !getDscr().equals(that.getDscr()) : that.getDscr() != null) return false;
        return getSum() != null ? getSum().equals(that.getSum()) : that.getSum() == null;
    }

    @Override
    public int hashCode() {
        int result = getFullDate() != null ? getFullDate().hashCode() : 0;
        result = 31 * result + (getDate() != null ? getDate().hashCode() : 0);
        result = 31 * result + (getTime() != null ? getTime().hashCode() : 0);
        result = 31 * result + (getCategory() != null ? getCategory().hashCode() : 0);
        result = 31 * result + (getDscr() != null ? getDscr().hashCode() : 0);
        result = 31 * result + (getSum() != null ? getSum().hashCode() : 0);
        return result;
    }
}
