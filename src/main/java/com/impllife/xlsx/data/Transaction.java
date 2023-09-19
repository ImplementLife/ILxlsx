package com.impllife.xlsx.data;

import java.util.StringJoiner;

public class Transaction {
    private String date;
    private String time;
    private String category;
    private String dscr;
    private String sum;

    public String getDate() {
        return date;
    }
    public void setDate(String date) {
        this.date = date;
    }

    public String getTime() {
        return time;
    }
    public void setTime(String time) {
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

    public String getSum() {
        return sum;
    }
    public void setSum(String sum) {
        this.sum = sum;
    }

    @Override
    public String toString() {
        return new StringJoiner("|", Transaction.class.getSimpleName() + "[", "]")
            .add("date='" + date + "'")
            .add("time='" + time + "'")
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

        if (getDate() != null ? !getDate().equals(that.getDate()) : that.getDate() != null) return false;
        if (getTime() != null ? !getTime().equals(that.getTime()) : that.getTime() != null) return false;
        if (getCategory() != null ? !getCategory().equals(that.getCategory()) : that.getCategory() != null)
            return false;
        if (getDscr() != null ? !getDscr().equals(that.getDscr()) : that.getDscr() != null) return false;
        return getSum() != null ? getSum().equals(that.getSum()) : that.getSum() == null;
    }

    @Override
    public int hashCode() {
        int result = getDate() != null ? getDate().hashCode() : 0;
        result = 31 * result + (getTime() != null ? getTime().hashCode() : 0);
        result = 31 * result + (getCategory() != null ? getCategory().hashCode() : 0);
        result = 31 * result + (getDscr() != null ? getDscr().hashCode() : 0);
        result = 31 * result + (getSum() != null ? getSum().hashCode() : 0);
        return result;
    }
}
