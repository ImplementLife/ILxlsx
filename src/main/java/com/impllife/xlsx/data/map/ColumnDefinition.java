package com.impllife.xlsx.data.map;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.impllife.xlsx.data.Transaction;
import org.apache.poi.ss.usermodel.Cell;

import java.util.Map;
import java.util.function.BiConsumer;

public class ColumnDefinition {
    private Integer index;
    private String setter;
    private Convert<?> convert;

    @JsonCreator
    public ColumnDefinition(@JsonProperty("index") Integer index,
                            @JsonProperty("convert") Object convert,
                            @JsonProperty("setter") String setter) {
        this.index = index;
        this.setter = setter;
        this.convert = ConvertFabric.create(convert);
    }

    public Integer getIndex() {
        return index;
    }
    public void setIndex(Integer index) {
        this.index = index;
    }

    public String getSetter() {
        return setter;
    }
    public void setSetter(String setter) {
        this.setter = setter;
    }

    public Convert<?> getConvert() {
        return convert;
    }
    public void setConvert(Convert<?> convert) {
        this.convert = convert;
    }

    public <T> T convert(Cell cell) {
        return convert.convert(cell);
    }
}
