package com.impllife.xlsx.data.map;

import java.util.Map;

public final class ConvertFabric {
    public static Convert<?> create(Object json) {
        if (json instanceof Map jsonAsMap) {
            String type = (String) jsonAsMap.get("type");
            if (type.equals("numeric")) {
                return new NumberConvert((int) jsonAsMap.get("scale"));
            } else if (type.equals("date")) {
                return new DateConvert((String) jsonAsMap.get("pattern"));
            }
        } else if (json instanceof String) {
            return new StringConvert();
        }
        throw new IllegalArgumentException("Unknown convert type.");
    }
}
