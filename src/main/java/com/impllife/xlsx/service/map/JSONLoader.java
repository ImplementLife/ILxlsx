package com.impllife.xlsx.service.map;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.type.CollectionType;
import com.impllife.xlsx.data.map.ColumnDefinition;

import java.io.File;
import java.io.IOException;
import java.util.List;

public class JSONLoader {
    private static final ObjectMapper objectMapper = new ObjectMapper();

    public static List<ColumnDefinition> loadColumnDefinitions(String jsonFilePath) {
        File jsonFile = new File(jsonFilePath);
        CollectionType columnDefinitionListType = objectMapper.getTypeFactory().constructCollectionType(List.class, ColumnDefinition.class);
        try {
            return objectMapper.readValue(jsonFile, columnDefinitionListType);
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }
    }
}

