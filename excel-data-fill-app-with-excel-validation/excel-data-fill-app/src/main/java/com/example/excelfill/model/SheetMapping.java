package com.example.excelfill.model;
import java.util.LinkedHashMap;
import java.util.Map;
public class SheetMapping {
    private final String objectName;
    private final Map<Integer, String> columnMappings;
    public SheetMapping(String objectName, Map<Integer, String> columnMappings) {
        this.objectName = objectName;
        this.columnMappings = new LinkedHashMap<>(columnMappings);
    }
    public String getObjectName() { return objectName; }
    public Map<Integer, String> getColumnMappings() { return columnMappings; }
}
