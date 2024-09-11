package com.fis.services;

import java.util.Map;

public class DynamicObject {
    private Map<String, Object> properties;
    private Map<String,String> columns;

    public DynamicObject(Map<String, Object> properties, Map<String,String> columns) {
        this.properties = properties;
        this.columns = columns;
    }

    public DynamicObject() {
    }

    public DynamicObject(Map<String, Object> properties) {
        this.properties = properties;
    }

    public Map<String, Object> getProperties() {
        return properties;
    }

    public void setProperties(Map<String, Object> properties) {
        this.properties = properties;
    }

    public Map<String, String> getColumns() {
        return columns;
    }

    public void setColumns(Map<String, String> columns) {
        this.columns = columns;
    }

    @Override
    public String toString() {
        return "DynamicObject{" +
                "properties=" + properties +
                ", columns=" + columns +
                '}';
    }
}
