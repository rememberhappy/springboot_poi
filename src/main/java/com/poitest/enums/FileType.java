package com.poitest.enums;

/**
 * 文件类型枚举
 */
public enum FileType {
//    BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER, DATE, NULL
    DATE("date"), DOUBLE("double"), STRING("string"), BIGINT("bigint"), BOOLEAN("boolean"),CHAR("char");

    private String value;

    FileType(String value) {
        this.value = value;
    }

    public String getValue() {
        return value;
    }
}
