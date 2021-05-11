package com.poitest.utils.easyexcel;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.property.ExcelContentProperty;

/**
 * easyExcel自定义的转换器
 */
public class CustomStringStringConverter implements Converter<String> {
    @Override
    public Class supportJavaTypeKey() {
        return String.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    /**
     * 这里读的时候会调用
     *
     * @param cellData            NotNull 单元格内容
     * @param contentProperty     Nullable 内容属性
     * @param globalConfiguration NotNull 全局配置
     * @return
     */
    @Override
    public String convertToJavaData(CellData cellData, ExcelContentProperty contentProperty,
                                    GlobalConfiguration globalConfiguration) {
        int numberValue = Integer.parseInt(cellData.getStringValue());
        if (numberValue < 1 || numberValue > 100) {
            return "年龄错误";
        }
        return String.valueOf(numberValue);
    }

    /**
     * 这里是写的时候会调用 不用管
     *
     * @param value               NotNull
     * @param contentProperty     Nullable
     * @param globalConfiguration NotNull
     * @return
     */
    @Override
    public CellData convertToExcelData(String value, ExcelContentProperty contentProperty,
                                       GlobalConfiguration globalConfiguration) {
        return new CellData(value);
    }
}
