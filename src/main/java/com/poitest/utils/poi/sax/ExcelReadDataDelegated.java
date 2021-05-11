package com.poitest.utils.poi.sax;

import java.util.List;

/**
 * @author qjwyss
 * @date 2018/12/19
 * @description 读取excel数据委托接口
 */
public interface ExcelReadDataDelegated {

    /**
     * 每获取一条记录，即写数据
     * 在flume里每获取一条记录即写，而不必缓存起来，可以大大减少内存的消耗，这里主要是针对flume读取大数据量excel来说的
     *
     * @param sheetIndex    sheet位置
     * @param totalRowCount 该sheet总行数
     * @param curRow        行号
     * @param cellList      行数据
     */
    public abstract void readExcelDate(int sheetIndex, int totalRowCount, int curRow, List<String> cellList);
}
