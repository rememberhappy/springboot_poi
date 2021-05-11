package com.poitest.utils.poi;

import com.poitest.utils.poi.sax.ExcelReadDataDelegated;
import com.poitest.utils.poi.sax.ExcelSAXWithDefaultHandler;
import org.springframework.util.StringUtils;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 导入的工具类，使用poi【默认是dom方式的解析】的sax解析方式
 * SAX解析是基于流来读取文件的，每次占用内存是非常小的，如果能做好解析完的数据就直接处理掉，基本是不怎么耗内存的
 * SAX解析数据可能会丢失精度，需要保留一下有效数字
 */
public class ExcelImportSAXUtil {
    public static void readExcel(String filePath, ExcelReadDataDelegated excelReadDataDelegated) throws Exception {
        int totalRows = 0;
        // 处理格式
        if (ExcelUtil.isExcel2007(filePath)) {// 如果是2007版【格式为xlsx】的
            ExcelSAXWithDefaultHandler excelXlsxReader = new ExcelSAXWithDefaultHandler(excelReadDataDelegated);
            totalRows = excelXlsxReader.process(filePath);
        } else {
            throw new Exception("文件格式错误，fileName的扩展名只能是xlsx!");
        }
        System.out.println("读取的数据总行数：" + totalRows);
    }

    public static void readExcel(String filePath, InputStream inputStream, ExcelReadDataDelegated excelReadDataDelegated) throws Exception {
        int totalRows = 0;
        // 处理格式
        if (ExcelUtil.isExcel2007(filePath)) {// 如果是2007版【格式为xlsx】的
            ExcelSAXWithDefaultHandler excelXlsxReader = new ExcelSAXWithDefaultHandler(excelReadDataDelegated);
            totalRows = excelXlsxReader.process(inputStream, filePath);
        } else {
            throw new Exception("文件格式错误，fileName的扩展名只能是xlsx!");
        }
        System.out.println("读取的数据总行数：" + totalRows);
    }

    /**
     * main 入口方法进行测试
     *
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {
        String filePath = "C:\\Users\\wd\\Desktop\\训练营列表 (4).xlsx";
        // 用来处理符合条件的数据
        List mobileManagerList = new ArrayList();
        int PER_READ_INSERT_BATCH_COUNT = 10000;
        readExcel(filePath, new ExcelReadDataDelegated() {
            // 定义的扩展类
            @Override
            public void readExcelDate(int sheetIndex, int totalRowCount, int curRow, List<String> cellList) {
//                System.out.println("总行数为：" + totalRowCount + " 行号为：" + curRow + " 数据：" + cellList);
                // TODO 校验数据合法性
                Boolean legalFlag = true;
                // 号码、成本号码费、成本低消费、客户号码费、客户低消费不能为空
                if (StringUtils.isEmpty(cellList.get(0))) {
                    legalFlag = false;
                }
                // TODO 很多个判断逻辑
                // 如果数据合法，则封装号码对象
                if (legalFlag) {
                    // 处理数据合法的对象，可以将cellList中的数据转换为对象
                    mobileManagerList.add(cellList);
                } else {
                    throw new RuntimeException("数据不合法");
                }
                // TODO 批量保存
                try {
                    // 每1000条 批量保存一次
                    if (mobileManagerList.size() >= PER_READ_INSERT_BATCH_COUNT) {
                        // 批量保存逻辑
                        mobileManagerList.clear();
                    } else if (mobileManagerList.size() < PER_READ_INSERT_BATCH_COUNT) {// 没有达到10000的数量时
                        // 计算总共能存多少次    总行数/10000
                        int lastInsertBatchCount = totalRowCount % PER_READ_INSERT_BATCH_COUNT == 0 ?
                                totalRowCount / PER_READ_INSERT_BATCH_COUNT :
                                totalRowCount / PER_READ_INSERT_BATCH_COUNT + 1;
                        // 计算 当前条数是否大于（最大存储次数-1的总共条数）
                        if ((curRow - 1) >= ((lastInsertBatchCount - 1) * PER_READ_INSERT_BATCH_COUNT + 1)
                                && (curRow - 1) < lastInsertBatchCount * PER_READ_INSERT_BATCH_COUNT) {
                            // 判断是不是最后一次存储的最后一条数据
                            if (curRow - 1 == totalRowCount) {
                                // 批量保存逻辑
                            }
                        }
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        });
    }
}
