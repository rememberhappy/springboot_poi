package com.poitest.utils.poi;

import com.poitest.domain.Student;
import com.poitest.utils.poi.sax.ExcelReadDataDelegated;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.util.StringUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 公共的excel工具类
 */
public class ExcelUtil {

    // *****************************************导出*****************************************

    /**
     * 根据excel的版本导出excel【使用DOM解析】
     *
     * @param httpServletResponse 响应请求
     * @param objects             数据集合
     * @param type                数据类型
     * @param columnNames         列名集合
     * @param columns             要打印的列
     */
    public static <T> void exportExcel(HttpServletResponse httpServletResponse, List<T> objects, Class type,
                                       String[] columnNames, String... columns) {
        String header = httpServletResponse.getHeader("Content-Disposition");
        String fileName = header.substring(header.indexOf("=") + 1, header.length());
        System.out.println("文件名称：" + fileName);
        //验证文件名是否合格
        validateExcel(fileName);
        //根据文件名判断文件是2003版本还是2007版本
        try {
            if (isExcel2007(fileName)) {// 文件类型是 .xlsx 格式的，2007版本的
                System.out.println("文件类型是 .xlsx 格式的，2007版本的");
                if (objects.size() > 500000) {// XSSFWorkbook可以导出104万，如果数量超过50万的时候，选择大批量的SXSSFWorkbook
                    System.out.println("大批量数据，选用SXSSFWorkbook");
                    ExcelSXSSFWorkbookUtil.exportObjectsToExcel(httpServletResponse.getOutputStream(), objects, type, columnNames, columns);
                } else {
                    System.out.println("大批量数据，选用XSSFWorkbook");
                    ExcelXSSFWorkbookUtil.exportObjectsToExcel(httpServletResponse.getOutputStream(), objects, type, columnNames, columns);
                }
            } else {// 文件类型是 .xls 格式的，2003版本的
                System.out.println("文件类型是 .xls 格式的，2003版本的");
                ExcelHSSFWorkbookUtil.exportObjectsToExcel(httpServletResponse.getOutputStream(), objects, type, columnNames, columns);
            }
        } catch (Exception e) {
            throw new RuntimeException("导出失败");
        }
    }

    // **************************************导入***********************************

    /**
     * 读取excel文件内容【使用DOM解析】
     *
     * @param fileName
     * @param is
     * @param isPriview 是否预览
     * @return
     */
    public static Map<String, Object> importExcelDOM(String fileName, InputStream is, boolean isPriview) {
        Map<String, Object> result = new HashMap<String, Object>();
        try {
            //验证文件名是否合格
            validateExcel(fileName);
            //根据文件名判断文件是2003版本还是2007版本
            boolean isExcel2003 = true;
            if (isExcel2007(fileName)) {
                isExcel2003 = false;
            }
            // 获取excel内容，根据传入的excel类型，指定使用XSSFWorkbook或HSSFWorkbook格式进行加载
            Workbook wb = ExcelImportPOIUtil.getExcelInfo(is, isExcel2003);
            // 读取标题信息 其中也设置了有效列数量
            List titleList = ExcelImportPOIUtil.readExcelTitle(wb);
            //读取Excel信息
            List customerList = ExcelImportPOIUtil.readExcelValue(wb, isPriview);
            //读取Excel信息，转化成Bean对象
            List beanList = ExcelImportPOIUtil.readExcelValueToBean(wb, Student.class);
            // 读取表头信息，{value，valueType}
            Map columnstypes = null;
            // 是不是预览
            if (isPriview) {
                // 读取表头，将表头插入到第一行
                columnstypes = ExcelImportPOIUtil.getColumnType(wb);
                customerList.add(0, columnstypes);
            }
            result.put("schema", titleList);
            result.put("data", customerList);
            result.put("columnstypes", columnstypes);
            result.put("beanList", beanList);
            is.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    is = null;
                    e.printStackTrace();
                }
            }
        }
        return result;
    }

    /**
     * 使用SAX解析读取上传的excel【使用SAX解析】
     *
     * @param fileName
     * @param is
     * @param isPriview 是否预览
     * @throws Exception
     */
    public static void importExcelSAX(String fileName, InputStream is, boolean isPriview) throws Exception {
        //验证文件名是否合格
        validateExcel(fileName);
        // 用来处理符合条件的数据
        List mobileManagerList = new ArrayList();
        int PER_READ_INSERT_BATCH_COUNT = 10000;
        ExcelImportSAXUtil.readExcel(fileName, is, new ExcelReadDataDelegated() {
            // 定义的扩展类
            @Override
            public void readExcelDate(int sheetIndex, int totalRowCount, int curRow, List<String> cellList) {
                // TODO 校验数据合法性
                Boolean legalFlag = true;
                // 号码、成本号码费、成本低消费、客户号码费、客户低消费不能为空
                if (StringUtils.isEmpty(cellList.get(0))) {
                    legalFlag = false;
                }
                // TODO 可以处理很多个判断逻辑
                // 如果数据合法，则封装号码对象
                if (legalFlag) {
                    // 处理数据合法的对象，可以将cellList中的数据转换为对象
                    mobileManagerList.add(cellList);
                } else {
                    throw new RuntimeException("数据不合法");
                }
                // TODO 每1000条批量保存
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
    // ********************************公用方法****************************************

    /**
     * 是否是2003的excel，返回true是2003
     *
     * @param filePath
     * @return
     */
    public static boolean isExcel2003(String filePath) {
        return filePath.matches("^.+\\.(?i)(xls)$");
    }

    /**
     * 是否是2007的excel，返回true是2007
     *
     * @param filePath
     * @return
     */
    public static boolean isExcel2007(String filePath) {
        return filePath.matches("^.+\\.(?i)(xlsx)$");
    }

    /**
     * 验证是不是EXCEL文件【文件后缀判断 xls/xlsx】
     *
     * @param filePath
     * @return
     */
    public static void validateExcel(String filePath) {
        if (filePath == null || !(isExcel2003(filePath) || isExcel2007(filePath))) {
            throw new RuntimeException("文件名不是excel格式");
        }
    }
}
