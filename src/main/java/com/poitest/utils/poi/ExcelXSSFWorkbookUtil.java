package com.poitest.utils.poi;

import org.apache.logging.log4j.util.Strings;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.lang.reflect.Field;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Optional;

/**
 * poi之XSSFWorkbook
 * HSSFWorkbook:是操作Excel2007的版本，扩展名是.xlsx；
 * 这种形式的出现是为了突破HSSFWorkbook的65535行局限。其对应的是excel2007(1048576行，16384列)扩展名为“.xlsx”，最多可以导出104万行，
 * 不过这样就伴随着一个问题---OOM内存溢出，原因是你所创建的book sheet row cell等此时是存在内存的并没有持久化。
 */
public class ExcelXSSFWorkbookUtil {

    // 日志
    private static final Logger logger = LoggerFactory.getLogger(ExcelXSSFWorkbookUtil.class);

    // threadLocal保证了每个线程中都有一份独立的数据
    private static final ThreadLocal<DateFormat> FORMAT = new ThreadLocal<DateFormat>() {
        @Override
        protected DateFormat initialValue() {
            return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        }
    };

    /**
     * 将对象集合转化为excel【导出】
     *
     * @param outputStream 要输出的流对象
     * @param objects      数据集合
     * @param type         数据类型
     * @param columnNames  列名集合
     * @param columns      要打印的列
     */
    public static final <T> void exportObjectsToExcel(OutputStream outputStream, List<T> objects, Class type,
                                                     String[] columnNames, String... columns) throws NoSuchFieldException, IOException, IllegalAccessException {
        // 创建excel
        XSSFWorkbook wb = new XSSFWorkbook();
        wirteXSSWorkbookData(wb, objects, type, columnNames, columns);
        // 将文件写入流
        wb.write(outputStream);
        // 关闭流
        outputStream.flush();
        outputStream.close();
    }

    /**
     * 将对象集合转化为excel的InputStream
     *
     * @param objects     数据集合
     * @param type        数据类型
     * @param columnNames 列名集合
     * @param columns     要打印的列
     */
    public static final <T> InputStream importObjectsExcelInputStream(List<T> objects, Class type,
                                                                   String[] columnNames, String... columns) throws NoSuchFieldException, IllegalAccessException {
        // 创建excel
        XSSFWorkbook wb = new XSSFWorkbook();
        wirteXSSWorkbookData(wb, objects, type, columnNames, columns);
        ByteArrayInputStream in = null;
        try {
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            wb.write(os);
            byte[] b = os.toByteArray();
            in = new ByteArrayInputStream(b);
            os.close();
        } catch (IOException e) {
            logger.error("ExcelUtils getExcelFile error:{}", e.toString());
            return null;
        }
        return in;
    }

    /**
     * 将数据写入到XSSFWorkbook中
     *
     * @param wb          XSSFWorkbook对象
     * @param objects     写入的数据
     * @param type        数据类型
     * @param columnNames 列名集合
     * @param columns     要写入Excel的列
     * @param <T>         泛型
     * @throws NoSuchFieldException
     * @throws IllegalAccessException
     */
    private static <T> void wirteXSSWorkbookData(XSSFWorkbook wb, List<T> objects, Class type, String[] columnNames, String[] columns) throws NoSuchFieldException, IllegalAccessException {
        // 创建表单
        XSSFSheet sheet = wb.createSheet("sheet (total " + objects.size() + ")");
        // 设置文本格式
        XSSFCellStyle style = wb.createCellStyle();
//        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setFillForegroundColor(HSSFColor.BRIGHT_GREEN.index);
//        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        // 写入列名
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell;
        for (int i = 0; i < columnNames.length; i++) {
            sheet.setColumnWidth(i, 20 * 256);
            cell = row.createCell(i);
            cell.setCellValue(columnNames[i]);
            cell.setCellStyle(style);
        }
        // 要打印的列
        List<Field> fieldList = new ArrayList<>();
        for (String column : columns) {
            fieldList.add(type.getDeclaredField(column));
        }
        // 写入数据
        for (int i = 0; i < objects.size(); i++) {
            row = sheet.createRow(i + 1);
            Object obj = objects.get(i);
            for (int j = 0; j < fieldList.size(); j++) {
                Field field = fieldList.get(j);
                field.setAccessible(true);
                Object fieldObj = Optional.ofNullable(field.get(obj)).orElse("");
                if (field.getGenericType().toString().endsWith("Date")) {
                    row.createCell(j).setCellValue(Strings.isBlank(fieldObj.toString()) ? fieldObj.toString() : FORMAT.get().format((Date) fieldObj));
                } else {
                    row.createCell(j).setCellValue(fieldObj.toString());
                }
            }
        }
    }

}
