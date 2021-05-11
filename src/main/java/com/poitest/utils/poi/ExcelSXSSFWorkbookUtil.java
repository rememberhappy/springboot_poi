package com.poitest.utils.poi;

import org.apache.logging.log4j.util.Strings;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
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
 * poi之SXSSFWorkbook
 * 从POI 3.8版本开始，提供了一种基于XSSF的低内存占用的API----SXSSFWorkbook
 * 当数据量超出65536条后，在使用HSSFWorkbook或XSSFWorkbook，程序会报OutOfMemoryError：Javaheap space;内存溢出错误。这时应该用SXSSFworkbook。
 * 注意：针对 SXSSFWorkbook Beta 3.8下，会有临时文件产生，比如：
 * poi-sxssf-sheet4654655121378979321.xml
 * 文件位置：java.io.tmpdir这个环境变量下的位置
 * Windows 7下是C:\Users\xxxxxAppData\Local\Temp
 * Linux下是 /var/tmp/
 * 要根据实际情况，看是否删除这些临时文件
 * 与XSSF的对比
 * 在一个时间点上，只可以访问一定数量的数据
 * 不再支持Sheet.clone()
 * 不再支持公式的求值
 * 在使用Excel模板下载数据时将不能动态改变表头，因为这种方式已经提前把excel写到硬盘的了就不能再改了
 */
public class ExcelSXSSFWorkbookUtil {

    // 日志
    private static final Logger logger = LoggerFactory.getLogger(ExcelSXSSFWorkbookUtil.class);

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
        long startTime = System.currentTimeMillis();
        // 创建excel
        SXSSFWorkbook wb = new SXSSFWorkbook(1000);
        wirteXSSWorkbookData(wb, objects, type, columnNames, columns);
        // 将文件写入流
        wb.write(outputStream);
        // 关闭流
        outputStream.flush();
        outputStream.close();
        // 在磁盘上释放备份此工作簿的临时文件
        wb.dispose();
        System.out.println("导出共耗时：" + ((System.currentTimeMillis() - startTime) / 1000) + "秒");// 412秒
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
        SXSSFWorkbook wb = new SXSSFWorkbook();
        wirteXSSWorkbookData(wb, objects, type, columnNames, columns);
        ByteArrayInputStream in = null;
        try {
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            wb.write(os);
            byte[] b = os.toByteArray();
            in = new ByteArrayInputStream(b);
            os.close();
            // 在磁盘上释放备份此工作簿的临时文件
            wb.dispose();
        } catch (IOException e) {
            logger.error("ExcelUtils getExcelFile error:{}", e.toString());
            return null;
        }
        return in;
    }

    /**
     * 将数据写入到SXSSFWorkbook中
     * 大批量数据的时候，如果超过一百万，不能全部放入一个sheet页中，需要动态的创建sheet页
     *
     * @param wb          SXSSFWorkbook对象
     * @param objects     写入的数据
     * @param type        数据类型
     * @param columnNames 列名集合
     * @param columns     要写入Excel的列
     * @param <T>         泛型
     * @throws NoSuchFieldException
     * @throws IllegalAccessException
     */
    private static <T> void wirteXSSWorkbookData(SXSSFWorkbook wb, List<T> objects, Class type, String[] columnNames, String[] columns) throws NoSuchFieldException, IllegalAccessException {
        // 创建表单
        SXSSFSheet sheet = wb.createSheet("我的第" + 1 + "个工作簿");
        // 设置文本格式
        CellStyle style = wb.createCellStyle();
//        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setFillForegroundColor(HSSFColor.BRIGHT_GREEN.index);
//        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        // 写入列名
        SXSSFRow row = sheet.createRow(0);
        SXSSFCell cell;
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
        int rowNum = 1;
        // 每页显示多少条数据【含表头】
        int pageNum = 1000001;
        // 写入数据
        for (int i = 0; i < objects.size(); i++) {
            //打印5条后切换到下个工作表【表头占一行，实际一页填充六行】，可根据需要自行拓展，1000，10000...数据一样操作，只要不超过1048576就可以
            if (rowNum % pageNum == 0) {
                System.out.println("自动创建一个sheet页，当前是第 " + (i / pageNum + 2) + " 个sheet页");
                sheet = wb.createSheet("我的第" + (i / pageNum + 2) + "个工作簿"); //建立新的sheet对象
                sheet = wb.getSheetAt(i / pageNum + 1); //动态指定当前的工作表，下面操作的就是这个指定的sheet页【根据下标进行选择sheet页】
                // 设置新创建的sheet页的表头信息
                setHeaderInformation(sheet, columnNames, style);
                rowNum = 1; //每当新建了工作表就将当前工作表的行号重置为1，第0行已经设置表头信息
            }
            row = sheet.createRow(rowNum++);
            Object obj = objects.get(i);
            // 处理单元格数据
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

    /**
     * 设置表头信息
     */
    private static void setHeaderInformation(SXSSFSheet sheet, String[] columnNames, CellStyle style) {
        SXSSFRow row = sheet.createRow(0);
        SXSSFCell cell;
        for (int i = 0; i < columnNames.length; i++) {
            sheet.setColumnWidth(i, 20 * 256);
            cell = row.createCell(i);
            cell.setCellValue(columnNames[i]);
            cell.setCellStyle(style);
        }
    }
}
