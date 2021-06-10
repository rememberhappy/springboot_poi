package com.poitest.utils.poi;

import com.poitest.enums.FileType;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 导入的工具类，使用poi默认的DOM解析
 * 直接上几十万条数据的excel文件，内存会直接溢出了
 * 解析xlsx大文件的时候，POI本身会占据较大内存，例如100W行15列，POI自身将消耗400M+的内存，加上解析出来的内容会大于这个值，以100W为例大概需要700M+内存
 * excel文件请使用第一行表头，其余行信息的标准格式，如果有合并单元格情况，可能会解析失败（可以包含空行和空单元格，会自动过滤，但必须有表头）
 */
public class ExcelImportPOIUtil {

    /**
     * 根据excel的版本，读取excel里面的内容
     *
     * @param is          输入流
     * @param isExcel2003 excel是2003还是2007版本
     * @return
     */
    public static Workbook getExcelInfo(InputStream is, boolean isExcel2003) {
        /** 根据版本选择创建Workbook的方式 */
        Workbook wb = null;
        try {
            //当excel是2003时
            if (isExcel2003) {
                wb = new HSSFWorkbook(is);
            } else {  //当excel是2007时
                wb = new XSSFWorkbook(is);
            }
            return wb;
        } catch (IOException e) {
            e.printStackTrace();
        }
        return wb;
    }

    /**
     * 读取Excel内容，转化成实体对象集合
     *
     * @param wb
     * @return
     */
    public static <T> List<T> readExcelValueToBean(Workbook wb, Class<T> clazz) throws IllegalAccessException, InstantiationException, ParseException {
        // 1. 验证类型
        if (clazz == null) {
            throw new RuntimeException("类型不能指定为空！");
        }
        T t = clazz.newInstance();
        Field[] fields = clazz.getDeclaredFields();
        // 2. 处理excel中的数据
        // 得到第一个shell
        Sheet sheet = wb.getSheetAt(0);
        // 得到sheet页中的Excel的行数
        int totalRows = sheet.getPhysicalNumberOfRows();
        System.out.println("读取到的行数：" + totalRows);
        // 得到Excel的列数(前提是 sheet中有行数，只有表头也算)
        int totalCells = 0;
        if (totalRows >= 1 && sheet.getRow(0) != null) {
            // 获取指定sheet页中第一行的列数
            totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
        }
        // 返回的值
        List<List<String>> valueList = new ArrayList();
        // 记录空行 规则：如果连续空行大于1行 下面的视为垃圾数据。【可以重新制定规则】
        int blankLine = 0;
        // 循环Excel行数,从第二行开始【下标是1】。标题不入库
        for (int r = 1; r < totalRows; r++) {
            // 取到sheet页中每一行的数据
            Row row = sheet.getRow(r);
            // 1. 处理空行逻辑
            if (row == null) {
                // 遇到空白行 获取的行数加1
                blankLine++;
                // sheet中的行数加1，获取的时候只获取到有数据的行数，中间有一行空白行是不计入的，所以加1
                totalRows++;
                if (blankLine > 1) {
                    // sheet页中总行数 重新定义，结束循环，规则规定，连续超过一行的空白行后，下面的数据视为垃圾数据
                    totalRows = r;
                    break;
                }
                continue;
            } else {  // 无空白行 重置计数器
                blankLine = 0;
            }
            // 2. 处理非空行的列数据
            List<String> temp = new ArrayList();
            // 标记是否为插入的空白行 识别规则 插入的数据后第一个单元格为空
            boolean addFlag = false;
            //循环Excel的列
            for (int c = 0; c < totalCells; c++) {
                // 获取到每一列的值
                Cell cell = row.getCell(c);
                if (null != cell) {// 列不为空的话
                    String cellValue = getCellValue(cell);
                    // 针对又见插入的行 poi默认它不算空行 判断该行如果有一个 不为空 该条记录视为有效
                    if ("".equals(cellValue) && (!addFlag)) {
                        addFlag = false;
                    } else {
                        addFlag = true;
                    }
                    if ("".equals(cellValue)) {
                        temp.add("");
                    } else {
                        temp.add(cellValue);
                    }
                } else {// 列为空的话
                    temp.add("");
                }
            }
            System.out.println("每一行数据：" + temp + ", 是否为有效数据：" + addFlag);
            if (addFlag) { // 判断是否为有效数据
                valueList.add(temp);
            }
        }
        // 判断要被转的类的字段与excel中的列的数量的一致性。存在优化空间，使用注解的方式，可以将列中的字段与实体中的字段对应起来，形成一对一的映射关系
        if (fields == null || fields.length > valueList.get(0).size()) {
            throw new RuntimeException("excel中与实体中的字段数量不一致");
        }
        List<T> beanList = new ArrayList<>();
        for (List<String> value : valueList) {
            t = clazz.newInstance();
            for (int i = 0; i < fields.length; i++) {
                fields[i].setAccessible(true);
                Class<?> type = fields[i].getType();
                String name = type.getName();
                name = name.substring(name.lastIndexOf(".") + 1, name.length());
                // 该字段为空的话
                if (StringUtils.isBlank(value.get(i))) {
                    fields[i].set(t, null);
                    continue;
                }
                switch (name) {
                    case "Integer":
                        fields[i].set(t, Integer.valueOf(value.get(i)));
                        break;
                    case "String":
                        fields[i].set(t, String.valueOf(value.get(i)));
                        break;
                    case "Double":
                        fields[i].set(t, Double.valueOf(value.get(i)));
                        break;
                    case "Float":
                        fields[i].set(t, Float.valueOf(value.get(i)));
                        break;
                    case "Boolean":
                        fields[i].set(t, Boolean.valueOf(value.get(i)));
                        break;
                    case "Short":
                        fields[i].set(t, Short.valueOf(value.get(i)));
                        break;
                    case "Long":
                        fields[i].set(t, Long.valueOf(value.get(i)));
                        break;
                    case "BigDecimal":
                        fields[i].set(t, new BigDecimal(value.get(i)));
                        break;
                    case "Date":
                        fields[i].set(t,
                                new SimpleDateFormat("yyyy-MM-dd").parse(value.get(i)));
                        break;
                }
            }
            beanList.add(t);
        }
        return beanList;
    }

    /**
     * 读取Excel内容
     *
     * @param wb
     * @param isPriview
     * @return List<List<String>>
     */
    public static List readExcelValue(Workbook wb, Boolean isPriview) {
        // 得到第一个shell
        Sheet sheet = wb.getSheetAt(0);
        // 得到sheet页中的Excel的行数
        int totalRows = sheet.getPhysicalNumberOfRows();
        System.out.println("读取到的行数：" + totalRows);
        // 处理excel读取的行数，如果是预览模式，只读取前一百行数据
        if (isPriview && totalRows > 100) {
            totalRows = 101;
        }
        // 得到Excel的列数(前提是 sheet中有行数，只有表头也算)
        int totalCells = 0;
        if (totalRows >= 1 && sheet.getRow(0) != null) {
            // 获取指定sheet页中第一行的列数
            totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
        }
        // 返回的值
        List<List<String>> valueList = new ArrayList<>();
        // 记录空行 规则：如果连续空行大于1行 下面的视为垃圾数据。【可以重新制定规则】
        int blankLine = 0;
        // 循环Excel行数,从第二行开始【下标是1】。标题不入库
        for (int r = 1; r < totalRows; r++) {
            // 取到sheet页中每一行的数据
            Row row = sheet.getRow(r);
            // 1. 处理空行逻辑
            if (row == null) {
                // 遇到空白行 获取的行数加1
                blankLine++;
                // sheet中的行数加1，获取的时候只获取到有数据的行数，中间有一行空白行是不计入的，所以加1
                totalRows++;
                if (blankLine > 1) {
                    // sheet页中总行数 重新定义，结束循环，规则规定，连续超过一行的空白行后，下面的数据视为垃圾数据
                    totalRows = r;
                    break;
                }
                continue;
            } else {  // 无空白行 重置计数器
                blankLine = 0;
            }
            // 2. 处理非空行的列数据
            List<String> temp = new ArrayList<>();
            // 标记是否为插入的空白行 识别规则 插入的数据后第一个单元格为空
            boolean addFlag = false;
            //循环Excel的列
            for (int c = 0; c < totalCells; c++) {
                // 获取到每一列的值
                Cell cell = row.getCell(c);
                if (null != cell) {// 列不为空的话
                    String cellValue = getCellValue(cell);
                    // 针对又见插入的行 poi默认它不算空行 判断该行如果有一个 不为空 该条记录视为有效
                    if ("".equals(cellValue) && (!addFlag)) {
                        addFlag = false;
                    } else {
                        addFlag = true;
                    }
                    if ("".equals(cellValue)) {
                        temp.add("\\N");
                    } else {
                        temp.add(cellValue);
                    }
                } else {// 列为空的话
                    temp.add("\\N");
                }
            }
            System.out.println("每一行数据：" + temp + ", 是否为有效数据：" + addFlag);
            if (addFlag) { // 判断是否为有效数据
                valueList.add(temp);
            }
        }
        return valueList;
    }

    /**
     * 读取Excel表头
     *
     * @param wb
     * @return
     */
    public static List<String> readExcelTitle(Workbook wb) {
        // 得到第一个shell
        Sheet sheet = wb.getSheetAt(0);
        // 得到Excel的行数
        int totalRows = sheet.getPhysicalNumberOfRows();

        // 得到Excel的列数(前提是有行数)
        int totalCells = 0;
        if (totalRows >= 1 && sheet.getRow(0) != null) {
            // 得到指定sheet页的第一行的列数
            totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
        }
        // 读取标题，只读取第一行的数据
        Row row = sheet.getRow(0);
        if (row == null) return null;
        // 返回的对象
        List<String> titleList = new ArrayList<>();
        //循环Excel的列
        for (int c = 0; c < totalCells; c++) {
            // 获取每一个单元格的信息
            Cell cell = row.getCell(c);
            if (null != cell) {
                titleList.add(getCellValue(cell));
            } else {
                // 表头列遇到一个空白的标题 结束，可以重设sheet中有多少列
                // totalCells = c;
                break;
            }
        }
        return titleList;
    }

    /**
     * 读取Excel表头详细信息，值以及类型
     *
     * @param wb
     * @return
     */
    public static Map getColumnType(Workbook wb) {
        //得到第一个shell
        Sheet sheet = wb.getSheetAt(0);
        //得到Excel的行数
        int totalRows = sheet.getPhysicalNumberOfRows();
        //得到Excel的列数(前提是有行数)
        int totalCells = 0;
        if (totalRows >= 1 && sheet.getRow(0) != null) {
            totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
        }
        if (totalRows > 101) {
            totalRows = 101;
        }
        // 0,string
        Map rowColumns = new HashMap();
        // 记录空行 规则 如果空行大于1行 下面的视为垃圾数据
        int blankLine = 0;

        //循环Excel行数,从第二行开始。标题不入库
        for (int r = 1; r < totalRows; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                totalRows++;
                blankLine++;
                if (blankLine > 1) {
                    // totalrows 重新定义总行数
                    totalRows = r;
                    break;
                }
                continue;
            } else {  // 无空白行 重置计数器
                blankLine = 0;
            }
            //循环Excel的列
            for (int c = 0; c < totalCells; c++) {
                Cell cell = row.getCell(c);
                if (null != cell) {
                    String cellValue = getCellValue(cell);
                    Object value = rowColumns.get(c);
                    String val = (String) value;
                    String valType = getType(cellValue);
                    if (!"string".equals(val)) {
                        if ("string".equals(valType)) {
                            rowColumns.put(c, valType);
                        } else if (!"double".equals(val)) {
                            rowColumns.put(c, valType);
                        }
                    }
                } else {
                    rowColumns.put(c, "string");
                }
            }
        }
        return rowColumns;
    }

    /**
     * 判断单元格的数据类型，获取指定单元格的值
     *
     * @param cell
     * @return
     */
    public static String getCellValue(Cell cell) {
        String value = "";
        if (cell != null) {
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:// 数值型0
                    if ("General".equals(cell.getCellStyle().getDataFormatString())) {// 常规的
                        // 数据格式
                        DecimalFormat df = new DecimalFormat("#.########");
                        value = df.format(cell.getNumericCellValue()) + "";
                    } else if ("m/d/yy".equals(cell.getCellStyle().getDataFormatString())) {// 时间格式
                        value = sdf.format(cell.getDateCellValue()) + "";
                    } else {
                        // 针对十位数以上的数字出现科学记数法的处理
                        value = new DecimalFormat("#").format(cell.getNumericCellValue());
                    }
                    break;
                case Cell.CELL_TYPE_STRING:// 字符串型1
                    value = cell.getRichStringCellValue().getString();
                    break;
                case Cell.CELL_TYPE_FORMULA: //公式2
                    value = String.valueOf(cell.getCellFormula());
                    break;
                case Cell.CELL_TYPE_BLANK://空值3
                    value = "";
                    break;
                case Cell.CELL_TYPE_BOOLEAN://布尔型4
                    value = String.valueOf(cell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_ERROR: //故障
                    value = "非法字符";
                    break;
                default:
                    value = cell.toString();
                    break;
            }
        }
        return value;
    }
//    poi版本不同，方法和枚举值也不同
//    public String getCellValue(Cell cell) {
//        String value = "";
//        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
//        switch (cell.getCellTypeEnum()) {
//            case STRING:
//                value = cell.getRichStringCellValue().getString();
//                break;
//            case NUMERIC:
//                if ("General".equals(cell.getCellStyle().getDataFormatString())) {
//                    // 数据格式
//                    DecimalFormat df = new DecimalFormat("#.########");
//                    value = df.format(cell.getNumericCellValue())+"";
//                } else if ("m/d/yy".equals(cell.getCellStyle().getDataFormatString())) {
//                    value = sdf.format(cell.getDateCellValue())+"";
//                } else {
//                    // 针对十位数以上的数字出现科学记数法的处理
//                    value =   new DecimalFormat("#").format(cell.getNumericCellValue());
//                }
//                break;
//            case BOOLEAN:
//                value = cell.getBooleanCellValue() + "";
//                break;
//            case BLANK:
//                value = "";
//                break;
//            default:
//                value = cell.toString();
//                break;
//        }
//        return value;
//    }

    /**
     * 获取字符串的类型
     *
     * @param str
     * @return
     */
    public static String getType(String str) {
        // 优先判断日期类型
        String PATTERNING = "\\d{4}(-)\\d{2}(-)\\d{2}\\s\\d{2}(:)\\d{2}(:)\\d{2}";
        String PATTERN_DATE = "\\d{4}(-)\\d{2}(-)\\d{2}";
        String CHAR_PATTERN = "[^0-9]";
        String INT_PATTERN = "^-?[1-9]\\d*$";
        String DOUBLE_PATTERN = "^[-]?[1-9]\\d*\\.\\d*|-0\\.\\d*[1-9]\\d*$";
        // 首先去除两边的空格
        str = str.trim();
        if (str.matches(PATTERNING) || str.matches(PATTERN_DATE)) {// 正则验证 时间格式
            return FileType.DATE.getValue();
        }
        if ("true".equalsIgnoreCase(str) || "false".equalsIgnoreCase(str)) {// 正则验证 布尔类型
            return FileType.BOOLEAN.getValue();
        }
        if (str.matches(CHAR_PATTERN)) {// 正则验证 char类型
            return FileType.CHAR.getValue();
        }
        if (str.matches(INT_PATTERN)) {// 正则验证 int类型
            return FileType.BIGINT.getValue();
        }
        if (str.matches(DOUBLE_PATTERN)) {// 正则验证 double类型
            return FileType.DOUBLE.getValue();
        }
        return FileType.STRING.getValue();
    }
}
