package com.poitest.domain;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.alibaba.excel.annotation.format.NumberFormat;
import com.alibaba.excel.annotation.write.style.*;
import com.poitest.utils.easyexcel.CustomStringStringConverter;
import org.apache.poi.ss.usermodel.FillPatternType;

@ContentRowHeight(15)// 设置 row 高度，不包含表头
@HeadRowHeight(15)// 设置 表头 高度(与 @ContentRowHeight 相反)
@ColumnWidth(15)// 设置列宽
@HeadStyle(fillPatternType = FillPatternType.SOLID_FOREGROUND, fillForegroundColor = 10)// 头背景设置成红色 IndexedColors.RED.getIndex()
@HeadFontStyle(fontHeightInPoints = 20)// 头字体设置成20
@ContentStyle(fillPatternType = FillPatternType.SOLID_FOREGROUND, fillForegroundColor = 17)// 内容的背景设置成绿色 IndexedColors.GREEN.getIndex()
@ContentFontStyle(fontHeightInPoints = 20)// 内容字体设置成20
@OnceAbsoluteMerge(firstRowIndex = 5, lastRowIndex = 6, firstColumnIndex = 1, lastColumnIndex = 2)// 将第6-7行的2-3列合并成一个单元格，下标从0开始
public class Student {
    /**
     * ExcelProperty.value:用名字去匹配，这里需要注意，如果名字重复，会导致只有一个字段读取到数据
     */
//    // 字符串的头背景设置成粉红 IndexedColors.PINK.getIndex()
//    @HeadStyle(fillPatternType = FillPatternType.SOLID_FOREGROUND, fillForegroundColor = 14)
//    // 字符串的头字体设置成20
//    @HeadFontStyle(fontHeightInPoints = 30)
//    // 字符串的内容的背景设置成天蓝 IndexedColors.SKY_BLUE.getIndex()
//    @ContentStyle(fillPatternType = FillPatternType.SOLID_FOREGROUND, fillForegroundColor = 40)
//    // 字符串的内容字体设置成20
//    @ContentFontStyle(fontHeightInPoints = 30)
//    @ContentLoopMerge(eachRow = 2)// 列上合并单元格，这一列 每隔2行 合并单元格
    @ExcelProperty(value = "姓名")
    private String name;
    /**
     * 强制读取第三个 这里不建议 index 和 value 同时用，要么一个对象只用index，要么一个对象只用name去匹配
     * index和name除了读的时候读取指定的列，在写的时候，还可以指定写入的位置或者表头的名称
     * CustomStringStringConverter：自定义 转换器，对年龄做处理，>100或<1，按照20岁算
     */
    @ExcelProperty(index = 0, converter = CustomStringStringConverter.class)
//    @ExcelProperty(value = {"主标题", "年龄"})// 复杂表头  这两个表头是双层的表头，第一层是主标题，第二层是住址
    private String age;
    @ExcelProperty("住址")
//    @ExcelProperty({"主标题", "住址"})// 复杂表头  这两个表头是双层的表头，第一层是主标题，第二层是住址
    private String address;
    @ExcelProperty("手机号")
    private String phone;
    /**
     * NumberFormat：想接收百分比的数字
     */
    @NumberFormat("#.##%")
    @ExcelProperty("概率")
    private String probabilityOf;
    /**
     * DateTimeFormat：这里用string 去接日期才能格式化。我想接收年月日格式
     */
    @DateTimeFormat("yyyy年MM月dd日HH时mm分ss秒")
    @ExcelProperty("日期")
    private String date;
    /**
     * ExcelIgnore：忽略这个字段
     */
    @ExcelIgnore
    private String ignore;

    public Student() {
    }

    public Student(String name, String age, String address, String phone) {
        this.name = name;
        this.age = age;
        this.address = address;
        this.phone = phone;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    @Override
    public String toString() {
        return "Student{" +
                "name='" + name + '\'' +
                ", age='" + age + '\'' +
                ", address='" + address + '\'' +
                ", phone='" + phone + '\'' +
                '}';
    }
}
