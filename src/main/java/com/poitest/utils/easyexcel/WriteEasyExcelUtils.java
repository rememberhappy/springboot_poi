package com.poitest.utils.easyexcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.util.FileUtils;
import com.alibaba.excel.write.merge.LoopMergeStrategy;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.WriteTable;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import com.poitest.domain.ImageData;
import com.poitest.domain.Student;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.*;

/**
 * 关于写的 EasyExcel 几种写法
 * 注意：EasyExcel.write和poi-ooxml包有冲突
 */
public class WriteEasyExcelUtils {
    private static final Logger LOGGER = LoggerFactory.getLogger(WriteEasyExcelUtils.class);

    public static void main(String[] args) {
        // 处理需要导出的数据
        List<Student> trainDtoList = new ArrayList<Student>();
        // 设置大批量数据，一万条
        for (int i = 1; i < 50; i++) {
            trainDtoList.add(new Student("张三" + i, "13" + i, "北京3" + i, "333333" + i));
        }
        String fileName = "F:\\训练营列表.xlsx";

//        // 有个很重要的点 DataListener 不能被spring管理，要每次读取excel都要new,然后里面用到spring可以构造方法传进去
//        // 在实体中加注解@ExcelProperty("姓名")可以将excel中的列与实体字段保持一致，加@ExcelIgnore忽略某一个字段
//        // 写法1：默认写入第一个sheet页
//        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
//        // 如果这里想使用03 则 传入excelType参数即可
//        EasyExcel.write(fileName, Student.class).sheet("模板").doWrite(trainDtoList);
//
//        // 写法2：指定要写入的sheet页
//        // 这里 需要指定写用哪个class去写
//        ExcelWriter excelWriter = null;
//        try {
//            excelWriter = EasyExcel.write(fileName, Student.class).build();
//            WriteSheet writeSheet = EasyExcel.writerSheet("模板").build();
//            excelWriter.write(trainDtoList, writeSheet);
//        } finally {
//            // 千万别忘记finish 会帮忙关闭流
//            if (excelWriter != null) {
//                excelWriter.finish();
//            }
//        }
//
//        // 写法3：指定不写入的sheet页的字段
//        // 根据用户传入字段 假设我们要忽略 date
//        Set<String> excludeColumnFiledNames = new HashSet<String>();
//        excludeColumnFiledNames.add("date");
//        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
//        EasyExcel.write(fileName, Student.class).excludeColumnFiledNames(excludeColumnFiledNames).sheet("模板")
//                .doWrite(trainDtoList);
//
//        // 写法4：指定写入的sheet页的字段
//        // 根据用户传入字段 假设我们只要导出 date
//        Set<String> includeColumnFiledNames = new HashSet<String>();
//        includeColumnFiledNames.add("date");
//        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
//        EasyExcel.write(fileName, Student.class).includeColumnFiledNames(includeColumnFiledNames).sheet("模板")
//                .doWrite(trainDtoList);
//
//        // 写法5：将同样的对象多次写入到同一个sheet页中
////        ExcelWriter excelWriter = null;
//        try {
//            // 这里 需要指定写用哪个class去写
//            excelWriter = EasyExcel.write(fileName, Student.class).build();
//            // 这里注意 如果同一个sheet只要创建一次
//            WriteSheet writeSheet = EasyExcel.writerSheet("模板").build();
//            // 去调用写入,这里我调用了五次，实际使用时根据数据库分页的总的页数来
//            for (int i = 0; i < 5; i++) {
//                // 分页去数据库查询数据 这里可以去数据库查询每一页的数据
//                List<Student> data = trainDtoList;
//                excelWriter.write(data, writeSheet);
//            }
//        } finally {
//            // 千万别忘记finish 会帮忙关闭流
//            if (excelWriter != null) {
//                excelWriter.finish();
//            }
//        }
//
//        // 方法6：将同样的对象多次写入到不同的sheet页中
//        try {
//            // 这里 指定文件
//            excelWriter = EasyExcel.write(fileName, Student.class).build();
//            // 去调用写入,这里我调用了五次，实际使用时根据数据库分页的总的页数来。这里最终会写到5个sheet里面
//            for (int i = 0; i < 5; i++) {
//                // 每次都要创建writeSheet 这里注意必须指定sheetNo 而且sheetName必须不一样
//                WriteSheet writeSheet = EasyExcel.writerSheet(i, "模板" + i).build();
//                // 分页去数据库查询数据 这里可以去数据库查询每一页的数据
//                List<Student> data = trainDtoList;
//                excelWriter.write(data, writeSheet);
//            }
//        } finally {
//            // 千万别忘记finish 会帮忙关闭流
//            if (excelWriter != null) {
//                excelWriter.finish();
//            }
//        }
//
//        // 方法7：将不同的对象多次写入到不同的sheet页中
//        try {
//            // 这里 指定文件
//            excelWriter = EasyExcel.write(fileName).build();
//            // 去调用写入,这里我调用了五次，实际使用时根据数据库分页的总的页数来。这里最终会写到5个sheet里面
//            for (int i = 0; i < 5; i++) {
//                // 每次都要创建writeSheet 这里注意必须指定sheetNo 而且sheetName必须不一样。这里注意Student.class 可以每次都变，我这里为了方便 所以用的同一个class 实际上可以一直变
//                WriteSheet writeSheet = EasyExcel.writerSheet(i, "模板" + i).head(Student.class).build();
//                // 分页去数据库查询数据 这里可以去数据库查询每一页的数据
//                List<Student> data = trainDtoList;
//                excelWriter.write(data, writeSheet);
//            }
//        } finally {
//            // 千万别忘记finish 会帮忙关闭流
//            if (excelWriter != null) {
//                excelWriter.finish();
//            }
//        }
//
//        // 方法8：图片导出
//        // 如果使用流 记得关闭
//        InputStream inputStream = null;
//        try {
//            List<ImageData> list = new ArrayList<ImageData>();
//            ImageData imageData = new ImageData();
//            list.add(imageData);
//            String imagePath = "F:\\image.jpg";
//            // 放入五种类型的图片 实际使用只要选一种即可
//            imageData.setFile(new File(imagePath));
//            imageData.setString(imagePath);// 字符串类型的，需要做转换处理
//            try {
//                imageData.setByteArray(FileUtils.readFileToByteArray(new File(imagePath)));
//                inputStream = FileUtils.openInputStream(new File(imagePath));
//                imageData.setInputStream(inputStream);
//                // 网上随便找的图片
//                imageData.setUrl(new URL(
//                        "https://shardingsphere.apache.org/document/current/img/shardingsphere-scope_cn.png"));
//            } catch (IOException e) {
//                e.printStackTrace();
//            }
//            EasyExcel.write(fileName, ImageData.class).sheet().doWrite(list);
//        } finally {
//            if (inputStream != null) {
//                try {
//                    inputStream.close();
//                } catch (IOException e) {
//                    e.printStackTrace();
//                }
//            }
//        }
//
//        // 方法9：根据模板写入。将templateFileName指定的文件内容写入到fileName指定的文件中，再续写trainDtoList数据到fileName指定的文件中
//        String templateFileName = "F:\\训练营列表1.xlsx";
//        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
//        EasyExcel.write(fileName, Student.class).withTemplate(templateFileName).sheet().doWrite(trainDtoList);
//
//        // 方法10：样式的设置，颜色，行高，字体等。两种设置方式，一种是基于注解的【Student】，一种是代码的方式【下面是代码的】
//        // 表头的策略
//        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
//        // 背景设置为红色
//        headWriteCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
//        WriteFont headWriteFont = new WriteFont();
//        headWriteFont.setFontHeightInPoints((short) 20);
//        headWriteCellStyle.setWriteFont(headWriteFont);
//        // 内容的策略
//        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
//        // 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND 不然无法显示背景颜色.头默认了 FillPatternType所以可以不指定
//        contentWriteCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
//        // 背景绿色
//        contentWriteCellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
//        WriteFont contentWriteFont = new WriteFont();
//        // 字体大小
//        contentWriteFont.setFontHeightInPoints((short) 20);
//        contentWriteCellStyle.setWriteFont(contentWriteFont);
//        // 这个策略是 头是头的样式 内容是内容的样式 其他的策略可以自己实现
//        HorizontalCellStyleStrategy horizontalCellStyleStrategy =
//                new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);
//        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
//        EasyExcel.write(fileName, Student.class).registerWriteHandler(horizontalCellStyleStrategy).sheet("模板")
//                .doWrite(trainDtoList);
//
//        // 方法11：合并单元格，两种方式，一种基于注解【Student】，另一种通过代码方式【下面是代码的】
//        // 每隔2行会合并 把eachColumn 设置成 3 也就是我们数据的长度，所以就第一列会合并。当然其他合并策略也可以自己写
//        LoopMergeStrategy loopMergeStrategy = new LoopMergeStrategy(2, 0);// 循环合并策略
//        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
//        EasyExcel.write(fileName, Student.class).registerWriteHandler(loopMergeStrategy).sheet("模板").doWrite(trainDtoList);

        // 方法12：使用table去写入
        // 这里 需要指定写用哪个class去写
        ExcelWriter excelWriter = null;
        try {
            excelWriter = EasyExcel.write(fileName, Student.class).build();
            // 把sheet设置为不需要头 不然会输出sheet的头 这样看起来第一个table 就有2个头了
            WriteSheet writeSheet = EasyExcel.writerSheet("模板").needHead(Boolean.FALSE).build();
            // 这里必须指定需要头，table 会继承sheet的配置，sheet配置了不需要，table 默认也是不需要
            WriteTable writeTable0 = EasyExcel.writerTable(0).needHead(Boolean.TRUE).build();
            WriteTable writeTable1 = EasyExcel.writerTable(1).needHead(Boolean.TRUE).build();
            // 第一次写入会创建头
            excelWriter.write(trainDtoList, writeSheet, writeTable0);
            // 第二次写如也会创建头，然后在第一次的后面写入数据
            excelWriter.write(trainDtoList, writeSheet, writeTable1);
        } finally {
            // 千万别忘记finish 会帮忙关闭流
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }

        // 方法13：可以动态的指定表头
        EasyExcel.write(fileName)
                // head()：这里放入动态头
                .head(head()).sheet("模板")
                // 当然这里数据也可以用 List<List<String>> 去传入
                .doWrite(trainDtoList);

        // 方法14：自动列宽(不太精确)
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        EasyExcel.write(fileName, Student.class)
                .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy()).sheet("模板").doWrite(trainDtoList);

        // 方法15：自定义拦截器
        EasyExcel.write(fileName, Student.class).registerWriteHandler(new CustomSheetWriteHandler())
                .registerWriteHandler(new CustomCellWriteHandler()).sheet("模板").doWrite(trainDtoList);

        // 方法15：自定义拦截器，增加批注
        EasyExcel.write(fileName, Student.class).registerWriteHandler(new CommentWriteHandler())
                .sheet("模板").doWrite(trainDtoList);

        // 方法16：不使用对象进行写
        EasyExcel.write(fileName).head(head()).sheet("模板").doWrite(dataList());
    }

    private static List<List<String>> head() {
        List<List<String>> list = new ArrayList<List<String>>();
        List<String> head0 = new ArrayList<String>();
        head0.add("字符串" + System.currentTimeMillis());
        List<String> head1 = new ArrayList<String>();
        head1.add("数字" + System.currentTimeMillis());
        List<String> head2 = new ArrayList<String>();
        head2.add("日期" + System.currentTimeMillis());
        list.add(head0);
        list.add(head1);
        list.add(head2);
        return list;
    }
    private static List<List<Object>> dataList() {
        List<List<Object>> list = new ArrayList<List<Object>>();
        for (int i = 0; i < 10; i++) {
            List<Object> data = new ArrayList<Object>();
            data.add("字符串" + i);
            data.add(new Date());
            data.add(0.56);
            list.add(data);
        }
        return list;
    }
}
