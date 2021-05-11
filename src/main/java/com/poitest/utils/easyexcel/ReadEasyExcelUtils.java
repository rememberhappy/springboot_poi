package com.poitest.utils.easyexcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.enums.CellExtraTypeEnum;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.fastjson.JSON;
import com.poitest.domain.Student;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.Map;

/**
 * 关于读的 EasyExcel 几种写法
 */
public class ReadEasyExcelUtils {
    private static final Logger LOGGER = LoggerFactory.getLogger(ReadEasyExcelUtils.class);

    public static void main(String[] args) {
        // 有个很重要的点 DataListener 不能被spring管理，要每次读取excel都要new,然后里面用到spring可以构造方法传进去
        // 在实体中加注解@ExcelProperty("姓名")可以将excel中的列与实体字段保持一致
        // 写法1：默认读取第一个sheet页
        String fileName = "C:\\Users\\wd\\Desktop\\训练营列表.xlsx";
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        EasyExcel.read(fileName, Student.class, new DataListener())
                // 这里注意 我们也可以registerConverter来指定自定义转换器， 但是这个转换变成全局了， 所有java为string,excel为string的都会用这个转换器。
                // 如果就想单个字段使用请使用@ExcelProperty 指定converter：@ExcelProperty(converter = CustomStringStringConverter.class)
//                 .registerConverter(new CustomStringStringConverter())// 设置全局的自定义转换器
                // 读取sheet
                .sheet()
                // 这里可以设置1，因为头就是一行。如果多行头，可以设置其他值。不传入也可以，因为默认会根据DemoData 来解析，他没有指定头，也就是默认1行
//                .headRowNumber(1)
                .doRead();

        // 写法2：指定要读取的sheet页
        ExcelReader excelReader = null;
        try {
            excelReader = EasyExcel.read(fileName, Student.class, new DataListener()).build();
            ReadSheet readSheet = EasyExcel.readSheet(0).build();
            excelReader.read(readSheet);
        } finally {
            if (excelReader != null) {
                // 这里千万别忘记关闭，读的时候会创建临时文件，到时磁盘会崩的
                excelReader.finish();
            }
        }

        // 写法3：读取全部sheet
        // 这里需要注意 DemoDataListener的doAfterAllAnalysed 会在每个sheet读取完毕后调用一次。然后所有sheet都会往同一个DemoDataListener里面写
        EasyExcel.read(fileName, Student.class, new DataListener()).doReadAll();
        // 读取部分sheet
//        ExcelReader excelReader = null;
        try {
            excelReader = EasyExcel.read(fileName).build();
            // 这里为了简单 所以注册了 同样的head 和Listener 自己使用功能必须不同的Listener
            ReadSheet readSheet1 =
                    EasyExcel.readSheet(0).head(Student.class).registerReadListener(new DataListener()).build();
            ReadSheet readSheet2 =
                    EasyExcel.readSheet(1).head(Student.class).registerReadListener(new DataListener()).build();
            // 这里注意 一定要把sheet1 sheet2 一起传进去，不然有个问题就是03版的excel 会读取多次，浪费性能
            excelReader.read(readSheet1, readSheet2);
        } finally {
            if (excelReader != null) {
                // 这里千万别忘记关闭，读的时候会创建临时文件，到时磁盘会崩的
                excelReader.finish();
            }
        }

        // 写法4：同步的返回，不推荐使用，如果数据量大会把数据放到内存里面
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 同步读取会自动finish
        List<Student> list = EasyExcel.read(fileName).head(Student.class).sheet().doReadSync();
        for (Student data : list) {
            LOGGER.info("读取到数据:{}", JSON.toJSONString(data));
        }
        // 这里 也可以不指定class，返回一个list，然后读取第一个sheet 同步读取会自动finish
        List<Map<Integer, String>> listMap = EasyExcel.read(fileName).sheet().doReadSync();
        for (Map<Integer, String> data : listMap) {
            // 返回每条数据的键值对 表示所在的列 和所在列的值
            LOGGER.info("读取到数据:{}", JSON.toJSONString(data));
        }

        // 写法5：不创建对象的读，使用List<Map<>>来接收读取的对象
        // 这里 只要，然后读取第一个sheet 同步读取会自动finish
        EasyExcel.read(fileName, new NoModelDataListener()).sheet().doRead();

        // 写法6：在controller中能直接获取到file对象，可以直接使用file对象获取到文件输入流
//        EasyExcel.read(FileInputStream, new NoModelDataListener()).sheet().doRead();

        // 写法6：读取额外信息（批注、超链接、合并单元格信息读取）
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet
        EasyExcel.read(fileName, Student.class, new DataListener())
                // 需要读取批注 默认不读取
                .extraRead(CellExtraTypeEnum.COMMENT)
                // 需要读取超链接 默认不读取
                .extraRead(CellExtraTypeEnum.HYPERLINK)
                // 需要读取合并单元格信息 默认不读取
                .extraRead(CellExtraTypeEnum.MERGE)
                .sheet().doRead();


    }
}
