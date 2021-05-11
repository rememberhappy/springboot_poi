package com.poitest.utils.poi.sax;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.springframework.util.StringUtils;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 通过SAX模式 读取EXCEL辅助类
 * 解决思路：
 *      通过继承DefaultHandler类，重写process()，startElement()，characters()，endElement()这四个方法。
 *      process()方式主要是遍历所有的sheet，并依次调用startElement()、characters()方法、endElement()这三个方法。
 *      startElement()用于设定单元格的数字类型（如日期、数字、字符串等等）。
 *      characters()用于获取该单元格对应的索引值或是内容值（如果单元格类型是字符串、INLINESTR、数字、日期则获取的是索引值；其他如布尔值、错误、公式则获取的是内容值）。
 *      endElement()根据startElement()的单元格数字类型和characters()的索引值或内容值，最终得出单元格的内容值，并打印出来。对于日期的处理详见代码
 */
public class ExcelSAXWithDefaultHandler extends DefaultHandler {

    private ExcelReadDataDelegated excelReadDataDelegated;

    public ExcelReadDataDelegated getExcelReadDataDelegated() {
        return excelReadDataDelegated;
    }

    public void ExcelSAXWithDefaultHandler(ExcelReadDataDelegated excelReadDataDelegated) {
        this.excelReadDataDelegated = excelReadDataDelegated;
    }

    public ExcelSAXWithDefaultHandler(ExcelReadDataDelegated excelReadDataDelegated) {
        this.excelReadDataDelegated = excelReadDataDelegated;
    }

    /**
     * 单元格中的数据可能的数据类型
     */
    enum CellDataType {
        BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER, DATE, NULL
    }

    /**
     * 共享字符串表
     */
    private SharedStringsTable sst;

    /**
     * 上一次的索引值
     */
    private String lastIndex;

    /**
     * 文件的绝对路径
     */
    private String filePath = "";

    /**
     * 工作表索引
     */
    private int sheetIndex = 0;

    /**
     * sheet名
     */
    private String sheetName = "";

    /**
     * 总行数
     */
    private int totalRows = 0;

    /**
     * 一行内cell集合
     */
    private List<String> cellList = new ArrayList<String>();

    /**
     * 判断整行是否为空行的标记
     */
    private boolean flag = false;

    /**
     * 当前行
     */
    private int curRow = 1;

    /**
     * 上一行id, 判断空行
     */
    private int lastRowid = 0;

    /**
     * 当前列
     */
    private int curCol = 0;

    /**
     * T元素标识
     */
    private boolean isTElement;

    /**
     * 异常信息，如果为空则表示没有异常
     */
    private String exceptionMessage;

    /**
     * 单元格数据类型，默认为字符串类型
     */
    private CellDataType nextDataType = CellDataType.SSTINDEX;

    private final DataFormatter formatter = new DataFormatter();

    /**
     * 单元格日期格式的索引
     */
    private short formatIndex;

    /**
     * 日期格式字符串
     */
    private String formatString;

    /**
     * 定义前一个元素和当前元素的位置，用来计算其中空的单元格数量，如A6和A8等
     */
    private String preRef = null, ref = null;

    /**
     * 定义该文档一行最大的单元格数，用来补全一行最后可能缺失的单元格
     */
    private String maxRef = null;

    /**
     * 单元格
     */
    private StylesTable stylesTable;

    /**
     * 总行号
     */
    private Integer totalRowCount;

    /**
     * 通过文件名称 遍历工作簿中所有的电子表格
     * 并缓存在mySheetList中
     *
     * @param filename
     * @throws Exception
     */
    public int process(String filename) throws Exception {
        filePath = filename;
        // 加载Excel的核心方法。打开具有读/写权限的程序包，XSSFWorkbook地层中就使用了此类进行打开整个文件
        OPCPackage pkg = OPCPackage.open(filename);// 文件名称
        // 文件流对内存依赖极大，所以实际应用时，如果只能获取文件流的话，建议先将文件通过流拷贝到本地，然后再使用解析工具类
//        OPCPackage pkg = OPCPackage.open(FileInputStream流);// 文件流
        // 通过此类，可以轻松获取OOXML .xlsx文件的各个部分，适用于低内存sax解析或类似操作。
        XSSFReader xssfReader = new XSSFReader(pkg);
        stylesTable = xssfReader.getStylesTable();
        // 打开共享字符串表，对其进行解析，然后返回用于处理共享字符串的便捷对象
        SharedStringsTable sst = xssfReader.getSharedStringsTable();
        // 使用 org.apache.xerces.parsers.SAXParser 进行解析，获取XML解析器，使用XMLReader进行读取xml文件
        // XmlReader读取Xml，需要通过Read()实例方法，不断读取Xml文档中的声明，节点开始，节点内容，节点结束，以及空白等等，直到文档结束，Read()方法返回false。
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        this.sst = sst;
        parser.setContentHandler(this);
        // 获取所有的sheet页
        XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        while (sheets.hasNext()) { //遍历sheet
            curRow = 1; //标记初始行为第一行
            sheetIndex++;
            InputStream sheet = sheets.next(); //sheets.next()和sheets.getSheetName()不能换位置，否则sheetName报错
            sheetName = sheets.getSheetName();
            // InputSource：XML实体的单个输入源。SAX 在解析时从 InputSource 抽取这些信息，从而能够解析外部实体以及其它特定于文档来源的资源。
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource); //解析excel的每条记录，在这个过程中startElement()、characters()、endElement()这三个函数会依次执行
            sheet.close();
        }
        return totalRows; //返回该excel文件的总行数，不包括首列和空行
    }

    public int process(InputStream inputStream, String filename) throws Exception {
        // SAX是处理大文件的解析模式，通过 OPCPackage.open(FileInputStream流)对内存依赖极大，一般先将文件通过流拷贝到本地，然后再使用解析工具类
        File file = new File("F:\\temp" + File.separator + filename);
        int index = filename.lastIndexOf(".");
        String filenamePre = filename.substring(0, index);// 前缀
        String filenameSuf = filename.substring(index + 1);// 后缀
        // 循环判断创建的这个临时文件是否存在
        for (int i = 1; file.exists(); i++) {
            file = new File("F:\\temp" + File.separator + filenamePre + "(" + i + ")." + filenameSuf);
        }
        // 绝对路径
        String absolutePath = file.getAbsolutePath();
        FileOutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream(file);
            byte[] buf = new byte[1024];
            int len = 0;
            while ((len = inputStream.read(buf)) >= 0) {
                fileOutputStream.write(buf, 0, len);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (inputStream != null) inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            try {
                if (fileOutputStream != null) fileOutputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        int process = process(absolutePath);
        // 将拷贝到本地的临时文件删除
        file.delete();
        return process;
    }

    /**
     * 第一个执行
     */
    @Override
    public void startDocument() {
        System.out.println("开始解析！");
    }

    /**
     * 第二个执行
     *
     * @param uri
     * @param localName
     * @param name
     * @param attributes
     * @throws SAXException
     */
    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
        System.out.println("对" + name + "开始解析！");
        System.out.println("uri：" + uri + " localName:" + localName + " 元素名：" + name + " attributes：" + attributes.toString());
        // 获取总行号，获取excel中有文字填充的区域，在A1:D5这个区域中取D5中的5，然后减去1【除去表头】后就是数据的行数
        if ("dimension".equals(name)) {
            String dimensionStr = attributes.getValue("ref");
            System.out.println("获取excel中的数据填充区域：" + dimensionStr);// A1:D5
            totalRowCount = Integer.parseInt(dimensionStr.substring(dimensionStr.indexOf(":") + 2)) - 1;
            System.out.println("计算后的总行数：" + totalRowCount);
        }
        // row => 行
        if ("row".equals(name)) {
            System.out.print("行解析开始  ");
            String rowNum = attributes.getValue("r");//获取单元格的位置，如A1,B1
            System.out.print("，rowNum：" + rowNum);
            //判断空行
            if (lastRowid > 0) {
                System.out.print("，不是空行：" + rowNum);
                //与上一行相差2, 说明中间有空行
                int gap = Integer.parseInt(rowNum) - lastRowid;
                if (gap > 1) {
                    System.out.print("，上一行是空行：" + rowNum);
                    gap -= 1;
                    while (gap > 0) {
                        //container.add(new ArrayList<>());
                        gap--;
                    }
                }
            }
            // 将上一行单元格的位置改成当前行的单元格的位置
            lastRowid = Integer.parseInt(attributes.getValue("r"));//获取单元格的位置，如A1,B1
            System.out.println();
        }
        //c => 单元格
        if ("c".equals(name)) {
            System.out.print("单元格解析开始  ");
            //前一个单元格为空
            if (preRef == null) {
                // 获取当前单元格的位置 给上一个单元格位置 赋值
                preRef = attributes.getValue("r");//获取单元格的位置，如A1,B1
            } else {
                // 前一个元素的位置等于当前元素的位置【这个时候，当前元素没有指定为最新的当前元素，实际的值是在这个单元格处理之前的值】
                preRef = ref;
            }

            //当前单元格的位置
            ref = attributes.getValue("r");//获取单元格的位置，如A1,B1
            //设定单元格类型
            this.setNextDataType(attributes);
        }

        //当元素为t时
        if ("t".equals(name)) {
            isTElement = true;
        } else {
            isTElement = false;
        }

        // sheet => sheet页
        if (name.equals("sheet")) {
            String sheetName = "sheet";
            Map<String, String> stringStringHashMap = new HashMap<>();
            if (!StringUtils.isEmpty(sheetName) && attributes.getValue("name").equals(sheetName)) {
                stringStringHashMap.put("r:id", attributes.getValue("r:id"));
            }
            System.out.println(" sheet页名称 =" + attributes.getValue("name"));
            System.out.println(" r:id =" + attributes.getValue("r:id"));
        }

        //置空
        lastIndex = "";
    }


    /**
     * 第三个执行【执行的是单元格解析的时候才会第三个执行，要不然不执行】
     * 得到单元格对应的索引值或是内容值
     * 如果单元格类型是字符串、INLINESTR、数字、日期，lastIndex则是索引值
     * 如果单元格类型是布尔值、错误、公式，lastIndex则是内容值
     *
     * @param ch
     * @param start
     * @param length
     * @throws SAXException
     */
    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        String s = new String(ch, start, length);
        System.out.println("得到单元格对应的索引值或是内容值：" + s);
        lastIndex += s;
//        lastIndex += new String(ch, start, length);
    }


    /**
     * 第四个执行，单元格解析完成时触发【格式类似<xx/>时才会触发】
     *
     * @param uri
     * @param localName
     * @param name
     * @throws SAXException
     */
    @Override
    public void endElement(String uri, String localName, String name) throws SAXException {
        System.out.println("对" + name + "的解析完成！");
        //t元素也包含字符串
        if (isTElement) {//这个程序没经过
            //将单元格内容加入rowlist中，在这之前先去掉字符串前后的空白符
            String value = lastIndex.trim();
            cellList.add(curCol, value);
            curCol++;
            isTElement = false;
            //如果里面某个单元格含有值，则标识该行不为空行
            if (value != null && !"".equals(value)) {
                flag = true;
            }
        } else if ("v".equals(name)) {
            //v => 单元格的值，如果单元格是字符串，则v标签的值为该字符串在SST中的索引
            String value = this.getDataValue(lastIndex.trim(), "");//根据索引值获取对应的单元格值
            //补全单元格之间的空单元格
            if (!ref.equals(preRef)) {
                int len = countNullCell(ref, preRef);
                for (int i = 0; i < len; i++) {
                    cellList.add(curCol, "");
                    curCol++;
                }
            }
            cellList.add(curCol, value);
            curCol++;
            //如果里面某个单元格含有值，则标识该行不为空行
            if (value != null && !"".equals(value)) {
                flag = true;
            }
        } else {
            //如果标签名称为row，这说明已到行尾，调用optRows()方法
            if ("row".equals(name)) {
                //默认第一行为表头，以该行单元格数目为最大数目
                if (curRow == 1) {
                    maxRef = ref;
                }
                //补全一行尾部可能缺失的单元格
                if (maxRef != null) {
                    int len = countNullCell(maxRef, ref);
                    for (int i = 0; i <= len; i++) {
                        cellList.add(curCol, "");
                        curCol++;
                    }
                }

                // 处理除了表头一行的数据
                if (flag && curRow != 1) { //该行不为空行且该行不是第一行，则发送（第一行为列名，不需要）
                    // 调用excel读数据委托类进行读取插入操作
                    System.out.println("总行数为：" + totalRowCount + " 行号为：" + curRow + " 数据：" + cellList);
                    excelReadDataDelegated.readExcelDate(sheetIndex, totalRowCount, curRow, cellList);
                    totalRows++;
                }
                cellList.clear();
                curRow++;
                curCol = 0;
                preRef = null;
                ref = null;
                flag = false;
            }
        }
    }

    /**
     * 最后一个执行
     * 补充最后的数据处理，此步很重要
     */
    public void endDocument() {
        System.out.println("解析完成！");
        //TODO 此处为最后的文件输出地方
        //TODO 如果此处不处理，可能会丢失最后的一行数据，如果是自己写逻辑按照行处理的话
        //TODO 最后一行一定要处理
    }

    /**
     * 处理数据类型
     *
     * @param attributes
     */
    private void setNextDataType(Attributes attributes) {
        nextDataType = CellDataType.NUMBER; //cellType为空，则表示该单元格类型为数字
        formatIndex = -1;
        formatString = null;
        String cellType = attributes.getValue("t"); //单元格类型
        String cellValue = attributes.getValue("v"); //单元格的值
        String cellStyleStr = attributes.getValue("s"); //单元格样式
        String columnData = attributes.getValue("r"); //获取单元格的位置，如A1,B1

        if ("b".equals(cellType)) { //处理布尔值
            nextDataType = CellDataType.BOOL;
        } else if ("e".equals(cellType)) {  //处理错误
            nextDataType = CellDataType.ERROR;
        } else if ("inlineStr".equals(cellType)) {
            nextDataType = CellDataType.INLINESTR;
        } else if ("s".equals(cellType)) { //处理字符串
            nextDataType = CellDataType.SSTINDEX;
        } else if ("str".equals(cellType)) {
            nextDataType = CellDataType.FORMULA;
        }

        if (cellStyleStr != null) { //处理日期
            int styleIndex = Integer.parseInt(cellStyleStr);
            XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
            formatIndex = style.getDataFormat();
            formatString = style.getDataFormatString();
            if (formatString.contains("m/d/yy") || formatString.contains("yyyy/mm/dd") || formatString.contains("yyyy/m/d")) {
                nextDataType = CellDataType.DATE;
                formatString = "yyyy-MM-dd hh:mm:ss";
            }
            if (formatString == null) {
                nextDataType = CellDataType.NULL;
                formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
            }
        }
    }

    /**
     * 对解析出来的数据进行类型处理
     *
     * @param value   单元格的值，
     *                value代表解析：BOOL的为0或1， ERROR的为内容值，FORMULA的为内容值，INLINESTR的为索引值需转换为内容值，
     *                SSTINDEX的为索引值需转换为内容值， NUMBER为内容值，DATE为内容值
     * @param thisStr 一个空字符串
     * @return
     */
    @SuppressWarnings("deprecation")
    private String getDataValue(String value, String thisStr) {
        switch (nextDataType) {
            // 这几个的顺序不能随便交换，交换了很可能会导致数据错误
            case BOOL: //布尔值
                char first = value.charAt(0);
                thisStr = first == '0' ? "FALSE" : "TRUE";
                break;
            case ERROR: //错误
                thisStr = "\"ERROR:" + value.toString() + '"';
                break;
            case FORMULA: //公式
                thisStr = '"' + value.toString() + '"';
                break;
            case INLINESTR:
                XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
                thisStr = rtsi.toString();
                rtsi = null;
                break;
            case SSTINDEX: //字符串
                String sstIndex = value.toString();
                try {
                    int idx = Integer.parseInt(sstIndex);
                    XSSFRichTextString rtss = new XSSFRichTextString(sst.getEntryAt(idx));//根据idx索引值获取内容值
                    thisStr = rtss.toString();
                    rtss = null;
                } catch (NumberFormatException ex) {
                    thisStr = value.toString();
                }
                break;
            case NUMBER: //数字
                if (formatString != null) {
                    thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString).trim();
                } else {
                    thisStr = value;
                }
                thisStr = thisStr.replace("_", "").trim();
                break;
            case DATE: //日期
                thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString);
                // 对日期字符串作特殊处理，去掉T
                thisStr = thisStr.replace("T", " ");
                break;
            default:
                thisStr = " ";
                break;
        }
        return thisStr;
    }

    private int countNullCell(String ref, String preRef) {
        //excel2007最大行数是1048576，最大列数是16384，最后一列列名是XFD
        String xfd = ref.replaceAll("\\d+", "");
        String xfd_1 = preRef.replaceAll("\\d+", "");

        xfd = fillChar(xfd, 3, '@', true);
        xfd_1 = fillChar(xfd_1, 3, '@', true);

        char[] letter = xfd.toCharArray();
        char[] letter_1 = xfd_1.toCharArray();
        int res = (letter[0] - letter_1[0]) * 26 * 26 + (letter[1] - letter_1[1]) * 26 + (letter[2] - letter_1[2]);
        return res - 1;
    }

    private String fillChar(String str, int len, char let, boolean isPre) {
        int len_1 = str.length();
        if (len_1 < len) {
            if (isPre) {
                for (int i = 0; i < (len - len_1); i++) {
                    str = let + str;
                }
            } else {
                for (int i = 0; i < (len - len_1); i++) {
                    str = str + let;
                }
            }
        }
        return str;
    }
}
