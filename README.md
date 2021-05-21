# springboot_poi

本项目是一个poi的demo测试

里面有poi的基本使用，基于DOM和SAX两种解析方式的使用
还有EasyExcel的常用示例


springboot集成poi,easypoi,java Excel三种方式操作excel

apache poi：XSSFWorkbook和HSSFWorkbook，SXSSFWorkbook的区别
用JavaPOI导出Excel时，我们会考虑到Excel版本及数据量的问题。针对不同的Excel版本，要采用不同的工具类。

HSSFWorkbook:是操作Excel2003以前（包括2003）的版本，扩展名是.xls；
    此种方式的局限就是导出的行数至多为65535行，超出65536条后系统就会报错。此方式因为行数不足七万行所以一般不会发生内存不足的情况（OOM）

XSSFWorkbook:是操作Excel2007的版本，扩展名是.xlsx；
    这种形式的出现是为了突破HSSFWorkbook的65535行局限。其对应的是excel2007(1048576行，16384列)扩展名为“.xlsx”，最多可以导出104万行，不过这样就伴随着一个问题---OOM内存溢出，原因是你所创建的book sheet row cell等此时是存在内存的并没有持久化。

对于不同版本的EXCEL文档要使用不同的工具类，如果使用错了，会提示如下错误信息。
    org.apache.poi.openxml4j.exceptions.InvalidOperationException
    org.apache.poi.poifs.filesystem.OfficeXmlFileException

从POI 3.8版本开始，提供了一种基于XSSF的低内存占用的API----SXSSF
    当数据量超出65536条后，在使用HSSFWorkbook或XSSFWorkbook，程序会报OutOfMemoryError：Javaheap space;内存溢出错误。这时应该用SXSSFworkbook。
    注意：针对 SXSSF Beta 3.8下，会有临时文件产生，比如：
    poi-sxssf-sheet4654655121378979321.xml
    文件位置：java.io.tmpdir这个环境变量下的位置
    Windows 7下是C:\Users\xxxxxAppData\Local\Temp
    Linux下是 /var/tmp/
    要根据实际情况，看是否删除这些临时文件

与XSSF的对比
    在一个时间点上，只可以访问一定数量的数据
    不再支持Sheet.clone()
    不再支持公式的求值
    在使用Excel模板下载数据时将不能动态改变表头，因为这种方式已经提前把excel写到硬盘的了就不能再改了

操作excel，csv，pdf，以及压缩文件