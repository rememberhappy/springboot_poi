package com.poitest.controller;

import com.poitest.domain.Student;
import com.poitest.utils.poi.ExcelHSSFWorkbookUtil;
import com.poitest.utils.poi.ExcelSXSSFWorkbookUtil;
import com.poitest.utils.poi.ExcelUtil;
import com.poitest.utils.poi.ExcelXSSFWorkbookUtil;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.List;

/**
 * 导出的控制层
 * 使用了 apache poi：XSSFWorkbook和HSSFWorkbook，SXSSFWorkbook三种方式
 */
@RestController
@RequestMapping("export")
public class ExportExcelController {

    /**
     * 使用poi导出excel，根据文件名称自动判断选用的类型【XSSFWorkbook/HSSFWorkbook/SXSSFWorkbook】
     *
     * @param httpServletResponse
     * @return
     */
    @RequestMapping(value = "/studentExcel", method = RequestMethod.GET)
    public String exportStudentExcel(HttpServletResponse httpServletResponse) {
        // 处理需要导出的数据
        List<Student> trainDtoList = new ArrayList<Student>();
        trainDtoList.add(new Student("张三", "13", "北京3", "333333"));
        trainDtoList.add(new Student("李四", "14", "北京4", "444444"));
        trainDtoList.add(new Student("王五", "15", "北京5", "555555"));
        trainDtoList.add(new Student("赵六", "16", "北京6", "666666"));
        //设置 excel
        try {
            // 设置响应请求中的媒体类型信息，本示例中设置的是：二进制流数据
            httpServletResponse.setContentType(MediaType.APPLICATION_OCTET_STREAM_VALUE);
            // 设置响应请求中的要处理的文件的名称以及编码格式
            // 【setHeader方法是用新值去替换旧值】【"Content-Disposition“可以控制用户请求所得的内容存为一个文件的时候提供一个默认的文件名，文件直接在浏览器上显示或者在访问时弹出文件下载对话框】
            httpServletResponse.setHeader("Content-Disposition",
                    "attachment;filename=" + new String(("训练营列表.xlsx").getBytes(), "iso-8859-1"));

            String[] columnNames = new String[]{"姓名", "年龄", "住址", "手机号"};
            String[] propertyNames = new String[]{"name", "age", "address", "phone"};
            // 工具类中会使用httpServletResponse响应请求中的输出流来输出要下载的文件内容
            ExcelUtil.exportExcel(httpServletResponse, trainDtoList, Student.class, columnNames, propertyNames);
        } catch (Exception e) {
            return "导出失败";
        }
        return "导出成功";
    }

    /**
     * 使用poi导出excel【XSSFWorkbook】
     *
     * @param httpServletResponse
     * @return
     */
    @RequestMapping(value = "/studentXSSExcel", method = RequestMethod.GET)
    public String exportStudentXSSExcel(HttpServletResponse httpServletResponse) {
        // 处理需要导出的数据
        List<Student> trainDtoList = new ArrayList<Student>();
        trainDtoList.add(new Student("张三", "13", "北京3", "333333"));
        trainDtoList.add(new Student("李四", "14", "北京4", "444444"));
        trainDtoList.add(new Student("王五", "15", "北京5", "555555"));
        trainDtoList.add(new Student("赵六", "16", "北京6", "666666"));
        //设置 excel
        try {
            // 设置响应请求中的媒体类型信息，本示例中设置的是：二进制流数据
            httpServletResponse.setContentType(MediaType.APPLICATION_OCTET_STREAM_VALUE);
            // 设置响应请求中的要处理的文件的名称以及编码格式
            // 【setHeader方法是用新值去替换旧值】【"Content-Disposition“可以控制用户请求所得的内容存为一个文件的时候提供一个默认的文件名，文件直接在浏览器上显示或者在访问时弹出文件下载对话框】
            httpServletResponse.setHeader("Content-Disposition",
                    "attachment;filename=" + new String(("训练营列表.xlsx").getBytes(), "iso-8859-1"));

            String[] columnNames = new String[]{"姓名", "年龄", "住址", "手机号"};
            String[] propertyNames = new String[]{"name", "age", "address", "phone"};
            // 工具类中会使用httpServletResponse响应请求中的输出流来输出要下载的文件内容
            ExcelXSSFWorkbookUtil.exportObjectsToExcel(httpServletResponse.getOutputStream(), trainDtoList, Student.class, columnNames, propertyNames);
        } catch (Exception e) {
            return "导出失败";
        }
        return "导出成功";
    }

    /**
     * 使用poi导出excel【HSSFWorkbook】
     *
     * @param httpServletResponse
     * @return
     */
    @RequestMapping(value = "/studentHSSExcel", method = RequestMethod.GET)
    public String exportStudentHSSExcel(HttpServletResponse httpServletResponse) {
        // 处理需要导出的数据
        List<Student> trainDtoList = new ArrayList<Student>();
        trainDtoList.add(new Student("张三", "13", "北京3", "333333"));
        trainDtoList.add(new Student("李四", "14", "北京4", "444444"));
        trainDtoList.add(new Student("王五", "15", "北京5", "555555"));
        trainDtoList.add(new Student("赵六", "16", "北京6", "666666"));
        //设置 excel
        try {
            // 设置响应请求中的媒体类型信息，本示例中设置的是：二进制流数据
            httpServletResponse.setContentType(MediaType.APPLICATION_OCTET_STREAM_VALUE);
            // 设置响应请求中的要处理的文件的名称以及编码格式
            // 【setHeader方法是用新值去替换旧值】【"Content-Disposition“可以控制用户请求所得的内容存为一个文件的时候提供一个默认的文件名，文件直接在浏览器上显示或者在访问时弹出文件下载对话框】
            httpServletResponse.setHeader("Content-Disposition",
                    "attachment;filename=" + new String(("训练营列表.xls").getBytes(), "iso-8859-1"));

            String[] columnNames = new String[]{"姓名", "年龄", "住址", "手机号"};
            String[] propertyNames = new String[]{"name", "age", "address", "phone"};
            // 工具类中会使用httpServletResponse响应请求中的输出流来输出要下载的文件内容
            ExcelHSSFWorkbookUtil.exportObjectsToExcel(httpServletResponse.getOutputStream(), trainDtoList, Student.class, columnNames, propertyNames);
        } catch (Exception e) {
            return "导出失败";
        }
        return "导出成功";
    }

    /**
     * 使用poi导出excel【SXSSFWorkbook】
     *
     * @param httpServletResponse
     * @return
     */
    @RequestMapping(value = "/studentSXSSExcel", method = RequestMethod.GET)
    public String exportStudentSXSSExcel(HttpServletResponse httpServletResponse) {
        // 处理需要导出的数据
        List<Student> trainDtoList = new ArrayList<Student>();
//        trainDtoList.add(new Student("张三", "13", "北京3", "333333"));
        // 设置大批量数据，一万条
        for (int i = 1; i < 5000000; i++) {
            trainDtoList.add(new Student("张三" + i, "13" + i, "北京3" + i, "333333" + i));
        }
        //设置 excel
        try {
            // 设置响应请求中的媒体类型信息，本示例中设置的是：二进制流数据
            httpServletResponse.setContentType(MediaType.APPLICATION_OCTET_STREAM_VALUE);
            // 设置响应请求中的要处理的文件的名称以及编码格式
            // 【setHeader方法是用新值去替换旧值】【"Content-Disposition“可以控制用户请求所得的内容存为一个文件的时候提供一个默认的文件名，文件直接在浏览器上显示或者在访问时弹出文件下载对话框】
            httpServletResponse.setHeader("Content-Disposition",
                    "attachment;filename=" + new String(("训练营列表123.xlsx").getBytes(), "iso-8859-1"));

            String[] columnNames = new String[]{"姓名", "年龄", "住址", "手机号"};
            String[] propertyNames = new String[]{"name", "age", "address", "phone"};
            // 工具类中会使用httpServletResponse响应请求中的输出流来输出要下载的文件内容
            ExcelSXSSFWorkbookUtil.exportObjectsToExcel(httpServletResponse.getOutputStream(), trainDtoList, Student.class, columnNames, propertyNames);
        } catch (Exception e) {
            return "导出失败";
        }
        return "导出成功";
    }
}
