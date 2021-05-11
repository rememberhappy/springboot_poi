package com.poitest.controller;

import com.alibaba.excel.EasyExcel;
import com.alibaba.fastjson.JSON;
import com.poitest.domain.Student;
import com.poitest.utils.easyexcel.DataListener;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
public class EasyExcelController {
    @PostMapping("upload")
    public String upload(MultipartFile file) throws IOException {
        EasyExcel.read(file.getInputStream(), Student.class, new DataListener()).sheet().doRead();
        return "success";
    }

    @GetMapping("download")
    public void download(HttpServletResponse response) throws IOException {
        // 处理需要导出的数据
        List<Student> trainDtoList = new ArrayList<Student>();
        // 设置大批量数据，一万条
        for (int i = 1; i < 50; i++) {
            trainDtoList.add(new Student("张三" + i, "13" + i, "北京3" + i, "333333" + i));
        }

        // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
        String fileName = URLEncoder.encode("测试", "UTF-8").replaceAll("\\+", "%20");
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
        EasyExcel.write(response.getOutputStream(), Student.class).sheet("模板").doWrite(trainDtoList);
    }
    /**
     * 文件下载并且失败的时候返回json（默认失败了会返回一个有部分数据的Excel）
     *
     * @since 2.1.1
     */
    @GetMapping("downloadFailedUsingJson")
    public void downloadFailedUsingJson(HttpServletResponse response) throws IOException {
        // 处理需要导出的数据
        List<Student> trainDtoList = new ArrayList<Student>();
        // 设置大批量数据，一万条
        for (int i = 1; i < 50; i++) {
            trainDtoList.add(new Student("张三" + i, "13" + i, "北京3" + i, "333333" + i));
        }

        // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        try {
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setCharacterEncoding("utf-8");
            // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
            String fileName = URLEncoder.encode("测试", "UTF-8").replaceAll("\\+", "%20");
            response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
            // 这里需要设置不关闭流
            EasyExcel.write(response.getOutputStream(), Student.class).autoCloseStream(Boolean.FALSE).sheet("模板")
                    .doWrite(trainDtoList);
        } catch (Exception e) {
            // 重置response
            response.reset();
            response.setContentType("application/json");
            response.setCharacterEncoding("utf-8");
            Map<String, String> map = new HashMap<String, String>();
            map.put("status", "failure");
            map.put("message", "下载文件失败" + e.getMessage());
            response.getWriter().println(JSON.toJSONString(map));
        }
    }
}
