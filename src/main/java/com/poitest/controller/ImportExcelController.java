package com.poitest.controller;

import com.poitest.utils.poi.ExcelUtil;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.Map;

/**
 * 导入的控制层
 * 使用了 apache poi：XSSFWorkbook和HSSFWorkbook，SXSSFWorkbook三种方式
 */
@RestController
@RequestMapping("import")
public class ImportExcelController {

    @PostMapping("/upload")
    public void upload(@RequestParam("file") MultipartFile file) throws IOException {
        if (file.isEmpty()) {
            throw new RuntimeException("上传失败，请选择文件");
        }
        // 获取文件名称
        String fileName = file.getOriginalFilename();
        // 获取系统参数，绝对地址
        String dir = System.getProperty("user.dir");
        String destFileName = dir + File.separator + "uploadedfiles_" + fileName;
        System.out.println("上传到文件服务器的地址：" + destFileName);
        Map<String, Object> excelInfo = new ExcelUtil().importExcelDOM(fileName, file.getInputStream(), true);
        System.out.println(excelInfo.get("data"));
        System.out.println(excelInfo.get("columnstypes"));
        System.out.println(excelInfo.get("beanList"));
    }

    @PostMapping("/import1")
    public void import1(@RequestParam("file") MultipartFile file) throws Exception {
        if (file.isEmpty()) {
            throw new RuntimeException("上传失败，请选择文件");
        }
        // 获取文件名称
        String fileName = file.getOriginalFilename();
        // 获取系统参数，绝对地址
        String dir = System.getProperty("user.dir");
        String destFileName = dir + File.separator + "uploadedfiles_" + fileName;
        System.out.println("上传到文件服务器的地址：" + destFileName);
        new ExcelUtil().importExcelSAX(fileName, file.getInputStream(), true);
    }
}
