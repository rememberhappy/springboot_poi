package com.poitest.controller;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;

/**
 * 下载excel模板
 *
 * @Author zhangdj
 * @Date 2021/9/26:14:57
 */
@RestController
public class DownloadExcelTemplate {

    // 日志
    private static final Logger logger = LoggerFactory.getLogger(DownloadExcelTemplate.class);

    @RequestMapping("/downloadJunmin")
    public void download(HttpServletResponse response) {
        try {
            response.setContentType("application/vnd.ms-excel;charset=utf-8");
            response.setHeader("Content-Disposition", "attachment;filename=" + new String(("军民用户模板.xlsx").getBytes(), StandardCharsets.ISO_8859_1));
            OutputStream os = response.getOutputStream();
            InputStream is = new BufferedInputStream(DownloadExcelTemplate.class.getClassLoader().getResourceAsStream("template/用户模板.xlsx"));
            byte[] buffer = new byte[1024];
            while (is.read(buffer) != -1) {
                os.write(buffer);
            }
            os.flush();
            os.close();
            is.close();
        } catch (Exception e) {
            logger.error("下载军民用户模板失败", e);
        }
    }

    @RequestMapping("/downloadExcel")
    public void downloadExcel(HttpServletResponse response, HttpServletRequest request) {
        // 直接下载路径下的文件模板（这种方式貌似在SpringCloud和Springboot中，打包成JAR包时，无法读取到指定路径下面的文件，不知道记错没，你们可以自己尝试下！！！）
        try {
            //获取要下载的模板名称
            String fileName = "用户模板.xlsx";
            //设置要下载的文件的名称
            response.setHeader("Content-disposition", "attachment;fileName=" + fileName);
            //通知客服文件的MIME类型
            response.setContentType("application/vnd.ms-excel;charset=UTF-8");
            //获取文件的路径 打成jar包后，getClass().getResource("")返回null
            String filePath = getClass().getResource("/template/" + fileName).getPath();
            FileInputStream input = new FileInputStream(filePath);
            OutputStream out = response.getOutputStream();
            byte[] b = new byte[2048];
            int len;
            while ((len = input.read(b)) != -1) {
                out.write(b, 0, len);
            }
            out.flush();
            //修正 Excel在“xxx.xlsx”中发现不可读取的内容。是否恢复此工作薄的内容？如果信任此工作簿的来源，请点击"是"
            response.setHeader("Content-Length", String.valueOf(input.getChannel().size()));
            input.close();
            logger.info("用户模板下载完成");
        } catch (Exception ex) {
            logger.error("getApplicationTemplate :", ex);
        }
    }
}