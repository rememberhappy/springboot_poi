package com.poitest.controller;

import com.example.common.CommonInfoHolder;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

@RestController
@RequestMapping
public class TestController {

    /**
     * 简单的index页面，会直接跳转到index网页，操作上传下载按钮
     * @param mo
     * @return
     */
    @RequestMapping("index")
    public ModelAndView index(ModelAndView mo){
        // 公共信息持有器中获取信息
        Integer userId = CommonInfoHolder.getUserId();
        System.out.println("poi 项目中获取公共信息 userId：" + userId);
        String token = CommonInfoHolder.getToken();
        System.out.println("poi 项目中获取公共信息 token：" + token);
        mo.addObject("user", "asd");
        mo.setViewName("index");
        return mo;
    }
}
