package com.poitest.controller;

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
        mo.addObject("user", "asd");
        mo.setViewName("index");
        return mo;
    }
}
