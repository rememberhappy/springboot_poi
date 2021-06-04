package com.poitest.controller;

import com.example.common.CommonInfoHolder;
import com.poitest.feign.RedisFeignClientAgent;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.annotation.Resource;

@RestController
@RequestMapping
public class TestController {
    @Resource
    RedisFeignClientAgent redisFeignClient;

    /**
     * 简单的index页面，会直接跳转到index网页，操作上传下载按钮
     *
     * @param mo
     * @return
     */
    @RequestMapping("index")
    public ModelAndView index(ModelAndView mo) {
        // 公共信息持有器中获取信息
        Integer userId = CommonInfoHolder.getUserId();
        System.out.println("poi 项目中获取公共信息 userId：" + userId);
        String token = CommonInfoHolder.getToken();
        System.out.println("poi 项目中获取公共信息 token：" + token);
        // feign 接口调用
        String s = redisFeignClient.redisTemplateTest();
        System.out.println(s);
        mo.addObject("user", "asd");
        mo.setViewName("index");
        return mo;
    }
}
