package com.poitest.controller;

import com.alibaba.csp.sentinel.annotation.SentinelResource;
import com.example.common.CommonInfoHolder;
import com.poitest.feign.RedisFeignClientAgent;
import com.poitest.handle.CustomerBlockHandler;
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
     * SentinelResource:注解定义资源,其中value值是资源名 blockHandlerClass 和 blockHandler分别是兜底类和兜底方法，
     * 采用兜底类较在业务类中为每个方法单独写兜底方法优点在于避免代码的侵入和膨胀。
     *
     * @param mo
     * @return
     */
    @RequestMapping("index")
    @SentinelResource(value = "index"
            , blockHandlerClass = CustomerBlockHandler.class// 兜底类
            , blockHandler = "handlerException")// 兜底方法
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
