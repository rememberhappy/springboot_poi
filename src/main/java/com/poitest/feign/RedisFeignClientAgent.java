package com.poitest.feign;

import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

@FeignClient(name = "redisFeignClient", url = "http://127.0.0.1:8089", fallback = RedisFeignClientIml.class)
public interface RedisFeignClientAgent {
    @RequestMapping(value = "/stringtest/findvalue", method = RequestMethod.POST)
    String redisTemplateTest();
}
