package com.poitest.feign;

import org.springframework.stereotype.Component;

/**
 * @Author zhangdj
 * @Date 2021/6/4:12:51
 * @Description
 */
@Component
public class RedisFeignClientIml implements RedisFeignClientAgent{
    @Override
    public String redisTemplateTest() {
        return "消费者在调用声场者的时候发生了熔断，做的降级处理";
    }
}