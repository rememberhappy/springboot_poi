package com.poitest;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.cloud.openfeign.EnableFeignClients;

@SpringBootApplication
@EnableFeignClients// feign启动回报错
public class SpringbootPoiApplication {

    public static void main(String[] args) {
        SpringApplication.run(SpringbootPoiApplication.class, args);
    }

}
