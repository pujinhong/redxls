package com.red.redxls;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.scheduling.annotation.EnableScheduling;


@SpringBootApplication
@EnableScheduling
public class RedxlsApplication {

    public static void main(String[] args) {
        SpringApplication.run(RedxlsApplication.class, args);
        System.out.println( "长期维护模块：简单报表 启动成功！ \n模块维护人： 浦金宏");
    }

}
