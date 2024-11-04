package com.red.redxls.config;

/**
 * @author pjh
 * @created 2024/7/22
 */
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.CorsRegistry;
import org.springframework.web.servlet.config.annotation.ResourceHandlerRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

import java.io.File;

@Configuration
public class WebConfig implements WebMvcConfigurer {

    @Value("${report.outPath}")
    private String outputPath;
    @Value("${report.publishPath}")
    private String publishPath;

    @Override
    public void addResourceHandlers(ResourceHandlerRegistry registry) {

        System.out.println("outputPath:"+outputPath+ File.separator);
        registry.addResourceHandler(publishPath+"/**")
                .addResourceLocations("file:"+outputPath+ File.separator);
    }
//    @Override
//    public void addCorsMappings(CorsRegistry registry) {
//        registry.addMapping(publishPath+"/**") // 只对/static/下的资源生效
//                .allowedOrigins("*")
//                .allowedMethods("*")
//                .allowedHeaders("*")
//                .allowCredentials(true);
//    }
}