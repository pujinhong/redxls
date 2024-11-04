package com.red.redxls.config;

/**
 * @author pjh
 * @created 2024/7/22
 */
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

@Component
public class FilePathGenerator implements ApplicationRunner {

    @Value("${report.template}")
    private String templatesPath;

    @Value("${report.outPath}")
    private String outputPath;


    @Override
    public void run(ApplicationArguments args) throws Exception {

        //创建模板文件夹
        File file = new File(templatesPath);
        if (!file.exists()) {
            file.mkdirs();
        }
        //创建生成目录
        File file1 = new File(outputPath);
        if (!file1.exists()) {
            file1.mkdirs();
        }

    }
}
