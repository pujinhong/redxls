package com.red.redxls.service.Impl;

import com.red.redxls.service.ITomcatService;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.File;

/**
 * @author pjh
 * @created 2024/7/22
 */
@Service
public class TomcatServiceImpl implements ITomcatService {


    @Value("${report.publishPath}")
    private String publishPath;

    @Value("${report.outPath}")
    private String outPath;
    @Override
    public String getPublishPath(File file) {

        if(file.getAbsolutePath().contains("\\")){
            String shortFileName = file.getAbsolutePath().replace("\\","/");
            int index = shortFileName.indexOf(outPath)  + outPath.length();
            String relativeFileName = shortFileName.substring(index);
            if(relativeFileName.startsWith("/")){
                return publishPath+relativeFileName ;
            }else{
                return publishPath+"/"+relativeFileName ;
            }
        }else{
            int index = file.getAbsolutePath().indexOf(outPath) + outPath.length();
            String relativeFileName = file.getAbsolutePath().substring(index);
            if(relativeFileName.startsWith("/")){
                return publishPath+relativeFileName;
            }else{
                return publishPath+"/"+relativeFileName;
            }
        }
    }
}
