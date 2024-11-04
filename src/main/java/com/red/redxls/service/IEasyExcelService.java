package com.red.redxls.service;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

public interface IEasyExcelService {

    File objFill(String fileName, String templateName, Map data);
    File listFill(String fileName, String templateName, List data);

    File complex(String fileName, String templateName, Map<String, Object> data, Map<String, Object> option) throws IOException;

    void complex(HttpServletResponse response, String fileName, String templateName, Map<String, Object> data) throws IOException;

    void objFill(HttpServletResponse response, String fileName, String templateName, Map data) throws IOException;
    void listFill(HttpServletResponse response, String fileName, String templateName, List data) throws IOException;


}
