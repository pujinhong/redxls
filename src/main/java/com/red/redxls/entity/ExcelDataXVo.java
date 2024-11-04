package com.red.redxls.entity;

import lombok.Data;

import java.util.Map;

/**
 * @author pjh
 * @created 2024/7/22
 */
@Data
public class ExcelDataXVo {

    String templateName;
    String fileName;
    Map data;
    Integer isStream;
    Map option;
}
