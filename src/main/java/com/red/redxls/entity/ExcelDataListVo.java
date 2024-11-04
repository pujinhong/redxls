package com.red.redxls.entity;

import lombok.Data;

import java.util.List;
import java.util.Map;

/**
 * @author pjh
 * @created 2024/7/22
 */
@Data
public class ExcelDataListVo {

    String templateName;
    String fileName;
    List<Map> data;
    Integer isStream;
}
