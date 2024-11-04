package com.red.redxls.controller;


import com.red.redxls.entity.*;
import com.red.redxls.service.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;

@RestController
@RequestMapping("/template")
@CrossOrigin(origins = "*")
public class TemplateController {

    @Autowired
    private IEasyExcelService easyExcelService;

    @Autowired
    private ITomcatService tomcatService;


    @CrossOrigin
    @PostMapping("/export")
    public Result expComplex(@RequestBody ExcelDataXVo excelDataVo) throws IOException {
        if(excelDataVo.getTemplateName() == null || excelDataVo.getTemplateName().equals("")
        ){
            return Result.error("缺少模板名称");
        }
        if(excelDataVo.getData() == null){
            return Result.error("缺少数据");
        }

        // 是数组
        if(excelDataVo.getIsStream() != null && excelDataVo.getIsStream()>0){
            try {
                easyExcelService.complex(excelDataVo.getFileName(), excelDataVo.getTemplateName(), (LinkedHashMap<String, Object>) excelDataVo.getData(), (LinkedHashMap<String, Object>) excelDataVo.getOption());
            }catch (FileNotFoundException e){
                return Result.error(e.getMessage());
            } catch (IOException e) {
                e.printStackTrace();
                return Result.error("导出失败");
            }
            return Result.ok();
        }else {
            try{
                File file = easyExcelService.complex(excelDataVo.getFileName(), excelDataVo.getTemplateName(), (LinkedHashMap<String,Object>) excelDataVo.getData(),(LinkedHashMap<String,Object>) excelDataVo.getOption());
                return Result.ok().setData( tomcatService.getPublishPath(file));
            }catch (FileNotFoundException e){
                return Result.error(e.getMessage());
            } catch (IOException e) {
                e.printStackTrace();
                return Result.error("导出失败");
            }
        }
    }

//    /**
//     * 废弃
//     * @param excelDataVo
//     * @param response
//     * @return
//     */
//    @CrossOrigin
//    @PostMapping("/export/obj")
//    public Result excel(@RequestBody ExcelDataVo excelDataVo, HttpServletResponse response) {
//
//        if(excelDataVo.getTemplateName() == null || excelDataVo.getTemplateName().equals("")
//        ){
//            return Result.error("缺少模板名称");
//        }
//        if(excelDataVo.getData() == null){
//            return Result.error("缺少数据");
//        }
//        // 是对象
//        if(excelDataVo.getIsStream() != null && excelDataVo.getIsStream()>0){
//            try {
//                easyExcelService.objFill(response, excelDataVo.getFileName(), excelDataVo.getTemplateName(), (LinkedHashMap) excelDataVo.getData());
//            } catch (IOException e) {
//                e.printStackTrace();
//                return Result.error("导出失败");
//            }
//            return Result.ok();
//        }else {
//            File file = easyExcelService.objFill(excelDataVo.getFileName(), excelDataVo.getTemplateName(), (LinkedHashMap) excelDataVo.getData());
//            return Result.ok().setData( tomcatService.getPublishPath(file));
//        }
//
//    }
//
//    /**
//     * 废弃
//     * @param excelDataVo
//     * @param response
//     * @return
//     */
//    @CrossOrigin
//    @PostMapping("/export/list")
//    public Result expList(@RequestBody ExcelDataListVo excelDataVo, HttpServletResponse response){
//        if(excelDataVo.getTemplateName() == null || excelDataVo.getTemplateName().equals("")
//        ){
//            return Result.error("缺少模板名称");
//        }
//        if(excelDataVo.getData() == null){
//            return Result.error("缺少数据");
//        }
//        // 是数组
//        if(excelDataVo.getIsStream() != null && excelDataVo.getIsStream()>0){
//            try {
//                easyExcelService.listFill(response, excelDataVo.getFileName(), excelDataVo.getTemplateName(), (ArrayList) excelDataVo.getData());
//            } catch (IOException e) {
//                e.printStackTrace();
//                return Result.error("导出失败");
//            }
//            return Result.ok();
//        }else {
//            File file = easyExcelService.listFill(excelDataVo.getFileName(), excelDataVo.getTemplateName(), (ArrayList) excelDataVo.getData());
//            return Result.ok().setData( tomcatService.getPublishPath(file));
//        }
//    }


}
