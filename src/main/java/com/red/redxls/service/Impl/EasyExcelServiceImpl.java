package com.red.redxls.service.Impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.metadata.fill.FillWrapper;
import com.red.redxls.config.MergeCellStrategy;
import com.red.redxls.service.IEasyExcelService;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @author pjh
 * @created 2024/7/22
 */
@Service
public class EasyExcelServiceImpl implements IEasyExcelService {

    @Value("${report.template}")
    private String templatesPath;

    @Value("${report.outPath}")
    private String outputPath;


    SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
    @Override
    public File objFill(String fileName ,String templateName,Map data) {
        try{
            File templateFile = getTemplateFile(templateName);
            File outFile = getOutFile(fileName,templateFile);
            EasyExcel.write(outFile).withTemplate(templateFile).sheet().doFill(data);
            return outFile;
        }catch (Exception e){
            throw  e;
        }

    }

    @Override
    public File listFill(String fileName,String templateName , List data) {
        try{
            File templateFile = getTemplateFile(templateName);
            File outFile = getOutFile(fileName,templateFile);
            EasyExcel.write(outFile).withTemplate(templateFile).sheet().doFill(data);
            return outFile;
        }catch (Exception e){
            throw  e;
        }
    }

    @Override
    public void listFill(HttpServletResponse response, String fileName, String templateName, List data) throws IOException {
        File templateFile = getTemplateFile(templateName);
        String extensName = templateFile.getName().substring(templateFile.getName().indexOf("."));
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + URLEncoder.encode(fileName, "UTF-8").replaceAll("\\+", "%20") + extensName);
        EasyExcel.write(response.getOutputStream()).withTemplate(templateFile).sheet().doFill(data);
    }


    @Override
    public File complex(String fileName, String templateName, Map<String, Object> data, Map<String, Object> option) throws IOException {
        File templateFile = getTemplateFile(templateName);
        if(templateFile == null){
            throw new FileNotFoundException("模板不存在");
        }

        File outFile = getOutFile(fileName,templateFile);

        //填充数据
        try {

            ExcelWriterBuilder write = EasyExcel.write(outFile);
            MergeCellStrategy mergeCellStrategy = new MergeCellStrategy();
            write.registerWriteHandler(mergeCellStrategy);
            ExcelWriter excelWriter = write.withTemplate(templateFile).build();
            WriteSheet writeSheet = EasyExcel.writerSheet().build();
            writeSheet.setUseDefaultStyle(false);
            Map<String, Object> map = new HashMap<String, Object>();
            for (Map.Entry<String, Object> item : data.entrySet()) {
                if (item.getValue() instanceof LinkedHashMap) {
                    List obj = new ArrayList();
                    obj.add(item.getValue());
                    for (Map.Entry<String, Object> item2 : ((LinkedHashMap<String, Object>) item.getValue()).entrySet()) {
                        //如果对象的属性仍是数组，转换成字符串
                        if (item2.getValue() instanceof ArrayList) {
                            String strItem2 = " ";
                            for (Object objItem2 : (ArrayList) item2.getValue()) {
                                strItem2 += (objItem2.toString() + "、");
                            }
                            ((LinkedHashMap<String, Object>) item.getValue()).put(item2.getKey(), strItem2.substring(0, strItem2.length() - 1));
                        }
                    }
                    FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
                    fillConfig.setAutoStyle(false);
                    excelWriter.fill(new FillWrapper(item.getKey(), obj), fillConfig, writeSheet);
                } else if (item.getValue() instanceof ArrayList) {
                    FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
                    excelWriter.fill(new FillWrapper(item.getKey(), (ArrayList) item.getValue()), fillConfig, writeSheet);
                } else {
                    map.put(item.getKey(), item.getValue());
                }
            }
            List attri = new ArrayList();
            attri.add(map);
            excelWriter.fill(new FillWrapper("data", attri), writeSheet);
            excelWriter.finish();
            excelWriter.close();
        }catch (Exception ex){
            throw ex;
        }
        return outFile;
    }

    @Override
    public void complex(HttpServletResponse response, String fileName, String templateName, Map<String, Object> data) throws IOException {
        File templateFile = getTemplateFile(templateName);
        String extensName = templateFile.getName().substring(templateFile.getName().indexOf("."));
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + URLEncoder.encode(fileName, "UTF-8").replaceAll("\\+", "%20") + extensName);
        try (OutputStream outputStream = response.getOutputStream();
             ExcelWriter excelWriter = EasyExcel.write(outputStream).withTemplate(templateFile).build()) {
            WriteSheet writeSheet = EasyExcel.writerSheet().build();
            Map<String, Object> map = new HashMap<String, Object>();
            for(Map.Entry<String, Object> item : data.entrySet()){
                if(item.getValue() instanceof LinkedHashMap){
                    List obj = new ArrayList();
                    obj.add(item.getValue());
                    excelWriter.fill(new FillWrapper(item.getKey(), obj), writeSheet);
                } else if(item.getValue() instanceof ArrayList){
                    excelWriter.fill(new FillWrapper(item.getKey(), (ArrayList)item.getValue()), writeSheet);
                } else {
                    map.put(item.getKey(), item.getValue());
                }
            }
            List attri = new ArrayList();
            attri.add(map);
            excelWriter.fill(new FillWrapper("data", attri), writeSheet);
        }
    }

    @Override
    public void objFill(HttpServletResponse response, String fileName, String templateName, Map data) throws IOException {
        File templateFile = getTemplateFile(templateName);
        String extensName = templateFile.getName().substring(templateFile.getName().indexOf("."));
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + URLEncoder.encode(fileName, "UTF-8").replaceAll("\\+", "%20") + extensName);
        EasyExcel.write(response.getOutputStream()).withTemplate(templateFile).sheet().doFill(data);
    }

    private File getTemplateFile(String fileName){

        File dir = new File(templatesPath);
        File[] files = dir.listFiles((dir1, name) -> name.substring(0, name.indexOf(".")).equals(fileName));
        if(files.length>0){
            return files[0];
        }else {
            return null;
        }
    }
    private File getOutFile(String fileName,File template){

        com.red.redxls.util.UUID uuid = com.red.redxls.util.UUID.fastUUID();
        String strUuid = uuid.toString().replace("-", "");
        File dirUuid = new File(outputPath + File.separator+strUuid+File.separator);
        dirUuid.mkdir();
        String extensName = template.getName().substring(template.getName().indexOf("."));
        if(fileName!= null){
            return new File(dirUuid.getAbsolutePath()+File.separator+fileName+extensName);
        }else{
            return  new File(dirUuid.getAbsolutePath()+File.separator+ sdf.format(new Date()) +extensName);
        }

    }
}
