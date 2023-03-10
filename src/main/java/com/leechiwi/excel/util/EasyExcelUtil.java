package com.leechiwi.excel.util;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.enums.WriteDirectionEnum;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.util.StringUtils;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.builder.ExcelWriterTableBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.WriteTable;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.leechiwi.excel.handler.DefaultCellStyleHandler;
import com.leechiwi.excel.handler.DefaultCellWriteHandler;
import com.leechiwi.excel.handler.DefaultMergeStrategy;
import com.leechiwi.excel.handler.DefaultWidthStyleStrategy;
import com.leechiwi.excel.interfaces.CellWriteCallBack;
import com.leechiwi.excel.listener.MyReadListener;
import com.leechiwi.excel.model.ExcelFillElement;
import com.leechiwi.excel.model.ExcelSheetElement;
import org.apache.commons.collections.CollectionUtils;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.function.Consumer;

public class EasyExcelUtil {
    public static void exportWebExcel(List<ExcelSheetElement> sheetList, HttpServletRequest request, HttpServletResponse response, String filename){
        OutputStream outputStream = getOutputStream(filename, request, response);
        writeToExcel(sheetList,outputStream);
    }
    public static void exportWebExcelInZip(List<List<ExcelSheetElement>> excelList,List<String> excelFileNames, HttpServletResponse response, String filename){
        writeToExcelInZip(excelList,excelFileNames, getZipOutputStream(filename, response));
    }
    public static void writeToExcel(List<ExcelSheetElement> sheetList, OutputStream outputStream){
        ExcelWriter excelWriter = EasyExcel.write(outputStream).build();
        int index=0;
        DefaultWidthStyleStrategy defaultWidthStyleStrategy = new DefaultWidthStyleStrategy();
        DefaultCellStyleHandler defaultCellStyleHandler = new DefaultCellStyleHandler();
        for (ExcelSheetElement excelElement : sheetList) {
            List<List<Object>> datas = excelElement.getData();
            if(CollectionUtils.isEmpty(datas)){
                continue;
            }
            String sheetName=excelElement.getSheetName();
            ExcelWriterSheetBuilder excelWriterSheetBuilder = EasyExcel.writerSheet(index, StringUtils.isBlank(sheetName) ? new StringBuilder("sheet").append(index + 1).toString() : sheetName);
            if(CollectionUtils.isNotEmpty(excelElement.getSpanList())){//单元格合并处理
                excelWriterSheetBuilder.registerWriteHandler(new DefaultMergeStrategy(excelElement.getSpanList()));
            }
            WriteSheet sheet = excelWriterSheetBuilder.build();
            int table=0;
            for (List<Object> data : datas) {
                if(CollectionUtils.isEmpty(data)){
                    continue;
                }
                ExcelWriterTableBuilder excelWriterTableBuilder = EasyExcel.writerTable(table);
                if(CollectionUtils.isNotEmpty(excelElement.getHead())){//代码自定义表头
                    //List<List<String>> headList = new ArrayList<>();
                    //excelElement.getHead().forEach(h->headList.add(Collections.singletonList(h)));
                    excelWriterTableBuilder.head(excelElement.getHead());
                    //sheet=excelWriterSheetBuilder.head(headList).build();
                }else{
                    excelWriterTableBuilder.head(data.get(0).getClass());
                    //sheet=excelWriterSheetBuilder.head(data.get(0).getClass()).build();//实体类注解形式
                }
                if(CollectionUtils.isNotEmpty(excelElement.getCellWriteCallBacks())) {
                    CellWriteCallBack cellWriteCallBack= excelElement.getCellWriteCallBacks().get(table);
                    if(Objects.nonNull(cellWriteCallBack)) {
                        excelWriterTableBuilder.registerWriteHandler(new DefaultCellWriteHandler(cellWriteCallBack));//自定义单元格格式样式（比如某一个单元格的样式）
                    }
                }
                excelWriterTableBuilder.registerWriteHandler(defaultCellStyleHandler);//全局默认单元格格式样式
                excelWriterTableBuilder.registerWriteHandler(defaultWidthStyleStrategy);//全局单元格宽度（宽度适应文字长度）
                WriteTable writeTable=excelWriterTableBuilder.build();
                if(Objects.isNull(excelElement.getFunction())){//实体类注解形式
                    excelWriter.write(data,sheet,writeTable);
                }else{//代码自定义数据
                    List<List<Object>> list =excelElement.getFunction().apply(data);
                    excelWriter.write(list,sheet,writeTable);
                }
                table++;
            }
            index++;
        }
        excelWriter.finish();
    }
    public static void writeToExcelInZip(List<List<ExcelSheetElement>> excelList,List<String> excelFileNames, OutputStream zipOutputStream){
        if(CollectionUtils.isEmpty(excelList)){
            return;
        }
        List<byte[]> streamList = new ArrayList<>();
        for (List<ExcelSheetElement> excelSheetElements : excelList) {
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            writeToExcel(excelSheetElements,byteArrayOutputStream);
            streamList.add(byteArrayOutputStream.toByteArray());
        }
        Zip.zip(zipOutputStream,streamList,excelFileNames,".xlsx");
    }
    /**
     * 自定义数据且自定义表头单个sheet导出
     * @param data
     * @param head
     * @param outputStream
     */
    public static void writeToExcel(List<List<Object>> data, List<String> head,OutputStream outputStream){
        ExcelWriterBuilder write = EasyExcel.write(outputStream);
        ExcelWriterSheetBuilder sheet = write.sheet("sheet1");
        List<List<String>> headList = new ArrayList<>();
        head.forEach(h->headList.add(Collections.singletonList(h)));
        sheet.head(headList).doWrite(data);
    }
    private static OutputStream getOutputStream(String filename,HttpServletRequest request,HttpServletResponse response){
        OutputStream outputStream=null;
        try {
            String userAgent=request.getHeader("User-Agent");
            if(userAgent.contains("MSIE")||userAgent.contains("Trident")){
                filename = URLEncoder.encode(filename, "UTF-8");
            }else{
                filename = new String(filename.getBytes("UTF-8"), "ISO-8859-1");
            }
            response.setContentType("application/vnd.ms-exce");
            response.setCharacterEncoding("utf-8");
            response.addHeader("Content-Disposition","filename="+filename+".xlsx");
            outputStream = response.getOutputStream();
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return outputStream;
    }
    private static OutputStream getZipOutputStream(String filename,HttpServletResponse response){
        OutputStream out = null;
        try {
            response.reset();
            response.setCharacterEncoding("UTF-8");
            response.setContentType("application/octet-stream");
            response.setHeader("Content-Disposition", "attachment;filename=" + filename+".zip");
            out = response.getOutputStream();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return out;
    }

    /**
     * 导入第一个sheet或全部sheet
     * (该方法公用一个listener且仅支持设定固定实体类来接收即不支持自定义，如果想自定义接收则可以使用List<ExcelSheetElement> getFromExcel更加灵活的实现)
     * @param inputStream
     * @param clazz
     * @param consumer
     * @param readAllSheet
     * @param <T>
     * @return
     */
    public  static <T> List<T> getFromExcel(InputStream inputStream, Class<T> clazz, Consumer<T> consumer, boolean readAllSheet){
        MyReadListener<T> readListener=new MyReadListener<>(consumer);
        ExcelReaderBuilder read = EasyExcel.read(inputStream, clazz, readListener);
        if(readAllSheet){
            // 这里需要注意 DemoDataListener的doAfterAllAnalysed 会在每个sheet读取完毕后调用一次。然后所有sheet都会往同一个DemoDataListener里面写
            read.doReadAll();
        }
        read.sheet().doRead();
        return readListener.getList();
    }
    public static List<ExcelSheetElement> getFromExcel(InputStream inputStream, List<ExcelSheetElement> sheetElements, List<String> includeSheets, Consumer<?> defaultConsumer,List<String> defaultHead){
        List<ExcelSheetElement> result = new ArrayList<>();
        ExcelReader excelReader = EasyExcel.read(inputStream).build();
        List<ReadSheet> readSheets = excelReader.excelExecutor().sheetList();
        int index=0;
        if (!readSheets.isEmpty()) {
            for (ReadSheet readSheet : readSheets) {
                String sheetName = readSheet.getSheetName();
                if(CollectionUtils.isEmpty(includeSheets)||includeSheets.contains(sheetName)){
                    ExcelSheetElement sheetElement = sheetElements.get(index);
                    Consumer<?> consumer =sheetElement.getConsumer();
                    if(Objects.isNull(consumer)&&Objects.nonNull(defaultConsumer)){//处理consumer中的泛型一定要和指定的接收实体类(即sheetElement.getClazz())一样（自定义的类型是LinkedHashMap<Integer,String>）
                        consumer=defaultConsumer;
                    }
                    MyReadListener<?> readListener=new MyReadListener<>(consumer);
                    Class<?> clazz = sheetElement.getClazz();
                    if(Objects.isNull(clazz)){//不指定接收实体类，则返回的类型是LinkedHashMap<Integer,String>
                        final List<List<String>> headList = new ArrayList<>();
                        if(CollectionUtils.isNotEmpty(sheetElement.getHead())){
                            //sheetElement.getHead().forEach(h->headList.add(Collections.singletonList(h)));
                            headList.addAll(sheetElement.getHead());
                        }else if(CollectionUtils.isNotEmpty(defaultHead)){
                            defaultHead.forEach(h->headList.add(Collections.singletonList(h)));
                        }else{
                            continue;
                        }
                        readSheet = EasyExcel.readSheet(readSheet.getSheetNo()).head(headList).registerReadListener(readListener).build();
                    }else{//指定接收实体类，则listener中返回的类型是指定的类
                        readSheet = EasyExcel.readSheet(readSheet.getSheetNo()).head(clazz).registerReadListener(readListener).build();
                    }
                    excelReader.read(readSheet);
                    sheetElement.setSheetName(sheetName);
                    List list = readListener.getList();
                    sheetElement.setData(list);
                    result.add(sheetElement);
                    index++;
                }
            }
        }
        excelReader.finish();
        return result;
    }
    public static void getWebExcelByfillWithTemplate(String targetFilename, ExcelFillElement excelFillElement, HttpServletRequest request, HttpServletResponse response){
        OutputStream outputStream = getOutputStream(targetFilename, request, response);
        fillWithTemplate(targetFilename,excelFillElement,outputStream);
    }
    public static void fillWithTemplate(String targetFilename,ExcelFillElement excelFillElement, OutputStream outputStream){
        InputStream templateFile=null;
        Object file=excelFillElement.getFile();
        if(file instanceof String ){
            if(excelFillElement.isTemplateExsit()) {
                templateFile = EasyExcelUtil.class.getClassLoader().getResourceAsStream((String) file);
            }else{
                try {
                    templateFile=new FileInputStream(new File((String)file));
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                }
            }
        }else if(file instanceof InputStream){
            templateFile=(InputStream)file;
        }else if(file instanceof File){
            try {
                templateFile=new FileInputStream((File)file);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
        }
        // 目标文件
        //String targetFileName = "组合数据填充.xlsx";
        //List<FillData> fillDatas = initData();
        // 生成工作簿对象
        ExcelWriter excelWriter = null;
        ExcelWriterBuilder excelWriterBuilder=null;
        if(Objects.nonNull(outputStream)){
            excelWriterBuilder = EasyExcel.write(outputStream);
        }else{
            excelWriterBuilder=EasyExcel.write(targetFilename);//下载到和templateFile同步录下
        }
        //excelWriterBuilder.registerWriteHandler(new DefaultCellStyleHandler());
        if(Objects.nonNull(excelFillElement.getCellWriteCallBack())){
            excelWriterBuilder.registerWriteHandler(new DefaultCellWriteHandler(excelFillElement.getCellWriteCallBack()));
        }
        if(Objects.nonNull(excelFillElement.getSpanList())) {
            excelWriterBuilder.registerWriteHandler(new DefaultMergeStrategy(excelFillElement.getSpanList()));
        }
        excelWriter=excelWriterBuilder.withTemplate(templateFile).build();
        // 生成工作表对象
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        // 组合填充时，因为多组填充的数据量不确定，需要在多组填充完之后另起一行
        FillConfig.FillConfigBuilder fillConfigBuilder = FillConfig.builder().forceNewRow(true);
        if(excelFillElement.isHorizontal()){
            fillConfigBuilder.direction(WriteDirectionEnum.HORIZONTAL);
        }
        FillConfig fillConfig = fillConfigBuilder.build();
        //FillConfig fillConfig = FillConfig.builder().direction(WriteDirectionEnum.HORIZONTAL).build();
        for (Object fillData : excelFillElement.getFillDatas()) {
            // 填充并换行
            excelWriter.fill(fillData, fillConfig, writeSheet);
        }
        // 关闭
        excelWriter.finish();
    }
}
