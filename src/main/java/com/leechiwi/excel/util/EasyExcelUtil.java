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
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

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
    private static final Logger logger = LoggerFactory.getLogger(EasyExcelUtil.class);
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
            if(CollectionUtils.isNotEmpty(excelElement.getSpanList())){//?????????????????????
                excelWriterSheetBuilder.registerWriteHandler(new DefaultMergeStrategy(excelElement.getSpanList()));
            }
            WriteSheet sheet = excelWriterSheetBuilder.build();
            int table=0;
            for (List<Object> data : datas) {
                if(CollectionUtils.isEmpty(data)){
                    continue;
                }
                ExcelWriterTableBuilder excelWriterTableBuilder = EasyExcel.writerTable(table);
                if(CollectionUtils.isNotEmpty(excelElement.getHead())){//?????????????????????
                    //List<List<String>> headList = new ArrayList<>();
                    //excelElement.getHead().forEach(h->headList.add(Collections.singletonList(h)));
                    excelWriterTableBuilder.head(excelElement.getHead());
                    //sheet=excelWriterSheetBuilder.head(headList).build();
                }else{
                    excelWriterTableBuilder.head(data.get(0).getClass());
                    //sheet=excelWriterSheetBuilder.head(data.get(0).getClass()).build();//?????????????????????
                }
                if(CollectionUtils.isNotEmpty(excelElement.getCellWriteCallBacks())) {
                    CellWriteCallBack cellWriteCallBack= excelElement.getCellWriteCallBacks().get(table);
                    if(Objects.nonNull(cellWriteCallBack)) {
                        excelWriterTableBuilder.registerWriteHandler(new DefaultCellWriteHandler(cellWriteCallBack));//?????????????????????????????????????????????????????????????????????
                    }
                }
                excelWriterTableBuilder.registerWriteHandler(defaultCellStyleHandler);//?????????????????????????????????
                excelWriterTableBuilder.registerWriteHandler(defaultWidthStyleStrategy);//???????????????????????????????????????????????????
                WriteTable writeTable=excelWriterTableBuilder.build();
                if(Objects.isNull(excelElement.getFunction())){//?????????????????????
                    excelWriter.write(data,sheet,writeTable);
                }else{//?????????????????????
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
     * ???????????????????????????????????????sheet??????
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
            /*String userAgent=request.getHeader("User-Agent");
            if(userAgent.contains("MSIE")||userAgent.contains("Trident")){
                filename = URLEncoder.encode(filename, "UTF-8");
            }else{
                filename = new String(filename.getBytes("UTF-8"), "ISO-8859-1");
            }*/
            response.setContentType("application/vnd.ms-excel");
            response.setCharacterEncoding("utf-8");
            response.addHeader("Content-Disposition","attachment; filename=\""+java.net.URLEncoder.encode(filename+".xlsx", "UTF-8"));
            outputStream = response.getOutputStream();
        } catch (UnsupportedEncodingException e) {
            logger.error("getOutputStream????????????????????????",e);
        } catch (IOException e) {
            logger.error("getOutputStream???????????????",e);
        }
        return outputStream;
    }
    private static OutputStream getZipOutputStream(String filename,HttpServletResponse response){
        OutputStream out = null;
        try {
            response.reset();
            response.setCharacterEncoding("UTF-8");
            response.setContentType("application/octet-stream");
            response.addHeader("Content-Disposition","attachment; filename=\""+java.net.URLEncoder.encode(filename+".zip", "UTF-8"));
            out = response.getOutputStream();
        } catch (IOException e) {
            logger.error("getZipOutputStream???????????????",e);
        }
        return out;
    }

    /**
     * ???????????????sheet?????????sheet
     * (?????????????????????listener?????????????????????????????????????????????????????????????????????????????????????????????????????????List<ExcelSheetElement> getFromExcel?????????????????????)
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
            // ?????????????????? DemoDataListener???doAfterAllAnalysed ????????????sheet??????????????????????????????????????????sheet??????????????????DemoDataListener?????????
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
                    if(Objects.isNull(consumer)&&Objects.nonNull(defaultConsumer)){//??????consumer????????????????????????????????????????????????(???sheetElement.getClazz())??????????????????????????????LinkedHashMap<Integer,String>???
                        consumer=defaultConsumer;
                    }
                    MyReadListener<?> readListener=new MyReadListener<>(consumer);
                    Class<?> clazz = sheetElement.getClazz();
                    if(Objects.isNull(clazz)){//????????????????????????????????????????????????LinkedHashMap<Integer,String>
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
                    }else{//???????????????????????????listener?????????????????????????????????
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
                    logger.error("fillWithTemplate???????????????",e);
                }
            }
        }else if(file instanceof InputStream){
            templateFile=(InputStream)file;
        }else if(file instanceof File){
            try {
                templateFile=new FileInputStream((File)file);
            } catch (FileNotFoundException e) {
                logger.error("fillWithTemplate???????????????",e);
            }
        }
        // ????????????
        //String targetFileName = "??????????????????.xlsx";
        //List<FillData> fillDatas = initData();
        // ?????????????????????
        ExcelWriter excelWriter = null;
        ExcelWriterBuilder excelWriterBuilder=null;
        if(Objects.nonNull(outputStream)){
            excelWriterBuilder = EasyExcel.write(outputStream);
        }else{
            excelWriterBuilder=EasyExcel.write(targetFilename);//????????????templateFile????????????
        }
        //excelWriterBuilder.registerWriteHandler(new DefaultCellStyleHandler());
        if(Objects.nonNull(excelFillElement.getCellWriteCallBack())){
            excelWriterBuilder.registerWriteHandler(new DefaultCellWriteHandler(excelFillElement.getCellWriteCallBack()));
        }
        if(Objects.nonNull(excelFillElement.getSpanList())) {
            excelWriterBuilder.registerWriteHandler(new DefaultMergeStrategy(excelFillElement.getSpanList()));
        }
        excelWriter=excelWriterBuilder.withTemplate(templateFile).build();
        // ?????????????????????
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        // ??????????????????????????????????????????????????????????????????????????????????????????????????????
        FillConfig.FillConfigBuilder fillConfigBuilder = FillConfig.builder().forceNewRow(true);
        if(excelFillElement.isHorizontal()){
            fillConfigBuilder.direction(WriteDirectionEnum.HORIZONTAL);
        }
        FillConfig fillConfig = fillConfigBuilder.build();
        //FillConfig fillConfig = FillConfig.builder().direction(WriteDirectionEnum.HORIZONTAL).build();
        for (Object fillData : excelFillElement.getFillDatas()) {
            // ???????????????
            excelWriter.fill(fillData, fillConfig, writeSheet);
        }
        // ??????
        excelWriter.finish();
    }
}
