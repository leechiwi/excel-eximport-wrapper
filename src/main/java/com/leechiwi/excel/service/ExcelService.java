package com.leechiwi.excel.service;

import com.leechiwi.excel.interfaces.CellWriteCallBack;
import com.leechiwi.excel.model.ExcelFillElement;
import com.leechiwi.excel.model.ExcelSheetElement;
import com.leechiwi.excel.model.RowAndColSpan;
import com.leechiwi.excel.style.CellStyleUtil;
import com.leechiwi.excel.util.EasyExcelUtil;
import com.leechiwi.excel.mapper.TestMapper;
import com.leechiwi.excel.pojo.Demo;
import com.leechiwi.excel.pojo.FillData;
import com.leechiwi.excel.pojo.Test;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.Stream;

@Service
public class ExcelService {
    @Autowired
    private TestMapper testMapper;
    public List<List<Object>> getOriginData(){
        List<List<Object>> result = new ArrayList<>();
        List<Object> list = new ArrayList<>();
       /* Map<String,Object> map1=new HashMap<>();
        map1.put("index","1");
        map1.put("name","刘备");
        map1.put("code","玄德");
        Map<String,Object> map2=new HashMap<>();
        map2.put("index","2");
        map2.put("name","关羽");
        map2.put("code","云长");
        Map<String,Object> map3=new HashMap<>();
        map3.put("index","3");
        map3.put("name","张飞");
        map3.put("code","翼德");
        list.add(map1);
        list.add(map2);
        list.add(map3);
        result.add(list);
        //如果需要另一个table的数据则重复上面的工作然后result.add(list);*/
        Test t1=new Test();
        t1.setIndex(1);
        t1.setName("刘备");
        t1.setCode("玄德");
        t1.setOther("刘备字玄德，皇叔，城府深爱哭");
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        Date parse=null;
        try {
            parse = simpleDateFormat.parse("1970-01-01 10:10:10");
            t1.setBirthday(parse);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        Test t2=new Test();
        t2.setIndex(2);
        t2.setName("关羽");
        t2.setCode("云长");
        try {
            parse = simpleDateFormat.parse("1970-01-02 10:10:10");
            t2.setBirthday(parse);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        t2.setOther("关羽字云长，战神，口头禅：插标卖首");
        Test t3=new Test();
        t3.setIndex(3);
        t3.setName("张飞");
        t3.setCode("翼德");
        t3.setOther("张飞字翼德，猛将，脾气不好爱喝酒");
        try {
            parse = simpleDateFormat.parse("1970-01-03 10:10:10");
            t3.setBirthday(parse);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        list.add(t1);
        list.add(t2);
        list.add(t3);
        result.add(list);
        List<Object> list4 = new ArrayList<>();
        Demo t4=new Demo();
        t4.setIndex(1);
        t4.setName("刘备4");
        t4.setCode("玄德4");
        t4.setOther("刘备字玄德，皇叔，城府深爱哭4");
        try {
            parse = simpleDateFormat.parse("1970-01-01 10:10:10");
            t4.setBirthday(parse);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        Demo t5=new Demo();
        t5.setIndex(2);
        t5.setName("关羽5");
        t5.setCode("云长5");
        try {
            parse = simpleDateFormat.parse("1970-01-02 10:10:10");
            t5.setBirthday(parse);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        t5.setOther("关羽字云长，战神，口头禅：插标卖首5");
        Demo t6=new Demo();
        t6.setIndex(3);
        t6.setName("张飞6");
        t6.setCode("翼德6");
        t6.setOther("张飞字翼德，猛将，脾气不好爱喝酒6");
        try {
            parse = simpleDateFormat.parse("1970-01-03 10:10:10");
            t6.setBirthday(parse);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        list4.add(t4);
        list4.add(t5);
        list4.add(t6);
        result.add(list4);
        return result;
    }
    public void getWebExcelData(HttpServletRequest request,HttpServletResponse response){
        ExcelSheetElement excelSheetElement = new ExcelSheetElement();
        List<List<Object>> originData = getOriginData();
        Function<List<Object>,List<List<Object>>> function=(list)->{
            List<List<Object>> result = new ArrayList<>();
            for (Object obj : list) {
                List<Object> lst = new ArrayList<>();
                Map<String,Object> map=(Map<String,Object>)obj;
                lst.add(map.get("index"));
                lst.add(map.get("name"));
                lst.add(map.get("code"));
                result.add(lst);
            }
            return result;
        };
        excelSheetElement.setData(originData);
        List<RowAndColSpan> splist = new ArrayList<>();
        RowAndColSpan rc1=new RowAndColSpan(3,4,0,0);
        splist.add(rc1);
        RowAndColSpan rc2=new RowAndColSpan(4,5,1,1);
        splist.add(rc2);
        RowAndColSpan rc3=new RowAndColSpan(10,11,2,2);
        splist.add(rc3);
        excelSheetElement.setSpanList(splist);
        CellWriteCallBack cellWriteCallBack1 = (context,cell,value,writeCellStyle) -> {
            if(cell.getColumnIndex()==1&&"关羽".equals(value)) {
                writeCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                writeCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
            }
            if(originData.get(0).size()==(context.getRelativeRowIndex()+1)){//一个sheet的table之间表格隔行
                if(cell.getColumnIndex()==0) {//最后一行的第一列的时候创建一个新行就可以了(不然每一列都会创建一行就会出现有多少列就创建多少个隔行的情况)
                    Sheet sheet = cell.getSheet();
                    sheet.createRow(sheet.getLastRowNum() + 1);//想创建多个隔行就重复多条该语句即可
                    sheet.createRow(sheet.getLastRowNum() + 1);
                }
            }
        };
        CellWriteCallBack cellWriteCallBack2 = (context,cell,value,writeCellStyle) -> {
            if (cell.getColumnIndex() == 1 && "刘备4".equals(value)) {
                writeCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                writeCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
            }
        };
        List<CellWriteCallBack> CellWriteHandlerList = Stream.of(cellWriteCallBack1, cellWriteCallBack2).collect(Collectors.toList());
        excelSheetElement.setCellWriteCallBacks(CellWriteHandlerList);
        /*excelSheetElement.setFunction(function);
        List<List<String>> headList = new ArrayList<>();
        Arrays.asList("排行","名称","字").forEach(h->headList.add(Collections.singletonList(h)));
        excelSheetElement.setHead(headList);*/
        excelSheetElement.setSheetName("桃园三兄弟");
        List<ExcelSheetElement> sheetList = new ArrayList<>();
        sheetList.add(excelSheetElement);
        //EasyExcelUtil.exportWebExcel(sheetList,request, response,"test");//此处是下载单个excel
        //此处往下是多个excel打包成zip下载
        List<List<ExcelSheetElement>> excelList = new ArrayList<>();
        excelList.add(sheetList);
        List<String> excelFileNames=Stream.of("第一个").collect(Collectors.toList());
        EasyExcelUtil.exportWebExcelInZip(excelList,excelFileNames, response,"test");
    }
    public  List<Test> setWebExcelData(InputStream inputStream){
        Consumer<Test> consumer= test->{
            testMapper.insert(test);
        };
        List<Test> fromExcel = EasyExcelUtil.getFromExcel(inputStream, Test.class, consumer,false);
        return fromExcel;
    }
    public  List<ExcelSheetElement> setWebExcelDatas(InputStream inputStream){
        //以下被注释掉的代码是自定义接收类型的代码形式
        Consumer<Test> consumer= test->{
            testMapper.insert(test);
        };
        /*Consumer<LinkedHashMap<Integer,String>> defaultConsumer= map->{
            System.out.println(map.get(1));
        };*/
        List<ExcelSheetElement> sheetElements = new ArrayList<>();
        ExcelSheetElement sheetElement1=new ExcelSheetElement(Test.class,consumer);
        //ExcelSheetElement sheetElement1=new ExcelSheetElement();
        sheetElements.add(sheetElement1);
        ExcelSheetElement sheetElement2=new ExcelSheetElement(Test.class,consumer);
        //ExcelSheetElement sheetElement2=new ExcelSheetElement();
        sheetElements.add(sheetElement2);
        List<ExcelSheetElement> fromExcel = EasyExcelUtil.getFromExcel(inputStream, sheetElements, null,null,null);
        //List<ExcelSheetElement> fromExcel = EasyExcelUtil.getFromExcel(inputStream, sheetElements, null,defaultConsumer,Arrays.asList("排行","名称","字","生日","介绍"));
        return fromExcel;
    }
    public void fillExcelData(HttpServletRequest request,HttpServletResponse response){
        List<Object> fillDatas = new ArrayList<>();
        HashMap<String, String> otherData = new HashMap<>();
        otherData.put("date", "2020-03-14");
        otherData.put("total", "100");
        fillDatas.add(otherData);
        List<RowAndColSpan> splist = new ArrayList<>();
        fillDatas.add(initFillData(splist));
        ExcelFillElement excelFillElement = new ExcelFillElement("D:/ideaworkspace/excel/fill_data_template1.xlsx",fillDatas,false,splist);
        excelFillElement.setCellWriteCallBack((context,cell,value,writeCellStyle) -> {
            if(Objects.isNull(context.getRelativeRowIndex())){
                return;
            }
            final Cell c=cell;
            final RowAndColSpan rowAndColSpan=new RowAndColSpan();
            boolean isTargetRow = splist.stream().anyMatch(adc -> {
                rowAndColSpan.setColSpanStart(adc.getColSpanStart());
                rowAndColSpan.setColSpanEnd(adc.getColSpanEnd());
                return c.getRowIndex() == adc.getRowSpanStart();
            });
            if(isTargetRow){
                Row row=cell.getRow();
                writeCellStyle.setBorderBottom(BorderStyle.THIN);//设置底边框;
                writeCellStyle.setBottomBorderColor((short) 0);//设置底边框颜色;
                writeCellStyle.setBorderLeft(BorderStyle.THIN);  //设置左边框;
                writeCellStyle.setLeftBorderColor((short) 0);//设置左边框颜色;
                writeCellStyle.setBorderRight(BorderStyle.THIN);//设置右边框;
                writeCellStyle.setRightBorderColor((short) 0);//设置右边框颜色;
                writeCellStyle.setBorderTop(BorderStyle.THIN);//设置顶边框;
                writeCellStyle.setTopBorderColor((short) 0); //设置顶边框颜色;
                writeCellStyle.setWrapped(true);  //设置自动换行;
                writeCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);//设置水平对齐的样式为居中对齐;
                writeCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);  //设置垂直对齐的样式为居中对齐;
                writeCellStyle.setShrinkToFit(true);//设置文本收缩至合适
                CellStyle cellStyle = CellStyleUtil.getCellStyle(cell,writeCellStyle);
                for (int j = rowAndColSpan.getColSpanStart(); j <rowAndColSpan.getColSpanEnd(); j++) {
                    cell = row.createCell(j+1);//已经存在的列保持不变，针对改了后面的列进行新建，因此+1
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue(" ");
                }
            }
        });
        EasyExcelUtil.getWebExcelByfillWithTemplate("组合数据填充",excelFillElement,request,response);
    }
    private static List<FillData> initFillData(List<RowAndColSpan> splist) {
        ArrayList<FillData> fillDatas = new ArrayList<FillData>();
        for (int i = 0; i <10; i++) {
            FillData fillData = new FillData();
            fillData.setName("0" + i);
            fillData.setAge(10 + i);
            fillDatas.add(fillData);
            if(i>0) {
                int row = i + 2;//去掉模板中已经手动合并的一行，即2所以i从3(第四行)开始合并
                RowAndColSpan rc = new RowAndColSpan(row, row, 1, 2);
                splist.add(rc);
            }
        }
        return fillDatas;
    }
}
