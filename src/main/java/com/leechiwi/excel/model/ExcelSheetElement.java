package com.leechiwi.excel.model;

import com.leechiwi.excel.interfaces.CellWriteCallBack;

import java.util.List;
import java.util.function.Consumer;
import java.util.function.Function;

public class ExcelSheetElement {
    private String sheetName;//sheet的名称
    private Function<List<Object>,List<List<Object>>> function;//自定义数据生成逻辑实现，针对data进行处理成最终sheet上要展示的数据
    private List<List<Object>> data;//原始数据（比如从数据库中查询得到的数据集合），可通过function的实现去自定义数据生成逻辑
    private List<List<String>> head;//自定义表头
    private List<RowAndColSpan> spanList;//数据跨行跨列数据设置
    private List<CellWriteCallBack> cellWriteCallBacks;//自定义单元格样式处理实现类
    private Class<?> clazz;//接收导入的实体类
    private Consumer<?> consumer;//导入接收处理逻辑
    public String getSheetName() {
        return sheetName;
    }

    public ExcelSheetElement() {

    }
    public ExcelSheetElement(Class<?> clazz, Consumer<?> consumer) {
        this.clazz = clazz;
        this.consumer = consumer;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public Function<List<Object>, List<List<Object>>> getFunction() {
        return function;
    }

    public void setFunction(Function<List<Object>, List<List<Object>>> function) {
        this.function = function;
    }

    public List<List<Object>> getData() {
        return data;
    }

    public void setData(List<List<Object>> data) {
        this.data = data;
    }

    public List<List<String>> getHead() {
        return head;
    }

    public void setHead(List<List<String>> head) {
        this.head = head;
    }

    public List<RowAndColSpan> getSpanList() {
        return spanList;
    }

    public void setSpanList(List<RowAndColSpan> spanList) {
        this.spanList = spanList;
    }

    public List<CellWriteCallBack> getCellWriteCallBacks() {
        return cellWriteCallBacks;
    }

    public void setCellWriteCallBacks(List<CellWriteCallBack> cellWriteCallBacks) {
        this.cellWriteCallBacks = cellWriteCallBacks;
    }

    public Class<?> getClazz() {
        return clazz;
    }

    public void setClazz(Class<?> clazz) {
        this.clazz = clazz;
    }

    public Consumer<?> getConsumer() {
        return consumer;
    }

    public void setConsumer(Consumer<?> consumer) {
        this.consumer = consumer;
    }
}
