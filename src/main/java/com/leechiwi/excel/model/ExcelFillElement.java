package com.leechiwi.excel.model;

import com.leechiwi.excel.interfaces.CellWriteCallBack;

import java.util.List;

public class ExcelFillElement {
    private Object file;//模板文件
    List<Object> fillDatas;//填充的数据
    boolean templateExsit;//模板文件是否是项目内部的模板
    private List<RowAndColSpan> spanList;//数据跨行跨列数据设置
    private CellWriteCallBack cellWriteCallBack;//自定义单元格样式处理实现类
    private boolean horizontal;//是否横向扩展填充

    public ExcelFillElement(Object file, List<Object> fillDatas, boolean templateExsit, List<RowAndColSpan> spanList) {
        this.file = file;
        this.fillDatas = fillDatas;
        this.templateExsit = templateExsit;
        this.spanList = spanList;
    }

    public Object getFile() {
        return file;
    }

    public void setFile(Object file) {
        this.file = file;
    }

    public List<Object> getFillDatas() {
        return fillDatas;
    }

    public void setFillDatas(List<Object> fillDatas) {
        this.fillDatas = fillDatas;
    }

    public boolean isTemplateExsit() {
        return templateExsit;
    }

    public void setTemplateExsit(boolean templateExsit) {
        this.templateExsit = templateExsit;
    }

    public List<RowAndColSpan> getSpanList() {
        return spanList;
    }

    public void setSpanList(List<RowAndColSpan> spanList) {
        this.spanList = spanList;
    }

    public CellWriteCallBack getCellWriteCallBack() {
        return cellWriteCallBack;
    }

    public void setCellWriteCallBack(CellWriteCallBack cellWriteCallBack) {
        this.cellWriteCallBack = cellWriteCallBack;
    }

    public boolean isHorizontal() {
        return horizontal;
    }

    public void setHorizontal(boolean horizontal) {
        this.horizontal = horizontal;
    }
}
