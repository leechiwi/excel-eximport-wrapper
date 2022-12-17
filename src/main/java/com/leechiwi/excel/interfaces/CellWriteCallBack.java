package com.leechiwi.excel.interfaces;

import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import org.apache.poi.ss.usermodel.Cell;

@FunctionalInterface
public interface CellWriteCallBack {
    public void callBack(CellWriteHandlerContext context,Cell cell, String cellValue, WriteCellStyle writeCellStyle);
}
