package com.leechiwi.excel.handler;

import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.leechiwi.excel.style.DefaultCellStyle;

public class DefaultCellStyleHandler extends HorizontalCellStyleStrategy {
    public DefaultCellStyleHandler(){
        this(DefaultCellStyle.getHeadStyle(), DefaultCellStyle.getContentStyle());
    }
    public DefaultCellStyleHandler(WriteCellStyle headCellStyle, WriteCellStyle contentCellStyle){
        super(headCellStyle,contentCellStyle);
    }
}
