package com.leechiwi.excel.handler;

import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.CellData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.style.column.AbstractColumnWidthStyleStrategy;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;

import java.util.ArrayList;
import java.util.List;

public class DefaultWidthStyleStrategy extends AbstractColumnWidthStyleStrategy {
    private static final int MAX_COLUMN_WIDTH = 50;
    private List <Integer> headColumnWidths=new ArrayList<>();//此处代表各表头列对应宽度
    @Override
    protected void setColumnWidth(WriteSheetHolder writeSheetHolder, List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        boolean needSetWidth = isHead || !CollectionUtils.isEmpty(cellDataList);
        if (needSetWidth) {
            Integer columnWidth = this.dataLength(cellDataList, cell, isHead);
            if (columnWidth >= 0) {
                if (columnWidth > MAX_COLUMN_WIDTH) {
                    columnWidth = MAX_COLUMN_WIDTH;
                }
                Integer currentHeadWidth = this.headColumnWidths.get(cell.getColumnIndex());
                if (columnWidth < currentHeadWidth) {//内容长度小于对应表头列的长度则以表头列长度为准，否则表头别里面内容会折行
                    columnWidth = currentHeadWidth;
                }
                writeSheetHolder.getSheet().setColumnWidth(cell.getColumnIndex(), columnWidth * 256);
            }
        }
    }
    private Integer dataLength(List<?> cellDataList, Cell cell, Boolean isHead) {
        if (isHead) {
            int length = cell.getStringCellValue().getBytes().length;
            if (this.headColumnWidths.size()>cell.getColumnIndex()) {//获取表头每列的最长宽度（保证导出时候表头每一列都展示完整且不折行）
                Integer currentHeadWidth =this.headColumnWidths.get(cell.getColumnIndex());
                if(length>currentHeadWidth){
                    this.headColumnWidths.set(cell.getColumnIndex(),length);
                }
            }else{
                this.headColumnWidths.add(length);
            }
            return length;
        } else {
            CellData cellData = (CellData)cellDataList.get(0);
            CellDataTypeEnum type = cellData.getType();
            if (type == null) {
                return -1;
            } else {
                switch(type) {
                    case STRING:
                        return cellData.getStringValue().getBytes().length;
                    case BOOLEAN:
                        return cellData.getBooleanValue().toString().getBytes().length;
                    case NUMBER:
                        return cellData.getNumberValue().toString().getBytes().length;
                    default:
                        return -1;
                }
            }
        }
    }
}
