package com.leechiwi.excel.handler;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.merge.AbstractMergeStrategy;
import com.leechiwi.excel.model.RowAndColSpan;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

public class DefaultMergeStrategy extends AbstractMergeStrategy {
    private List<RowAndColSpan> spanList;
    private Map<String,RowAndColSpan> rowIndexMap;
    public DefaultMergeStrategy(){

    }
    public DefaultMergeStrategy(List<RowAndColSpan> spanList){
        this.spanList = spanList;
        rowIndexMap=new HashMap<>();
        for (RowAndColSpan rowAndColSpan : spanList) {
            StringBuilder position = new StringBuilder();
            position.append(rowAndColSpan.getRowSpanStart());
            position.append(";").append(rowAndColSpan.getColSpanStart());
            rowIndexMap.put(position.toString(),rowAndColSpan);
        }
    }
    @Override
    protected void merge(Sheet sheet, Cell cell, Head head, Integer integer) {
        if(Objects.isNull(integer)){
            return;
        }
        StringBuilder position = new StringBuilder();
        position.append(cell.getRowIndex()).append(";").append(cell.getColumnIndex());
        RowAndColSpan rowAndColSpan = rowIndexMap.get(position.toString());
        if(Objects.nonNull(rowAndColSpan)){//判断是否是合并的行，是的话进行设置合并
            CellRangeAddress cellRangeAddress = new CellRangeAddress(rowAndColSpan.getRowSpanStart(), rowAndColSpan.getRowSpanEnd(), rowAndColSpan.getColSpanStart(), rowAndColSpan.getColSpanEnd());
            sheet.addMergedRegionUnsafe(cellRangeAddress);
        }
    }
}
