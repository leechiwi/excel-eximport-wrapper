package com.leechiwi.excel.style;

import com.alibaba.excel.util.StyleUtil;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public class CellStyleUtil {
    public static CellStyle getCellStyle(Cell cell, WriteCellStyle writeCellStyle){
        Workbook workbook = cell.getSheet().getWorkbook();
        CellStyle cellStyle = StyleUtil.buildCellStyle(workbook, cell.getCellStyle(), writeCellStyle);
        return cellStyle;
    }
}
