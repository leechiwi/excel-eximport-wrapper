package com.leechiwi.excel.handler;

import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.util.StyleUtil;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.leechiwi.excel.interfaces.CellWriteCallBack;
import org.apache.poi.ss.usermodel.*;

import java.util.List;
import java.util.function.Consumer;

/**
 * 自定义拦截器（指定样式）
 */
public class DefaultCellWriteHandler implements CellWriteHandler {
    private CellWriteCallBack cellWriteCallBack;
    public DefaultCellWriteHandler(){

    }
    public DefaultCellWriteHandler(CellWriteCallBack cellWriteCallBack) {
        this.cellWriteCallBack = cellWriteCallBack;
    }

    /**
     * 在单元上的所有操作完成后调用,此方法在3.0版本设置不生效，改为afterCellDispose(CellWriteHandlerContext context)
     * @param writeSheetHolder
     * @param writeTableHolder
     * @param cellDataList
     * @param cell
     * @param head
     * @param relativeRowIndex
     * @param isHead
     */

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
       /* Workbook workbook = cell.getSheet().getWorkbook();
        //CellStyle cellStyle = workbook.createCellStyle();
        CellStyle cellStyle = cell.getCellStyle();
        Sheet sheet = writeSheetHolder.getSheet();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellStyle.setBorderTop(BorderStyle.THIN);//上边框
        cellStyle.setBorderRight(BorderStyle.THIN);//右边框
        //Row row = sheet.getRow(cell.getRowIndex());
        //cell = row.createCell(cell.getColumnIndex());
        //cell.setCellStyle(cellStyle);
*//*        WriteCellStyle writeCellStyle=new WriteCellStyle();
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
        CellStyle cellStyle = StyleUtil.buildCellStyle(workbook, cell.getCellStyle(), writeCellStyle);*//*
        //cell.getRow().getCell(cell.getColumnIndex()).setCellStyle(cellStyle);
        cell.getSheet().getRow(cell.getRowIndex()).getCell(cell.getColumnIndex()).setCellStyle(cellStyle);*/
    }

    /**
     * 针对某一个单元格进行样式设置(当前事件会在 数据设置到poi的cell里面才会回调)
     * @param context
     */
    @Override
    public void afterCellDispose(CellWriteHandlerContext context) {
        if(!context.getHead()){
            Cell cell = context.getCell();
            //获取第一个单元格对象
            // 只要不是头 一定会有数据 当然fill(填充)的情况 可能要context.getCellDataList()
            // 这个需要看模板，因为一个单元格会有多个 WriteCellData
            WriteCellData<?> cellData = context.getFirstCellData();
            String data = this.getData(cellData);
            // cellData 可以获取样式/数据，也可以直接设置样式/数据，设置后会立即生效
            // 这里也需要用cellData去获取样式
            // 很重要的一个原因是 WriteCellStyle 和 dataFormatData绑定的
            // 简单的说 比如你加了 DateTimeFormat，已经将writeCellStyle里面的dataFormatData改了 			  // 如果你自己new了一个WriteCellStyle，可能注解的样式就失效了
            // 然后getOrCreateStyle 用于返回一个样式，如果为空，则创建一个后返回
            // （总之记住用这个方法获取样式即可）
            WriteCellStyle writeCellStyle = cellData.getOrCreateStyle();
            this.cellWriteCallBack.callBack(context,cell,data,writeCellStyle);
            /*if(cell.getColumnIndex()==1&&"关羽".equals(data)) {

                writeCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                // 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND
                // 要不然背景色不会生效
                writeCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);// 这样样式就设置好了 后面有个FillStyleCellWriteHandler 默认会将 WriteCellStyle 设置			 到 cell里面去 所以可以不用管了
                // 获取当前单元格的数据，必要的时候可以根据数据来设置单元格的颜色
                // 比如当前列为状态列，数据为1是正常，背景色为绿色，反正不正常，背景色设置为红
                // 这种需求的实现将变得可能
            }*/
        }
    }
    private String getData(WriteCellData<?> cellData){
        CellDataTypeEnum type=cellData.getType();
        switch(type) {
            case STRING:
                return cellData.getStringValue();
            case BOOLEAN:
                return cellData.getBooleanValue().toString();
            case NUMBER:
                return cellData.getNumberValue().toString();
            default:
                return "-1";
        }
    }
}
