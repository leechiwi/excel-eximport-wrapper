package com.leechiwi.excel.model;

public class RowAndColSpan {
    private Integer rowSpanStart;
    private Integer rowSpanEnd;
    private Integer colSpanStart;
    private Integer colSpanEnd;
    public RowAndColSpan(){

    }
    public RowAndColSpan(Integer rowSpanStart, Integer rowSpanEnd, Integer colSpanStart, Integer colSpanEnd) {
        this.rowSpanStart = rowSpanStart;
        this.rowSpanEnd = rowSpanEnd;
        this.colSpanStart = colSpanStart;
        this.colSpanEnd = colSpanEnd;
    }

    public Integer getRowSpanStart() {
        return rowSpanStart;
    }

    public void setRowSpanStart(Integer rowSpanStart) {
        this.rowSpanStart = rowSpanStart;
    }

    public Integer getRowSpanEnd() {
        return rowSpanEnd;
    }

    public void setRowSpanEnd(Integer rowSpanEnd) {
        this.rowSpanEnd = rowSpanEnd;
    }

    public Integer getColSpanStart() {
        return colSpanStart;
    }

    public void setColSpanStart(Integer colSpanStart) {
        this.colSpanStart = colSpanStart;
    }

    public Integer getColSpanEnd() {
        return colSpanEnd;
    }

    public void setColSpanEnd(Integer colSpanEnd) {
        this.colSpanEnd = colSpanEnd;
    }
}
