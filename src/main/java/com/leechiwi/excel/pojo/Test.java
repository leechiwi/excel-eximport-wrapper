package com.leechiwi.excel.pojo;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;

import java.util.Date;

@HeadRowHeight(value = 20)
public class Test {
    @ExcelProperty(value={"桃园结义","桃园结义","排行"},index = 0)
    @ColumnWidth(value = 10)
    private Integer index;
    @ExcelProperty(value={"桃园结义","桃园结义","名称"},index = 1)
    @ColumnWidth(value = 20)
    private String name;
    @ExcelProperty(value={"桃园结义","桃园结义","字"},index = 2)
    @ColumnWidth(value = 20)
    private String code;
    @ExcelProperty(value={"桃园结义","桃园结义","生日"},index = 3)
    @DateTimeFormat(value="yyyy-MM-dd HH:mm:ss")
    @ColumnWidth(value = 30)
    private Date birthday;
    @ExcelProperty(value={"桃园结义","桃园结义","介绍"},index = 4)
     //@ExcelIgnore
    //@ColumnWidth(value = 50)
    private String other;

    public Integer getIndex() {
        return index;
    }

    public void setIndex(Integer index) {
        this.index = index;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public Date getBirthday() {
        return birthday;
    }

    public void setBirthday(Date birthday) {
        this.birthday = birthday;
    }

    public String getOther() {
        return other;
    }

    public void setOther(String other) {
        this.other = other;
    }
}
