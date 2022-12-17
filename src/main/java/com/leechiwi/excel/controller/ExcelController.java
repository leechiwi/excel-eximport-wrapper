package com.leechiwi.excel.controller;

import com.leechiwi.excel.model.ExcelSheetElement;
import com.leechiwi.excel.pojo.Test;
import com.leechiwi.excel.service.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

@Controller
public class ExcelController {
    @Autowired
    private ExcelService excelService;
    @GetMapping("/export")
    public void export(HttpServletRequest request, HttpServletResponse response,String ids){
        excelService.getWebExcelData(request,response);
    }
    @GetMapping("/import")
    @ResponseBody
    public List<Test> imports(MultipartFile file){
        List<Test> tests=new ArrayList<>();
        try {
            InputStream inputStream = null;
            if(Objects.nonNull(file)){
                inputStream=file.getInputStream();
            }
            inputStream = new FileInputStream(new File("D:/ideaworkspace/excel/test.xlsx"));
            tests = excelService.setWebExcelData(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return tests;
    }
    @GetMapping("/imports")
    @ResponseBody
    public  List<ExcelSheetElement> importss(MultipartFile file){
        List<ExcelSheetElement> result = new ArrayList<>();
        try {
            InputStream inputStream = null;
            if(Objects.nonNull(file)){
                inputStream=file.getInputStream();
            }
            inputStream = new FileInputStream(new File("D:/ideaworkspace/excel/tests.xlsx"));
            result = excelService.setWebExcelDatas(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return result;
    }
    @GetMapping("/fill")
    public void fill(HttpServletRequest request, HttpServletResponse response,String ids){
        excelService.fillExcelData(request, response);
    }
}
