package com.example.demo;

import com.example.demo.utils.ExcelUtil;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.*;

public class ExcelUtilTests {

    public String readFilePath = "C:/Users/Administrator/Desktop/test.xlsx";
    public String writeFilePath = "C:/Users/Administrator/Desktop/test2.xlsx";

    @Test
    public void readExcelDefault() {
        try {
            List<Map<String, Object>> list = ExcelUtil.readExcel(readFilePath);
            System.out.println(list);
        } catch (IOException e) {
            System.err.println("读取excel异常: "+e.getMessage());
        }
    }

    @Test
    public void readExcelBySheetIndex() {
        try {
            List<Map<String, Object>> list = ExcelUtil.readExcel(readFilePath, 1); // 第二个sheet
            System.out.println(list);
        } catch (IOException e) {
            System.err.println("读取excel异常: "+e.getMessage());
        }
    }

    @Test
    public void readExcelBySheetName() {
        try {
            List<Map<String, Object>> list4 = ExcelUtil.readExcel(readFilePath, "sheet3");
            System.out.println(list4);
            List<Map<String, Object>> list5 = ExcelUtil.readExcel(readFilePath, "sheet4");
            System.out.println(list5);
        } catch (IOException e) {
            System.err.println("读取excel异常: "+e.getMessage());
        }
    }

    @Test
    public void writeExcelDefault() {
        try {
            List<Map<String, Object>> list = ExcelUtil.readExcel(readFilePath);
            LinkedHashMap<String, String> headRowMap = new LinkedHashMap<>();
            headRowMap.put("keyA", "第A列");
            headRowMap.put("keyB", "第B列");
            headRowMap.put("keyC", "第C列");
            headRowMap.put("KeyD", "第D列");
            headRowMap.put("KeyE", "第E列");
            headRowMap.put("KeyF", "第F列");
            ExcelUtil.writeExcel(writeFilePath, headRowMap, list);
        } catch (IOException e) {
            System.err.println("写入excel异常: "+e.getMessage());
        }
    }

    @Test
    public void writeExcelBySheetName() {
        try {
            List<Map<String, Object>> list = ExcelUtil.readExcel(readFilePath);
            LinkedHashMap<String, String> headRowMap = new LinkedHashMap<>();
            headRowMap.put("KeyD", "第D列");
            headRowMap.put("KeyF", "第F列");
            headRowMap.put("keyB", "第B列");
            headRowMap.put("keyC", "第C列");
            headRowMap.put("keyA", "第A列");
            headRowMap.put("KeyE", "第E列");
            ExcelUtil.writeExcel(writeFilePath, "报表", headRowMap, list);
        } catch (IOException e) {
            System.err.println("写入excel异常: "+e.getMessage());
        }
    }

}
