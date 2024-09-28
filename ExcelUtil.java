package com.example.demo.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * excel操作工具类，依赖poi、poi-ooxml包(v4.1.2)
 * @author LiHK
 */
public class ExcelUtil {

    /**
     * 工具类无需实例化
     */
    private ExcelUtil() {}

    private static final SimpleDateFormat defaultSdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

    /**
     * 读取excel文件，支持xlsx、xls格式，仅读取第一个Sheet。
     *
     * @param filePath 文件绝对路径
     * @return {@link List}
     * @throws IOException 包括{@link FileNotFoundException}
     * @see #readSheet(Sheet)
     */
    public static List<Map<String, Object>> readExcel(String filePath) throws IOException {
        return readExcel(filePath, 0); // 默认读取第一个Sheet
    }

    /**
     * 读取excel文件，支持xlsx、xls格式，指定Sheet名称。
     * @see #readSheet(Sheet)
     * @param filePath 文件路径
     * @param sheetName 指定的Sheet名称
     * @return {@link List}
     * @throws IOException 包括{@link FileNotFoundException}
     */
    public static List<Map<String, Object>> readExcel(String filePath, String sheetName) throws IOException {
        if (isBlank(filePath)) {
            throw new IllegalArgumentException("filePath is null/empty");
        }
        if (isBlank(sheetName)) {
            throw new IllegalArgumentException("sheetName is null/empty");
        }

        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = null;

        try {
            workbook = getWorkbook(filePath, fis);
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                return new ArrayList<>();
            }
            return readSheet(sheet);
        } finally {
            close(workbook, fis);
        }
    }

    /**
     * 读取excel文件，支持xlsx、xls格式，指定第几个Sheet。
     * @see #readSheet(Sheet)
     * @param filePath 文件绝对路径
     * @param sheetIndex 读取第几个Sheet
     * @return {@link List}
     * @throws IOException 包括{@link FileNotFoundException}
     */
    public static List<Map<String, Object>> readExcel(String filePath, int sheetIndex) throws IOException {
        if (isBlank(filePath)) {
            throw new IllegalArgumentException("filePath is null/empty");
        }

        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = null;

        try {
            workbook = getWorkbook(filePath, fis);
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            return readSheet(sheet);
        } finally {
            close(workbook, fis);
        }
    }

    /**
     * 将数据写入excel文件(xlsx格式)，需提供表头映射关系
     * @param filePath 指定的文件路径
     * @param headRowMap 表头映射关系
     * @param dataList 待写入的数据
     * @throws IOException 包括{@link FileNotFoundException}
     */
    public static void writeExcel(String filePath, LinkedHashMap<String, String> headRowMap,
                                  List<Map<String, Object>> dataList) throws IOException {
        writeExcel(filePath, "Sheet1", headRowMap, dataList);
    }

    /**
     * 将数据写入excel文件(xlsx格式)，需提供表头映射关系
     * @param filePath 指定的文件路径
     * @param sheetName Sheet名称
     * @param headRowMap 表头映射关系
     * @param dataList 待写入的数据
     * @throws IOException 包括{@link FileNotFoundException}
     */
    public static void writeExcel(String filePath, String sheetName, LinkedHashMap<String, String> headRowMap,
                                  List<Map<String, Object>> dataList) throws IOException {
        if (isBlank(filePath)) {
            throw new IllegalArgumentException("filePath is null/empty");
        }
        if (headRowMap == null || headRowMap.isEmpty()) {
            throw new IllegalArgumentException("headRowMap is null/empty");
        }
        if (isBlank(sheetName)) {
            sheetName = "Sheet1";
        }

        // 写入内存
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(sheetName);
        CellStyle cellStyle = setDefaultStyle(workbook);
        writeSheet(sheet, cellStyle, headRowMap, dataList);

        // 写入硬盘
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(filePath);
            workbook.write(fos);
        } finally {
            close(fos, workbook);
        }
    }

    /**
     * 读取Sheet里的数据。默认首行用作标题行，首行为空将获取不到数据
     * @param sheet excel里指定的Sheet
     * @return {@link List}
     */
    private static List<Map<String, Object>> readSheet(Sheet sheet) {
        List<Map<String, Object>> resList = new ArrayList<>();
        Row headRow = sheet.getRow(0); // 标题行用作对象key
        if (headRow == null) {
            return resList;
        }

        // 从第二行开始，处理每一行
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue; // 跳过空行，继续读取后面的行数据
            }
            Map<String, Object> rowMap = new HashMap<>();
            // 处理每个单元格
            for (int j = 0; j < row.getLastCellNum(); j++) {
                Cell headCell = headRow.getCell(j);
                if (headCell == null) {
                    continue; // 跳过空列，继续读取后面的列数据
                }
                String key = headCell.getStringCellValue();
                Object value = getCellValue(row.getCell(j));
                rowMap.put(key, value); // 将读取到的数据封装成对象
            }
            resList.add(rowMap);
        }
        return resList;
    }

    /**
     * 数据写入指定Sheet
     * @param sheet 指定的Sheet
     * @param cellStyle 单元格样式
     * @param headRowMap 表头映射关系
     * @param dataList 待写入数据
     */
    private static void writeSheet(Sheet sheet, CellStyle cellStyle, LinkedHashMap<String, String> headRowMap,
                                   List<Map<String, Object>> dataList) {
        createHeadRow(sheet, cellStyle, headRowMap); // 首行

        // 处理后面具体的数据
        for (int i = 0; i < dataList.size(); i++) {
            Map<String, Object> dataMap = dataList.get(i);
            Object[] headNames = headRowMap.keySet().toArray();
            Row row = sheet.createRow(i + 1); // 跳过首行
            for (int j = 0; j < headNames.length; j++) {
                Cell cell = row.createCell(j);
                cell.setCellStyle(cellStyle);
                Object value = dataMap.get((String) headNames[j]);
                // System.out.printf("第%d行，第%d列，value:%s%n", i + 2, j + 1, value);
                setCellValue(cell, value);
            }
        }
    }

    /**
     * 根据单元格数据类型返回对应数据
     * @param cell 列单元格
     * @return {@link Object}
     */
    private static Object getCellValue(Cell cell) {
        if (cell == null) {
            return null; // 跳过空单元格，继续读取后面的单元格数据
        }

        Object value = null;
        switch (cell.getCellType()) {
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) { // 日期类型
                    value = cell.getDateCellValue();
                } else {
                    value = cell.getNumericCellValue();
                }
                break;
            case STRING:
                value = cell.getStringCellValue();
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case ERROR:
                value = cell.getErrorCellValue();
                break;
            case FORMULA: // 公式
                value = "=" + cell.getCellFormula();
                break;
            case BLANK:
                break;
            default:
                value = "[无法识别]";
        }
        return value;
    }

    /**
     * 给单元格设置数据
     * @param cell 指定的单元格
     * @param value 具体的数据
     */
    private static void setCellValue(Cell cell, Object value) {
        if (value == null) {
            cell.setBlank();
        } else if (value instanceof Number) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Date) {
            cell.setCellValue(defaultSdf.format((Date) value));
        } else {
            cell.setCellValue(value.toString());
        }
    }

    /**
     * 根据文件格式，创建工作簿
     * @param filePath 文件路径
     * @param fis 文件流
     * @return {@link Workbook}
     * @throws IOException IO操作异常
     */
    private static Workbook getWorkbook(String filePath, FileInputStream fis) throws IOException {
        if (filePath.endsWith(".xlsx")) {
            return new XSSFWorkbook(fis);
        }
        if (filePath.endsWith(".xls")) {
            return new HSSFWorkbook(fis);
        }
        throw new RuntimeException("invalid file format");
    }

    /**
     * 设置默认单元格样式
     * @param workbook 指定工作簿
     * @return {@link CellStyle}
     */
    private static CellStyle setDefaultStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    /**
     * 根据传入的表头(映射关系)创建收首行
     * @param sheet 指定的Sheet
     * @param cellStyle 单元格样式
     * @param headRowMap 表头映射关系
     */
    private static void createHeadRow(Sheet sheet, CellStyle cellStyle, LinkedHashMap<String, String> headRowMap) {
        Row row = sheet.createRow(0); // 首行
        Object[] headNames = headRowMap.values().toArray();
        for (int i = 0; i < headNames.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellStyle(cellStyle);
            cell.setCellValue((String) headNames[i]);
        }
    }

    /**
     * 释放未关闭的流，注意先后关系
     * @param objs 未关闭的流
     */
    public static void close(Closeable... objs) {
        for (Closeable obj : objs) {
            if (obj != null) {
                try {
                    obj.close();
                } catch (Exception e) {
                    System.err.println(obj.getClass() + ".close() failed");
                }
            }
        }
    }

    /**
     * 字符串是否为空或空串
     * @param str 字符串
     * @return {@link Boolean}
     */
    public static boolean isBlank(String str) {
        return str == null || str.isEmpty();
    }

}
