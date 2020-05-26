package com.example.minalz.minalzcloudtest.utils;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelUtils {
  
    /**     
     * @Description: 导出Excel 
     * @param workbook  
     * @param sheetNum (sheet的位置，0表示第一个表格中的第一个sheet) 
     * @param sheetTitle  （sheet的名称） 
     * @param headers    （表格的列标题） 
     * @param result   （表格的数据） 
     * @param out  （输出流） 
     * @throws Exception 
     */  
    public void exportExcel(HSSFWorkbook workbook, int sheetNum,
                            String sheetTitle, String[] headers, List<List<String>> result,
                            OutputStream out) throws Exception {
        // 生成一个表格  
        HSSFSheet sheet = workbook.createSheet();
        workbook.setSheetName(sheetNum, sheetTitle);  
        // 设置表格默认列宽度为20个字节  
        sheet.setDefaultColumnWidth((short) 20);  
        // 生成一个样式  
        HSSFCellStyle style = workbook.createCellStyle();
        // 设置这些样式  
        style.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);  
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);  
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);  
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);  
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);  
        // 生成一个字体  
        HSSFFont font = workbook.createFont();
        font.setColor(HSSFColor.BLACK.index);  
        font.setFontHeightInPoints((short) 12);  
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);  
        // 把字体应用到当前的样式  
        style.setFont(font);  
  
        // 指定当单元格内容显示不下时自动换行  
        style.setWrapText(true);  
  
        // 产生表格标题行  
        HSSFRow row = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {  
            HSSFCell cell = row.createCell((short) i);  
          
            cell.setCellStyle(style);  
            HSSFRichTextString text = new HSSFRichTextString(headers[i]);  
            cell.setCellValue(text.toString());  
        }  
        // 遍历集合数据，产生数据行  
        if (result != null) {
            int index = 1;
            for (List<String> m : result) {
                row = sheet.createRow(index);
                int cellIndex = 0;
                for (String str : m) {
                    HSSFCell cell = row.createCell((short) cellIndex);
                    cell.setCellValue(str.toString());
                    cellIndex++;
                }
                index++;
            }
        }
    }

    public static void main(String[] args) {
        String fileName = "到货数量统计1.xlsx";
        //excel导出的路径和名称
        OutputStream out = null;
        try {
            out = new FileOutputStream("/Users/zhouwei/Desktop/temp/" + fileName);

            List<Map<String,String>> sheetList1 = new ArrayList<>();
            List<Map<String,String>> sheetList2 = new ArrayList<>();
            List<Map<String,String>> sheetList3 = new ArrayList<>();

            Map<String,String> itemMap = new HashMap<>();
            List<String> headers = new ArrayList<String>(){{
                add("客户");
                add("物料");
                add("Attribute");
                add("DOI类别");
                add("Unit Of Measure");
                add("4/25/20-5/1/20");
                add("5/2/20-5/8/20");
                add("5/9/20-5/15/20");
                add("5/16/20-5/22/20");
                add("5/23/20-5/29/20");
                add("5/30/20-6/5/20");
                add("6/6/20-6/12/20");
                add("6/13/20-6/19/20");
                add("6/20/20-6/26/20");
                add("6/27/20-7/3/20");
                add("7/4/20-7/10/20");
                add("7/11/20-7/17/20");
                add("7/18/20-7/24/20");
                add("7/25/20-7/31/20");
            }};


            for (int j = 0; j < 10; j++) {
                itemMap = new HashMap<>();
                for (int i = 0; i < headers.size(); i++) {
                    itemMap.put(headers.get(i),j+"word值是否了解了"+headers.get(i) + "i");
                }
                sheetList1.add(itemMap);
                sheetList2.add(itemMap);
                sheetList3.add(itemMap);
            }

            ExcelUtils eeu = new ExcelUtils();
            HSSFWorkbook workbook = new HSSFWorkbook();
            eeu.exportExcel(workbook, 0, "最大DOI数据", headers, sheetList1, out);
            eeu.exportExcel(workbook, 1, "安全DOI数据", headers, sheetList1, out);
            eeu.exportExcel(workbook, 2, "目标DOI数据", headers, sheetList2, out);
            //将所有的数据一起写入，然后再关闭输入流。
            workbook.write(out);
            System.out.println("导出结束");
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            if(out != null){
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public void exportExcel(HSSFWorkbook workbook, int sheetNum, String sheetTitle, List<String> headers, List<Map<String, String>> result, OutputStream out) {
        // 生成一个表格
        HSSFSheet sheet = workbook.createSheet();
        workbook.setSheetName(sheetNum, sheetTitle);
        // 设置表格默认列宽度为20个字节
        sheet.setDefaultColumnWidth((short) 20);
        // 生成一个样式
        HSSFCellStyle style = workbook.createCellStyle();
        // 设置这些样式
        style.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 生成一个字体
        HSSFFont font = workbook.createFont();
        font.setColor(HSSFColor.BLACK.index);
        font.setFontHeightInPoints((short) 12);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        // 把字体应用到当前的样式
        style.setFont(font);

        // 指定当单元格内容显示不下时自动换行
        style.setWrapText(true);

        // 产生表格标题行
        HSSFRow row = sheet.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            HSSFCell cell = row.createCell((short) i);

            cell.setCellStyle(style);
            HSSFRichTextString text = new HSSFRichTextString(headers.get(i));
            cell.setCellValue(text.toString());
        }
        // 遍历集合数据，产生数据行
        if (result != null) {
            int index = 1;
            for (Map<String,String> m : result) {
                row = sheet.createRow(index);
                int cellIndex = 0;
                for (int i = 0; i < headers.size(); i++) {
                    String val = m.get(headers.get(i));
                    HSSFCell cell = row.createCell((short) cellIndex);
                    cell.setCellValue(val);
                    cellIndex++;
                }
//                for (String str : m) {
//                    HSSFCell cell = row.createCell((short) cellIndex);
//                    cell.setCellValue(str.toString());
//                    cellIndex++;
//                }
                index++;
            }
        }
    }
}  