package com.example.minalz.minalzcloudtest.controller;

import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import io.swagger.annotations.ApiParam;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

@Api(description = "客户到货量处理")
@RestController
@RequestMapping(value = "/temp/dealing")
public class TempController {

    /**
     * 输入：将附件excel，第一个页为客户doi数据，第二个页为客户预测量数据
     * 逻辑：
     * 理想期末 = 未来（当周doi天）的预测量
     * 期初+到货 = 预测 + 期末===> 到货 = 期末 + 预测 - 期初
     * 输出：安全doi下的客户到货量，目标doi下的客户到货量，最大doi下的客户到货量；分为3个sheet页
     * 注意：此卡只是应客户对数据需求临时开发的方案，不发布测试机。
     * @return
     */
    @ApiOperation(value = "客户到货量处理-excel处理", notes = "文件路径必须填写")
    @PostMapping(value = "/excel",produces = "application/json; utf-8")
    public String dealExcel(@ApiParam(required = true,  value = "文件路径") @RequestParam(value = "filePath") String filePath,
                            HttpServletRequest request, HttpServletResponse response){
        try {
            System.out.println("文件路径 -->" + filePath);
            // 获取文件对象
            if(filePath == null){
                return "文件路径必须填写";
            }
            // 判断文件格式是否是Excel
            if (!filePath.endsWith("xlsx") && !filePath.endsWith("xls")) {
                return "文件格式不对,请上传excel文件";
            }
            // 读入Excel
            readExcel(filePath,response);

        } catch (Exception e) {
            e.printStackTrace();
        }
        String fileName = "到货处理数据.xlsx";
        String newFilePath = filePath.substring(0, filePath.lastIndexOf("\\") + 1);
        return "生成的文件路径是-->" + newFilePath + fileName;
    }

    public static void readExcel(String filePath,HttpServletResponse response) throws Exception {

        InputStream is = new FileInputStream(new File(filePath));
        Workbook hssfWorkbook = null;
        if (filePath.endsWith("xlsx")) {
            hssfWorkbook = new XSSFWorkbook(is);// Excel 2007
        } else if (filePath.endsWith("xls")) {
            hssfWorkbook = new HSSFWorkbook(is);// Excel 2003
        }
        Map<String,String> itemMap1 = new HashMap<>();
        Map<String,String> itemMap2 = new HashMap<>();
        Map<String,String> itemMap3 = new HashMap<>();
        // 存储sheet1表数据
        List<Map<String,String>> itemList = new ArrayList<>();
        // 存储sheet2表数据
        Map<String,Map<String,String>> keyItemMap2 = new HashMap();
        // 存储sheet3表数据
        Map<String,Map<String,String>> keyItemMap3 = new HashMap();
        // 存放标题行的Map
        Map<String,List<String>> titleMap = new HashMap<>();
        // 循环工作表Sheet
        for (int numSheet = 1; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
            // HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
            Sheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
            if (hssfSheet == null) {
                continue;
            }

            List<String> titleList = new ArrayList<>();
            titleMap.put("sheet"+numSheet,titleList);
            // 循环行Row
            for (int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                // 第0行 是标题行
                Row hssfRow = hssfSheet.getRow(rowNum);
                if(rowNum == 0){
                    // 该sheet有多少列
                    int columnNumber = hssfRow.getPhysicalNumberOfCells();
                    System.out.println("excel总列数 -->" + columnNumber);
                    // 标题行
                    for (int i = 0; i < columnNumber; i++) {
                        String cellValue = getStringValueFromCell(hssfRow.getCell(i));
                        titleList.add(cellValue);
                    }
                    continue;
                }
                // 数据行
                if (hssfRow != null) {
                    itemMap1 = new HashMap();
                    if(numSheet == 3){
                        itemMap2 = new HashMap();
                        itemMap3 = new HashMap();
                    }
                    for (int i = 0,len = titleList.size(); i < len; i++) {
                        String cellValue = getStringValueFromCell(hssfRow.getCell(i));
                        itemMap1.put(titleList.get(i),cellValue);
                        if(numSheet == 3){
                            itemMap2.put(titleList.get(i),cellValue);
                            itemMap3.put(titleList.get(i),cellValue);
                        }
                    }
                    if(numSheet == 1){
                        // sheet1表数据
                        itemList.add(itemMap1);
                    }else if(numSheet == 2){
                        // sheet2表数据
                        keyItemMap2.put(itemMap1.get(titleList.get(0)) + "," + itemMap1.get(titleList.get(1)),itemMap1);
                    }else if(numSheet == 3){
                        // sheet3表数据
                        keyItemMap3.put(itemMap1.get(titleList.get(0)) + "," + itemMap1.get(titleList.get(1)) + ",最大DOI",itemMap1);
                        keyItemMap3.put(itemMap2.get(titleList.get(0)) + "," + itemMap2.get(titleList.get(1)) + ",安全DOI",itemMap2);
                        keyItemMap3.put(itemMap3.get(titleList.get(0)) + "," + itemMap3.get(titleList.get(1)) + ",目标DOI",itemMap3);
                    }
                }
            }
        }
        // excelMap中的数据就是在Excel中读取的内容
        // 客户DOI数据 itemList
        // 预测量 keyItemMap2
        // 期初值 keyItemMap3
        // 标题行
        List<String> titleList1 = titleMap.get("sheet1");
        List<String> titleList2 = titleMap.get("sheet2");
        List<String> titleList3 = titleMap.get("sheet3");

        System.out.println("sheet1:" + itemList.size() + " -- titleList1 -- " + titleList1);
        System.out.println("sheet2:" + keyItemMap2.size() + " -- titleList2 -- " + titleList2);
        System.out.println("sheet3:" + keyItemMap3.size() + " -- titleList3 -- " + titleList3);

        /**
         * 处理数据
         * 理想期末 = 未来（当周doi天）的预测量
         * 期初+到货 = 预测 + 期末===> 到货 = 期末 + 预测 - 期初
         */
        // 最大DOI
        List<Map<String,String>> maxList = new ArrayList<>();
        // 安全DOI
        List<Map<String,String>> safeList = new ArrayList<>();
        // 目标DOI
        List<Map<String,String>> targetList = new ArrayList<>();
        // 存储元素的map
        Map<String,String> newMap = new HashMap<>();

        for (int i = 0; i < itemList.size(); i++) {
            newMap = new HashMap<>();
            Map<String, String> item = itemList.get(i);

            String title0 = titleList1.get(0);
            String title1 = titleList1.get(1);
            String title2 = titleList1.get(2);
            String title3 = titleList1.get(3);
            String title4 = titleList1.get(4);

            // 获取期初值的key
            String expectTitle = titleList3.get(2);

            String customer = item.get(title0); // 客户
            String material = item.get(title1); // 物料
            String attribute = item.get(title2); // Attribute
            String DOIType = item.get(title3); // DOI类别
            String unitType = item.get(title4); // 单位 Day Or Case

            newMap.put(title0,customer);
            newMap.put(title1,material);
            newMap.put(title2,attribute);
            newMap.put(title3,DOIType);
            newMap.put(title4,unitType);

            // 从第5列开始 往后都是日期列
            for (int i1 = 5; i1 < titleList1.size(); i1++) {
                String dayStr = titleList1.get(i1);
                // 期初
                String start = keyItemMap3.get(customer + "," + material + "," + DOIType).get(expectTitle);
                BigDecimal startValue = new BigDecimal(start);
//                Integer startValue = Integer.valueOf(start);
                // 预计
                String expect = keyItemMap2.get(customer + "," + material).get(dayStr);
                BigDecimal expectValue = new BigDecimal(expect);
//                Integer expectValue = Integer.valueOf(expect);
                // Day or Case的值 来计算期末
                String dayOrCase = item.get(dayStr);
                Integer dayOrCaseValue = Integer.valueOf(dayOrCase);

                // 期末 -- 等价于 a > week
                BigDecimal endValue = new BigDecimal(0);
                if("Day".equals(unitType)){
                    // 天
                    int week = dayOrCaseValue / 7; // 一共几周
                    int over = dayOrCaseValue % 7; // 余几天
                    // 期末 = 到日doi计算得到
                    // 这是计算不足一周的
                    if(week == 0 && over > 0){
                        // 这里是计算余数
                        if(i1 + 1 > titleList1.size() - 1){
                            continue;
                        }
                        String newDayStr = titleList1.get(i1 + 1);
                        String newExpect = keyItemMap2.get(customer + "," + material).get(newDayStr);
                        //                            Integer newExpecteValue = Integer.valueOf(newExpect);
                        BigDecimal newExpecteValue = new BigDecimal(newExpect);
                        //                            Integer newExpecteValue2 = newExpecteValue * over / 7;
                        // 四舍五入
                        BigDecimal newExpecteValue2 = newExpecteValue.multiply(new BigDecimal(over)).divide(new BigDecimal(7), 1, BigDecimal.ROUND_HALF_UP);
                        //                            endValue += newExpecteValue2;
                        endValue = endValue.add(newExpecteValue2);

                    }else if(week > 0){
                        // 期初 + 到货 = 预计 + 期末
                        int a = 1;
                        while(a <= week){
                            // 这是计算整周的
                            if(i1 + a <= titleList1.size() - 1){
                                String newDayStr = titleList1.get(i1 + a);
                                String newExpect = keyItemMap2.get(customer + "," + material).get(newDayStr);
//                            Integer newExpecteValue = Integer.valueOf(newExpect);
                                BigDecimal newExpecteValue = new BigDecimal(newExpect);
//                            endValue += newExpecteValue;
                                endValue = endValue.add(newExpecteValue);
                            }

                            if( a == week && over > 0){
                                // 这里是计算余数
                                if(i1 + a + 1 > titleList1.size() - 1){
                                    a++;
                                    continue;
                                }
                                String newDayStr = titleList1.get(i1 + a + 1);
                                String newExpect = keyItemMap2.get(customer + "," + material).get(newDayStr);
//                            Integer newExpecteValue = Integer.valueOf(newExpect);
                                BigDecimal newExpecteValue = new BigDecimal(newExpect);
//                            Integer newExpecteValue2 = newExpecteValue * over / 7;
                                // 四舍五入
                                BigDecimal newExpecteValue2 = newExpecteValue.multiply(new BigDecimal(over)).divide(new BigDecimal(7), 1, BigDecimal.ROUND_HALF_UP);
//                            endValue += newExpecteValue2;
                                endValue = endValue.add(newExpecteValue2);

                            }
                            a++;
                        }
                    }

                }else{
                    // Case 箱
//                    String newExpect = keyItemMap2.get(customer + "," + material).get(dayStr);
                    endValue = new BigDecimal(dayOrCaseValue);
                }

                // 期初 + 到货 = 预计 + 期末
                // 到货
                BigDecimal arrival = new BigDecimal(0);
//                if(startValue > expectValue + endValue){
                if(startValue.compareTo(expectValue.add(endValue)) == 1){
                    arrival = new BigDecimal(0);
                    // 如果等式不成立 那么需要重新计算
                    // 第一周的期末 是第二周的期初
                    // 期初
                    String start1 = keyItemMap3.get(customer + "," + material + "," + DOIType).get(expectTitle);
                    BigDecimal startValue1 = new BigDecimal(start1);
                    // 预计
                    String expect1 = keyItemMap2.get(customer + "," + material).get(dayStr);
                    BigDecimal expectValue1 = new BigDecimal(expect1);

                    // 实际期末值 就是下一周的期初
                    BigDecimal endValue1 = startValue1.subtract(expectValue1);

                    keyItemMap3.get(customer + "," + material + "," + DOIType).put(expectTitle,String.valueOf(endValue1));
                }else{
//                    arrival = expectValue + endValue - startValue;
                    arrival = expectValue.add(endValue).subtract(startValue);
                    // 第一周的期末 是第二周的期初
                    keyItemMap3.get(customer + "," + material + "," + DOIType).put(expectTitle,String.valueOf(endValue));
                }


                newMap.put(dayStr,String.valueOf(arrival));

            }

            switch (DOIType){
                case "最大DOI":
                    maxList.add(newMap);
                    break;
                case "安全DOI":
                    safeList.add(newMap);
                    break;
                case "目标DOI":
                    targetList.add(newMap);
                    break;
            }

        }
        // 数据封装完成 开始导出
        OutputStream out = null;
        try {
            String fileName = "到货处理数据NEW.xlsx";
            String newFilePath = filePath.substring(0, filePath.lastIndexOf("\\") + 1);
            out = new FileOutputStream(newFilePath + fileName);
            ExcelUtils eeu = new ExcelUtils();
            XSSFWorkbook workbook = new XSSFWorkbook();
            eeu.exportExcel(workbook, 0, "最大DOI数据", titleList1, maxList);
            eeu.exportExcel(workbook, 1, "安全DOI数据", titleList1, safeList);
            eeu.exportExcel(workbook, 2, "目标DOI数据", titleList1, targetList);
            //将所有的数据一起写入，然后再关闭输入流。
            workbook.write(out);
            out.flush();
            workbook.close();
            System.out.println("导出结束,文件路径在-->"+newFilePath + fileName);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if(out != null){
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }


    public static String getStringValueFromCell(Cell cell) {
        SimpleDateFormat sFormat = new SimpleDateFormat("MM/dd/yyyy");
        DecimalFormat decimalFormat = new DecimalFormat("#.#");
        String cellValue = "";
        if(cell == null) {
            return cellValue;
        }
        else if(cell.getCellType() == Cell.CELL_TYPE_STRING) {
            cellValue = cell.getStringCellValue();
        }

        else if(cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
            if(HSSFDateUtil.isCellDateFormatted(cell)) {
                double d = cell.getNumericCellValue();
                Date date = HSSFDateUtil.getJavaDate(d);
                cellValue = sFormat.format(date);
            }
            else {
                cellValue = decimalFormat.format((cell.getNumericCellValue()));
            }
        }
        else if(cell.getCellType() == Cell.CELL_TYPE_BLANK) {
            cellValue = "";
        }
        else if(cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
            cellValue = String.valueOf(cell.getBooleanCellValue());
        }
        else if(cell.getCellType() == Cell.CELL_TYPE_ERROR) {
            cellValue = "";
        }
        else if(cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
            cellValue = cell.getCellFormula().toString();
        }
        return cellValue;
    }

    public void exportExcel(){
        String fileName = "";
    }
}

class ExcelUtils {

    public void exportExcel(XSSFWorkbook workbook, int sheetNum, String sheetTitle, List<String> headers, List<Map<String, String>> result) {
        // 生成一个表格
        XSSFSheet sheet = workbook.createSheet();
        workbook.setSheetName(sheetNum, sheetTitle);
        // 设置表格默认列宽度为20个字节
        sheet.setDefaultColumnWidth((short) 20);
        // 生成一个样式
        XSSFCellStyle style = workbook.createCellStyle();
        // 设置这些样式
        style.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 生成一个字体
        XSSFFont font = workbook.createFont();
        font.setColor(HSSFColor.BLACK.index);
        font.setFontHeightInPoints((short) 12);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        // 把字体应用到当前的样式
        style.setFont(font);

        // 指定当单元格内容显示不下时自动换行
        style.setWrapText(true);

        // 产生表格标题行
        XSSFRow row = sheet.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            XSSFCell cell = row.createCell((short) i);

            cell.setCellStyle(style);
            HSSFRichTextString text = new HSSFRichTextString(headers.get(i));
            cell.setCellValue(text.toString());
        }
        // 遍历集合数据，产生数据行
        if (result != null) {
            int index = 1;
            for (Map<String, String> m : result) {
                row = sheet.createRow(index);
                int cellIndex = 0;
//                System.out.print("**"+sheetNum+"**");
                for (int i = 0; i < headers.size(); i++) {
                    String val = m.get(headers.get(i));
                    XSSFCell cell = row.createCell((short) cellIndex);
//                    System.out.print(headers.get(i) + " = " + val + "   ");
                    cell.setCellValue(val);
                    cellIndex++;
                }
//                System.out.println();
                index++;
            }
        }
    }


    public static void main(String[] args) {

        List<Map<String,String>> itemList = new ArrayList<>();
        Map<String,Map<String,String>> keyItemMap2 = new HashMap<>();
        Map<String,Map<String,String>> keyItemMap3 = new HashMap<>();

        List<String> titleList1 = new ArrayList<String>(){{
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
        Map<String,String> itemMap = new HashMap<>();
        itemMap.put(titleList1.get(0),"800500");
        itemMap.put(titleList1.get(1),"100815");
        itemMap.put(titleList1.get(2),"Max On Hand Constraint");
        itemMap.put(titleList1.get(3),"最大DOI");
        itemMap.put(titleList1.get(4),"Day");
        itemMap.put(titleList1.get(5),"9");
        itemMap.put(titleList1.get(6),"9");
        itemMap.put(titleList1.get(7),"9");
        itemMap.put(titleList1.get(8),"9");
        itemMap.put(titleList1.get(9),"9");
        itemMap.put(titleList1.get(10),"9");
        itemMap.put(titleList1.get(11),"9");
        itemMap.put(titleList1.get(12),"9");
        itemMap.put(titleList1.get(13),"9");
        itemMap.put(titleList1.get(14),"9");
        itemMap.put(titleList1.get(15),"9");
        itemMap.put(titleList1.get(16),"9");
        itemMap.put(titleList1.get(17),"9");
        itemMap.put(titleList1.get(18),"9");
        itemList.add(itemMap);
        itemMap = new HashMap<>();
        itemMap.put(titleList1.get(0),"800500");
        itemMap.put(titleList1.get(1),"100815");
        itemMap.put(titleList1.get(2),"Max On Hand Constraint");
        itemMap.put(titleList1.get(3),"目标DOI");
        itemMap.put(titleList1.get(4),"Day");
        itemMap.put(titleList1.get(5),"19");
        itemMap.put(titleList1.get(6),"19");
        itemMap.put(titleList1.get(7),"19");
        itemMap.put(titleList1.get(8),"19");
        itemMap.put(titleList1.get(9),"19");
        itemMap.put(titleList1.get(10),"19");
        itemMap.put(titleList1.get(11),"19");
        itemMap.put(titleList1.get(12),"19");
        itemMap.put(titleList1.get(13),"19");
        itemMap.put(titleList1.get(14),"19");
        itemMap.put(titleList1.get(15),"19");
        itemMap.put(titleList1.get(16),"19");
        itemMap.put(titleList1.get(17),"19");
        itemMap.put(titleList1.get(18),"19");
        itemList.add(itemMap);
        itemMap = new HashMap<>();
        itemMap.put(titleList1.get(0),"800500");
        itemMap.put(titleList1.get(1),"100815");
        itemMap.put(titleList1.get(2),"Max On Hand Constraint");
        itemMap.put(titleList1.get(3),"安全DOI");
        itemMap.put(titleList1.get(4),"Case (箱)");
        itemMap.put(titleList1.get(5),"50000");
        itemMap.put(titleList1.get(6),"50000");
        itemMap.put(titleList1.get(7),"50000");
        itemMap.put(titleList1.get(8),"50000");
        itemMap.put(titleList1.get(9),"50000");
        itemMap.put(titleList1.get(10),"50000");
        itemMap.put(titleList1.get(11),"50000");
        itemMap.put(titleList1.get(12),"50000");
        itemMap.put(titleList1.get(13),"50000");
        itemMap.put(titleList1.get(14),"50000");
        itemMap.put(titleList1.get(15),"50000");
        itemMap.put(titleList1.get(16),"50000");
        itemMap.put(titleList1.get(17),"50000");
        itemMap.put(titleList1.get(18),"50000");
        itemList.add(itemMap);



        itemMap = new HashMap<>();
        List<String> titleList2 = new ArrayList<String>(){{
            add("客户");
            add("物料");
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

        itemMap.put(titleList2.get(0),"800500");
        itemMap.put(titleList2.get(1),"100815");
        itemMap.put(titleList2.get(2),"34000");
        itemMap.put(titleList2.get(3),"30000");
        itemMap.put(titleList2.get(4),"45021");
        itemMap.put(titleList2.get(5),"58946");
        itemMap.put(titleList2.get(6),"23762");
        itemMap.put(titleList2.get(7),"65364");
        itemMap.put(titleList2.get(8),"55745");
        itemMap.put(titleList2.get(9),"69879");
        itemMap.put(titleList2.get(10),"86055");
        itemMap.put(titleList2.get(11),"100576");
        itemMap.put(titleList2.get(12),"178587");
        itemMap.put(titleList2.get(13),"122715");
        itemMap.put(titleList2.get(14),"151661");
        itemMap.put(titleList2.get(15),"0");
        keyItemMap2.put("800500"+","+"100815",itemMap);

        List<String> titleList3 = new ArrayList<String>(){{
            add("客户");
            add("物料");
            add("客户期初库存");
        }};
        Map<String,String> itemMap1 = new HashMap<>();
        itemMap1.put(titleList3.get(0),"800500");
        itemMap1.put(titleList3.get(1),"100815");
        itemMap1.put(titleList3.get(2),"82697");

        Map<String,String> itemMap2 = new HashMap<>();
        itemMap2.put(titleList3.get(0),"800500");
        itemMap2.put(titleList3.get(1),"100815");
        itemMap2.put(titleList3.get(2),"82697");

        Map<String,String> itemMap3 = new HashMap<>();
        itemMap3.put(titleList3.get(0),"800500");
        itemMap3.put(titleList3.get(1),"100815");
        itemMap3.put(titleList3.get(2),"82697");

        keyItemMap3.put("800500"+","+"100815"+",最大DOI",itemMap1);
        keyItemMap3.put("800500"+","+"100815"+",安全DOI",itemMap2);
        keyItemMap3.put("800500"+","+"100815"+",目标DOI",itemMap3);

        // 最大DOI
        List<Map<String,String>> maxList = new ArrayList<>();
        // 安全DOI
        List<Map<String,String>> safeList = new ArrayList<>();
        // 目标DOI
        List<Map<String,String>> targetList = new ArrayList<>();
        // 存储元素的map
        Map<String,String> newMap = new HashMap<>();

        for (int i = 0; i < itemList.size(); i++) {
            newMap = new HashMap<>();
            Map<String, String> item = itemList.get(i);

            String title0 = titleList1.get(0);
            String title1 = titleList1.get(1);
            String title2 = titleList1.get(2);
            String title3 = titleList1.get(3);
            String title4 = titleList1.get(4);

            // 获取期初值的key
            String expectTitle = titleList3.get(2);

            String customer = item.get(title0); // 客户
            String material = item.get(title1); // 物料
            String attribute = item.get(title2); // Attribute
            String DOIType = item.get(title3); // DOI类别
            String unitType = item.get(title4); // 单位 Day Or Case

            newMap.put(title0,customer);
            newMap.put(title1,material);
            newMap.put(title2,attribute);
            newMap.put(title3,DOIType);
            newMap.put(title4,unitType);

            // 从第5列开始 往后都是日期列
            for (int i1 = 5; i1 < titleList1.size(); i1++) {
                String dayStr = titleList1.get(i1);
                // 期初
                String start = keyItemMap3.get(customer + "," + material + "," + DOIType).get(expectTitle);
                BigDecimal startValue = new BigDecimal(start);
//                Integer startValue = Integer.valueOf(start);
                // 预计
                String expect = keyItemMap2.get(customer + "," + material).get(dayStr);
                BigDecimal expectValue = new BigDecimal(expect);
//                Integer expectValue = Integer.valueOf(expect);
                // Day or Case的值 来计算期末
                String dayOrCase = item.get(dayStr);
                Integer dayOrCaseValue = Integer.valueOf(dayOrCase);

                // 期末 -- 等价于 a > week
                BigDecimal endValue = new BigDecimal(0);
                if("Day".equals(unitType)){
                    // 天
                    int week = dayOrCaseValue / 7; // 一共几周
                    int over = dayOrCaseValue % 7; // 余几天
                    // 期末 = 到日doi计算得到
                    // 期初 + 到货 = 预计 + 期末
                    int a = 1;
                    while(a <= week){
                        // 这是计算整周的
                        if(i1 + a <= titleList1.size() - 1){
                            String newDayStr = titleList1.get(i1 + a);
                            String newExpect = keyItemMap2.get(customer + "," + material).get(newDayStr);
//                            Integer newExpecteValue = Integer.valueOf(newExpect);
                            BigDecimal newExpecteValue = new BigDecimal(newExpect);
//                            endValue += newExpecteValue;
                            endValue = endValue.add(newExpecteValue);
                        }

                        if( a == week && over > 0){
                            // 这里是计算余数
                            if(i1 + a + 1 > titleList1.size() - 1){
                                a++;
                                continue;
                            }
                            String newDayStr = titleList1.get(i1 + a + 1);
                            String newExpect = keyItemMap2.get(customer + "," + material).get(newDayStr);
//                            Integer newExpecteValue = Integer.valueOf(newExpect);
                            BigDecimal newExpecteValue = new BigDecimal(newExpect);
//                            Integer newExpecteValue2 = newExpecteValue * over / 7;
                            // 四舍五入
                            BigDecimal newExpecteValue2 = newExpecteValue.multiply(new BigDecimal(over)).divide(new BigDecimal(7), 1, BigDecimal.ROUND_HALF_UP);
//                            endValue += newExpecteValue2;
                            endValue = endValue.add(newExpecteValue2);

                        }
                        a++;
                    }

                }else{
                    // Case 箱
//                    String newExpect = keyItemMap2.get(customer + "," + material).get(dayStr);
                    endValue = new BigDecimal(dayOrCaseValue);
                }

                // 期初 + 到货 = 预计 + 期末
                // 到货
                BigDecimal arrival = new BigDecimal(0);
//                if(startValue > expectValue + endValue){
                if(startValue.compareTo(expectValue.add(endValue)) == 1){
                    arrival = new BigDecimal(0);
                    // 如果等式不成立 那么需要重新计算
                    // 第一周的期末 是第二周的期初
                    // 期初
                    String start1 = keyItemMap3.get(customer + "," + material + "," + DOIType).get(expectTitle);
                    BigDecimal startValue1 = new BigDecimal(start1);
                    // 预计
                    String expect1 = keyItemMap2.get(customer + "," + material).get(dayStr);
                    BigDecimal expectValue1 = new BigDecimal(expect1);

                    // 实际期末值 就是下一周的期初
                    BigDecimal endValue1 = startValue1.subtract(expectValue1);

                    keyItemMap3.get(customer + "," + material + "," + DOIType).put(expectTitle,String.valueOf(endValue1));
                }else{
//                    arrival = expectValue + endValue - startValue;
                    arrival = expectValue.add(endValue).subtract(startValue);
                    // 第一周的期末 是第二周的期初
                    keyItemMap3.get(customer + "," + material + "," + DOIType).put(expectTitle,String.valueOf(endValue));
                }

                newMap.put(dayStr,String.valueOf(arrival));

            }

            switch (DOIType){
                case "最大DOI":
                    maxList.add(newMap);
                    break;
                case "安全DOI":
                    safeList.add(newMap);
                    break;
                case "目标DOI":
                    targetList.add(newMap);
                    break;
            }

        }
        // 数据封装完成 开始导出
        OutputStream out = null;
        try {
            String filePath = "D:\\temp\\testdata.xlsx";
            String fileName = "到货处理数据-test3.xlsx";
            String newFilePath = filePath.substring(0, filePath.lastIndexOf("\\") + 1);
            out = new FileOutputStream(newFilePath + fileName);
            ExcelUtils eeu = new ExcelUtils();
            XSSFWorkbook workbook = new XSSFWorkbook();
            eeu.exportExcel(workbook, 0, "最大DOI数据", titleList1, maxList);
            eeu.exportExcel(workbook, 1, "安全DOI数据", titleList1, safeList);
            eeu.exportExcel(workbook, 2, "目标DOI数据", titleList1, targetList);
            //将所有的数据一起写入，然后再关闭输入流。
            workbook.write(out);
            out.flush();
            workbook.close();
            System.out.println("导出结束,文件路径在-->"+newFilePath + fileName);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if(out != null){
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }


    }
}