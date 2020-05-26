package com.example.minalz.minalzcloudtest.controller;

import com.example.minalz.minalzcloudtest.utils.ExcelUtils;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiImplicitParam;
import io.swagger.annotations.ApiOperation;
import io.swagger.annotations.ApiParam;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @program: Ehance
 * @description:
 * @author: oldman
 * @date: 2020-05-26 19:32
 **/
@Api(value = "到货数据处理")
@RestController("/temp/dealing")
public class POIController {

    @ApiOperation(value = "导入excel", nickname = "导入excel文件")
    @ApiImplicitParam(name = "filePath", value = "文件路径", required = true, dataType = "String")
    @PostMapping("/deal")
    public String dealdata(String filePath){
        return "文件路径-->" + filePath;
    }

    @ApiOperation(value = "导出excel", nickname = "导出excel文件")
    @PostMapping("/exportExcel")
    public String exportExcel(){

        return "导出文件" + true;
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
            cellValue = cell.getCellFormula();
        }
        return cellValue;
    }


    public static void main(String[] args) {

        List<Map<String,String>> resultList = new ArrayList<>();
        Map<String,Map<String,String>> itemMap1 = new HashMap<>();
        Map<String,Map<String,String>> itemMap2 = new HashMap<>();

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
        resultList.add(itemMap);
        itemMap = new HashMap<>();
        itemMap.put(titleList1.get(0),"800500");
        itemMap.put(titleList1.get(1),"100815");
        itemMap.put(titleList1.get(2),"Max On Hand Constraint");
        itemMap.put(titleList1.get(3),"安全DOI");
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
        resultList.add(itemMap);
        itemMap = new HashMap<>();
        itemMap.put(titleList1.get(0),"800500");
        itemMap.put(titleList1.get(1),"100815");
        itemMap.put(titleList1.get(2),"Max On Hand Constraint");
        itemMap.put(titleList1.get(3),"目标DOI");
        itemMap.put(titleList1.get(4),"Case (箱)");
        itemMap.put(titleList1.get(5),"3000");
        itemMap.put(titleList1.get(6),"3000");
        itemMap.put(titleList1.get(7),"3000");
        itemMap.put(titleList1.get(8),"3000");
        itemMap.put(titleList1.get(9),"3000");
        itemMap.put(titleList1.get(10),"3000");
        itemMap.put(titleList1.get(11),"3000");
        itemMap.put(titleList1.get(12),"3000");
        itemMap.put(titleList1.get(13),"3000");
        itemMap.put(titleList1.get(14),"3000");
        itemMap.put(titleList1.get(15),"3000");
        itemMap.put(titleList1.get(16),"3000");
        itemMap.put(titleList1.get(17),"3000");
        itemMap.put(titleList1.get(18),"3000");
        resultList.add(itemMap);



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
        itemMap.put(titleList2.get(2),"0");
        itemMap.put(titleList2.get(3),"0");
        itemMap.put(titleList2.get(4),"0");
        itemMap.put(titleList2.get(5),"0");
        itemMap.put(titleList2.get(6),"0");
        itemMap.put(titleList2.get(7),"17120");
        itemMap.put(titleList2.get(8),"40007");
        itemMap.put(titleList2.get(9),"44410");
        itemMap.put(titleList2.get(10),"42540");
        itemMap.put(titleList2.get(11),"41734");
        itemMap.put(titleList2.get(12),"41143");
        itemMap.put(titleList2.get(13),"40494");
        itemMap.put(titleList2.get(14),"44509");
        itemMap.put(titleList2.get(15),"0");
        itemMap1.put("800500"+","+"100815",itemMap);

        List<String> titleList3 = new ArrayList<String>(){{
            add("客户");
            add("物料");
            add("客户期初库存");
        }};
        itemMap = new HashMap<>();
        itemMap.put(titleList3.get(0),"800500");
        itemMap.put(titleList3.get(1),"100815");
        itemMap.put(titleList3.get(2),"31823");
        itemMap2.put("800500"+","+"100815",itemMap);

        // 最大DOI
        List<Map<String,String>> maxList = new ArrayList<>();
        // 安全DOI
        List<Map<String,String>> safeList = new ArrayList<>();
        // 目标DOI
        List<Map<String,String>> targetList = new ArrayList<>();
        // 存储元素的map
        Map<String,String> newMap = new HashMap<>();

        for (int i = 0; i < resultList.size(); i++) {
            newMap = new HashMap<>();
            Map<String, String> item = resultList.get(i);

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
                String start = itemMap2.get(customer + "," + material).get(expectTitle);
                Integer startValue = Integer.valueOf(start);
                // 预计
                String expect = itemMap1.get(customer + "," + material).get(dayStr);
                Integer expectValue = Integer.valueOf(expect);
                // Day or Case的值 来计算期末
                String dayOrCase = item.get(dayStr);
                Integer dayOrCaseValue = Integer.valueOf(dayOrCase);

                // 期末 -- 等价于 a > week
                Integer endValue = expectValue;
                if("Day".equals(unitType)){
                    // 天
                    int week = dayOrCaseValue / 7;// 一共几周
                    int over = dayOrCaseValue % 7;// 余几天
                    // 期末 = 到日doi计算得到
                    // 期初 + 到货 = 预计 + 期末
                    int a = 1;
                    while(a <= week){
                        // 这是计算整周的 不需要单独去除
                        if(i1 + a <= titleList1.size() - 1){
                            String newDayStr = titleList1.get(i1 + a);
                            String newExpect = itemMap1.get(customer + "," + material).get(newDayStr);
                            Integer newExpecteValue = Integer.valueOf(newExpect);
                            if( a >= week){
                                newExpecteValue = newExpecteValue * over / 7;
                            }
                            endValue += newExpecteValue;
                        }
                        a++;
                    }

                }else{
                    // Case 箱
                    String newExpect = itemMap1.get(customer + "," + material).get(dayStr);
                    endValue = Integer.valueOf(newExpect);
                }

                // 期初 + 到货 = 预计 + 期末
                // 到货
                Integer arrival = 0;
                if(startValue > expectValue + endValue){
                    arrival = 0;
                }else{
                    arrival = expectValue + endValue - startValue;
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
        // 。。。
        OutputStream out = null;
        try {
            String fileName = "最终处理数据2.xlsx";
            out = new FileOutputStream("/Users/zhouwei/Desktop/temp/" + fileName);
            ExcelUtils eeu = new ExcelUtils();
            HSSFWorkbook workbook = new HSSFWorkbook();
            eeu.exportExcel(workbook, 0, "最大DOI数据", titleList1, maxList, out);
            eeu.exportExcel(workbook, 1, "安全DOI数据", titleList1, safeList, out);
            eeu.exportExcel(workbook, 2, "目标DOI数据", titleList1, targetList, out);
            //将所有的数据一起写入，然后再关闭输入流。
            workbook.write(out);
            System.out.println("导出结束");
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
