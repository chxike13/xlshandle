package com.bicon.xlshandle.Service.impl;

import com.bicon.xlshandle.Service.HandleService;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

@Service("handleServiceImpl")
public class HandleServiceImpl implements HandleService {
    @Override
    public void format(String inputName)throws IOException{
        //      读取文件，创建工作空间
        String inputURL = "D:\\shuakahuizong\\"+inputName+".xls";//指定的文件目录，需要在此目录放入要格式化的文件。
        String outputURL = "D:\\shuakahuizong\\"+inputName+"标记.xls";
        FileOutputStream output = new FileOutputStream(outputURL);
        FileInputStream input = new FileInputStream(inputURL);
        HSSFWorkbook workbook = new HSSFWorkbook(input);
        HSSFSheet sheet = workbook.getSheetAt(0);

//      多数据异常项样式
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.RED.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        HSSFDataFormat format = workbook.createDataFormat();

//      单数据异常项样式
        HSSFCellStyle style0 = workbook.createCellStyle();
        style0.setFillForegroundColor(IndexedColors.RED.getIndex());
        style0.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style0.setAlignment(HorizontalAlignment.CENTER);
        style0.setBorderBottom(BorderStyle.THIN);
        style0.setBorderRight(BorderStyle.THIN);
        style0.setBorderLeft(BorderStyle.THIN);
        style0.setBorderTop(BorderStyle.THIN);
//      遍历文件
        for (int i = 0;i < sheet.getPhysicalNumberOfRows();i++){
            HSSFRow row = sheet.getRow(i);
            for (int j = 0;j < row.getPhysicalNumberOfCells();j++){
                HSSFCell cell = row.getCell(j);
                String s = row.getCell(j).toString();

//                日期处理成当月日期并设置样式
                if (i==1 &&j > 4){
                    cell.setCellValue(sheet.getRow(2).getCell(4).getNumericCellValue()+j-5);
                    HSSFDataFormat format1 = workbook.createDataFormat();
                    HSSFFont font = workbook.createFont();
                    font.setColor(IndexedColors.WHITE.getIndex());
                    HSSFCellStyle style1 = workbook.createCellStyle();
                    style1.setFillForegroundColor(IndexedColors.ROYAL_BLUE.getIndex());
                    style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    style1.setDataFormat(format1.getFormat("yyyy-mm-dd"));
                    style1.setAlignment(HorizontalAlignment.CENTER);
                    style1.setBorderBottom(BorderStyle.THIN);
                    style1.setBorderRight(BorderStyle.THIN);
                    style1.setBorderLeft(BorderStyle.THIN);
                    style1.setBorderTop(BorderStyle.THIN);
                    style1.setFont(font);
                    cell.setCellStyle(style1);
                }
//              月份格式
                if (i>1 &&j == 4){
                    HSSFDataFormat format1 = workbook.createDataFormat();
                    HSSFCellStyle style1 = workbook.createCellStyle();
                    style1.setDataFormat(format1.getFormat("yyyy-mm"));
                    style1.setAlignment(HorizontalAlignment.CENTER);
                    style1.setVerticalAlignment(VerticalAlignment.CENTER);
                    style1.setBorderBottom(BorderStyle.THIN);
                    style1.setBorderRight(BorderStyle.THIN);
                    style1.setBorderLeft(BorderStyle.THIN);
                    style1.setBorderTop(BorderStyle.THIN);
                    cell.setCellStyle(style1);
                }

//                设置列宽
                if (j>=5){
                    sheet.setColumnWidth(j,20*256);
                }
//                去除多余数据
                if (i>=2 && j >= 5){
                    if (s.indexOf(" ")>0) {
                        String temp1 = s.substring(0, s.indexOf(" "));
                        String temp2 = s.substring(s.lastIndexOf(" ") + 1);
                        String s1 = temp1 + " " + temp2;
                        if (temp1.compareTo("08:15") > 0) {
                            cell.setCellStyle(style);
                        }
                        if (temp2.compareTo("18:00") < 0) {
                            cell.setCellStyle(style);
                        }
                        if (isWeekend(sheet.getRow(1).getCell(j))){
                            HSSFCellStyle style2 = workbook.createCellStyle();
                            style2.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                            style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            style2.setAlignment(HorizontalAlignment.CENTER);
                            style2.setBorderBottom(BorderStyle.THIN);
                            style2.setBorderRight(BorderStyle.THIN);
                            style2.setBorderLeft(BorderStyle.THIN);
                            style2.setBorderTop(BorderStyle.THIN);
                            cell.setCellStyle(style2);
                        }
                        cell.setCellValue(s1);
                    }
                    else {
                        style0.setDataFormat(format.getFormat("hh:mm"));
                        cell.setCellStyle(style0);
                        if (isWeekend(sheet.getRow(1).getCell(j))){
                            HSSFCellStyle style3 = workbook.createCellStyle();
                            style3.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                            style3.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            style3.setAlignment(HorizontalAlignment.CENTER);
                            style3.setBorderBottom(BorderStyle.THIN);
                            style3.setBorderRight(BorderStyle.THIN);
                            style3.setBorderLeft(BorderStyle.THIN);
                            style3.setBorderTop(BorderStyle.THIN);
                            style3.setDataFormat(format.getFormat("hh:mm"));
                            cell.setCellStyle(style3);
                        }
                    }
                }
            }
        }
//        输出文件
        workbook.write(output);
        output.close();
        System.out.println(inputName+"：格式转换完成!");
    }

    @Override
    public String splitByDepartment(String inputPath, String departmentName) throws IOException {
        FileInputStream inputStream = new FileInputStream(inputPath);
        FileOutputStream outputStream = new FileOutputStream(inputPath+"/"+departmentName+".xls");//拆分后的文件的输出路径
        HSSFWorkbook workbookIn = new HSSFWorkbook(inputStream);
        HSSFWorkbook workbookOut = new HSSFWorkbook();
        HSSFSheet sheetIn = workbookIn.getSheetAt(0);
        List<List> list = new ArrayList<List>();
        for(int rowNum = 0;rowNum < sheetIn.getPhysicalNumberOfRows();rowNum++){
            HSSFRow row = sheetIn.getRow(rowNum);
            if (rowNum < 2){
                List<HSSFCell> listTemp = new ArrayList<HSSFCell>();
                for (int i = 0;i < row.getPhysicalNumberOfCells();i++){
                    listTemp.add(row.getCell(i));
                }
                list.add(listTemp);
            }
            for (int cellNum = 0;cellNum < row.getPhysicalNumberOfCells();cellNum++){
                HSSFCell cell = row.getCell(cellNum);
                String s = cell.toString();
                if (s.equals(departmentName)){
                    List<HSSFCell> listTemp = new ArrayList<HSSFCell>();
                    for (int i = 0;i < row.getPhysicalNumberOfCells();i++){
                        listTemp.add(row.getCell(i));
                    }
                    list.add(listTemp);
                }
            }
        }

        HSSFSheet sheetOut = workbookOut.createSheet(departmentName);
        CellRangeAddress region = new CellRangeAddress(0,0,0,35);
        sheetOut.addMergedRegion(region);
        for (int rowNum = 0;rowNum < list.size();rowNum++){
            HSSFRow row = sheetOut.createRow(rowNum);
            if (rowNum<2) {
                row.setHeight((short) (25 * 20));
            }
            for (int cellNum = 0;cellNum < list.get(rowNum).size();cellNum++){
                row.createCell(cellNum);
                HSSFCell cell = row.getCell(cellNum);
                HSSFCell cellTemp = (HSSFCell) list.get(rowNum).get(cellNum);
                CellType s = cellTemp.getCellTypeEnum();
                switch (s){
                    case STRING: cell.setCellValue(cellTemp.getStringCellValue());break;
                    case NUMERIC: cell.setCellValue(cellTemp.getNumericCellValue());break;
                    case BLANK: cell.setCellValue(cellTemp.toString());break;
                }
                HSSFCellStyle style = workbookOut.createCellStyle();
                style.cloneStyleFrom(cellTemp.getCellStyle());
                cell.setCellStyle(style);
            }
        }
        workbookOut.write(outputStream);
        outputStream.close();
        System.out.println(departmentName+"：分拆完成！");
        format(departmentName);
        return departmentName+"：处理完成！";
    }

    @Override
    public boolean isWeekend(HSSFCell cell) {
        Date date  = cell.getDateCellValue();
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        int i = calendar.get(Calendar.DAY_OF_WEEK)-1;
//        System.out.println(date+"\n"+i);
        if (i==0||i==6) {
            return true;
        }
        return false;
    }

}
