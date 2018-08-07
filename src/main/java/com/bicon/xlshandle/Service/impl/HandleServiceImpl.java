package com.bicon.xlshandle.Service.impl;

import com.bicon.xlshandle.Service.HandleService;
import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import org.apache.http.NameValuePair;
import org.apache.http.client.entity.UrlEncodedFormEntity;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.message.BasicNameValuePair;
import org.apache.http.util.EntityUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

@Service("handleServiceImpl")
public class HandleServiceImpl implements HandleService {
    @Override
    public void format(String inputPath, String inputName)throws IOException{
        //      读取文件，创建工作空间
        File file = new File(inputPath+"/result");
        file.mkdir();
        FileOutputStream output = new FileOutputStream(file+"\\"+inputName+"result.xls");
        FileInputStream input = new FileInputStream(inputPath+"\\"+inputName+".xls");
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
//        周末样式
        HSSFCellStyle style2 = workbook.createCellStyle();
        style2.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style2.setAlignment(HorizontalAlignment.CENTER);
        style2.setBorderBottom(BorderStyle.THIN);
        style2.setBorderRight(BorderStyle.THIN);
        style2.setBorderLeft(BorderStyle.THIN);
        style2.setBorderTop(BorderStyle.THIN);
        System.out.println("行数："+sheet.getPhysicalNumberOfRows());

        List<Map<String, Long>> leaveInfo = null;
//      遍历文件
        for (int rowNum = 0;rowNum < sheet.getPhysicalNumberOfRows();rowNum++){
            HSSFRow row = sheet.getRow(rowNum);
            for (int cellNum = 0;cellNum < row.getPhysicalNumberOfCells();cellNum++){
                HSSFCell cell = row.getCell(cellNum);
                String s = row.getCell(cellNum).toString();
//       获取请假信息
                if (rowNum > 1 && cellNum == 1){
                    Date date = sheet.getRow(2).getCell(4).getDateCellValue();
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    String timeS = sdf.format(date);
                    String nameS = sheet.getRow(rowNum).getCell(cellNum).toString();
                    leaveInfo = getLeaveInfo(timeS,nameS);

                }

//                日期处理成当月日期并设置样式
                if (rowNum==1 &&cellNum > 4&& sheet.getPhysicalNumberOfRows() > 2){
                    cell.setCellValue(sheet.getRow(2).getCell(4).getNumericCellValue()+cellNum-5);
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
                if (rowNum>1 && cellNum == 4 && sheet.getPhysicalNumberOfRows() > 2){
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
                if (cellNum>=5){
                    sheet.setColumnWidth(cellNum,20*256);
                }
//                去除多余数据
                if (rowNum>=2 && cellNum >= 5){
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
                        if (isWeekend(sheet.getRow(1).getCell(cellNum))){
                            cell.setCellStyle(style2);
                        }
                        Date date = sheet.getRow(1).getCell(cellNum).getDateCellValue();
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        String timeS = sdf.format(date);
//                        System.out.println(nameS);
                        if(isAskForLeave(timeS, leaveInfo)){
                            HSSFCellStyle styleAskForLeave = workbook.createCellStyle();
                            styleAskForLeave.setFillForegroundColor(IndexedColors.GREEN.getIndex());
                            styleAskForLeave.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            styleAskForLeave.setAlignment(HorizontalAlignment.CENTER);
                            styleAskForLeave.setBorderBottom(BorderStyle.THIN);
                            styleAskForLeave.setBorderRight(BorderStyle.THIN);
                            styleAskForLeave.setBorderLeft(BorderStyle.THIN);
                            styleAskForLeave.setBorderTop(BorderStyle.THIN);
                            cell.setCellStyle(styleAskForLeave);
                        }
                        cell.setCellValue(s1);
                    }
                    else {
                        style0.setDataFormat(format.getFormat("hh:mm"));
                        cell.setCellStyle(style0);
                        if (isWeekend(sheet.getRow(1).getCell(cellNum))){
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
                        Date date = sheet.getRow(1).getCell(cellNum).getDateCellValue();
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        String timeS = sdf.format(date);
//                        System.out.println(timeS);
                        if (isAskForLeave(timeS, leaveInfo)){
                            HSSFCellStyle styleAskForLeave = workbook.createCellStyle();
                            styleAskForLeave.setFillForegroundColor(IndexedColors.GREEN.getIndex());
                            styleAskForLeave.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            styleAskForLeave.setAlignment(HorizontalAlignment.CENTER);
                            styleAskForLeave.setBorderBottom(BorderStyle.THIN);
                            styleAskForLeave.setBorderRight(BorderStyle.THIN);
                            styleAskForLeave.setBorderLeft(BorderStyle.THIN);
                            styleAskForLeave.setBorderTop(BorderStyle.THIN);
                            styleAskForLeave.setDataFormat(format.getFormat("hh:mm"));
                            cell.setCellStyle(styleAskForLeave);
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
    public String splitByDepartment(String inputPath,String inputName, String departmentName) throws IOException {
        File inputFile = new File(inputPath+"/"+inputName);
        inputFile.mkdir();
//        System.out.println(inputFile.getPath());
        String inputs = inputPath+"/temp";
        File outputFile = new File(inputs);
        outputFile.mkdir();
        String outPath = outputFile.getPath();
        FileInputStream inputStream = new FileInputStream(inputFile);
//        System.out.println(outPath);
        FileOutputStream outputStream = new FileOutputStream(outputFile+"\\"+departmentName+".xls");//拆分后的文件的输出路径


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
//        inputStream.close();
        outputStream.close();
//        System.out.println(departmentName+"：分拆完成！");
        format(outPath,departmentName);
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

    @Override
    public List<Map<String, Long>> getLeaveInfo(String times, String names){
        String startTime = times.substring(0,times.lastIndexOf("-"));
        String userName = names;
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM");
        Date date = new Date();
        try {
            date = sdf.parse(startTime);
        }catch (Exception e){
            e.printStackTrace();
        }

        CloseableHttpClient client = HttpClients.createDefault();
        HttpPost post = new HttpPost("https://app.botong.tech:443/approval/approval/ApprovalHistoryAttr");
        List<NameValuePair> parms = new ArrayList<NameValuePair>();
        parms.add(new BasicNameValuePair("companyName","新沂必康新医药产业综合体"));
        parms.add(new BasicNameValuePair("startTime",Long.toString(date.getTime())));
        parms.add(new BasicNameValuePair("approvalName","请假"));
        parms.add(new BasicNameValuePair("userName",userName));
        List<Map<String,Long>> result = new ArrayList<Map<String, Long>>();
        try {
            post.setHeader("Content-Type","application/x-www-form-urlencoded");
            post.setEntity(new UrlEncodedFormEntity(parms,"utf-8"));
            CloseableHttpResponse response = client.execute(post);
            String jsonString = EntityUtils.toString(response.getEntity());
            Gson gson = new Gson();
            JsonObject jsonObject = gson.fromJson(jsonString,JsonObject.class);
            if (jsonObject.get("statusCode").getAsInt()==200) {
                JsonArray jsonArray = jsonObject.getAsJsonArray("data");
                if (jsonArray == null){
                    result=null;
                }else {
                    for (int i = 0; i < jsonArray.size(); i++) {
                        Map<String, Long> map = new HashMap<String, Long>();
                        JsonObject time = jsonArray.get(i).getAsJsonObject().get("field2").getAsJsonObject();
                        String valueS = time.get("value").toString();
                        String startTimeS = valueS.substring(valueS.indexOf("\"") + 1, valueS.indexOf(","));
                        String endTimeS = valueS.substring(valueS.indexOf(",") + 1, valueS.lastIndexOf("\""));
                        System.out.println("startTime:" + startTimeS);
                        System.out.println("endTime:" + endTimeS);
                        Long startTimeL = Long.parseLong(startTimeS);
                        Long endTimeL = Long.parseLong(endTimeS);
                        SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
                        Date startDate, endDate;
                        startDate = format.parse(format.format(startTimeL));
                        endDate = format.parse(format.format(endTimeL));
                        map.put("startTime", startDate.getTime());
                        map.put("endTime", endDate.getTime());
                        System.out.println("startTime:" + format.format(startDate) + "\tendTime:" + format.format(endDate));
                        result.add(map);
                    }
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            return result;
        }
    }
    @Override
    public boolean isAskForLeave(String times, List<Map<String, Long>> list){
        String dateStr = times;
        try {
            SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
            Date date = format.parse(dateStr);
            Long dateL = date.getTime();
            if (list == null){
                System.out.println("NULL!!!");
            }else{
                for (int i = 0; i < list.size(); i++) {
                    Map<String, Long> map = list.get(i);
                    Long l1 = map.get("startTime");
                    Long l2 = map.get("endTime");
                    if (dateL >= l1 && dateL <= l2) {
                        return true;
                    }
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        return false;
    }

}
