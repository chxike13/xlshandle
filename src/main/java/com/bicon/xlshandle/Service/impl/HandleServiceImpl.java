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
    /**
     * @Description: 格式化excel文件，主要为去除多余打卡数据，迟到早退标红，周末标黄，请假标绿。
     * @Param:  inputPath输入文件的路径；inputName输入文件名（部门名）
     * @return:  无返回值但会在输入路径的目录下新建一个result文件夹输出处理后的文件。
     * @Author: xike
     * @Date: 2018/8/7
     */
    @Override
    public void format(String inputPath, String inputName)throws IOException{
        //      读取文件，创建工作空间
        File file = new File(inputPath+"/result");
        file.mkdir();
//        创建输入流和输出流
        FileOutputStream output = new FileOutputStream(file+"\\"+inputName+"result.xls");
        FileInputStream input = new FileInputStream(inputPath+"\\"+inputName+".xls");
        HSSFWorkbook workbook = new HSSFWorkbook(input);
        HSSFSheet sheet = workbook.getSheetAt(0);

//      多数据(含有两个或两个以上时间数据)异常项样式
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.RED.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        HSSFDataFormat format = workbook.createDataFormat();

//      单数据（含有0个或1个时间数据）异常项样式
        HSSFCellStyle style0 = workbook.createCellStyle();
        style0.setFillForegroundColor(IndexedColors.RED.getIndex());
        style0.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style0.setAlignment(HorizontalAlignment.CENTER);
        style0.setBorderBottom(BorderStyle.THIN);
        style0.setBorderRight(BorderStyle.THIN);
        style0.setBorderLeft(BorderStyle.THIN);
        style0.setBorderTop(BorderStyle.THIN);
//        周末样式（背景色标黄）
        HSSFCellStyle style2 = workbook.createCellStyle();
        style2.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style2.setAlignment(HorizontalAlignment.CENTER);
        style2.setBorderBottom(BorderStyle.THIN);
        style2.setBorderRight(BorderStyle.THIN);
        style2.setBorderLeft(BorderStyle.THIN);
        style2.setBorderTop(BorderStyle.THIN);
//        用于装载请假信息的list
        List<Map<String, Long>> leaveInfo = null;
//      遍历文件
        for (int rowNum = 0;rowNum < sheet.getPhysicalNumberOfRows();rowNum++){
            HSSFRow row = sheet.getRow(rowNum);
            for (int cellNum = 0;cellNum < row.getPhysicalNumberOfCells();cellNum++){
                HSSFCell cell = row.getCell(cellNum);
                String s = row.getCell(cellNum).toString();
//       当出现新员工（遍历到第3列）时获取该员工请假信息
                if (rowNum > 1 && cellNum == 1){
//                    获取月份（第5列的数据）并转换成("yyyy-MM-dd")格式的字符串
                    Date date = sheet.getRow(2).getCell(4).getDateCellValue();
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    String timeS = sdf.format(date);
//                    获取姓名（本单元格的数据）
                    String nameS = sheet.getRow(rowNum).getCell(cellNum).toString();
//                    以月份和姓名作为参数调用getLeaveInfo（）获取请假信息。
                    leaveInfo = getLeaveInfo(timeS,nameS);

                }

//                对第二行第六列以后的日期设置格式，因为原数据不是当月的数据而是1970年1月1日，所以需要先以第三行第五列的月份数据为基准获取当月日期再设置样式。
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
//              设置第五列的月份数据的样式
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
                if (rowNum>=2 && cellNum >= 5){
//                    有两个或以上打卡数据时，去除多余打卡数据，判断是否迟到早退。
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
//                        判断当日是否为周末
                        if (isWeekend(sheet.getRow(1).getCell(cellNum))){
                            cell.setCellStyle(style2);
                        }
//                       处理日期格式以便传参给 isAskForLeave（）判断是否请假。
                        Date date = sheet.getRow(1).getCell(cellNum).getDateCellValue();
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        String timeS = sdf.format(date);
//                       判断当日是否请假
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
//                    只有一个或0个打卡数据，直接判定为早退或旷工
                    else {
                        style0.setDataFormat(format.getFormat("hh:mm"));
                        cell.setCellStyle(style0);
//                      判定是否周末
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
//                       预处理准备传参
                        Date date = sheet.getRow(1).getCell(cellNum).getDateCellValue();
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        String timeS = sdf.format(date);
//                      判定是否请假
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
    /**
     * @Description: 按部门分拆excel文件
     * @Param:  inputPath输入文件路径；inputName输入文件名；departmentName需要分拆的部门，也是输出的文件名。
     * @return:  返回值无意义懒得改void了，会在输入路径的目录下新建一个temp文件夹输出处理后的文件
     * @Author: xike
     * @Date: 2018/8/7
     */
    @Override
    public String splitByDepartment(String inputPath,String inputName, String departmentName) throws IOException {
//        创建输入输出流
        File inputFile = new File(inputPath+"/"+inputName);
        inputFile.mkdir();
        String inputs = inputPath+"/temp";
        File outputFile = new File(inputs);
        outputFile.mkdir();
        String outPath = outputFile.getPath();
        FileInputStream inputStream = new FileInputStream(inputFile);
//        System.out.println(outPath);
        FileOutputStream outputStream = new FileOutputStream(outputFile+"\\"+departmentName+".xls");//拆分后的文件的输出路径

//      创建工作间
        HSSFWorkbook workbookIn = new HSSFWorkbook(inputStream);
        HSSFWorkbook workbookOut = new HSSFWorkbook();
        HSSFSheet sheetIn = workbookIn.getSheetAt(0);
//      用一个队列装载列
        List<List> list = new ArrayList<List>();
//       遍历输入文件
        for(int rowNum = 0;rowNum < sheetIn.getPhysicalNumberOfRows();rowNum++){
            HSSFRow row = sheetIn.getRow(rowNum);
//          公有表头直接装载
            if (rowNum < 2){
                List<HSSFCell> listTemp = new ArrayList<HSSFCell>();
                for (int i = 0;i < row.getPhysicalNumberOfCells();i++){
                    listTemp.add(row.getCell(i));
                }
                list.add(listTemp);
            }
//            部门名称一致则装载
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
//      创建输出文件
        HSSFSheet sheetOut = workbookOut.createSheet(departmentName);
//       合并第一行的单元格
        CellRangeAddress region = new CellRangeAddress(0,0,0,35);
        sheetOut.addMergedRegion(region);
//        将上述队列遍历输出到新文件
        for (int rowNum = 0;rowNum < list.size();rowNum++){
            HSSFRow row = sheetOut.createRow(rowNum);
//            设置行高
            if (rowNum<2) {
                row.setHeight((short) (25 * 20));
            }
            for (int cellNum = 0;cellNum < list.get(rowNum).size();cellNum++){
                row.createCell(cellNum);
                HSSFCell cell = row.getCell(cellNum);
                HSSFCell cellTemp = (HSSFCell) list.get(rowNum).get(cellNum);
//               获取单元格数据格式并按格式输出到新文件
                CellType s = cellTemp.getCellTypeEnum();
                switch (s){
                    case STRING: cell.setCellValue(cellTemp.getStringCellValue());break;
                    case NUMERIC: cell.setCellValue(cellTemp.getNumericCellValue());break;
                    case BLANK: cell.setCellValue(cellTemp.toString());break;
                }
//                克隆输入文件的样式到输出文件
                HSSFCellStyle style = workbookOut.createCellStyle();
                style.cloneStyleFrom(cellTemp.getCellStyle());
                cell.setCellStyle(style);
            }
        }

        workbookOut.write(outputStream);
//        inputStream.close();
        outputStream.close();
//        将拆解后的部门文件传给format()格式化。
        format(outPath,departmentName);
        return departmentName+"：处理完成！";
    }
    /**
     * @Description: 判断当日是否请假
     * @Param:  times当日日期转换成的字符串（格式为“yyyy-MM-dd”）；list含有该员工当月请假信息的队列
     * @return:  true或者false
     * @Author: xike
     * @Date: 2018/8/7
     */
    @Override
    public boolean isWeekend(HSSFCell cell) {
//        获取单元格中的日期
        Date date  = cell.getDateCellValue();
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
//        判定日期是否为周末
        int i = calendar.get(Calendar.DAY_OF_WEEK)-1;
        if (i==0||i==6) {
            return true;
        }
        return false;
    }
    /**
     * @Description: 获取员工当月请假信息并用队列返回
     * @Param:  times月份字符串；names员工姓名
     * @return:  装有员工请假信息的队列，每个map中含有一个startTime和一个endTime，对应着起止时间的13位时间戳，若无请假为null。
     * @Author: xike
     * @Date: 2018/8/7
     */
    @Override
    public List<Map<String, Long>> getLeaveInfo(String times, String names){
//      将"yyyy-MM-dd"切割成"yyyy-MM"只保留月份
        String startTime = times.substring(0,times.lastIndexOf("-"));
        String userName = names;
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM");
        Date date = new Date();
        try {
            date = sdf.parse(startTime);
        }catch (Exception e){
            e.printStackTrace();
        }
//      利用httpclient向伯通app发起请求获取请假数据
        CloseableHttpClient client = HttpClients.createDefault();
        HttpPost post = new HttpPost("https://app.botong.tech:443/approval/approval/ApprovalHistoryAttr");
        List<NameValuePair> parms = new ArrayList<NameValuePair>();
//        设置请求参数
        parms.add(new BasicNameValuePair("companyName","新沂必康新医药产业综合体"));
        parms.add(new BasicNameValuePair("startTime",Long.toString(date.getTime())));
        parms.add(new BasicNameValuePair("approvalName","请假"));
        parms.add(new BasicNameValuePair("userName",userName));
        List<Map<String,Long>> result = new ArrayList<Map<String, Long>>();
        try {
            post.setHeader("Content-Type","application/x-www-form-urlencoded");
            post.setEntity(new UrlEncodedFormEntity(parms,"utf-8"));
            CloseableHttpResponse response = client.execute(post);
//            获取请求返回数据中的json字符串
            String jsonString = EntityUtils.toString(response.getEntity());
//            利用gson解析json
            Gson gson = new Gson();
//            将上述获取的json字符串转换成json格式
            JsonObject jsonObject = gson.fromJson(jsonString,JsonObject.class);
//            判定请求是否成功
            if (jsonObject.get("statusCode").getAsInt()==200) {
                JsonArray jsonArray = jsonObject.getAsJsonArray("data");
//                判定json中是否含有请假数据
                if (jsonArray == null){
                    result=null;
                }else {
                    for (int i = 0; i < jsonArray.size(); i++) {
//                        获取请假数据（起止时间）
                        Map<String, Long> map = new HashMap<String, Long>();
                        JsonObject time = jsonArray.get(i).getAsJsonObject().get("field2").getAsJsonObject();
                        String valueS = time.get("value").toString();
                        String startTimeS = valueS.substring(valueS.indexOf("\"") + 1, valueS.indexOf(","));
                        String endTimeS = valueS.substring(valueS.indexOf(",") + 1, valueS.lastIndexOf("\""));
//                        System.out.println("startTime:" + startTimeS);
//                        System.out.println("endTime:" + endTimeS);
//                        将请假数据由字符串转换成13位时间戳
                        Long startTimeL = Long.parseLong(startTimeS);
                        Long endTimeL = Long.parseLong(endTimeS);
//                        以下代码只是测试用
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
//            将当前日期转换成13位时间戳
            SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
            Date date = format.parse(dateStr);
            Long dateL = date.getTime();
            if (list == null){
                System.out.println("NULL!!!");
            }else{
//                遍历队列判断当前日期是否处于某次请假的起止时间内
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
