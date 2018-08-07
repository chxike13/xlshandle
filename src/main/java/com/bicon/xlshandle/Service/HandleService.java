package com.bicon.xlshandle.Service;

import org.apache.poi.hssf.usermodel.HSSFCell;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public interface  HandleService {
    /**
    * @Description: 格式化excel文件，主要为去除多余打卡数据，迟到早退标红，周末标黄，请假标绿。
    * @Param:  inputPath输入文件的路径；inputName输入文件名（部门名）
    * @return:  无返回值但会在输入路径的目录下新建一个result文件夹输出处理后的文件。
    * @Author: xike
    * @Date: 2018/8/7
    */
    public void format(String inputPath, String inputName)throws IOException;
    /**
    * @Description: 按部门分拆excel文件
    * @Param:  inputPath输入文件路径；inputName输入文件名；departmentName需要分拆的部门，也是输出的文件名。
    * @return:  返回值无意义懒得改void了，会在输入路径的目录下新建一个temp文件夹输出处理后的文件
    * @Author: xike
    * @Date: 2018/8/7
    */
    public String splitByDepartment(String inputPath,String inputName, String departmentName)throws IOException;
    /**
    * @Description: 判断当日是否为周末
    * @Param:  cell当前单元格所在日期的单元格（excel文件中第二行第六列之后的单元格）
    * @return: true或者false
    * @Author: xike
    * @Date: 2018/8/7
    */
    public boolean isWeekend(HSSFCell cell);
    /**
    * @Description: 判断当日是否请假
    * @Param:  times当日日期转换成的字符串（格式为“yyyy-MM-dd”）；list含有该员工当月请假信息的队列
    * @return:  true或者false
    * @Author: xike
    * @Date: 2018/8/7
    */
    public boolean isAskForLeave(String times,List<Map<String, Long>> list);
    /**
    * @Description: 获取员工当月请假信息并用队列返回
    * @Param:  times月份字符串；names员工姓名
    * @return:  装有员工请假信息的队列，每个map中含有一个startTime和一个endTime，对应着起止时间的13位时间戳，若无请假为null。
    * @Author: xike
    * @Date: 2018/8/7
    */
    public List<Map<String, Long>> getLeaveInfo(String times, String names);
}