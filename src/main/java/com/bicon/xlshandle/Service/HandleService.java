package com.bicon.xlshandle.Service;

import org.apache.poi.hssf.usermodel.HSSFCell;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public interface  HandleService {
    public void format(String inputPath, String inputName)throws IOException;
    public String splitByDepartment(String inputPath,String inputName, String departmentName)throws IOException;
    public boolean isWeekend(HSSFCell cell);
    public boolean isAskForLeave(String times,List<Map<String, Long>> list);
    public List<Map<String, Long>> getLeaveInfo(String times, String names);
}
