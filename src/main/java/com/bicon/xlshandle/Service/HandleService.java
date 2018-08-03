package com.bicon.xlshandle.Service;

import org.apache.poi.hssf.usermodel.HSSFCell;

import java.io.IOException;

public interface  HandleService {
    public void format(String inputName)throws IOException;
    public String splitByDepartment(String inputPath, String departmentName)throws IOException;
    public boolean isWeekend(HSSFCell cell);
}
