package com.bicon.xlshandle.Controller;

import com.bicon.xlshandle.Service.HandleService;
import com.sun.org.apache.xerces.internal.xs.datatypes.ObjectList;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Controller
public class HandleController {
    @Resource(name = "handleServiceImpl")
    HandleService handleService;
    @RequestMapping("/handle")
    public Object handle(HttpServletRequest request, HttpServletResponse response){
        Map<String, Object> retData = new HashMap<String, Object>();
        String filePath = request.getServletContext().getRealPath("/resource/input");
        File file = new File(filePath);
//        String[] department = {"安保部","GMP","必康综合体",
//                "财务部","采购部","工程部",
//                "工程管理办公室","国际贸易部","行政部",
//                "行政部（后勤）","护理品项目（生产）","护理品项目（营销）",
//                "企宣部","人力资源部","软件部（信息大部）",
//                "数据中心","项目申报办公室","新阳",
//                "信息部（信息大部）","证券部","智能化部",
//                "资金运营管理中心","自动化部"
//        };
//        try {
//            List<String> messsageList = new ArrayList<String>();
//            for (int i = 0; i < department.length; i++){
//                messsageList.add(handleService.splitByDepartment(filePath,department[i]));
//            }
//            retData.put("success",1);
//            retData.put("msg",messsageList);
//        }catch (Exception e){
//            System.out.println(e.toString());
//            retData.put("success",-1);
//        }
//        finally {
//            return retData;
//        }
        return "Hello";
    }
}
