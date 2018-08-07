package com.bicon.xlshandle.Controller;

import com.bicon.xlshandle.Service.HandleService;
import com.sun.org.apache.xerces.internal.xs.datatypes.ObjectList;
import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.*;

@RestController
public class HandleController {
    @Resource(name = "handleServiceImpl")
    HandleService handleService;
    @RequestMapping(value = "/upload", method = RequestMethod.POST)
    public Object uploadFile(HttpServletRequest request, HttpServletResponse response){
        Map<String, Object> retData = new HashMap<String, Object>();
        response.setHeader("Access-Control-Allow-Origin", "*");
        response.addHeader("Access-Control-Allow-Methods","*");
        ArrayList<String> list = new ArrayList<>();
        String savePath = request.getServletContext().getRealPath("/UploadFiles");
        File file = new File(savePath);
        if (!file.exists()){
            System.out.println(savePath);
            file.mkdir();
        }
        try {
            DiskFileItemFactory factory = new DiskFileItemFactory();
            ServletFileUpload upload = new ServletFileUpload(factory);
            upload.setHeaderEncoding("UTF-8");
            if (!ServletFileUpload.isMultipartContent(request)){
                System.out.println("isn't multipart!");
                return null;
            }
            List<FileItem> fileItemList = upload.parseRequest(request);
            System.out.println("file number："+fileItemList.size());
            for (FileItem fileItem : fileItemList){
                if (fileItem.isFormField()){
                    String fileName = fileItem.getName();
                    String value = fileItem.getString("UTF-8");
                    System.out.println(fileName+"="+value);
                }else {
                    String fileName = fileItem.getName();
                    System.out.println("fileName："+fileName);
//                    fileName = fileName.substring(fileName.lastIndexOf("\\"+1));
                    InputStream inputStream = fileItem.getInputStream();
                    FileOutputStream outputStream = new FileOutputStream(savePath+"\\"+fileName);
                    byte Buffer[] = new byte[1024];
                    int len = 0;
                    while ((len=inputStream.read(Buffer))>0){
                        outputStream.write(Buffer,0,len);
                    }
                    inputStream.close();
                    outputStream.close();
                    fileItem.delete();
                    System.out.println("upload done!");

                    String[] department = {"安保部","GMP","必康综合体",
                            "财务部","采购部","工程部",
                            "工程管理办公室","国际贸易部","行政部",
                            "行政部（后勤）","护理品项目（生产）","护理品项目（营销）",
                            "企宣部","人力资源部","软件（信息大部）",
                            "数据中心","项目申报办公室","新阳",
                            "信息部（信息大部）","证券部","智能化部",
                            "资金运营管理中心","自动化部"
                    };
                    for (int i = 0; i < department.length; i++){
                        handleService.splitByDepartment(savePath,fileName,department[i]);
                        list.add(department[i]+"result.xls");
                    }
                    retData.put("success","上传成功！");
                    retData.put("filePath",savePath);
                    retData.put("resultName",list);
                }

            }
        }catch (Exception e){
            System.out.println(e.toString());
            retData.put("sucess","上传失败："+e.toString());
        }finally {
            return retData;
        }

    }
}
