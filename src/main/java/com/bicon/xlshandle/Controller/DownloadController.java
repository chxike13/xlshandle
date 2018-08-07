package com.bicon.xlshandle.Controller;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLEncoder;

/**
* @Description: 提供下载接口下载处理后的文件。
* @Param:  request
* @return:
* @Author: xike
* @Date: 2018/8/7
*/
@RestController
public class DownloadController {
    @RequestMapping("/download")
    public void downloadFile(HttpServletRequest request, HttpServletResponse response){
        try {
            request.setCharacterEncoding("UTF-8");
            String name = request.getParameter("fileName");
            response.setContentType("application/force-download");
            String path = request.getServletContext().getRealPath("/UploadFiles/temp/result/"+name);
            InputStream inputStream = new FileInputStream(path);
            name = URLEncoder.encode(name, "UTF-8");
            response.setHeader("Content-Disposition", "attachment;filename="+name);
            response.setContentLength(inputStream.available());
            OutputStream outputStream = response.getOutputStream();
            byte[] b = new byte[1024];
            int len = 0;
            while((len = inputStream.read(b))!=-1){
                outputStream.write(b, 0, len);
            }
            outputStream.flush();
            outputStream.close();
            inputStream.close();
        }catch (Exception e){
            System.out.println(e.toString());
        }
    }
}
