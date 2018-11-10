package com.gzcdc.officeserver.controller;

import com.gzcdc.officeserver.util.OfficeFileUtils;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;
import java.util.zip.ZipOutputStream;

/**
 * Created by lanshg on 2017/10/26.
 */

@Api(value = "文档下载服务", description = "根据id文件名下载word/excel/pdf等文档的相关API")
@RestController
@RequestMapping(value = "filedownload")
public class FileDownloadController {

    @ApiOperation("下载文档")
    @RequestMapping(value = "/download", method = RequestMethod.GET)
    public void download(@RequestParam String fileUrl, HttpServletRequest request, HttpServletResponse response)  throws Exception {

        if (fileUrl.isEmpty() || fileUrl.lastIndexOf('.') < 1)
            return;
        int index = fileUrl.lastIndexOf('.');
        String contentType = fileUrl.substring(index + 1);
        // 告诉浏览器用什么软件可以打开此文件
        response.setHeader("content-Type", "application/"+contentType);
        // 下载文件的默认名称
        String fileName = fileUrl.substring(fileUrl.lastIndexOf("/") + 1);
        response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "utf-8"));
        response.setCharacterEncoding("UTF-8");

        //String rootPath = request.getSession().getServletContext().getRealPath("/");
        //System.out.println(rootPath);
        String filePath = OfficeFileUtils.LINUXROOTPATH + fileUrl;
        System.out.println(filePath);

        ServletOutputStream out;
        try {
            FileInputStream inputStream = new FileInputStream(filePath);
            //3.通过response获取ServletOutputStream对象(out)
            out = response.getOutputStream();

            int b = 0;
            byte[] buffer = new byte[1024];
            while (b != -1){
                b = inputStream.read(buffer);
                //4.写到输出流(out)中
                out.write(buffer,0, b);
            }
            inputStream.close();
            out.close();
            out.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

//    /**
//     * 打包下载
//     * 单个文件直接下载
//     * @param request
//     * @param response
//     * @return
//     * @throws ServletException
//     * @throws IOException
//     */
//    @ApiOperation("打包下载")
//    @RequestMapping(value = "/downloadZip", method = RequestMethod.GET)
//    public String downloadFiles(@RequestParam String[] fileUrls, HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
//
//        //String ppp = "E://temp";
//
//        if (fileUrls.length == 0) {
//            return "没有文件";
//        }
//        if (fileUrls.length == 1) {
//            try {
//                singleFile(fileUrls[0], request, response);
//            } catch (Exception e) {
//                e.printStackTrace();
//            }
//
//            return "文件不存在!";
//        }
//        // 1:创建文件集合，将需要下载的文件全部装入集合中。
//        List<File> files = new ArrayList<>();
//        for (String url : fileUrls) {
//            // 1.1查找文件是否存在，存在，及添加到文件集合中
//            if (url.indexOf("/upload") == -1) {
//                continue;
//            }
//            // linux环境下使用
//            String subUrl = url.substring(url.indexOf("/upload"));
//            // window环境下使用
//            //String subUrl = url.substring(url.indexOf("/upload")).replace("/", "\\");
//            File file = new File(OfficeFileUtils.LINUXROOTPATH + subUrl);
//            //File file = new File(ppp + subUrl);
//            if (file.exists()) {
//                files.add(file);
//            }
//        }



        // // 判断文件存在情况;
        // if (files.size() == 0) {
        // return "选择文件不存在!";
        // }

//        // 2:下载文件名字，及临时打包目录
//        String fileName = UUID.randomUUID().toString() + ".zip";
//        // 在服务器端创建打包下载的临时文件
//        String outFilePath = OfficeFileUtils.LINUXROOTPATH + "/upload/";
//        //String outFilePath = ppp + "/upload/";
//
//        // 3:开始创建文件流并下载
//        File fileZip = new File(outFilePath + fileName);
//        // 文件输出流
//        FileOutputStream outStream = new FileOutputStream(fileZip);
//        // 压缩流
//        ZipOutputStream toClient = new ZipOutputStream(outStream);
//
//        // 4:调用操作函数，下载文件
//        OfficeFileUtils.zipFile(files, toClient);
//        toClient.close();
//        outStream.close();
//        OfficeFileUtils.downloadFile(fileZip, response, true);
//
//        // if (files.size() != fileUrls.length) {
//        // return "存在文件：" + files.size() + "个，不存在文件：";
//        // } else {
//        // return "下载成功!";
//        // }
//        return null;
//    }

    /**
     * 打包下载
     * 单个文件直接下载
     * @param request
     * @param response
     * @return
     * @throws ServletException
     * @throws IOException
     */
    @ApiOperation("打包下载")
    @RequestMapping(value = "/downloadZip", method = {RequestMethod.POST, RequestMethod.GET})
    public String downloadFiles(@RequestParam String[] fileUrls, String fileName, HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        //fileName = URLDecoder.decode(fileName, "UTF-8");
        //String ppp = "E://temp";

        if (fileUrls.length == 0) {
            return "没有文件";
        }
        if (fileUrls.length == 1) {
            try {
                //String fileUrl = URLDecoder.decode(fileUrls[0], "UTF-8");
                singleFile(fileUrls[0], request, response);
            } catch (Exception e) {
                e.printStackTrace();
            }

            return "文件不存在!";
        }
        // 1:创建文件集合，将需要下载的文件全部装入集合中。
        List<File> files = new ArrayList<>();
        for (String url : fileUrls) {
            //url = URLDecoder.decode(url, "UTF-8");
            // 1.1查找文件是否存在，存在，及添加到文件集合中
            if (url.indexOf("/upload") == -1) {
                continue;
            }
            // linux环境下使用
            String subUrl = url.substring(url.indexOf("/upload"));
            // window环境下使用
            //String subUrl = url.substring(url.indexOf("/upload")).replace("/", "\\");
            File file = new File(OfficeFileUtils.LINUXROOTPATH + subUrl);
            //File file = new File(ppp + subUrl);
            if (file.exists()) {
                files.add(file);
            }
        }



        // // 判断文件存在情况;
        // if (files.size() == 0) {
        // return "选择文件不存在!";
        // }

        // 2:下载文件名字，及临时打包目录
        if (fileName == null || fileName.isEmpty()) {
            fileName = UUID.randomUUID().toString() + ".zip";
        } else {
            fileName = fileName + ".zip";
        }
        //fileName = fileName.isEmpty() ?  UUID.randomUUID().toString() + ".zip" : fileName+".zip";

        // 在服务器端创建打包下载的临时文件
        String outFilePath = OfficeFileUtils.LINUXROOTPATH + "/upload/";
        //String outFilePath = ppp + "/upload/";

        // 3:开始创建文件流并下载
        File fileZip = new File(outFilePath + fileName);
        // 文件输出流
        FileOutputStream outStream = new FileOutputStream(fileZip);
        // 压缩流
        ZipOutputStream toClient = new ZipOutputStream(outStream);

        // 4:调用操作函数，下载文件
        OfficeFileUtils.zipFile(files, toClient);
        toClient.close();
        outStream.close();
        OfficeFileUtils.downloadFile(fileZip, response, true);

        // if (files.size() != fileUrls.length) {
        // return "存在文件：" + files.size() + "个，不存在文件：";
        // } else {
        // return "下载成功!";
        // }
        return null;
    }

    public void singleFile(String fileUrl, HttpServletRequest request, HttpServletResponse response) throws Exception{
        //String ppp = "E://temp";

        if (fileUrl.isEmpty() || fileUrl.lastIndexOf('.') < 1)
            return;
        int index = fileUrl.lastIndexOf('.');
        String contentType = fileUrl.substring(index + 1);
        // 告诉浏览器用什么软件可以打开此文件
        response.setHeader("content-Type", "application/"+contentType);
        // 下载文件的默认名称
        String fileName = fileUrl.substring(fileUrl.lastIndexOf("/") + 1);
        response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "utf-8"));
        response.setCharacterEncoding("UTF-8");

        // linux环境下使用
        String subUrl = fileUrl.substring(fileUrl.indexOf("/upload"));
        // window环境下使用
        // String subUrl = fileUrl.substring(fileUrl.indexOf("/upload")).replace("/", "\\");
        String filePath =  OfficeFileUtils.LINUXROOTPATH + subUrl;
        //String filePath =  ppp + subUrl;

        ServletOutputStream outputStream;
        try {
            FileInputStream inputStream = new FileInputStream(filePath);
            //3.通过response获取ServletOutputStream对象(out)
            outputStream = response.getOutputStream();
            int b = 0;
            byte[] buffer = new byte[1024];
            while (b != -1) {
                b = inputStream.read(buffer);
                //4.写到输出流(out)中
                outputStream.write(buffer, 0, b);
            }
            inputStream.close();
            outputStream.close();
            outputStream.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
