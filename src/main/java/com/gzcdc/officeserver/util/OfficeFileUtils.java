package com.gzcdc.officeserver.util;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * Created by User: skh.
 * Date: 2017/10/31 Time: 9:16.
 * Description: 文件工具类
 */
public class OfficeFileUtils {

    // Linux环境下的根目录
   // public static final String LINUXROOTPATH = "/home/pmis";
    // window
   public static final String LINUXROOTPATH = "E:/website/GzcdcPmis/AllFile";//正式环境
  //  public static final String LINUXROOTPATH = "D:/WebSite";//测试环境

    /**
     * 打包文件
     *
     * @param files
     * @param outputStream
     * @throws IOException
     * @throws ServletException
     */
    public static void zipFile(List<File> files, ZipOutputStream outputStream) throws IOException, ServletException {
        try {
            int size = files.size();
            // 压缩列表中的文件
            for (int i = 0; i < size; i++) {
                File file = (File) files.get(i);
                zipFile(file, outputStream);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // 单个文件
    public static void zipFile(File inputFile, ZipOutputStream outputstream) throws IOException, ServletException {
        try {
            if (inputFile.exists()) {
                if (inputFile.isFile()) {
                    FileInputStream inStream = new FileInputStream(inputFile);
                    BufferedInputStream bInStream = new BufferedInputStream(inStream);
                    ZipEntry entry = new ZipEntry(inputFile.getName());
                    outputstream.putNextEntry(entry);

                    final int MAX_BYTE = 10 * 1024 * 1024; // 最大的流为10M
                    long streamTotal = 0; // 接受流的容量
                    int streamNum = 0; // 流需要分开的数量
                    int leaveByte = 0; // 文件剩下的字符数
                    byte[] inOutbyte; // byte数组接受文件的数据

                    streamTotal = bInStream.available(); // 通过available方法取得流的最大字符数
                    streamNum = (int) Math.floor(streamTotal / MAX_BYTE); // 取得流文件需要分开的数量
                    leaveByte = (int) streamTotal % MAX_BYTE; // 分开文件之后,剩余的数量

                    if (streamNum > 0) {
                        for (int j = 0; j < streamNum; ++j) {
                            inOutbyte = new byte[MAX_BYTE];
                            // 读入流,保存在byte数组
                            bInStream.read(inOutbyte, 0, MAX_BYTE);
                            outputstream.write(inOutbyte, 0, MAX_BYTE); // 写出流
                        }
                    }
                    // 写出剩下的流数据
                    inOutbyte = new byte[leaveByte];
                    bInStream.read(inOutbyte, 0, leaveByte);
                    outputstream.write(inOutbyte);
                    outputstream.closeEntry(); // Closes the current ZIP entry
                    // and positions the stream for
                    // writing the next entry
                    bInStream.close(); // 关闭
                    inStream.close();
                }
            } else {
                throw new ServletException("文件不存在！");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 下载文件
     *  ISO-8859-1
     * @param file
     * @param response
     * @param isDelete
     */
    public static void downloadFile(File file, HttpServletResponse response, boolean isDelete) {
        try {
            // 以流的形式下载文件。
            BufferedInputStream fis = new BufferedInputStream(new FileInputStream(file.getPath()));

            // 清空response
            response.reset();
            OutputStream toClient = new BufferedOutputStream(response.getOutputStream());
           // response.setContentType("application/octet-stream");
            response.setContentType("application/x-download");
            response.addHeader("Content-Length",""+file.length());
            //response.setHeader("Content-Disposition", "attachment;filename=" + new String(file.getName().getBytes("UTF-8"), "UTF-8"));
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(file.getName(), "utf-8"));
            response.setContentType("application/octet-stream");
            // byte[] buffer = new byte[fis.available()];
            byte[] buffer = new byte[1024];
            int i=-1;
            while (((i=fis.read(buffer))!=-1)){
                toClient.write(buffer,0,i);
            }
          //  fis.read(buffer);
            fis.close();
           // toClient.write(buffer);
            toClient.flush();
            toClient.close();
            if (isDelete) {
                file.delete(); // 是否将生成的服务器端文件删除
            }
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }
}
