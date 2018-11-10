package com.gzcdc.officeserver.util;

import java.io.File;
import java.text.DecimalFormat;

/**
 * Created by skh on 2017/10/26.
 * 文件大小计算工具类
 */
public class FileSizeUtil {

    /**
     * 转换文件大小
     * 缺点丢失精度
     * @param filePath
     * @return
     */
    public static String fileSize(String filePath) {
        File file = new File(filePath);
        long size;
        if (file.exists() && file.isFile()) {
            size = file.length();
        } else {
            return "没有文件";
        }
        //long size = filesize;
        //如果字节数少于1024，则直接以B为单位，否则先除于1024，后3位因太少无意义
        if (size < 1024) {
            return String.valueOf(size) + "B";
        } else {
            size = size / 1024;
        }
        //如果原字节数除于1024之后，少于1024，则可以直接以KB作为单位
        //因为还没有到达要使用另一个单位的时候
        //接下去以此类推
        if (size < 1024) {
            return String.valueOf(size) + "KB";
        } else {
            size = size / 1024;
        }
        if (size < 1024) {
            //因为如果以MB为单位的话，要保留最后1位小数，
            //因此，把此数乘以100之后再取余
            size = size * 100;
            return String.valueOf((size / 100)) + "."
                    + String.valueOf((size % 100)) + "MB";
        } else {
            //否则如果要以GB为单位的，先除于1024再作同样的处理
            size = size * 100 / 1024;
            return String.valueOf((size / 100)) + "."
                    + String.valueOf((size % 100)) + "GB";
        }
    }

    /**
     * 转换文件大小
     * @param filePath
     * @return
     */
    public static String FormatFileSize(String filePath)
    {
        long fileS;
        DecimalFormat df = new DecimalFormat("#.0");
        File file = new File(filePath);
        if (file.exists() && file.isFile()) {
            fileS = file.length();
        } else {
            return "没有文件";
        }
        String fileSizeString = "";
        String wrongSize="0B";
        if(fileS == 0){
            return wrongSize;
        }
        if (fileS < 1024){
            fileSizeString = df.format((double) fileS) + "B";
        } else if (fileS < 1048576){
            fileSizeString = df.format((double) fileS / 1024) + "KB";
        } else if (fileS < 1073741824){
            fileSizeString = df.format((double) fileS / 1048576) + "MB";
        } else{
            fileSizeString = df.format((double) fileS / 1073741824) + "GB";
        }
        return fileSizeString;
    }
}
