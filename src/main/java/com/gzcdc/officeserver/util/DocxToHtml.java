package com.gzcdc.officeserver.util;

import fr.opensagres.poi.xwpf.converter.core.ImageManager;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;

/**
 * Created by User: skh.
 * Date: 2018/1/25 Time: 16:11.
 * Description:Docx转HTML
 */
public class DocxToHtml {
    public static String wordToHtml(String sourceFileName) {
        File sourceFile = new File(sourceFileName);
        String fileName = sourceFile.getName();

        int index = sourceFileName.lastIndexOf('.');
        String prefix = sourceFileName.substring(0, index);
        String htmlName = prefix + ".html";
        File htmlFile = new File(htmlName);

        if (!sourceFile.exists()) {
            System.out.println("Sorry File does not Exists!");
            return null;
        }
        InputStream in = null;
        OutputStream out = null;
        try {
            File fileDir = new File(prefix);
            if (!fileDir.exists()) {
                fileDir.mkdirs();
            }

            in = new FileInputStream(sourceFile);
            out = new FileOutputStream(htmlFile);
            XWPFDocument document = new XWPFDocument(in);

            // 解析 XHTML配置 (这里设置IURIResolver来设置图片存放的目录)
            ImageManager imageManager = new ImageManager(fileDir, "image");
            XHTMLOptions options = XHTMLOptions.create();
            options.setImageManager(imageManager);
            options.setIgnoreStylesIfUnused(false);

            // 将 XWPFDocument转换成XHTML
            //XWPF2XHTMLConverter.getInstance().convert(in, out, options);
            XHTMLConverter.getInstance().convert(document, out, options);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (in != null) {
                    in.close();
                }
                if (out != null) {
                    out.close();
                }

            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        String content = readHtmlContent(htmlFile);

        // 取得html代码后将html文件删除
//            File file = new File(htmlName);
//            if (file.exists()) {
//                if (!file.delete()) {
//                    System.out.println("Doc util =======> delete temp html file fail");
//                }
//            }

        return content;
    }

    private static String readHtmlContent(File file) {
        //File file = new File(filename);
        Long length = file.length();
        byte[] content = new byte[length.intValue()];

        try {
            FileInputStream in = new FileInputStream(file);
            in.read(content);
            in.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            return null;
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }

        return new String(content);
    }
}
