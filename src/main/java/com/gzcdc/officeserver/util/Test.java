package com.gzcdc.officeserver.util;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


public class Test {  
      
    public static void main(String[] args) throws Exception {
        WordUtil wordUtil=new WordUtil();
    	 Map<String, Object> param = new HashMap<String, Object>();

         /*param.put("${name}", "panxianfeng");
       param.put("${zhuanye}", "信息管理与信息系统");  
        
        param.put("${school_name}", "山东财经大学");  
        param.put("${date}", new Date().toString());*/  
          
    	 param.put("sex", "男");
        param.put("name", "蓝善根");

        Map<String,Object> header = new HashMap<String, Object>();  
        header.put("width", 100);  
        header.put("height", 150);
        header.put("type", "jpg");
        //header.put("content", wordUtil.inputStream2ByteArray(new FileInputStream("D:\\1.jpg"), true));
        param.put("${header}",header);

        List<Map<String, Object>> paramSubs = new ArrayList<Map<String, Object>>();
        Map<String,Object> row = new HashMap<String, Object>();
        row.put("col1", "andy1");
        row.put("col2", "andy2");
        row.put("col3", "andy3");
        paramSubs.add(row);

        row = new HashMap<String, Object>();
        row.put("col1", "jack1");
        row.put("col2", "jack2");
        row.put("col3", "jack3");
        paramSubs.add(row);

        row = new HashMap<String, Object>();
        row.put("col1", "fk1");
        row.put("col2", "fk2");
        row.put("col3", "fk3");
        paramSubs.add(row);

        param.put("alist",paramSubs);
          
        /*Map<String,Object> twocode = new HashMap<String, Object>();  
        twocode.put("width", 100);  
        twocode.put("height", 100);  
        twocode.put("type", "png");  
        twocode.put("content", ZxingEncoderHandler.getTwoCodeByteArray("测试二维码,huangqiqing", 100,100));  
        param.put("${twocode}",twocode);  */
          
        CustomXWPFDocument doc = wordUtil.generateWordMultiRows(param, "D:\\GzcdcJL12.docx");
        doc.write(new FileOutputStream("D:\\2.docx"));
        FileOutputStream fopts = new FileOutputStream("D:\\2.pdf");  
        
       PdfOptions options = PdfOptions.create();
        PdfConverter.getInstance().convert(doc, fopts, options);
        
        //doc.write(fopts);  
        fopts.close();  
    }  
}  