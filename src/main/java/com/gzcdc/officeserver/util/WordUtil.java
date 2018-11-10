package com.gzcdc.officeserver.util;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.math.BigInteger;
import java.util.*;
import java.util.Map.Entry;


public class WordUtil {  
  
    /** 
     * 根据指定的参数值、模板，生成 word 文档 
     * @param param 需要替换的变量 
     * @param template 模板 
     */  
    public CustomXWPFDocument generateWord(Map<String, Object> param, String template) {
        CustomXWPFDocument doc = null;  
        try {  
            OPCPackage pack = POIXMLDocument.openPackage(template);  
            doc = new CustomXWPFDocument(pack);  
            if (param != null && param.size() > 0) {
                //处理段落  
                List<XWPFParagraph> paragraphList = doc.getParagraphs();  
                processParagraphs(paragraphList, param, doc);  
                  
                //处理表格  
                Iterator<XWPFTable> it = doc.getTablesIterator();  
                while (it.hasNext()) {  
                    XWPFTable table = it.next();  
                    List<XWPFTableRow> rows = table.getRows();
                    for (XWPFTableRow row : rows) {
                        List<XWPFTableCell> cells = row.getTableCells();
                        for (XWPFTableCell cell : cells) {  
                            List<XWPFParagraph> paragraphListTable =  cell.getParagraphs();  
                            processParagraphs(paragraphListTable, param, doc);  
                        }  
                    }  
                }  
            }  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        return doc;  
    }

    // 处理项目协作申请表
    public CustomXWPFDocument generateCoopApplyWord(Map<String, Object> param, String template) {
        CustomXWPFDocument doc = null;
        try {
            OPCPackage pack = POIXMLDocument.openPackage(template);
            doc = new CustomXWPFDocument(pack);
            if (param != null && param.size() > 0) {
                //处理段落
                List<XWPFParagraph> paragraphList = doc.getParagraphs();
                processParagraphs(paragraphList, param, doc);

                //处理表格
                Iterator<XWPFTable> it = doc.getTablesIterator();
                while (it.hasNext()) {
                    XWPFTable table = it.next();
                    List<XWPFTableRow> rows = table.getRows();
                    for (XWPFTableRow row : rows) {
                        List<XWPFTableCell> cells = row.getTableCells();
                        for (XWPFTableCell cell : cells) {
                            List<XWPFParagraph> paragraphListTable =  cell.getParagraphs();
                            getKeyInsertTable(param, cell);
                            processParagraphs(paragraphListTable, param, doc);
                            replacePlaceholder(paragraphListTable, doc);
                        }
                    }

                    // 删除行（固定比例，固定总价显示一个）
                    if (param.get("calcWay").equals("fxdProportion")) {
                        boolean a = table.removeRow(6);
                        System.out.println(a);
                    } else if (param.get("calcWay").equals("fxdPrice")) {
                        boolean b = table.removeRow(5);
                        System.out.println(b);
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return doc;
    }

    // 替换掉插入表格的占位符
    private void replacePlaceholder(List<XWPFParagraph> paragraphs, CustomXWPFDocument doc) {
        Map<String, Object> data = new HashMap<>();
        data.put("fxdProportionTable", "");
        data.put("proTable", "");
        data.put("proTechTable", "");
        processParagraphs(paragraphs, data, doc);
    }

    // 根据协作费计算方式（calcWay）或 是否同意部门建议的协作单位（isAgreeSuggestCompany）
    // 在相应位置插入表格
    private void getKeyInsertTable(Map<String, Object> param, XWPFTableCell tableCell) {
        if (param.get("calcWay").equals("fxdProportion")) {
            insertTable("${fxdProportionTable}", param, tableCell, 1);
            insertTable("${proTable}", param, tableCell, 2);
            if (param.get("isAgreeSuggestCompany").equals(3)) {
                insertTable("${proTechTable}", param, tableCell, 2);
            }
        } else if (param.get("calcWay").equals("fxdPrice")) {
            insertTable("${proTable}", param, tableCell, 3);
            if (param.get("isAgreeSuggestCompany").equals(3)) {
                insertTable("${proTechTable}", param, tableCell, 3);
            }
        }
    }

    // 在key处插入表格
    private void insertTable(String key, Map<String, Object> param, XWPFTableCell cell, int kind) {
        List<XWPFParagraph> paragraphList = cell.getParagraphs();
        if (paragraphList != null && paragraphList.size() > 0) {
            for (XWPFParagraph paragraph : paragraphList) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    String text = run.getText(0);
                    if (text != null) {
                        if (text.indexOf(key) >= 0) {
                            insertWhereKindTable(paragraph, param, cell, kind);
                        }
                    }
                }
            }
        }
    }

    // 根据isFxdProportion来确定插入哪种表格
    private void insertWhereKindTable(XWPFParagraph paragraph, Map<String, Object> param,
                                      XWPFTableCell cell, int kind) {
        if (kind == 1) {
            setFxdProportionTableValue(paragraph, param, cell);
        } else if (kind == 2) {
            setFxdProportionOtherTableValue(paragraph, param, cell);
        } else {
            setFixedPriceOtherTableValue(paragraph, param, cell);
        }
    }

    // 固定比列时插入并设置表格的值
    private void setFxdProportionTableValue(XWPFParagraph paragraph, Map<String, Object> param, XWPFTableCell cell) {
        XmlCursor cursor = paragraph.getCTP().newCursor();
        XWPFTable tableOne = cell.insertNewTbl(cursor);// ---这个是关键

        boolean a = tableOne.removeRow(0);

        CTTblGrid grid = tableOne.getCTTbl().addNewTblGrid();
        grid.addNewGridCol().setW(BigInteger.valueOf(2000));
        grid.addNewGridCol().setW(BigInteger.valueOf(2000));
        grid.addNewGridCol().setW(BigInteger.valueOf(2000));

        tableBorderStyle(tableOne);

        List<Map<String, Object>> tableValueList = (List<Map<String, Object>>) param.get("fore_com_list");

        XWPFTableRow tableOneRowOne = tableOne.createRow();
        tableOneRowOne.setHeight(380);
        tableOneRowOne.addNewTableCell().setText("专业");
        tableOneRowOne.addNewTableCell().setText("工作量");
        tableOneRowOne.addNewTableCell().setText("协作比例");
        tableOneRowOne.addNewTableCell().setText("协作费用");

        for (int i = 0; i < tableValueList.size(); i++) {
            Map<String, Object> map = tableValueList.get(i);
            XWPFTableRow tableRow = tableOne.createRow();
            tableRow.setHeight(380);
            tableRow.getCell(0).setText((String) map.get("CoopProfessionName"));
            tableRow.getCell(1).setText(String.valueOf(map.get("CoopWorkProportion")));
            tableRow.getCell(2).setText(String.valueOf(map.get("CoopProportion")));
            tableRow.getCell(3).setText(String.valueOf(map.get("CoopPrice")));
        }

        tableCellStyle(tableOne, 1500);
    }

    // 固定比例时写，在部门领导审核和公司协调时生产技术部审核插入表格
    private void setFxdProportionOtherTableValue(XWPFParagraph paragraph, Map<String, Object> param, XWPFTableCell cell) {
        XmlCursor cursor = paragraph.getCTP().newCursor();
        XWPFTable tableOne = cell.insertNewTbl(cursor);// ---这个是关键

        boolean a = tableOne.removeRow(0);

        CTTblGrid grid = tableOne.getCTTbl().addNewTblGrid();
        grid.addNewGridCol().setW(BigInteger.valueOf(2000));
        grid.addNewGridCol().setW(BigInteger.valueOf(2000));
        grid.addNewGridCol().setW(BigInteger.valueOf(2000));

        tableBorderStyle(tableOne);

        List<Map<String, Object>> tableValueList = (List<Map<String, Object>>) param.get("product_coms");

        XWPFTableRow tableOneRowOne = tableOne.createRow();
        tableOneRowOne.setHeight(380);
        tableOneRowOne.addNewTableCell().setText("专业");
        tableOneRowOne.addNewTableCell().setText("协作费用");
        tableOneRowOne.addNewTableCell().setText("协作单位");

        for (int i = 0; i < tableValueList.size(); i++) {
            Map<String, Object> map = tableValueList.get(i);
            XWPFTableRow tableRow = tableOne.createRow();
            tableRow.setHeight(380);
            tableRow.getCell(0).setText((String) map.get("CoopProfessionName"));
            tableRow.getCell(1).setText(String.valueOf(map.get("CoopPrice")));
            tableRow.getCell(2).setText((String) map.get("CoopCompanyName"));
        }

        tableCellStyle(tableOne, 2000);
    }

    // 固定总价时写，在部门领导审核和公司协调时生产技术部审核插入表格
    private void setFixedPriceOtherTableValue(XWPFParagraph paragraph, Map<String, Object> param, XWPFTableCell cell) {
        XmlCursor cursor = paragraph.getCTP().newCursor();
        XWPFTable tableOne = cell.insertNewTbl(cursor);// ---这个是关键
        boolean a = tableOne.removeRow(0);

        CTTblGrid grid = tableOne.getCTTbl().addNewTblGrid();
        grid.addNewGridCol().setW(BigInteger.valueOf(2000));
        grid.addNewGridCol().setW(BigInteger.valueOf(2000));

        tableBorderStyle(tableOne);

        List<Map<String, Object>> tableValueList = (List<Map<String, Object>>) param.get("product_coms");

        XWPFTableRow tableOneRowOne = tableOne.createRow();
        tableOneRowOne.setHeight(380);
        tableOneRowOne.addNewTableCell().setText("协作单位");
        tableOneRowOne.addNewTableCell().setText("协作费用");
        //tableOneRowOne.addNewTableCell().getParagraphs().get(0).getCTP().addNewR().addNewT().setStringValue("协作费用");

        for (int i = 0; i < tableValueList.size(); i++) {
            Map<String, Object> map = tableValueList.get(i);
            XWPFTableRow tableRow = tableOne.createRow();
            tableRow.setHeight(380);
            tableRow.getCell(0).setText((String) map.get("CoopCompanyName"));
            tableRow.getCell(1).setText(String.valueOf(map.get("CoopPrice")));
        }

        tableCellStyle(tableOne, 2000);
    }

    // 添加单元格属性
    private void tableCellStyle(XWPFTable table, int cellWidth) {
        List<XWPFTableRow> tableRows = table.getRows();
        for (int i = 0; i < tableRows.size(); i++) {
            XWPFTableRow xwpfTableRow = tableRows.get(i);
            List<XWPFTableCell> cellList = xwpfTableRow.getTableCells();
            for (XWPFTableCell tableCell : cellList) {
                CTTc ctTc = tableCell.getCTTc();
                CTTcPr ctTcPr = ctTc.addNewTcPr();
                ctTcPr.addNewTcW().setW(BigInteger.valueOf(cellWidth));
                ctTcPr.addNewVAlign().setVal(STVerticalJc.CENTER);
                ctTcPr.addNewGridSpan().setVal(BigInteger.valueOf(1));

                // 添加单元格里的run属性
                CTP ctp = ctTc.getPList().get(0);
                CTPPr ctpPr = ctp.addNewPPr();
                ctpPr.addNewWidowControl();
                ctpPr.addNewJc().setVal(STJc.CENTER);
                CTParaRPr ctParaRPr = ctpPr.addNewRPr();
                CTFonts ctFontsp = ctParaRPr.addNewRFonts();
                ctFontsp.setAscii("宋体");
                ctFontsp.setHAnsi("宋体");
                ctFontsp.setCs("宋体");
                ctParaRPr.addNewSzCs().setVal(BigInteger.valueOf(21));


                CTR ctr = ctp.getRList().get(0);
                //CTR ctr = ctp.addNewR();
                CTRPr ctrPr = ctr.addNewRPr();
                CTFonts ctFonts = ctrPr.addNewRFonts();
                ctFonts.setHint(STHint.Enum.forString("eastAsia"));
                ctFonts.setEastAsia("宋体");
                ctFonts.setAscii("宋体");
                ctFonts.setHAnsi("宋体");
                ctFonts.setCs("宋体");
                ctrPr.addNewSzCs().setVal(BigInteger.valueOf(21));

                XWPFParagraph paragraph = tableCell.getParagraphs().get(0);
                String s = tableCell.getText();
                System.out.println("s= " + s);
            }
        }
    }

    // 设置表格边框
    private void tableBorderStyle(XWPFTable table){
        //表格属性
        CTTblPr tablePr = table.getCTTbl().addNewTblPr();
        //表格宽度
        //CTTblWidth width = tablePr.addNewTblW();
        CTJc ctJc = tablePr.addNewJc();
        ctJc.setVal(STJc.Enum.forString("center"));
        //width.setW(BigInteger.valueOf(7000));
        //width.setType(STTblWidth.DXA);

        //表格颜色
        CTTblBorders borders=table.getCTTbl().getTblPr().addNewTblBorders();
        //表格内部横向表格颜色
        CTBorder hBorder=borders.addNewInsideH();
        hBorder.setVal(STBorder.Enum.forString("single"));
        hBorder.setSz(new BigInteger("1"));
        hBorder.setColor("000000");
        //表格内部纵向表格颜色
        CTBorder vBorder=borders.addNewInsideV();
        vBorder.setVal(STBorder.Enum.forString("single"));
        vBorder.setSz(new BigInteger("1"));
        vBorder.setColor("000000");
        //表格最左边一条线的样式
        CTBorder lBorder=borders.addNewLeft();
        lBorder.setVal(STBorder.Enum.forString("single"));
        lBorder.setSz(new BigInteger("1"));
        lBorder.setColor("000000");
        //表格最左边一条线的样式
        CTBorder rBorder=borders.addNewRight();
        rBorder.setVal(STBorder.Enum.forString("single"));
        rBorder.setSz(new BigInteger("1"));
        rBorder.setColor("000000");
        //表格最上边一条线（顶部）的样式
        CTBorder tBorder=borders.addNewTop();
        tBorder.setVal(STBorder.Enum.forString("single"));
        tBorder.setSz(new BigInteger("1"));
        tBorder.setColor("000000");
        //表格最下边一条线（底部）的样式
        CTBorder bBorder=borders.addNewBottom();
        bBorder.setVal(STBorder.Enum.forString("single"));
        bBorder.setSz(new BigInteger("1"));
        bBorder.setColor("000000");
    }

    public CustomXWPFDocument generateWordMultiRows(Map<String, Object> param, String template) {
        CustomXWPFDocument doc = null;
        try {
            OPCPackage pack = POIXMLDocument.openPackage(template);
            doc = new CustomXWPFDocument(pack);
            int newRowCount = 0;
            List<NeedDealRow> needDealRows = new ArrayList<NeedDealRow>();
            if (param != null && param.size() > 0) {
                //处理段落
                List<XWPFParagraph> paragraphList = doc.getParagraphs();
                processParagraphsMultiRows(paragraphList, param, doc);

                //处理表格
                Iterator<XWPFTable> it = doc.getTablesIterator();
                while (it.hasNext()) {
                    XWPFTable table = it.next();
                    List<XWPFTableRow> rows = table.getRows();
                    int rowNum = 0;
                    for (XWPFTableRow row : rows) {
                        rowNum++;
                        String rowText = getRowText(row).trim();
                        if (rowText.indexOf("${") == -1)
                            continue;
                        if (rowText.indexOf("${iterationrow:") != -1) {
                            String key = rowText.split("iterationrow:")[1];
                            key = key.substring(0, key.indexOf("}"));
                            if (param.containsKey(key)) {
                                List<Map<String, Object>> paramSubs = (List<Map<String, Object>>) param.get(key);
                                if (paramSubs.size() < 1) continue;
                                NeedDealRow needDealRow = new NeedDealRow();
                                needDealRow.setParamSubs(paramSubs);
                                needDealRow.setDoc(doc);
                                needDealRow.setTable(table);
                                needDealRow.setRow(row);
                                needDealRow.setRowNum(rowNum + newRowCount);
                                needDealRow.setKey(key);
                                needDealRows.add(needDealRow);
                                newRowCount += paramSubs.size();
                            }
                            continue;
                        }
                        processRow(row, param, doc);
                    }
                }
                for (NeedDealRow needDealRow : needDealRows) {
                    List<Map<String, Object>> paramSubs = needDealRow.getParamSubs();
                    //先移除第一行；
                    Map<String, Object> paramSubT = paramSubs.remove(0);
                    //循环添加其他行，并替换；
                    for (Map<String, Object> paramSub : paramSubs) {
                        paramSub.put("iterationrow:" + needDealRow.getKey(), "");
                        XWPFTableRow rowT = needDealRow.getTable().insertNewTableRow(needDealRow.getRowNum());
                        copyPro(needDealRow.getRow(), rowT);
                        processRow(rowT, paramSub, doc);
                    }
                    //再处理第一行；
            paramSubT.put("iterationrow:" + needDealRow.getKey(), "");
            processRow(needDealRow.getRow(), paramSubT, doc);
        }
    }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return doc;
    }

    public void processRow(XWPFTableRow row,Map<String, Object> param,CustomXWPFDocument doc) {
        List<XWPFTableCell> cells = row.getTableCells();
        for (XWPFTableCell cell : cells) {
            List<XWPFParagraph> paragraphListTable = cell.getParagraphs();
            processParagraphsMultiRows(paragraphListTable, param, doc);
        }
    }

    private String getRowText(XWPFTableRow row) {
        String result = "";
        for (XWPFTableCell cell : row.getTableCells()) {
            result +=cell.getText();
        }
        return result;
    }

    public void processParagraphsMultiRows(List<XWPFParagraph> paragraphList,Map<String, Object> param,CustomXWPFDocument doc){
        if(paragraphList != null && paragraphList.size() > 0){
            for(XWPFParagraph paragraph : paragraphList){
                if(paragraph.getParagraphText().indexOf("${") == -1)
                    continue;
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    String text = run.getText(0);
                    if(text != null){
                        text=text.toLowerCase();
                        boolean isSetText = false;
                        for (Entry<String, Object> entry : param.entrySet()) {
                            String key = "${"+ entry.getKey().toLowerCase() +"}";
                            if(text.indexOf(key) != -1){
                                isSetText = true;
                                Object value = entry.getValue();
                                if (value instanceof String) {//文本替换
                                    text = text.replace(key, value.toString());
                                    //System.out.println(text);
                                } else if (value instanceof Double) {
                                    text = text.replace(key, value.toString());
                                } else if (value instanceof Map) {//图片替换
                                    text = text.replace(key, "");
                                    Map pic = (Map)value;
                                    int width = Integer.parseInt(pic.get("width").toString());
                                    int height = Integer.parseInt(pic.get("height").toString());
                                    int picType = getPictureType(pic.get("type").toString());
                                    byte[] byteArray = (byte[]) pic.get("content");
                                    ByteArrayInputStream byteInputStream = new ByteArrayInputStream(byteArray);
                                    try {
                                        String ind = doc.addPictureData(byteInputStream, picType);
                                        int id = doc.getNextPicNameNumber(picType);
                                        createPicture(id, ind, width , height, run);
                                        //System.out.println("图片替换");
                                    } catch (Exception e) {
                                        e.printStackTrace();
                                    }
                                }
                            }
                        }
                        if(isSetText){
                            run.setText(text,0);
                        }
                    }
                }
            }
        }
    }

    private void copyPro(XWPFTableRow sourceRow,XWPFTableRow targetRow) {
        //复制行属性
        targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
        List<XWPFTableCell> cellList = sourceRow.getTableCells();
        if (null == cellList) {
            return;
        }
        //添加列、复制列以及列中段落属性
        XWPFTableCell targetCell = null;
        for (XWPFTableCell sourceCell : cellList) {
            targetCell = targetRow.addNewTableCell();
            //列属性
            targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
            //段落属性
            while (targetCell.getParagraphs().size()<sourceCell.getParagraphs().size())
            {
                targetCell.addParagraph();
            }
            //拷贝段落和run
            int index = 0;
            for(XWPFParagraph sourceParagraph : sourceCell.getParagraphs()) {
                targetCell.getParagraphs().get(index).getCTP().setPPr(sourceParagraph.getCTP().getPPr());
                for (XWPFRun run : sourceParagraph.getRuns()) {
                    //System.out.println(run.getText(0));
                    XWPFRun runT = targetCell.getParagraphs().get(index).createRun();
                    runT.getCTR().setRPr(run.getCTR().getRPr());
                    runT.setText(run.getText(0));
                }
                index++;
            }
        }
    }

    /** 
     * 处理段落 
     * @param paragraphList 
     */  
    public void processParagraphs(List<XWPFParagraph> paragraphList,Map<String, Object> param,CustomXWPFDocument doc){
        if(paragraphList != null && paragraphList.size() > 0){  
            for(XWPFParagraph paragraph : paragraphList){
                List<XWPFRun> runs = paragraph.getRuns();  
                for (XWPFRun run : runs) {  
                    String text = run.getText(0);  
                    if(text != null){  
                        boolean isSetText = false;  
                        for (Entry<String, Object> entry : param.entrySet()) {  
                            String key = "${"+ entry.getKey() +"}";
                            if(text.indexOf(key) != -1){
                                isSetText = true;  
                                Object value = entry.getValue();  
                                if (value instanceof String) {//文本替换  
                                    text = text.replace(key, value.toString());  
                                    //System.out.println(text);
                                } else if (value instanceof Double) {
                                    text = text.replace(key, value.toString());
                                } else if (value instanceof Map) {//图片替换
                                    text = text.replace(key, "");  
                                    Map pic = (Map)value;  
                                    int width = Integer.parseInt(pic.get("width").toString());  
                                    int height = Integer.parseInt(pic.get("height").toString());  
                                    int picType = getPictureType(pic.get("type").toString());  
                                    byte[] byteArray = (byte[]) pic.get("content");  
                                    ByteArrayInputStream byteInputStream = new ByteArrayInputStream(byteArray);  
                                    try {  
                                        String ind = doc.addPictureData(byteInputStream, picType);
                                        int id = doc.getNextPicNameNumber(picType);
                                        createPicture(id, ind, width , height, run);
                                        //System.out.println("图片替换");
                                    } catch (Exception e) {  
                                        e.printStackTrace();  
                                    }  
                                }  
                            }  
                        }  
                        if(isSetText){  
                            run.setText(text,0);  
                        }  
                    }  
                }  
            }  
        }  
    }
    /** 
     * 根据图片类型，取得对应的图片类型代码 
     * @param picType 
     * @return int 
     */  
    private int getPictureType(String picType){
        int res = CustomXWPFDocument.PICTURE_TYPE_PICT;  
        if(picType != null){  
            if(picType.equalsIgnoreCase("png")){  
                res = CustomXWPFDocument.PICTURE_TYPE_PNG;
            }else if(picType.equalsIgnoreCase("dib")){  
                res = CustomXWPFDocument.PICTURE_TYPE_DIB;  
            }else if(picType.equalsIgnoreCase("emf")){  
                res = CustomXWPFDocument.PICTURE_TYPE_EMF;  
            }else if(picType.equalsIgnoreCase("jpg") || picType.equalsIgnoreCase("jpeg")){  
                res = CustomXWPFDocument.PICTURE_TYPE_JPEG;  
            }else if(picType.equalsIgnoreCase("wmf")){  
                res = CustomXWPFDocument.PICTURE_TYPE_WMF;  
            }  
        }  
        return res;
    }  
    /** 
     * 将输入流中的数据写入字节数组 
     * @param in 
     * @return 
     */  
    public byte[] inputStream2ByteArray(InputStream in,boolean isClose){
        byte[] byteArray = null;  
        try {  
            int total = in.available();  
            byteArray = new byte[total];  
            in.read(byteArray);  
        } catch (IOException e) {  
            e.printStackTrace();  
        }finally{  
            if(isClose){  
                try {  
                    in.close();  
                } catch (Exception e2) {  
                    System.out.println("关闭流失败");  
                }  
            }  
        }  
        return byteArray;  
    }  
    
    /** 
     * @param id 
     * @param width 宽 
     * @param height 高 
     * @param run  段落
     */  
    public void createPicture(int id, String blipId, int width, int height,XWPFRun run) {
        final int EMU = 9525;    
        width *= EMU;    
        height *= EMU;    
          
        CTInline inline = run.getCTR().addNewDrawing().addNewInline();    
        String picXml = ""    
                + "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"    
                + "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"    
                + "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"    
                + "         <pic:nvPicPr>" + "            <pic:cNvPr id=\""    
                + id    
                + "\" name=\"Generated\"/>"    
                + "            <pic:cNvPicPr/>"    
                + "         </pic:nvPicPr>"    
                + "         <pic:blipFill>"    
                + "            <a:blip r:embed=\""    
                + blipId    
                + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>"    
                + "            <a:stretch>"    
                + "               <a:fillRect/>"    
                + "            </a:stretch>"    
                + "         </pic:blipFill>"    
                + "         <pic:spPr>"    
                + "            <a:xfrm>"    
                + "               <a:off x=\"0\" y=\"0\"/>"    
                + "               <a:ext cx=\""    
                + width    
                + "\" cy=\""    
                + height    
                + "\"/>"    
                + "            </a:xfrm>"    
                + "            <a:prstGeom prst=\"rect\">"    
                + "               <a:avLst/>"    
                + "            </a:prstGeom>"    
                + "         </pic:spPr>"    
                + "      </pic:pic>"    
                + "   </a:graphicData>" + "</a:graphic>";    
    
        inline.addNewGraphic().addNewGraphicData();    
        XmlToken xmlToken = null;    
        try {    
            xmlToken = XmlToken.Factory.parse(picXml);    
        } catch (XmlException xe) {    
            xe.printStackTrace();    
        }    
        inline.set(xmlToken);   
          
        inline.setDistT(0);      
        inline.setDistB(0);      
        inline.setDistL(0);      
        inline.setDistR(0);      
          
        CTPositiveSize2D extent = inline.addNewExtent();    
        extent.setCx(width);    
        extent.setCy(height);    
          
        CTNonVisualDrawingProps docPr = inline.addNewDocPr();      
        docPr.setId(id);      
        docPr.setName("图片" + id);      
        docPr.setDescr("测试");   
    }

    public static void main (String[] args) throws Exception {
        InputStream in1 = null;
        InputStream in2 = null;
        OPCPackage src1Package = null;
        OPCPackage src2Package = null;

        OutputStream dest = new FileOutputStream("E:\\dest.docx");
        try {
            in1 = new FileInputStream("E:\\tmp.docx");
            in2 = new FileInputStream("E:\\tmp2.docx");
            src1Package = OPCPackage.open(in1);
            src2Package = OPCPackage.open(in2);
        } catch (Exception e) {
            e.printStackTrace();
        }

        XWPFDocument src1Document = new XWPFDocument(src1Package);
        CTBody src1Body = src1Document.getDocument().getBody();
        XWPFDocument src2Document = new XWPFDocument(src2Package);
        CTBody src2Body = src2Document.getDocument().getBody();
        appendBody(src1Body, src2Body);
        src1Document.write(dest);

    }

    private static void appendBody(CTBody src, CTBody append) throws Exception {
        XmlOptions optionsOuter = new XmlOptions();
        optionsOuter.setSaveOuter();
        String appendString = append.xmlText(optionsOuter);
        String srcString = src.xmlText();
        String prefix = srcString.substring(0,srcString.indexOf(">")+1);
        String mainPart = srcString.substring(srcString.indexOf(">")+1,srcString.lastIndexOf("<"));
        String sufix = srcString.substring( srcString.lastIndexOf("<") );
        String addPart = appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));
        CTBody makeBody = CTBody.Factory.parse(prefix+mainPart+addPart+sufix);
        src.set(makeBody);
    }

}

class NeedDealRow
{
    public String getKey() {
        return key;
    }

    public void setKey(String key) {
        this.key = key;
    }

    String key;
    XWPFTableRow row;
    XWPFTable table;

    public XWPFTableRow getRow() {
        return row;
    }

    public void setRow(XWPFTableRow row) {
        this.row = row;
    }

    public XWPFTable getTable() {
        return table;
    }

    public void setTable(XWPFTable table) {
        this.table = table;
    }

    public CustomXWPFDocument getDoc() {
        return doc;
    }

    public void setDoc(CustomXWPFDocument doc) {
        this.doc = doc;
    }

    public int getRowNum() {
        return rowNum;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public List<Map<String, Object>> getParamSubs() {
        return paramSubs;
    }

    public void setParamSubs(List<Map<String, Object>> paramSubs) {
        this.paramSubs = paramSubs;
    }

    CustomXWPFDocument doc = null;
    int rowNum = 0;
    List<Map<String, Object>> paramSubs;
}