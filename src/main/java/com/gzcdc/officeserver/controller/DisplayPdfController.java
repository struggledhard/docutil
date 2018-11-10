package com.gzcdc.officeserver.controller;

import com.gzcdc.officeserver.util.CustomXWPFDocument;
import com.gzcdc.officeserver.util.JsonConvertUtil;
import com.gzcdc.officeserver.util.WordUtil;
import com.lowagie.text.Font;
import com.lowagie.text.pdf.BaseFont;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import fr.opensagres.xdocreport.itext.extension.font.IFontProvider;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.awt.*;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

/**
 * Created by User: skh.
 * Date: 2018/1/26 Time: 10:37.
 * Description：表单docx的pdf生成及在线预览
 */
@Api(value = "文档生成服务", description = "根据模板生成word/excel/pdf等文档的相关API")
@RestController
@RequestMapping(value = "v1/word")
public class DisplayPdfController {
    @ApiOperation("PDF生成")
    @RequestMapping(value = "/genpdf")
    public void genPdf(@RequestBody String data, @RequestParam String id, HttpServletRequest request, HttpServletResponse response) {
        if (id.equals("GzcdcJL104")) {
            // 处理协作申请表
            projectIsCoopApplyTable(data, id, request, response);
        } else {
            projectNotCoopApplyTable(data, id, request, response);
        }
    }

    private void projectIsCoopApplyTable(String data, String id, HttpServletRequest request, HttpServletResponse response) {
        OutputStream fopts = null;
        WordUtil wordUtil = new WordUtil();
        try {
            String rootPath = request.getSession().getServletContext().getRealPath("/");
            // 模板路径
            String tempPath = rootPath + "tabletemplate/" + id + ".docx";
            //String realPath = OfficeFileUtils.LINUXROOTPATH + filePath;

            PdfOptions options =getPdfOptions(rootPath);

            data = data.replace("\n","\\n").replace("\r", "\\r");
            //Map<String, Object> map = JsonConvertUtil.jsonToMap(data);
            List<Map<String, Object>> maps = JsonConvertUtil.jsonConvert(data);
            CustomXWPFDocument doc =null;
            for (Map<String, Object> map : maps) {
                doc = wordUtil.generateCoopApplyWord(map, tempPath);
            }

            response.setHeader("Content-Disposition", "attachment;fileName=" + id + ".docx");
            response.setContentType("multipart/form-data");
            fopts = response.getOutputStream();
            //doc.write(new FileOutputStream("E://andy.docx"));
            PdfConverter.getInstance().convert(doc, fopts, options);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (null != fopts) {
                try {
                    fopts.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

    }

    private void projectNotCoopApplyTable(String data, String id, HttpServletRequest request, HttpServletResponse response) {
        OutputStream fopts = null;
        WordUtil wordUtil = new WordUtil();
        try {
            //Gson gson = new Gson();

            //JsonObject jObject = new JsonParser().parse(data).getAsJsonObject();
            //JsonArray array = jObject.get("groupList").getAsJsonArray();
            String rootPath = request.getSession().getServletContext().getRealPath("/");
            // 模板路径
            String tempPath = rootPath + "tabletemplate/" + id + ".docx";
            //String realPath = OfficeFileUtils.LINUXROOTPATH + filePath;

            PdfOptions options =getPdfOptions(rootPath);

            data = data.replace("\n","\\n").replace("\r", "\\r");

            List<Map<String, Object>> maps = JsonConvertUtil.jsonConvert(data);
            CustomXWPFDocument doc =null;
            int index=0;
            for (Map dataMap:maps) {
                CustomXWPFDocument docT = wordUtil.generateWordMultiRows(dataMap, tempPath);
                //doc.write(new FileOutputStream("E://temp.docx"));
                //fopts = new FileOutputStream(genFile);
                if (doc == null)
                    doc = docT;
                else
                {
//                    XmlOptions optionsOuter = new XmlOptions();
//                    optionsOuter.setSaveOuter();
//                    String appendString = docT.getDocument().getBody().xmlText(optionsOuter);
//                    String srcString = doc.getDocument().getBody().xmlText();
//                    String prefix = srcString.substring(0,srcString.indexOf(">")+1);
//                    String mainPart = srcString.substring(srcString.indexOf(">")+1,srcString.lastIndexOf("<"));
//                    String sufix = srcString.substring( srcString.lastIndexOf("<") );
//                    String addPart = appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));
//                    CTBody makeBody = CTBody.Factory.parse(prefix+mainPart+addPart+sufix);
//                    doc.getDocument().unsetBody();
//                    doc.getDocument().setBody(makeBody);
                    //appendBody(doc.getDocument()..getBody(), docT.getDocument().getBody());
                }

                //doc.write(new FileOutputStream("E://temp.docx"));
            }

            //outputPdf(realPath, request, response);
            response.setHeader("Content-Disposition", "attachment;fileName=" + id + ".docx");
            response.setContentType("multipart/form-data");
            fopts = response.getOutputStream();
            //doc.write(new FileOutputStream("E://andy.docx"));
            PdfConverter.getInstance().convert(doc, fopts, options);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (null != fopts) {
                try {
                    fopts.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private PdfOptions getPdfOptions(String rootPath)
    {
        PdfOptions options = PdfOptions.create();
        //中文字体处理
        /**
         * 将word文档， 转换成pdf
         * 宋体：STSong-Light
         *
         * @param fontParam1 可以字体的路径，也可以是itextasian-1.5.2.jar提供的字体，比如宋体"STSong-Light"
         * @param fontParam2 和fontParam2对应，fontParam1为路径时，fontParam2=BaseFont.IDENTITY_H，为itextasian-1.5.2.jar提供的字体时，fontParam2="UniGB-UCS2-H"
         * @throws Exception
         */
        options.fontProvider(new IFontProvider() {
            @Override
            public Font getFont(String s, String s1, float v, int i, Color color) {
                try {
                    BaseFont bfChinese = BaseFont.createFont(rootPath + "wordtemplate/simsun.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                    Font fontChinese = new Font(bfChinese, v, i, color);
                    if (s != null) {
                        fontChinese.setFamily(s);
                    }
                    return fontChinese;
                } catch (Exception e) {
                    e.printStackTrace();
                    return null;
                }
            }
        });
        return options;
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
