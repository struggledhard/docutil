package com.gzcdc.officeserver.controller;

/**
 * Created by lanshg on 2017/10/20.
 */

import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;
import com.gzcdc.officeserver.model.DocServiceBackEntrustModel;
import com.gzcdc.officeserver.model.ReturnJsonBase;
import com.gzcdc.officeserver.util.CustomXWPFDocument;
import com.gzcdc.officeserver.util.OfficeFileUtils;
import com.gzcdc.officeserver.util.WordUtil;
import com.lowagie.text.Font;
import com.lowagie.text.pdf.BaseFont;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import fr.opensagres.xdocreport.itext.extension.font.IFontProvider;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;

/**
 * 描述: TODO:
 * 包名: com.gzcdc.officeserver
 * 作者: 蓝善根.
 * 日期: 17-10-20.
 * 项目名称: officeserver
 * 版本: 1.0
 * JDK: since 1.8
 */

@Api(value = "文档生成服务", description = "根据模板生成word/excel/pdf等文档的相关API")
@RestController
@RequestMapping(value = "v1/word")
public class WordDealController {

    @ApiOperation("生成协作委托书")
    @RequestMapping(value = "/gencooperpdf")
    public ReturnJsonBase genCooperPdf(@RequestParam String data, HttpServletRequest request) {
        ReturnJsonBase result = new ReturnJsonBase();
        FileOutputStream fopts = null;
        try {
            List<DocServiceBackEntrustModel> models = new ArrayList<>();
            result.setData(models);

            Gson gson = new Gson();

            List<Map> coopProxyModels = gson.fromJson(data, new TypeToken<List<Map>>() {}.getType());

            WordUtil wordUtil = new WordUtil();
            for (Map<String, Object> param : coopProxyModels) {

                String uuid = UUID.randomUUID().toString();
                // 获取data数据里的委托书编号
                String fileName = null;
                Map<String, Object> map = param;
                Object value = map.get("EntrustNum");
                if (value != null) {
                    fileName = "协作项目【" + value.toString() + "】委托书" + getUploadCurrentTime();
                } else {
                    fileName = uuid;
                }

                String rootPath = request.getSession().getServletContext().getRealPath("/");
                String filePath = "/upload/projectmanager/" + fileName + ".pdf";
                String realPath = OfficeFileUtils.LINUXROOTPATH + filePath;
                System.out.println(realPath);
                //处理每个委托
                File genFile = new File(realPath);
                genFile.setWritable(true, false);    //设置写权限，windows下不用此语句
                if (!genFile.exists()) {
                    // 先得到文件的上级目录，并创建上级目录，在创建文件
                    // uploadFile.getParentFile().mkdir();
                    // 如果路径不存在,则创建
                    if (!genFile.getParentFile().exists()) {
                        genFile.getParentFile().mkdirs();
                    }
                    // 创建文件
                    genFile.createNewFile();
                }
                CustomXWPFDocument doc = wordUtil.generateWord(param, rootPath + "wordtemplate/xiezuoweit-v1.docx");
                fopts = new FileOutputStream(genFile);
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
                PdfConverter.getInstance().convert(doc, fopts, options);
                //doc.write(fopts);

                DocServiceBackEntrustModel model = new DocServiceBackEntrustModel();

                // 文件大小
                //String fileSize = FileSizeUtil.fileSize(filePath);
                //String fileSize = FileSizeUtil.FormatFileSize(filePath);

                //File file = new File(realPath);

                model.setCoopCompanyId(param.get("CoopCompanyId").toString());
                model.setFileSize(genFile.length());
                model.setFilePath(filePath);
                models.add(model);
            }
            result.setSuccess(true);
        } catch (IOException e) {
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

        return result;
    }

    @ApiOperation("协作项目后评估打分表")
    @RequestMapping(value = "/gencooperscorepdf")
    public ReturnJsonBase genCooperScorePdf(@RequestParam String data, HttpServletRequest request) {
        ReturnJsonBase result = new ReturnJsonBase();
        FileOutputStream fopts = null;
        try {
            List<DocServiceBackEntrustModel> models = new ArrayList<>();
            result.setData(models);

            Gson gson = new Gson();

            List<Map> coopProxyModels = gson.fromJson(data,
                    new TypeToken<List<Map>>() {}.getType());

            WordUtil wordUtil = new WordUtil();
            for (Map param : coopProxyModels) {
                String uuid = UUID.randomUUID().toString();

                // 获取data数据里的设计编号
                String fileName = null;
                Map<String, Object> map = param;
                String projectValue = map.get("ProjectId").toString();
                if (projectValue != null) {
                    fileName = "协作项目【" + projectValue + "】评分表" + getUploadCurrentTime();
                } else {
                    fileName = uuid;
                }

                String rootPath = request.getSession().getServletContext().getRealPath("/");
                String filePath = "/upload/projectmanager/" + fileName + ".pdf";
                String realPath = OfficeFileUtils.LINUXROOTPATH + filePath;
                //String realPath = rootPath + filePath;
                //处理每个委托
                File genFile = new File(realPath);
                genFile.setWritable(true, false);    //设置写权限，windows下不用此语句
                if (!genFile.exists()) {
                    // 先得到文件的上级目录，并创建上级目录，在创建文件
                    // uploadFile.getParentFile().mkdir();
                    // 如果路径不存在,则创建
                    if (!genFile.getParentFile().exists()) {
                        genFile.getParentFile().mkdirs();
                    }
                    // 创建文件
                    genFile.createNewFile();
                }
                CustomXWPFDocument doc = wordUtil.generateWord(param, rootPath + "wordtemplate/xiezuopinggu.docx");
                fopts = new FileOutputStream(genFile);
                PdfOptions options = PdfOptions.create();
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
                PdfConverter.getInstance().convert(doc, fopts, options);
                //doc.write(fopts);

                DocServiceBackEntrustModel model = new DocServiceBackEntrustModel();

                // 文件大小
                //String fileSize = FileSizeUtil.fileSize(filePath);
                //String fileSize = FileSizeUtil.FormatFileSize(filePath);
                File file = new File(realPath);

                model.setCoopCompanyId(param.get("CoopCompanyId").toString());
                model.setFileSize(file.length());
                model.setFilePath(filePath);
                models.add(model);
            }
            result.setSuccess(true);
        } catch (IOException e) {
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

        return result;
    }

    @ApiOperation("设计修改通知单")
    @RequestMapping(value = "/gendesignmodifypdf")
    public ReturnJsonBase genDesignModifyPdf(@RequestParam String data, HttpServletRequest request) {
        ReturnJsonBase result = new ReturnJsonBase();
        FileOutputStream fopts = null;
        try {
                DocServiceBackEntrustModel model = new DocServiceBackEntrustModel();
                result.setData(model);
                Gson gson = new Gson();
                Map<String, Object> coopProxyModels = gson.fromJson(data,
                        new TypeToken<Map>() {}.getType());

                WordUtil wordUtil = new WordUtil();
                String uuid = UUID.randomUUID().toString();

                // 获取data数据里的设计编号
                String fileName = null;
                Map<String, Object> map = coopProxyModels;
                String projectValue = map.get("ProjectId").toString();
                if (projectValue != null) {
                    fileName = "项目【" + projectValue + "】设计修改通知单" + getUploadCurrentTime();
                } else {
                    fileName = uuid;
                }

                String rootPath = request.getSession().getServletContext().getRealPath("/");
                String filePath = "/upload/projectmanager/" + fileName + ".pdf";
                String realPath = OfficeFileUtils.LINUXROOTPATH + filePath;
                //处理每个委托
                File genFile = new File(realPath);
                genFile.setWritable(true, false);    //设置写权限，windows下不用此语句
                if (!genFile.exists()) {
                    // 先得到文件的上级目录，并创建上级目录，在创建文件
                    // uploadFile.getParentFile().mkdir();
                    // 如果路径不存在,则创建
                    if (!genFile.getParentFile().exists()) {
                        genFile.getParentFile().mkdirs();
                    }
                    // 创建文件
                    genFile.createNewFile();
                }
                CustomXWPFDocument doc = wordUtil.generateWord(coopProxyModels, rootPath + "wordtemplate/designmodify-v1.docx");
                fopts = new FileOutputStream(genFile);
                PdfOptions options = PdfOptions.create();
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
                PdfConverter.getInstance().convert(doc, fopts, options);
                //doc.write(fopts);

                // 文件大小
                //String fileSize = FileSizeUtil.fileSize(filePath);
                //String fileSize = FileSizeUtil.FormatFileSize(filePath);
                File file = new File(realPath);

                model.setFileSize(file.length());
                model.setFilePath(filePath);
            result.setSuccess(true);
        } catch (IOException e) {
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
        return result;
    }

    // 获取当前时间
    private String getUploadCurrentTime() {
        return new SimpleDateFormat("yyyyMMddHHmmssSSS").format(new Date());
    }
}