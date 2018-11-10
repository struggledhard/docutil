package com.gzcdc.officeserver.controller;

import com.gzcdc.officeserver.model.FileEntity;
import com.gzcdc.officeserver.util.OfficeFileUtils;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.apache.commons.io.FileUtils;
import org.springframework.util.FileCopyUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by User: skh.
 * Date: 2017/10/27 Time: 17:09.
 * Description:上传文件
 */

@Api(value = "文档上传服务", description = "上传文件")
@RestController
@RequestMapping(value = "/uploadfile")
public class UploadFileController {

    private MultipartFile imageFile;// 文件
    private String extendName; // 文件扩展名
    private final int MAX_SIZE = 1024 * 1024 * 100; // 限制用户头像的最大值为1M
    private String rootPath; // 文件根路径
    private String imageNewPath; // 头像新路径（包含头像名以及扩展名）
    private String imageNames; // 头像的新名字（时间+用户名），时间精确到毫秒
    private String newPath;
    public static final String FILE_SEPARATOR = System.getProperty("file.separator");

    @ApiOperation("上传文件")
    @RequestMapping(value = "/upload", method = RequestMethod.POST)
    public Map<String, Object> uploadFiles(@RequestParam MultipartFile[] files,
                                            HttpServletRequest request, HttpServletResponse response) {
        int i = 0;
        List<FileEntity> data = new ArrayList<>();
        Map<String, Object> map = new HashMap<>();
        // 获取附件所属的项目名称和流程名称，存放到对应的文件夹下
        String projectName = request.getParameter("projectName");
        String flowName = request.getParameter("flowName");

        try {
            // 将文件信息插入到数据库中
            for (MultipartFile multipartFile : files) {
                imageFile = multipartFile;
                // 获取上传头像的文件名
                String fileName = imageFile.getOriginalFilename();

                // 获取文件扩展名
                extendName = fileName.substring(fileName.lastIndexOf("."));
                // 获取上传头像的大小
                int fileSize = (int) imageFile.getSize();
                // 判断文件大小，不应大于100M
                if (fileSize >= MAX_SIZE) {
                    return getMap(false, "文件不能大于100M!", null, map);
                }
                String url = "/upload";
                if (projectName != null) {
                    if (!projectName.isEmpty()) {
                        url = url + "/" + projectName;
                        if (flowName != null) {
                            if (!flowName.isEmpty()) {
                                url = url + "/" + flowName;
                            }
                        }
                    } else {
                        url += "/other";
                    }
                } else {
                    url += "/other";
                }

                // 获取文件路径
                //rootPath = request.getSession().getServletContext().getRealPath(url);

                imageNames = getUploadCurrentTime() + fileName; // 重新命名上传名称:时间+文件名
                //imageNewPath = rootPath + "/" + imageNames;

                newPath = OfficeFileUtils.LINUXROOTPATH + url + "/" + imageNames;

                System.out.println("newPath = " + newPath);
                //System.out.println("path = " + imageNewPath);

                File uploadFile = new File(newPath);
                uploadFile.setWritable(true, false);    //设置写权限，windows下不用此语句
                if (!uploadFile.exists()) {
                    // 先得到文件的上级目录，并创建上级目录，在创建文件
                    // uploadFile.getParentFile().mkdir();
                    // 如果路径不存在,则创建
                    if (!uploadFile.getParentFile().exists()) {
                        uploadFile.getParentFile().mkdirs();
                    }
                    // 创建文件
                    uploadFile.createNewFile();
                }
                FileCopyUtils.copy(imageFile.getBytes(), uploadFile);
                FileUtils.copyInputStreamToFile(imageFile.getInputStream(), new File(newPath));

                // Linux环境下使用
                String subStr = newPath.substring(newPath.indexOf("/upload"));
                // windows环境下使用
                // String subStr = imageNewPath.substring(imageNewPath.indexOf("\\upload")).replace("\\", "/");
                FileEntity file = new FileEntity();
                file.setName(fileName);
                file.setUrl(subStr);
                file.setFileSize(fileSize + "");

                System.out.println("subStr = " + subStr);

                data.add(file);
                ++i;
            }
        } catch (Exception e) {
            e.printStackTrace();
            map = getMap(false, "文件上传异常!", null, map);
            return map;
        }
        // 保存情况
        if (i == files.length) {
            map = getMap(true, "文件保存完成", data, map);
        } else {
            map = getMap(false, "保存成功" + i + "个，出错" + (files.length - i) + "个", data, map);
        }

        return map;
    }

    /**
     * 返回map信息
     *
     * @param result
     * @param message
     * @param data
     * @param map
     * @return
     */
    public Map<String, Object> getMap(Boolean result, String message, Object data, Map<String, Object> map) {
        map.put("result", result);
        map.put("messageInfo", message);
        map.put("data", data);
        return map;
    }

    // 获取头上上传的当前时间
    private String getUploadCurrentTime() {
        return new SimpleDateFormat("yyyyMMddHHmmssSSS").format(new Date());
    }
}
