package com.gzcdc.officeserver.model;

/**
 * Created by User: admin.
 * Date: 2017/10/27 Time: 16:45.
 * Description:
 */
public class FileEntity {
    private String name;
    private String url;
    private String fileSize;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getUrl() {
        return url;
    }

    public void setUrl(String url) {
        this.url = url;
    }

    public String getFileSize() {
        return fileSize;
    }

    public void setFileSize(String fileSize) {
        this.fileSize = fileSize;
    }
}
