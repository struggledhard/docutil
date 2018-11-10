package com.gzcdc.officeserver.model;

/**
 * Created with User: skh.
 * Date: 2017/10/27 Time: 14:46.
 * Description:
 */
public class TbFile {
    private String id;

    private String projectname;

    private String flowname;

    private String url;

    private String thumbUrl;

    private Integer size;

    private String extendname;

    private Integer downloadsamount;

    private String createtime;

    private Integer isdelete;

    private String extendtag1;

    private String extendtag2;

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id == null ? null : id.trim();
    }

    public String getProjectname() {
        return projectname;
    }

    public void setProjectname(String projectname) {
        this.projectname = projectname == null ? null : projectname.trim();
    }

    public String getFlowname() {
        return flowname;
    }

    public void setFlowname(String flowname) {
        this.flowname = flowname == null ? null : flowname.trim();
    }

    public String getUrl() {
        return url;
    }

    public void setUrl(String url) {
        this.url = url == null ? null : url.trim();
    }

    public String getThumbUrl() {
        return thumbUrl;
    }

    public void setThumbUrl(String thumbUrl) {
        this.thumbUrl = thumbUrl == null ? null : thumbUrl.trim();
    }

    public Integer getSize() {
        return size;
    }

    public void setSize(Integer size) {
        this.size = size;
    }

    public String getExtendname() {
        return extendname;
    }

    public void setExtendname(String extendname) {
        this.extendname = extendname == null ? null : extendname.trim();
    }

    public Integer getDownloadsamount() {
        return downloadsamount;
    }

    public void setDownloadsamount(Integer downloadsamount) {
        this.downloadsamount = downloadsamount;
    }

    public String getCreatetime() {
        return createtime;
    }

    public void setCreatetime(String createtime) {
        this.createtime = createtime == null ? null : createtime.trim();
    }

    public Integer getIsdelete() {
        return isdelete;
    }

    public void setIsdelete(Integer isdelete) {
        this.isdelete = isdelete;
    }

    public String getExtendtag1() {
        return extendtag1;
    }

    public void setExtendtag1(String extendtag1) {
        this.extendtag1 = extendtag1 == null ? null : extendtag1.trim();
    }

    public String getExtendtag2() {
        return extendtag2;
    }

    public void setExtendtag2(String extendtag2) {
        this.extendtag2 = extendtag2 == null ? null : extendtag2.trim();
    }
}
