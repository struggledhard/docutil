package com.gzcdc.officeserver.model;

/**
 * Created by lanshg on 2017/10/20.
 */
public class DocServiceBackEntrustModel {
    /// <summary>
    /// 协作单位编号
    /// </summary>
    private String CoopCompanyId;
    /// <summary>
    /// 返回服务器文件地址
    /// </summary>
    private String FilePath;
    /// <summary>
    /// 文件大小
    /// </summary>
    private long FileSize;

    public String getCoopCompanyId() {
        return CoopCompanyId;
    }

    public void setCoopCompanyId(String coopCompanyId) {
        CoopCompanyId = coopCompanyId;
    }

    public String getFilePath() {
        return FilePath;
    }

    public void setFilePath(String filePath) {
        FilePath = filePath;
    }

    public long getFileSize() {
        return FileSize;
    }

    public void setFileSize(long fileSize) {
        FileSize = fileSize;
    }
}
