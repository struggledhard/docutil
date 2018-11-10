package com.gzcdc.officeserver.model;

/**
 * Created by lanshg on 2017/10/20.
 */
public class ReturnJsonBase {

    /// <summary>
    /// 接口调用是否成功（默认为true）
    /// </summary>
    private boolean IsSuccess;

    /// <summary>
    /// 接口调用失败时返回的错误提示消息
    /// </summary>
    private String ResultMsg;

    /// <summary>
    /// 接口调用时返回的接口调用数据
    /// </summary>
    private Object Data;

    public ReturnJsonBase()
    {
        this.IsSuccess = true;
    }

    public boolean isSuccess() {
        return IsSuccess;
    }

    public void setSuccess(boolean success) {
        IsSuccess = success;
    }

    public String getResultMsg() {
        return ResultMsg;
    }

    public void setResultMsg(String resultMsg) {
        ResultMsg = resultMsg;
    }

    public Object getData() {
        return Data;
    }

    public void setData(Object data) {
        Data = data;
    }
}
