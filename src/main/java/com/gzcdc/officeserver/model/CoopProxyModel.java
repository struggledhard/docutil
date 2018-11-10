package com.gzcdc.officeserver.model;

/**
 * Created by lanshg on 2017/10/20.
 */
public class CoopProxyModel {

    /// <summary>
    /// 项目名称
    /// </summary>
    private String ProjectName;
    /// <summary>
    /// 项目设计编号
    /// </summary>
    private String ProjectId;
    /// <summary>
    /// 生成部门名称
    /// </summary>
    private String ProductDeartmentName;
    /// <summary>
    /// 要求完成时间
    /// </summary>
    private String CustReqFinishDate;
    /// <summary>
    /// 协作流程编号
    /// </summary>
    private String CoopFlowId;
    /// <summary>
    /// 协作申请或修改编号
    /// 编号生成规则：部门代号-年份-三位流水号，其中部门代号分别如下：
    ///本公司代号：Gzcdc；通信设计一所：TXY；综合部：ZHB；通信设计二所：TXE；市场部：SCB；通信设计三所：TXS；财务部：CWB；信息技术研发部：XJB；生产技术部：SJB；规划咨询研究所：GHS；监理部：JLB；数码建筑设计所：JZS；情报资料室：ZLS；监理项目处/办事处：JLCn；人力资源部：RZB。
    ///例如：GHS-2017-006，代表规划咨询研究所2017年的第六个协作申请(注：协作修改编号在最后加上/XG，如GHS-2017-006/XG)
    /// </summary>
    private String CoopId;
    /// <summary>
    /// 协作内容
    /// </summary>
    private String CoopContent;
    /// <summary>
    /// 流程类型：1协作申请流程；2协作修改流程。
    /// </summary>
    private int FlowType;
    /// <summary>
    /// 协作单位编号
    /// </summary>
    private String CoopCompanyId;
    /// <summary>
    /// 协作单位
    /// </summary>
    private String CoopComanyName;
    /// <summary>
    /// 协作工作量记录日期
    /// </summary>
    private String CreateDate;
    /// <summary>
    /// 协作委托书编号
    /// </summary>
    private String EntrustNum;
    /// <summary>
    /// 委托时间。分管领导审核通过时间
    /// </summary>
    private String EntrustDate;


    public CoopProxyModel() {}

    public String getProjectName() {
        return ProjectName;
    }

    public void setProjectName(String projectName) {
        ProjectName = projectName;
    }

    public String getProjectId() {
        return ProjectId;
    }

    public void setProjectId(String projectId) {
        ProjectId = projectId;
    }

    public String getProductDeartmentName() {
        return ProductDeartmentName;
    }

    public void setProductDeartmentName(String productDeartmentName) {
        ProductDeartmentName = productDeartmentName;
    }

    public String getCustReqFinishDate() {
        return CustReqFinishDate;
    }

    public void setCustReqFinishDate(String custReqFinishDate) {
        CustReqFinishDate = custReqFinishDate;
    }

    public String getCoopFlowId() {
        return CoopFlowId;
    }

    public void setCoopFlowId(String coopFlowId) {
        CoopFlowId = coopFlowId;
    }

    public String getCoopId() {
        return CoopId;
    }

    public void setCoopId(String coopId) {
        CoopId = coopId;
    }

    public String getCoopContent() {
        return CoopContent;
    }

    public void setCoopContent(String coopContent) {
        CoopContent = coopContent;
    }

    public int getFlowType() {
        return FlowType;
    }

    public void setFlowType(int flowType) {
        FlowType = flowType;
    }

    public String getCoopCompanyId() {
        return CoopCompanyId;
    }

    public void setCoopCompanyId(String coopCompanyId) {
        CoopCompanyId = coopCompanyId;
    }

    public String getCoopComanyName() {
        return CoopComanyName;
    }

    public void setCoopComanyName(String coopComanyName) {
        CoopComanyName = coopComanyName;
    }

    public String getCreateDate() {
        return CreateDate;
    }

    public void setCreateDate(String createDate) {
        CreateDate = createDate;
    }

    public String getEntrustNum() {
        return EntrustNum;
    }

    public void setEntrustNum(String entrustNum) {
        EntrustNum = entrustNum;
    }

    public String getEntrustDate() {
        return EntrustDate;
    }

    public void setEntrustDate(String entrustDate) {
        EntrustDate = entrustDate;
    }
}