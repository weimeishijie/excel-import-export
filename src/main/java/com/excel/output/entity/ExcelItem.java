package com.excel.output.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;

import java.util.Date;

/**
 * Created by li wen ya on 2020/5/9
 */
public class ExcelItem {

    /**
     * 主键id
     */
    @Excel(name = "用户id")
    private Long userId;

    /**
     * 头像
     */
    @Excel(name = "头像")
    private String avatar;

    /**
     * 账号
     */
    @Excel(name = "账号")
    private String account;

    /**
     * 名字
     */
    @Excel(name = "姓名")
    private String name;

    /**
     * 生日
     */
    @Excel(name = "生日")
    private Date birthday;

    /**
     * 性别(字典)
     */
    @Excel(name = "性别")
    private String sex;

    /**
     * 电子邮件
     */
    @Excel(name = "邮箱")
    private String email;

    /**
     * 电话
     */
    @Excel(name = "电话")
    private String phone;

    /**
     * 角色id(多个逗号隔开)
     */
    @Excel(name = "角色id")
    private String roleId;

    /**
     * 部门id(多个逗号隔开)
     */
    @Excel(name = "部门id")
    private Long deptId;

    /**
     * 状态(字典)
     */
    @Excel(name = "状态")
    private String status;

    /**
     * 创建时间
     */
    @Excel(name = "创建时间")
    private Date createTime;

    public Long getUserId() {
        return userId;
    }

    public void setUserId(Long userId) {
        this.userId = userId;
    }

    public String getAvatar() {
        return avatar;
    }

    public void setAvatar(String avatar) {
        this.avatar = avatar;
    }

    public String getAccount() {
        return account;
    }

    public void setAccount(String account) {
        this.account = account;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Date getBirthday() {
        return birthday;
    }

    public void setBirthday(Date birthday) {
        this.birthday = birthday;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public String getRoleId() {
        return roleId;
    }

    public void setRoleId(String roleId) {
        this.roleId = roleId;
    }

    public Long getDeptId() {
        return deptId;
    }

    public void setDeptId(Long deptId) {
        this.deptId = deptId;
    }

    public String getStatus() {
        return status;
    }

    public void setStatus(String status) {
        this.status = status;
    }

    public Date getCreateTime() {
        return createTime;
    }

    public void setCreateTime(Date createTime) {
        this.createTime = createTime;
    }

    @Override
    public String toString() {
        return "ExcelItem{" +
                "userId=" + userId +
                ", avatar='" + avatar + '\'' +
                ", account='" + account + '\'' +
                ", name='" + name + '\'' +
                ", birthday=" + birthday +
                ", sex='" + sex + '\'' +
                ", email='" + email + '\'' +
                ", phone='" + phone + '\'' +
                ", roleId='" + roleId + '\'' +
                ", deptId=" + deptId +
                ", status='" + status + '\'' +
                ", createTime=" + createTime +
                '}';
    }
}
