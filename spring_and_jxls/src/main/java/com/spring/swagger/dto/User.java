package com.spring.swagger.dto;


import io.swagger.v3.oas.annotations.media.Schema;

public class User {
    @Schema(description = "學員編號")
    private String userCode;
    @Schema(description = "學員姓名")
    private String userName;
    @Schema(description = "學員年齡")
    private String userDept;
    @Schema(description = "學員部門")
    private Integer userAge;
    @Schema(description = "學員薪水")
    private Integer userSalary;

    public Integer getUserSalary() {
        return userSalary;
    }

    public void setUserSalary(Integer userSalary) {
        this.userSalary = userSalary;
    }

    public String showUserData() {
        String msg = "學員編號: " + userCode + "\n學員姓名: " + userName + "\n學員年齡: " + userAge + "\n學員部門: " + userDept;
        return msg;
    }

    public String getUserCode() {
        return userCode;
    }

    public void setUserCode(String userCode) {
        this.userCode = userCode;
    }

    public String getUserName() {
        return userName;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }

    public String getUserDept() {
        return userDept;
    }

    public void setUserDept(String userDept) {
        this.userDept = userDept;
    }

    public Integer getUserAge() {
        return userAge;
    }

    public void setUserAge(Integer userAge) {
        this.userAge = userAge;
    }
}
