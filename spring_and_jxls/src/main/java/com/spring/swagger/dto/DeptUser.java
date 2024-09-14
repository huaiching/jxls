package com.spring.swagger.dto;


import io.swagger.v3.oas.annotations.media.Schema;

import java.util.List;

public class DeptUser {
    private String userDept;
    private List<User> userList;

    private Integer totle;

    public Integer getTotle() {
        return totle;
    }

    public void setTotle(Integer totle) {
        this.totle = totle;
    }

    public String getUserDept() {
        return userDept;
    }

    public void setUserDept(String userDept) {
        this.userDept = userDept;
    }

    public List<User> getUserList() {
        return userList;
    }

    public void setUserList(List<User> userList) {
        this.userList = userList;
    }
}
