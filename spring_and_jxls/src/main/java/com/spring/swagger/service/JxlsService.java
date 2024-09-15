package com.spring.swagger.service;

import com.spring.swagger.dto.User;
import jakarta.servlet.http.HttpServletResponse;

import java.util.List;

public interface JxlsService {

    /**
     * Jxls 範例: 產出單一工作表
     * 樣版: test.xlsx
     * @param userList
     * @param response
     */
    public void expostExcel1(List<User> userList, HttpServletResponse response);

    /**
     * Jxls 範例: 依照部門拆分多個工作表
     * 樣版: test2.xlsx
     * @param userList
     * @param response
     */
    public void expostExcel2(List<User> userList, HttpServletResponse response);

    /**
     * Jxls 範例: 多個不同的工作表
     * 樣版: test3.xlsx
     * @param userList
     * @param response
     */
    public void expostExcel3(List<User> userList, HttpServletResponse response);

    /**
     * 設定 Excel 產出的檔案資訊
     * @param response
     * @param fileName
     * @return
     */
    public HttpServletResponse getExcelResponseData(HttpServletResponse response, String fileName);

}
