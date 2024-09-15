package com.spring.swagger.controller;

import com.spring.swagger.dto.User;
import com.spring.swagger.service.JxlsService;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@Tag(name = "Jxls", description = "Jxls APIs")
@RestController
@RequestMapping("/api/jxls")
public class JxlsController {
    @Autowired
    private JxlsService jxlsService;

    /**
     * Jxls 範例: 產出單一工作表
     * 樣版: test.xlsx
     * @param userList
     * @param response
     */
    @Operation(summary = "expostExcel1", description = "Jxls 範例: 產出單一工作表")
    @PostMapping("/expostExcel1")
    public void expostExcel1(
            @RequestBody List<User> userList, HttpServletResponse response
    ) {
        try {
            jxlsService.expostExcel1(userList, response);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Jxls 範例: 依照部門拆分多個工作表
     * 樣版: test2.xlsx
     * @param userList
     * @param response
     */
    @Operation(summary = "expostExcel2", description = "Jxls 範例: 依照部門拆分多個工作表")
    @PostMapping("/expostExcel2")
    public void expostExcel2(
            @RequestBody List<User> userList, HttpServletResponse response
    ) {
        try {
            jxlsService.expostExcel2(userList, response);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Jxls 範例: 多個不同的工作表
     * 樣版: test3.xlsx
     * @param userList
     * @param response
     */
    @Operation(summary = "expostExcel3", description = "Jxls 範例: 多個不同的工作表")
    @PostMapping("/expostExcel3")
    public void expostExcel3(
            @RequestBody List<User> userList, HttpServletResponse response
    ) {
        try {
            jxlsService.expostExcel3(userList, response);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
