package com.spring.swagger.service;

import com.spring.swagger.dto.User;
import jakarta.servlet.http.HttpServletResponse;

public interface PoiTlService {

    /**
     * Poi-tl 範例: 產出 Word
     * 樣版: wordTest.docx
     * @param user
     * @param response
     */
    public void expostWord(User user, HttpServletResponse response);
}
