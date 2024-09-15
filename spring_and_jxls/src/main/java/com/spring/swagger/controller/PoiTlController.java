package com.spring.swagger.controller;

import com.deepoove.poi.XWPFTemplate;
import com.spring.swagger.dto.User;
import com.spring.swagger.service.JxlsService;
import com.spring.swagger.service.PoiTlService;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.BufferedOutputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Tag(name = "Poi-tl", description = "Poi-tl APIs")
@RestController
@RequestMapping("/api/Poi-tl")
public class PoiTlController {
    @Autowired
    private PoiTlService poiTlService;

    /**
     * Poi-tl 範例: 產出 Word
     * 樣版: wordTest.docx
     * @param user
     * @param response
     */
    @Operation(summary = "expostWord", description = "Poi-tl 範例: 產出 Word")
    @PostMapping("/expostWord")
    public void expostWord(
            @RequestBody User user, HttpServletResponse response
    ) {
        try {
            poiTlService.expostWord(user, response);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
