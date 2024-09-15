package com.spring.swagger.controller;

import com.deepoove.poi.XWPFTemplate;
import com.spring.swagger.dto.User;
import com.spring.swagger.service.JxlsService;
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
            // 設定內容
            Map<String, Object> context = new HashMap<>();
            context.put("userCode", user.getUserCode());
            context.put("userName", user.getUserName());
            context.put("userAge", user.getUserAge());
            context.put("userDept",user.getUserDept());
            context.put("userSalary",user.getUserSalary());
            // 渲染數據
            XWPFTemplate template = XWPFTemplate.compile(new ClassPathResource("/templates/wordTest.docx").getFile());
            template.render(context);
            // 設定網路流
            String wordFileName = "測試文件.docx";
            String fileName = URLEncoder.encode(wordFileName, "UTF-8");
            response.setHeader("Content-Disposition", "attachment; filename=" +
                    new String(fileName.getBytes("UTF-8"), "ISO-8859-1"));
            response.setContentType("application/octet-stream");
            OutputStream outputStream = response.getOutputStream();
            // 導出文件
            BufferedOutputStream bufferedOutputStream = new BufferedOutputStream(outputStream);
            template.write(bufferedOutputStream);
            // 釋放資源
            bufferedOutputStream.flush();
            outputStream.flush();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
