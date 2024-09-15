package com.spring.swagger.service.impl;

import com.deepoove.poi.XWPFTemplate;
import com.spring.swagger.dto.User;
import com.spring.swagger.service.PoiTlService;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.io.BufferedOutputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.Map;

@Service
public class PoiTlServiceImpl implements PoiTlService {
    /**
     * Poi-tl 範例: 產出 Word
     * 樣版: wordTest.docx
     *
     * @param user
     * @param response
     */
    @Override
    public void expostWord(User user, HttpServletResponse response) {
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
            bufferedOutputStream.close();
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
