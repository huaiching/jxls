package com.spring.swagger.service.impl;

import com.deepoove.poi.XWPFTemplate;
import com.spring.swagger.dto.DeptUser;
import com.spring.swagger.dto.User;
import com.spring.swagger.service.JxlsService;
import jakarta.servlet.http.HttpServletResponse;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;
import org.jxls.util.TransformerFactory;
import org.springframework.stereotype.Service;

import java.io.BufferedOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
@Service
public class JxlsServiceImpl implements JxlsService {
    /**
     * Jxls 範例: 產出單一工作表
     * 樣版: test.xlsx
     *
     * @param userList
     * @param response
     */
    @Override
    public void expostExcel1(List<User> userList, HttpServletResponse response) {
        // 資料處理: 計算總數
        Integer totle = userList.size();

        // Jxls 生成 Excel
        try {
            // 導入 報表樣版
            InputStream inputStream = this.getClass().getResourceAsStream("/templates/test.xlsx");
            // 設定 檔案資訊
            response = getExcelResponseData(response, "測試表");

            // 設定 輸出檔案
            OutputStream outputStream = response.getOutputStream();
            // 設定內容: 第一個參數 要跟 樣版的items 相同
            Context context = new Context();               // 內容物件
            context.putVar("userList",userList);     // 寫入分頁資料: 內容
            context.putVar("totle", totle);          // 寫入分頁資料: 總數
            // 產出 Excel: 樣版檔, 輸出檔, 內容
            JxlsHelper.getInstance().setEvaluateFormulas(true).processTemplate(inputStream, outputStream, context);

            // 釋放資源
            outputStream.flush();
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Jxls 範例: 依照部門拆分多個工作表
     * 樣版: test2.xlsx
     *
     * @param userList
     * @param response
     */
    @Override
    public void expostExcel2(List<User> userList, HttpServletResponse response) {
        // 資料處理
        List<DeptUser> deptUserList = new ArrayList<>();
        userList.stream().map(User::getUserDept).distinct().collect(Collectors.toList())
                .forEach(dept->{
                    List userDeptList = userList.stream().filter(x -> dept.equals(x.getUserDept())).collect(Collectors.toList());
                    DeptUser deptUser = new DeptUser();
                    deptUser.setUserDept(dept);
                    deptUser.setUserList(userDeptList);
                    deptUser.setTotle(userDeptList.size());
                    deptUserList.add(deptUser);
                });

        // Jxls 生成 Excel
        try {

            // 導入 報表樣版
            InputStream inputStream = this.getClass().getResourceAsStream("/templates/test2.xlsx");
            // 設定 檔案資訊
            response = getExcelResponseData(response, "測試表");

            // 設定 輸出檔案
            OutputStream outputStream = response.getOutputStream();
            // 設定內容
            Context context = new Context();
            context.putVar("userData",deptUserList); // 重點: 第一個參數 要跟 樣版的items 相同
            // 產出 Excel: 樣版檔, 輸出檔, 內容
            JxlsHelper.getInstance().setEvaluateFormulas(true).processTemplate(inputStream, outputStream, context);

            // 釋放資源
            outputStream.flush();
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Jxls 範例: 多個不同的工作表
     * 樣版: test3.xlsx
     *
     * @param userList
     * @param response
     */
    @Override
    public void expostExcel3(List<User> userList, HttpServletResponse response) {
        // 計算 總數
        Integer totle = userList.size();

        // Jxls 生成 Excel
        try {
            // 導入 報表樣版
            InputStream inputStream = this.getClass().getResourceAsStream("/templates/test3.xlsx");
            // 設定 檔案資訊
            response = getExcelResponseData(response, "測試表");

            // 設定 輸出檔案
            OutputStream outputStream = response.getOutputStream();

            // 設定內容
            // 1. 建立 Excel 物件
            Transformer transformer = TransformerFactory.createTransformer(inputStream, outputStream);
            AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer, true);
            List<Area> xlsAreaList = areaBuilder.build();

            // 2. 設定分頁資訊: 樣版有三個分頁
            // 2.1. 設定 第一個分頁
            // 2.1.1. 取得分頁資訊
            Area sheet1Area = xlsAreaList.get(0);
            // 2.1.2. 設定分頁內容: 第一個參數 要跟 樣版的items 相同
            Context context1 = new Context();               // 內容物件
            context1.putVar("userList",userList);     // 寫入分頁資料: 內容
            context1.putVar("totle", totle);          // 寫入分頁資料: 總數
            // 2.1.3. 完成分頁設定
            sheet1Area.applyAt(sheet1Area.getStartCellRef(), context1);
            sheet1Area.processFormulas();

            // 2.2. 設定 第二個分頁
            // 2.2.1. 取得分頁資訊
            Area sheet2Area = xlsAreaList.get(1);
            // 2.2.2. 設定分頁內容: 第一個參數 要跟 樣版的items 相同
            Context context2 = new Context();               // 內容物件
            context2.putVar("userList",userList);     // 寫入分頁資料: 內容
            context2.putVar("totle", totle);          // 寫入分頁資料: 總數
            // 2.2.3. 完成分頁設定
            sheet2Area.applyAt(sheet2Area.getStartCellRef(), context2);
            sheet2Area.processFormulas();

            // 2.3. 設定 第三個分頁
            // 2.3.1. 取得分頁資訊
            Area sheet3Area = xlsAreaList.get(2);
            // 2.3.2. 設定分頁內容: 第一個參數 要跟 樣版的items 相同
            Context context3 = new Context();               // 內容物件
            context3.putVar("userList",userList);     // 寫入分頁資料: 內容
            context3.putVar("totle", totle);          // 寫入分頁資料: 總數
            // 2.3.3. 完成分頁設定
            sheet3Area.applyAt(sheet3Area.getStartCellRef(), context3);
            sheet3Area.processFormulas();

            // 3. Excel 輸出
            transformer.write();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 設定 Excel 產出的檔案資訊
     * @param response
     * @param fileName
     * @return
     */
    @Override
    public HttpServletResponse getExcelResponseData(HttpServletResponse response, String fileName)  {
        try {
            fileName = URLEncoder.encode(fileName, "UTF-8");
            response.setHeader("Content-Disposition", "attachment; filename=" +
                    new String(fileName.getBytes("UTF-8"), "ISO-8859-1") + ".xlsx");
            response.setContentType("application/vnd.ms-excel;charset=utf8");
        } catch (Exception e) {
            e.printStackTrace();
        }

        return response;
    }
}
