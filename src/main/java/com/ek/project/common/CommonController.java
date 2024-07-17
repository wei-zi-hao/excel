package com.ek.project.common;


import com.ek.common.utils.DateUtils;
import com.ek.framework.web.damain.AjaxResult;
import com.ek.project.common.dto.ExcelDetail;
import com.ek.project.service.ExcelService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.List;

/**
 * 通用请求处理
 *
 * @author ruoyi
 */
@Controller
public class CommonController {
    private static final Logger log = LoggerFactory.getLogger(CommonController.class);

    @Autowired
    private ExcelService excelService;

    public static void main(String[] args) {
        log.info(DateUtils.dateTimeNow());
    }
    /**
     * 通用下载请求
     *
     */
    @GetMapping("common/download")
    public void fileDownload(String fileName, HttpServletResponse response) {
        try {
           excelService.downloadExcel(fileName, response);
        } catch (Exception e) {
            log.error("下载文件失败", e);
        }
    }

    /**
     * 通用下载请求
     *
     */
    @PostMapping("excel/handle")
    @ResponseBody
    public AjaxResult excelHandle(@RequestBody ExcelDetail excelDetail) {
        try {
            return excelService.handleExcel(excelDetail);
        } catch (Exception e) {
            log.error("下载文件失败", e);
        }
        return AjaxResult.error();
    }



    /**
     * 通用上传请求
     */
    @PostMapping("/common/upload")
    @ResponseBody
    public AjaxResult uploadFile(List<MultipartFile> files, HttpServletResponse response) throws Exception {
        response.setHeader("Access-Control-Allow-Origin", "*");
        for (MultipartFile file : files) {
            log.info("正在上传:"+file.getOriginalFilename());
        }
        try {
            AjaxResult ajaxResult = excelService.getExcel(files);
            ajaxResult.put("oldExcelName", files.get(0).getOriginalFilename());
            ajaxResult.put("newExcelName", files.get(1).getOriginalFilename());
            return ajaxResult;
        } catch (Exception e) {
            return AjaxResult.error(e.getMessage());
        }
    }

    /**
     * 通用上传请求
     */
//    @PostMapping("/excel/handel")
//    @ResponseBody
//    public AjaxResult handelExcel(@RequestBody ExcelDetail excelDetail) throws Exception {
//        try {
//            return excelService.handleExcel(excelDetail);
//        } catch (Exception e) {
//            return AjaxResult.error(e.getMessage());
//        }
//    }



}
