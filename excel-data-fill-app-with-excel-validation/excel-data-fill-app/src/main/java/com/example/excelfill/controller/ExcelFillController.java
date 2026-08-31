package com.example.excelfill.controller;
import com.example.excelfill.service.ExcelFillService;
import io.swagger.v3.oas.annotations.Operation;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
@RestController
@RequestMapping("/api/excel")
public class ExcelFillController {
    private final ExcelFillService excelFillService;
    public ExcelFillController(ExcelFillService excelFillService) { this.excelFillService = excelFillService; }
    @Operation(summary = "Fill Excel template using Excel file and JSON file")
    @PostMapping(value = "/fill", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<byte[]> fillExcel(@RequestPart("templateFile") MultipartFile templateFile,
                                            @RequestPart("jsonFile") MultipartFile jsonFile) throws Exception {
        byte[] output = excelFillService.fillExcelTemplate(templateFile, jsonFile);
        String name = templateFile.getOriginalFilename() == null ? "filled-template.xlsx" : templateFile.getOriginalFilename().replace(".xlsx", "-filled.xlsx");
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + name + "\"")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(output);
    }
}
