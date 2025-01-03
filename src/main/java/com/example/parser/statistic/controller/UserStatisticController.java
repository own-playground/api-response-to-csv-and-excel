package com.example.parser.statistic.controller;

import com.example.parser.user.entity.User;
import com.example.parser.user.repository.UserRepository;
import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;
import java.util.List;

@Slf4j
@RestController
@RequiredArgsConstructor
@RequestMapping("/v1")
public class UserStatisticController {

    private final UserRepository userRepository;

    @GetMapping("/users/user-statistic")
    public ResponseEntity<InputStreamResource> downloadUserStatistic(
            final HttpServletResponse response
    ) throws IOException {
        final List<User> users = userRepository.findAll();

        // 엑셀 파일
        Workbook workbook = new SXSSFWorkbook();

        // 엑셀 시트
        Sheet sheet = workbook.createSheet("사용자 포인트 통계");

        // 로우 & 열 (n x m)
        final Row row = sheet.createRow(0);
        final Cell cell = row.createCell(0);
        cell.setCellValue("사용자 ID");

        File tmpFile = File.createTempFile("temp", ".xlsx");
        try (OutputStream fos = new FileOutputStream(tmpFile)) {
            workbook.write(fos);
        }
        InputStream res = new FileInputStream(tmpFile) {
            @Override
            public void close() throws IOException {
                super.close();
                if (tmpFile.delete()) {
                    log.info("삭제완료");
                }
            }
        };
        return ResponseEntity.ok()
                .contentLength(tmpFile.length())
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .header("Content-Disposition", String.format("attachment;filename=\"%s.xlsx\";", "temp"))
                .body(new InputStreamResource(res));
    }


}
