package com.example.parser.statistic.controller;

import com.example.parser.user.entity.User;
import com.example.parser.user.repository.UserRepository;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
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

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.List;

@Slf4j
@RestController
@RequiredArgsConstructor
@RequestMapping("/v1")
public class UserStatisticController {

    private final UserRepository userRepository;

    private static final String[] HEADERS = {"사용자 ID", "사용자 이름"};

    @GetMapping("/users/user-statistic")
    public ResponseEntity<InputStreamResource> downloadUserStatistic() {
        final List<User> users = userRepository.findAll();

        try (final Workbook workbook = new SXSSFWorkbook();
             final ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            final Sheet sheet = createExcelSheet(workbook);
            setSheetHeader(sheet);
            setSheetBody(users, sheet);
            workbook.write(out);

            return createDownloadResponse(out);
        } catch (IOException e) {
            log.error("파일 생성에 실패하였습니다.", e);
            throw new RuntimeException(e);
        }
    }

    private static ResponseEntity<InputStreamResource> createDownloadResponse(final ByteArrayOutputStream out) {
        InputStreamResource resource = new InputStreamResource(new ByteArrayInputStream(out.toByteArray()));

        return ResponseEntity.ok()
                .contentLength(out.size())
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .header("Content-Disposition", "attachment; filename=\"user-statistic.xlsx\"")
                .body(resource);
    }

    private static void setSheetBody(List<User> users, Sheet sheet) {
        int rowNum = 1;
        for (User user : users) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(user.getId());
            row.createCell(1).setCellValue(user.getName());
        }
    }

    private static void setSheetHeader(Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < HEADERS.length; i++) {
            headerRow.createCell(i).setCellValue(HEADERS[i]);
        }
    }

    private static Sheet createExcelSheet(Workbook workbook) {
        String sheetName = getSheetName();
        return workbook.createSheet(sheetName);
    }

    private static String getSheetName() {
        LocalDate today = LocalDate.now();
        return String.format("탈리월드 사용자 통계(%d-%02d)", today.getYear(), today.getMonthValue());
    }
}
