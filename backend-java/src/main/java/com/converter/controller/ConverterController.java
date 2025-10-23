package com.converter.controller;

import com.converter.service.DocxToHwpConverter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;
import java.util.UUID;

/**
 * DOCX to HWP 변환 API 컨트롤러
 *
 * <p>이 컨트롤러는 DOCX 파일 업로드, 변환, 다운로드 API를 제공합니다.</p>
 */
@RestController
@RequestMapping("/api")
public class ConverterController {

    private static final Logger logger = LoggerFactory.getLogger(ConverterController.class);

    private static final String UPLOAD_DIR = "backend-java/uploads/";
    private static final String OUTPUT_DIR = "backend-java/outputs/";

    @Autowired
    private DocxToHwpConverter converter;

    public ConverterController() {
        // 디렉토리 생성
        try {
            Files.createDirectories(Paths.get(UPLOAD_DIR));
            Files.createDirectories(Paths.get(OUTPUT_DIR));
            logger.info("업로드 디렉토리 생성: {}", UPLOAD_DIR);
            logger.info("출력 디렉토리 생성: {}", OUTPUT_DIR);
        } catch (IOException e) {
            logger.error("디렉토리 생성 실패", e);
        }
    }

    /**
     * Health Check API
     *
     * <p>서버 상태를 확인합니다.</p>
     *
     * @return 서버 상태 정보
     */
    @GetMapping("/health")
    public ResponseEntity<Map<String, Object>> health() {
        logger.info("Health check 요청");

        Map<String, Object> response = new HashMap<>();
        response.put("status", "ok");
        response.put("converter_type", "hwplib (Java)");
        response.put("hwp_installed", false); // hwplib은 HWP 설치가 필요 없음
        response.put("hwp_version", "hwplib 1.1.10");
        response.put("message", "서버가 정상 작동 중입니다");

        logger.info("Health check 응답: {}", response);
        return ResponseEntity.ok(response);
    }

    /**
     * 파일 업로드 API
     *
     * <p>DOCX 파일을 서버에 업로드합니다.</p>
     *
     * @param file 업로드할 DOCX 파일
     * @return 업로드 결과 (파일 ID 포함)
     */
    @PostMapping("/upload")
    public ResponseEntity<Map<String, Object>> uploadFile(@RequestParam("file") MultipartFile file) {
        logger.info("=".repeat(80));
        logger.info("파일 업로드 요청");
        logger.info("파일명: {}", file.getOriginalFilename());
        logger.info("파일 크기: {} bytes", file.getSize());
        logger.info("=".repeat(80));

        Map<String, Object> response = new HashMap<>();

        try {
            // 파일 확장자 검증
            String originalFilename = file.getOriginalFilename();
            if (originalFilename == null || (!originalFilename.endsWith(".docx") && !originalFilename.endsWith(".doc"))) {
                logger.error("잘못된 파일 형식: {}", originalFilename);
                response.put("success", false);
                response.put("message", "DOCX 또는 DOC 파일만 업로드 가능합니다");
                return ResponseEntity.badRequest().body(response);
            }

            // 고유한 파일 ID 생성
            String fileId = UUID.randomUUID().toString();
            String extension = originalFilename.substring(originalFilename.lastIndexOf("."));
            String savedFilename = fileId + extension;
            Path uploadPath = Paths.get(UPLOAD_DIR + savedFilename);

            // 파일 저장
            Files.write(uploadPath, file.getBytes());

            logger.info("✓ 파일 저장 완료: {}", uploadPath);
            logger.info("✓ File ID: {}", fileId);

            response.put("success", true);
            response.put("file_id", fileId);
            response.put("original_filename", originalFilename);
            response.put("saved_filename", savedFilename);
            response.put("message", "파일이 성공적으로 업로드되었습니다");

            return ResponseEntity.ok(response);

        } catch (IOException e) {
            logger.error("파일 업로드 실패", e);
            response.put("success", false);
            response.put("message", "파일 업로드에 실패했습니다: " + e.getMessage());
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(response);
        }
    }

    /**
     * DOCX to HWP 변환 API
     *
     * <p>업로드된 DOCX 파일을 HWP 파일로 변환합니다.</p>
     *
     * @param fileId 변환할 파일의 ID
     * @return 변환 결과
     */
    @PostMapping("/convert")
    public ResponseEntity<Map<String, Object>> convertFile(@RequestParam("file_id") String fileId) {
        logger.info("=".repeat(80));
        logger.info("파일 변환 요청");
        logger.info("File ID: {}", fileId);
        logger.info("=".repeat(80));

        Map<String, Object> response = new HashMap<>();

        try {
            // DOCX 파일 찾기
            Path docxPath = findFile(UPLOAD_DIR, fileId);
            if (docxPath == null) {
                logger.error("파일을 찾을 수 없습니다: {}", fileId);
                response.put("success", false);
                response.put("message", "파일을 찾을 수 없습니다");
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body(response);
            }

            // HWP 출력 경로
            Path hwpPath = Paths.get(OUTPUT_DIR + fileId + ".hwp");

            logger.info("입력 파일: {}", docxPath);
            logger.info("출력 파일: {}", hwpPath);

            // 변환 실행
            logger.info("변환 시작...");
            boolean success = converter.convert(docxPath.toFile(), hwpPath.toFile());

            if (success) {
                logger.info("✓ 변환 성공!");
                response.put("success", true);
                response.put("message", "HWP 파일이 생성되었습니다");
                response.put("output_file", hwpPath.toString());
            } else {
                logger.error("✗ 변환 실패");
                response.put("success", false);
                response.put("message", "변환에 실패했습니다");
            }

            return ResponseEntity.ok(response);

        } catch (Exception e) {
            logger.error("변환 중 오류 발생", e);
            response.put("success", false);
            response.put("detail", e.getMessage());
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(response);
        }
    }

    /**
     * HWP 파일 다운로드 API
     *
     * <p>변환된 HWP 파일을 다운로드합니다.</p>
     *
     * @param fileId 다운로드할 파일의 ID
     * @return HWP 파일
     */
    @GetMapping("/download/{fileId}")
    public ResponseEntity<Resource> downloadFile(@PathVariable String fileId) {
        logger.info("파일 다운로드 요청: {}", fileId);

        try {
            Path hwpPath = Paths.get(OUTPUT_DIR + fileId + ".hwp");

            if (!Files.exists(hwpPath)) {
                logger.error("파일을 찾을 수 없습니다: {}", hwpPath);
                return ResponseEntity.status(HttpStatus.NOT_FOUND).build();
            }

            Resource resource = new FileSystemResource(hwpPath.toFile());

            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"converted_" + fileId + ".hwp\"");
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);

            logger.info("✓ 파일 다운로드 시작: {}", hwpPath);

            return ResponseEntity.ok()
                    .headers(headers)
                    .body(resource);

        } catch (Exception e) {
            logger.error("파일 다운로드 실패", e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }

    /**
     * 디렉토리에서 파일 찾기
     *
     * @param directory 디렉토리 경로
     * @param fileId    파일 ID
     * @return 파일 경로 (없으면 null)
     */
    private Path findFile(String directory, String fileId) throws IOException {
        Path dir = Paths.get(directory);
        if (!Files.exists(dir)) {
            return null;
        }

        return Files.list(dir)
                .filter(path -> path.getFileName().toString().startsWith(fileId))
                .findFirst()
                .orElse(null);
    }
}
