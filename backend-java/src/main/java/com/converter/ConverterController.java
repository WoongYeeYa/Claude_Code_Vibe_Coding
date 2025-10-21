package com.converter;

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
import java.nio.file.StandardCopyOption;
import java.util.HashMap;
import java.util.Map;
import java.util.UUID;

/**
 * REST API Controller for DOCX to HWP conversion
 */
@RestController
@CrossOrigin(origins = "*")
public class ConverterController {

    private static final Logger logger = LoggerFactory.getLogger(ConverterController.class);

    private final DocxToHwpConverter converter;

    // Directory paths
    private final Path uploadDir = Paths.get("../uploads");
    private final Path outputDir = Paths.get("../outputs");
    private final Path templateDir = Paths.get("../templates");

    @Autowired
    public ConverterController(DocxToHwpConverter converter) {
        this.converter = converter;

        // Create directories if they don't exist
        try {
            Files.createDirectories(uploadDir);
            Files.createDirectories(outputDir);
            Files.createDirectories(templateDir);
            logger.info("Directories created/verified: uploads, outputs, templates");
        } catch (IOException e) {
            logger.error("Failed to create directories: {}", e.getMessage());
        }
    }

    /**
     * Health check endpoint
     */
    @GetMapping("/api/health")
    public ResponseEntity<Map<String, Object>> health() {
        Map<String, Object> response = new HashMap<>();
        response.put("status", "ok");
        response.put("converter_type", "hwplib (Java)");
        response.put("hwp_installed", "not required");
        response.put("hwp_version", "hwplib 1.1.10");
        return ResponseEntity.ok(response);
    }

    /**
     * Upload DOCX file
     */
    @PostMapping("/api/upload")
    public ResponseEntity<Map<String, Object>> uploadFile(@RequestParam("file") MultipartFile file) {
        try {
            logger.info("Received file upload: {}", file.getOriginalFilename());

            // Validate file extension
            String filename = file.getOriginalFilename();
            if (filename == null || (!filename.endsWith(".docx") && !filename.endsWith(".doc"))) {
                Map<String, Object> error = new HashMap<>();
                error.put("success", false);
                error.put("detail", "DOCX 또는 DOC 파일만 업로드 가능합니다");
                return ResponseEntity.badRequest().body(error);
            }

            // Generate unique file ID
            String fileId = UUID.randomUUID().toString();
            String extension = filename.substring(filename.lastIndexOf("."));
            String savedFilename = fileId + extension;

            // Save file
            Path filePath = uploadDir.resolve(savedFilename);
            Files.copy(file.getInputStream(), filePath, StandardCopyOption.REPLACE_EXISTING);

            logger.info("File saved: {} -> {}", filename, savedFilename);

            Map<String, Object> response = new HashMap<>();
            response.put("success", true);
            response.put("file_id", fileId);
            response.put("original_filename", filename);
            response.put("saved_filename", savedFilename);
            response.put("message", "파일이 성공적으로 업로드되었습니다");

            return ResponseEntity.ok(response);

        } catch (Exception e) {
            logger.error("File upload failed: {}", e.getMessage(), e);
            Map<String, Object> error = new HashMap<>();
            error.put("success", false);
            error.put("detail", "파일 업로드 실패: " + e.getMessage());
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(error);
        }
    }

    /**
     * Convert DOCX to HWP
     */
    @PostMapping("/api/convert")
    public ResponseEntity<Map<String, Object>> convert(@RequestParam("file_id") String fileId) {
        try {
            logger.info("Starting conversion for file ID: {}", fileId);

            // Find uploaded file
            File[] uploadedFiles = uploadDir.toFile().listFiles((dir, name) -> name.startsWith(fileId + "."));

            if (uploadedFiles == null || uploadedFiles.length == 0) {
                logger.error("Uploaded file not found: {}", fileId);
                Map<String, Object> error = new HashMap<>();
                error.put("success", false);
                error.put("detail", "업로드된 파일을 찾을 수 없습니다");
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body(error);
            }

            File docxFile = uploadedFiles[0];
            String hwpFilename = fileId + ".hwp";
            Path hwpPath = outputDir.resolve(hwpFilename);

            logger.info("Converting: {} -> {}", docxFile.getName(), hwpFilename);

            // Perform conversion
            boolean success = converter.convert(docxFile.getAbsolutePath(), hwpPath.toString());

            if (!success) {
                logger.error("Conversion failed for: {}", fileId);
                Map<String, Object> error = new HashMap<>();
                error.put("success", false);
                error.put("detail", "HWP 변환에 실패했습니다");
                return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(error);
            }

            logger.info("Conversion completed: {}", hwpFilename);

            Map<String, Object> response = new HashMap<>();
            response.put("success", true);
            response.put("file_id", fileId);
            response.put("hwp_filename", hwpFilename);
            response.put("message", "HWP 파일로 변환되었습니다");

            return ResponseEntity.ok(response);

        } catch (Exception e) {
            logger.error("Conversion failed: {}", e.getMessage(), e);
            Map<String, Object> error = new HashMap<>();
            error.put("success", false);
            error.put("detail", "변환 실패: " + e.getMessage());
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(error);
        }
    }

    /**
     * Download HWP file
     */
    @GetMapping("/api/download/{file_id}")
    public ResponseEntity<Resource> download(@PathVariable("file_id") String fileId) {
        try {
            Path hwpPath = outputDir.resolve(fileId + ".hwp");

            if (!Files.exists(hwpPath)) {
                logger.error("HWP file not found: {}", fileId);
                return ResponseEntity.notFound().build();
            }

            Resource resource = new FileSystemResource(hwpPath);

            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"converted_" + fileId + ".hwp\"");
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);

            logger.info("Downloading file: {}", fileId);

            return ResponseEntity.ok()
                    .headers(headers)
                    .body(resource);

        } catch (Exception e) {
            logger.error("Download failed: {}", e.getMessage(), e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }

    /**
     * Serve frontend index.html
     */
    @GetMapping("/")
    public ResponseEntity<Resource> index() {
        try {
            Path indexPath = Paths.get("../frontend/index.html");
            if (Files.exists(indexPath)) {
                Resource resource = new FileSystemResource(indexPath);
                return ResponseEntity.ok()
                        .contentType(MediaType.TEXT_HTML)
                        .body(resource);
            } else {
                // Return simple JSON response if frontend not found
                logger.warn("Frontend index.html not found at: {}", indexPath.toAbsolutePath());
                return ResponseEntity.notFound().build();
            }
        } catch (Exception e) {
            logger.error("Error serving index: {}", e.getMessage());
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }
}
