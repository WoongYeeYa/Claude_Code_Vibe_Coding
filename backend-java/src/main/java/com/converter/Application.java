package com.converter;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;
import org.springframework.web.servlet.config.annotation.CorsRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

/**
 * DOCX to HWP 변환 서버 메인 애플리케이션
 *
 * <p>이 애플리케이션은 DOCX 파일을 HWP 파일로 변환하는 REST API 서버입니다.
 * Apache POI를 사용하여 DOCX 파일을 파싱하고, hwplib을 사용하여 HWP 파일을 생성합니다.</p>
 *
 * <p>주요 기능:</p>
 * <ul>
 *   <li>DOCX 파일 업로드</li>
 *   <li>텍스트 및 스타일 추출 (폰트, 색상, 크기, 굵기, 이탤릭 등)</li>
 *   <li>테이블 구조 및 스타일 추출</li>
 *   <li>HWP 파일 생성 및 다운로드</li>
 * </ul>
 */
@SpringBootApplication
public class Application {

    private static final Logger logger = LoggerFactory.getLogger(Application.class);

    public static void main(String[] args) {
        logger.info("=".repeat(80));
        logger.info("DOCX to HWP 변환 서버 시작");
        logger.info("=".repeat(80));
        logger.info("✓ Apache POI: DOCX 파일 파싱");
        logger.info("✓ hwplib: HWP 파일 생성");
        logger.info("✓ 포트: 8000");
        logger.info("=".repeat(80));

        SpringApplication.run(Application.class, args);

        logger.info("서버가 성공적으로 시작되었습니다!");
        logger.info("프론트엔드: http://localhost:8000");
        logger.info("API 엔드포인트: http://localhost:8000/api/");
    }

    /**
     * CORS 설정
     * 프론트엔드와의 통신을 위해 CORS를 허용합니다.
     */
    @Bean
    public WebMvcConfigurer corsConfigurer() {
        return new WebMvcConfigurer() {
            @Override
            public void addCorsMappings(CorsRegistry registry) {
                registry.addMapping("/api/**")
                        .allowedOrigins("*")
                        .allowedMethods("GET", "POST", "PUT", "DELETE", "OPTIONS")
                        .allowedHeaders("*");
            }
        };
    }
}
