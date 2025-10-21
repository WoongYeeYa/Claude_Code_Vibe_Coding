package com.converter;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * DOCX to HWP Converter Application
 * Spring Boot main class
 */
@SpringBootApplication
public class Application {

    private static final Logger logger = LoggerFactory.getLogger(Application.class);

    public static void main(String[] args) {
        logger.info("=".repeat(60));
        logger.info("DOCX to HWP Converter Starting...");
        logger.info("Using hwplib for conversion");
        logger.info("=".repeat(60));

        SpringApplication.run(Application.class, args);

        logger.info("Server started successfully");
        logger.info("API endpoints available at http://localhost:8000");
    }
}
