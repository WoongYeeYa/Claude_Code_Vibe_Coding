package com.converter.config;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.ResourceHandlerRegistry;
import org.springframework.web.servlet.config.annotation.ViewControllerRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

/**
 * Spring MVC 설정
 *
 * <p>정적 리소스 및 뷰 매핑을 설정합니다.</p>
 */
@Configuration
public class WebConfig implements WebMvcConfigurer {

    private static final Logger logger = LoggerFactory.getLogger(WebConfig.class);

    /**
     * 정적 리소스 핸들러 설정
     *
     * <p>프론트엔드 파일을 서빙합니다.</p>
     */
    @Override
    public void addResourceHandlers(ResourceHandlerRegistry registry) {
        logger.info("정적 리소스 핸들러 설정");

        // /static/** 경로로 요청 시 frontend 디렉토리의 파일 제공
        registry.addResourceHandler("/static/**")
                .addResourceLocations("file:C:/workspace/Claude_Code_Vibe_Coding/frontend/", "classpath:/static/");

        logger.info("✓ /static/** → frontend/ 디렉토리");
    }

    /**
     * 뷰 컨트롤러 설정
     *
     * <p>루트 경로로 요청 시 index.html을 제공합니다.</p>
     */
    @Override
    public void addViewControllers(ViewControllerRegistry registry) {
        // 루트 경로를 index.html로 포워딩
        registry.addViewController("/").setViewName("forward:/static/index.html");
        logger.info("✓ / → /static/index.html");
    }
}
