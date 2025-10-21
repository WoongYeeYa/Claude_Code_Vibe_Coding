package com.converter;

import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.ResourceHandlerRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

/**
 * Web MVC Configuration for serving static frontend files
 */
@Configuration
public class WebConfig implements WebMvcConfigurer {

    @Override
    public void addResourceHandlers(ResourceHandlerRegistry registry) {
        // Serve frontend static files
        registry.addResourceHandler("/static/**")
                .addResourceLocations("file:../frontend/", "classpath:/static/");

        // Serve frontend files directly
        registry.addResourceHandler("/**")
                .addResourceLocations("file:../frontend/", "classpath:/static/")
                .resourceChain(false);
    }
}
