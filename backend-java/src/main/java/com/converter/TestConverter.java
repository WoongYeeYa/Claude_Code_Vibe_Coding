package com.converter;

import java.io.File;

/**
 * Simple test class to test DOCX to HWP conversion
 */
public class TestConverter {
    public static void main(String[] args) {
        try {
            System.out.println("=== DOCX to HWP Conversion Test ===");

            DocxToHwpConverter converter = new DocxToHwpConverter();

            String inputFile = "backend-java/templates/공문양식.docx";
            String outputFile = "outputs/공문양식_test.hwp";

            File input = new File(inputFile);
            if (!input.exists()) {
                System.err.println("Input file not found: " + inputFile);
                System.err.println("Absolute path: " + input.getAbsolutePath());
                return;
            }

            System.out.println("Input file: " + inputFile);
            System.out.println("Output file: " + outputFile);
            System.out.println("Converting...");

            boolean success = converter.convert(inputFile, outputFile);

            if (success) {
                System.out.println("✓ Conversion successful!");
                File output = new File(outputFile);
                if (output.exists()) {
                    System.out.println("Output file size: " + output.length() + " bytes");
                }
            } else {
                System.err.println("✗ Conversion failed!");
            }

        } catch (Exception e) {
            System.err.println("ERROR: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
