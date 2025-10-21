package com.converter;

import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.tool.blankfilemaker.BlankFileMaker;
import kr.dogfoot.hwplib.writer.HWPWriter;

public class SimpleHwpTest {
    public static void main(String[] args) {
        try {
            System.out.println("Creating blank HWP file using BlankFileMaker...");
            HWPFile hwpFile = BlankFileMaker.make();

            if (hwpFile != null) {
                String outputPath = "../outputs/test_blank.hwp";
                HWPWriter.toFile(hwpFile, outputPath);
                System.out.println("✓ Blank HWP file created successfully: " + outputPath);
            } else {
                System.err.println("✗ Failed to create blank HWP file");
            }
        } catch (Exception e) {
            System.err.println("ERROR: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
