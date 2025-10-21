import com.converter.DocxToHwpConverter;

public class test_convert {
    public static void main(String[] args) {
        try {
            DocxToHwpConverter converter = new DocxToHwpConverter();
            String inputFile = "templates/공문양식.docx";
            String outputFile = "outputs/공문양식_test.hwp";
            
            System.out.println("Converting: " + inputFile + " -> " + outputFile);
            boolean success = converter.convert(inputFile, outputFile);
            
            if (success) {
                System.out.println("Conversion successful!");
            } else {
                System.out.println("Conversion failed!");
            }
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
