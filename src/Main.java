import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {
    public static void main(String[] args) {
        try {
            // Path to the input file
            String filePath = "src\\file.docx";

            // Extract variables from the DOCX file
            List<String> variables = extractVariablesFromDocx(filePath);

            // Print the found variables
            System.out.println("Found variables:");
            for (String variable : variables) {
                System.out.println(variable);
            }

        } catch (Exception e) {
            System.err.println("Error processing file: " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * Extracts variables from a DOCX file.
     * Variables are defined as text that starts with "./ and ends with "/
     * 
     * @param filePath Path to the DOCX file
     * @return List of variables found in the file
     * @throws IOException If there's an error reading the file
     */
    private static List<String> extractVariablesFromDocx(String filePath) throws IOException {
        List<String> variables = new ArrayList<>();

        // Create a pattern to match variables
        // Variables start with "./ and ends with "/
        Pattern pattern = Pattern.compile("\\\"\\\\./(.*?)/\\\"");

        try (FileInputStream fis = new FileInputStream(new File(filePath));
             XWPFDocument document = new XWPFDocument(fis)) {

            // Extract text from paragraphs
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                String text = paragraph.getText();
                Matcher matcher = pattern.matcher(text);

                // Find all matches in the current paragraph
                while (matcher.find()) {
                    // Add the variable (without the "./ and "/ delimiters)
                    variables.add(matcher.group(1));
                }
            }
        }

        return variables;
    }
}
