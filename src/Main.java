import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public class Main {
    public static void main(String[] args) {
        try {
            // Path to the input file
            String filePath = "input\\file";

            // Read the DOCX file and print its content
            readDocxFile(filePath);

        } catch (Exception e) {
            System.err.println("Error processing file: " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * Reads a DOCX file and prints its content to the console.
     * 
     * @param filePath Path to the DOCX file
     * @throws IOException If there's an error reading the file
     */
    private static void readDocxFile(String filePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             ZipInputStream zis = new ZipInputStream(fis)) {

            System.out.println("DOCX file contents:");

            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                // Look for the main document content
                if (entry.getName().equals("word/document.xml")) {
                    // Read the XML content
                    String xmlContent = readInputStream(zis);

                    // Extract text from XML (simple approach)
                    List<String> textLines = extractTextFromXml(xmlContent);

                    // Print all lines
                    for (String line : textLines) {
                        System.out.println(line);
                    }

                    break; // We found what we needed
                }
            }
        }
    }

    /**
     * Reads an input stream into a string.
     */
    private static String readInputStream(InputStream is) throws IOException {
        StringBuilder sb = new StringBuilder();
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(is, "UTF-8"))) {
            String line;
            while ((line = reader.readLine()) != null) {
                sb.append(line).append("\n");
            }
        }
        return sb.toString();
    }

    /**
     * Extracts template placeholders from XML by looking for text between <w:t> tags
     * that contains XML-like content (e.g., &lt;Content Select="./VS"/&gt;).
     */
    private static List<String> extractTextFromXml(String xml) {
        List<String> lines = new ArrayList<>();

        // Simple approach: extract all text between <w:t> and </w:t> tags
        int pos = 0;
        while (pos < xml.length()) {
            int startTag = xml.indexOf("<w:t", pos);
            if (startTag == -1) break;

            int contentStart = xml.indexOf(">", startTag);
            if (contentStart == -1) break;
            contentStart++; // Move past the '>'

            int contentEnd = xml.indexOf("</w:t>", contentStart);
            if (contentEnd == -1) break;

            // Extract the text content
            String content = xml.substring(contentStart, contentEnd);

            // Look for template placeholders (text that looks like XML entities)
            if (content.contains("&lt;") && content.contains("&gt;")) {
                // Extract the template placeholder
                int placeholderStart = content.indexOf("&lt;");
                int placeholderEnd = content.lastIndexOf("&gt;") + 4; // Length of "&gt;"

                if (placeholderStart >= 0 && placeholderEnd > placeholderStart) {
                    String placeholder = content.substring(placeholderStart, placeholderEnd);

                    // Convert XML entities to actual characters
                    placeholder = placeholder.replace("&lt;", "<").replace("&gt;", ">");

                    // Add to lines
                    lines.add(placeholder);
                }
            }

            // Move to the next position
            pos = contentEnd + 6; // Length of "</w:t>"
        }

        return lines;
    }
}
