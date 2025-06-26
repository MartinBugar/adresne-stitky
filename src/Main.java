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

                    // Find and list variable placeholders
                    findVariablePlaceholders(textLines);

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

    /**
     * Finds variable placeholders in the extracted text lines that start with "./" and end with "/"
     * and lists them in the console.
     * 
     * @param lines List of extracted text lines
     */
    private static void findVariablePlaceholders(List<String> lines) {
        System.out.println("\nVariable placeholders found:");

        for (String line : lines) {
            // Look for Select="./something" pattern which is common in the XML
            if (line.contains("Select=\"./")) {
                int selectIndex = line.indexOf("Select=\"");
                int start = selectIndex + 8; // Length of "Select=\"" is 8
                int end = line.indexOf("\"", start);

                if (end != -1) {
                    // Extract the placeholder including "./" prefix
                    String fullValue = line.substring(start, end);

                    // Check if the value contains a variable placeholder
                    if (fullValue.startsWith("./")) {
                        // Remove "./" prefix before printing
                        System.out.println(fullValue.substring(2));
                    }
                }
            } else {
                // General case for finding "./" patterns
                int startIndex = 0;
                while (true) {
                    // Find the start of a placeholder
                    int start = line.indexOf("./", startIndex);
                    if (start == -1) break;

                    // Find the end of the placeholder (next "/" after "./")
                    int end = line.indexOf("/", start + 2);
                    if (end == -1) break;

                    // Extract the placeholder
                    String placeholder = line.substring(start, end + 1);
                    // Remove "./" prefix before printing
                    System.out.println(placeholder.substring(2));

                    // Move to the next position
                    startIndex = end + 1;
                }
            }
        }
    }
}
