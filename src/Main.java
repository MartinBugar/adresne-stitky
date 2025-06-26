import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public class Main {
    public static void main(String[] args) {
        try {
            // Path to the input file
            String filePath = "input\\file";

            // Create a map with placeholder values
            Map<String, String> input = new HashMap<>();
            input.put("Label", "LabelZ");
            input.put("VS", "VSZ");
            input.put("Oslovenie", "OslovenieZ");
            input.put("Adresat", "AdresatZ");
            input.put("AdresnyRiadok1", "AdresnyRiadok1Z");
            input.put("AdresnyRiadok2", "AdresnyRiadok2Z");
            input.put("Stat", "StatZ");

            // Read the DOCX file, replace placeholders, and export to XML file
            readDocxFile(filePath, input);

            System.out.println("XML file with replaced placeholders has been created and exported to src folder.");
        } catch (Exception e) {
            System.err.println("Error processing file: " + e.getMessage());
            e.printStackTrace();
        }
    }


    /**
     * Reads a DOCX file, replaces placeholders with values from the input map, and prints its content to the console.
     *
     * @param filePath Path to the DOCX file
     * @param input    Map containing key-value pairs for placeholder replacement
     * @return
     * @throws IOException If there's an error reading the file
     */
    private static String readDocxFile(String filePath, Map<String, String> input) throws IOException {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             ZipInputStream zis = new ZipInputStream(fis)) {

            System.out.println("DOCX file contents:");

            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                // Look for the main document content
                if (entry.getName().equals("word/document.xml")) {
                    // Read the XML content
                    String xmlContent = readInputStream(zis);

                    // Extract text from XML and replace placeholders
                    String modifiedXml = xmlContent;
                    List<String> textLines = extractTextFromXml(xmlContent);

                    // Print all lines
                    for (String line : textLines) {
                        System.out.println(line);
                    }

                    // Find and replace variable placeholders
                    modifiedXml = replacePlaceholders(xmlContent, input);

                    // Print the modified XML content
                    System.out.println("\nModified XML content:");
                    System.out.println(modifiedXml);

                    return modifiedXml;
                }
            }
        }
        return filePath;
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

    /**
     * Replaces placeholders in the XML content with values from the input map.
     *
     * @param xmlContent The original XML content
     * @param input      Map containing key-value pairs for placeholder replacement
     * @return The modified XML content with placeholders replaced
     */
    private static String replacePlaceholders(String xmlContent, Map<String, String> input) {
        String modifiedXml = xmlContent;
        // Process the XML content to find and replace placeholders
        int pos = 0;
        while (pos < modifiedXml.length()) {
            int startTag = modifiedXml.indexOf("<w:t", pos);
            if (startTag == -1) break;

            int contentStart = modifiedXml.indexOf(">", startTag);
            if (contentStart == -1) break;
            contentStart++; // Move past the '>'

            int contentEnd = modifiedXml.indexOf("</w:t>", contentStart);
            if (contentEnd == -1) break;

            // Extract the text content
            String content = modifiedXml.substring(contentStart, contentEnd);
            String originalContent = content;

            // Look for template placeholders (text that looks like XML entities)
            if (content.contains("&lt;") && content.contains("&gt;")) {
                // Extract the template placeholder
                int placeholderStart = content.indexOf("&lt;");
                int placeholderEnd = content.lastIndexOf("&gt;") + 4; // Length of "&gt;"

                if (placeholderStart >= 0 && placeholderEnd > placeholderStart) {
                    String placeholder = content.substring(placeholderStart, placeholderEnd);

                    // Convert XML entities to actual characters for processing
                    String processedPlaceholder = placeholder.replace("&lt;", "<").replace("&gt;", ">");

                    // Extract the variable name from the placeholder
                    String variableName = null;

                    // Look for Select="./something" pattern
                    if (processedPlaceholder.contains("Select=\"./")) {
                        int selectIndex = processedPlaceholder.indexOf("Select=\"");
                        int start = selectIndex + 8; // Length of "Select=\"" is 8
                        int end = processedPlaceholder.indexOf("\"", start);

                        if (end != -1) {
                            String fullValue = processedPlaceholder.substring(start, end);

                            if (fullValue.startsWith("./")) {
                                variableName = fullValue.substring(2).trim(); // Remove "./" prefix and trim spaces
                            }
                        }
                    } else {
                        // General case for finding "./" patterns
                        int start = processedPlaceholder.indexOf("./");
                        if (start != -1) {
                            int end = processedPlaceholder.indexOf("/", start + 2);
                            if (end != -1) {
                                variableName = processedPlaceholder.substring(start + 2, end).trim(); // Remove "./" prefix and trim spaces
                            }
                        }
                    }

                    // If we found a variable name and it exists in the input map, replace it
                    if (variableName != null && input.containsKey(variableName)) {
                        // Replace the placeholder with the value from the input map
                        String replacementValue = input.get(variableName);
                        content = content.replace(placeholder, replacementValue);

//                        System.out.println("Replaced placeholder: " + variableName + " with value: " + replacementValue);
                        System.out.println(replacementValue);
                        // Update the XML content
                        modifiedXml = modifiedXml.substring(0, contentStart) + content + modifiedXml.substring(contentEnd);

                        // Adjust position based on the length difference
                        int lengthDifference = content.length() - originalContent.length();
                        contentEnd += lengthDifference;
                    }
                }
            }

            // Move to the next position
            pos = contentEnd + 6; // Length of "</w:t>"
        }

        return modifiedXml;
    }


}
