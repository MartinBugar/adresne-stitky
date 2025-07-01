import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

/**
 * Main application class for processing DOCX files and replacing placeholders with actual values.
 * This utility reads a DOCX template file, identifies XML placeholders in the format &lt;Content Select="./KEY"/&gt;,
 * and replaces them with corresponding values from a data list. The program can process multiple data records
 * and generate a single output file with all records or individual files for each record.
 */
public class Main {
    /**
     * XML tags that need to be removed from the document.
     * These tags are used for template control but should not appear in the final document.
     * The list includes both HTML-encoded (&lt; &gt;) and regular XML versions of the tags.
     */
    private static final String[] XML_TAGS_TO_REMOVE = {
        "&lt;EndRepeat/&gt;", "<EndRepeat/>", 
        "&lt;Repeat Select=\"./Label \"/&gt;", "<Repeat Select=\"./Label \"/>"
    };

    /**
     * Main entry point of the application.
     * Creates sample data records, processes the template file, and generates output documents.
     *
     * @param args Command line arguments (not used)
     */
    public static void main(String[] args) {
        try {
            // Path to the input file (DOCX template)
            String filePath = "input\\file";

            // Create a list to hold all data records for processing
            List<Map<String, String>> dataList = new ArrayList<>();

            // Sample record 1 - Each key corresponds to a placeholder in the template
            // Keys match the variable names in the XML placeholders: <Content Select="./KEY"/>
            Map<String,String> rec1 = new HashMap<>();
            rec1.put("Label",       "ID-12345");     // Identifier label
            rec1.put("VS",          "2025");         // Variable symbol
            rec1.put("Oslovenie",   "Vážený pán Novák,"); // Salutation
            rec1.put("Adresat",     "Ján Novák");    // Recipient name
            rec1.put("AdresnyRiadok1", "Hlavná 10"); // Address line 1
            rec1.put("AdresnyRiadok2", "010 01 Žilina"); // Address line 2 (postal code and city)
            rec1.put("Stat",        "Slovensko");    // Country

            // Sample record 2
            Map<String,String> rec2 = new HashMap<>();
            rec2.put("Label",       "ID-12345");
            rec2.put("VS",          "2026");
            rec2.put("Oslovenie",   "Vážený pán Hurban,");
            rec2.put("Adresat",     "Ján Hurban");
            rec2.put("AdresnyRiadok1", "Hlavná 25");
            rec2.put("AdresnyRiadok2", "010 01 Martin");
            rec2.put("Stat",        "Madarsko");

            // Sample record 3
            Map<String,String> rec3 = new HashMap<>();
            rec3.put("Label",       "ID-12345");
            rec3.put("VS",          "2026");
            rec3.put("Oslovenie",   "Vážený pán Hurban,");
            rec3.put("Adresat",     "Ján Hurban");
            rec3.put("AdresnyRiadok1", "Hlavná 25");
            rec3.put("AdresnyRiadok2", "010 01 Martin");
            rec3.put("Stat",        "Madarsko");

            // Add all records to the data list for processing
            dataList.add(rec1);
            dataList.add(rec2);
            dataList.add(rec3);

            // Read the DOCX file, replace placeholders, and export to XML file
            readDocxFile(filePath, dataList);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    /**
     * Removes specific XML control tags from the document content.
     * This method cleans up the XML by removing template control tags that should not
     * appear in the final document, such as &lt;Repeat&gt; and &lt;EndRepeat&gt; tags.
     * These tags are used for template processing but are not part of the actual content.
     * 
     * @param content The XML content to process
     * @return The cleaned XML content with control tags removed
     */
    private static String removeXmlTags(String content) {
        // Iterate through all tags that need to be removed
        for (String tag : XML_TAGS_TO_REMOVE) {
            // Replace each tag with an empty string
            content = content.replace(tag, "");
        }
        return content;
    }

    /**
     * Reads a DOCX file, replaces placeholders with values from the dataList, and creates a new DOCX file
     * with the replaced content. This method handles the core document processing logic.
     *
     * @param filePath Path to the DOCX file (template)
     * @param dataList List of maps containing key-value pairs for placeholder replacement
     * @throws IOException If there's an error reading or writing files
     */
    private static void readDocxFile(String filePath, List<Map<String, String>> dataList) throws IOException {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             ZipInputStream zis = new ZipInputStream(fis)) {

            // Extract the document.xml file from the DOCX (which is a ZIP file)
            String originalXml = null;
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                if (entry.getName().equals("word/document.xml")) {
                    // Found the main document XML content
                    originalXml = readInputStream(zis);
                }
            }

            if (originalXml != null) {
                // Remove control tags from the template
                String modifiedXml = removeXmlTags(originalXml);

                // Process each dataset (record) in the data list
                for (int i = 0; i < dataList.size(); i++) {
                    // Replace placeholders with actual values from the current dataset
                    modifiedXml = replacePlaceholders(modifiedXml, dataList.get(i));

                    // If this isn't the last dataset, duplicate the template for the next record
                    if (i < dataList.size() - 1) {
                        int bodyEndPos = modifiedXml.lastIndexOf("</w:body>");
                        if (bodyEndPos != -1) {
                            // Add a paragraph break between records
                            String paragraphBreak = "<w:p><w:r><w:t xml:space=\"preserve\">&#10;</w:t></w:r></w:p>";
                            // Extract the body content from the original template
                            String templateContent = removeXmlTags(originalXml.substring(
                                originalXml.indexOf("<w:body>") + 8, 
                                originalXml.lastIndexOf("</w:body>")
                            ));

                            // Insert the template content for the next record
                            modifiedXml = modifiedXml.substring(0, bodyEndPos) + 
                                          paragraphBreak + 
                                          templateContent + 
                                          modifiedXml.substring(bodyEndPos);
                        }
                    }
                }

                // Create the output DOCX file with all processed records
                createDocxFile(filePath, removeXmlTags(modifiedXml));
            }
        }
    }

    /**
     * Reads an input stream into a string without closing the stream.
     * This is used to read the XML content from the DOCX file's ZIP entries.
     * 
     * @param is The input stream to read from
     * @return A string containing all the content from the input stream
     * @throws IOException If there's an error reading from the stream
     */
    private static String readInputStream(InputStream is) throws IOException {
        StringBuilder sb = new StringBuilder();
        // Use UTF-8 encoding to properly handle international characters
        BufferedReader reader = new BufferedReader(new InputStreamReader(is, StandardCharsets.UTF_8));
        String line;
        // Read the stream line by line and append to the StringBuilder
        while ((line = reader.readLine()) != null) {
            sb.append(line).append("\n");
        }
        return sb.toString();
    }

    /**
     * Replaces placeholders in the XML content with values from the input map.
     * This method scans through the XML content, identifies placeholders in the format
     * &lt;Content Select="./KEY"/&gt;, and replaces them with corresponding values from the input map.
     *
     * @param xmlContent The original XML content from the DOCX template
     * @param input Map containing key-value pairs for placeholder replacement
     * @return The modified XML content with placeholders replaced by actual values
     */
    private static String replacePlaceholders(String xmlContent, Map<String, String> input) {
        String modifiedXml = xmlContent;

        int pos = 0;
        // Iterate through the XML content looking for text tags that might contain placeholders
        while (pos < modifiedXml.length()) {
            // Find the next text tag
            int startTag = modifiedXml.indexOf("<w:t", pos);
            if (startTag == -1) break; // No more text tags

            // Find the end of the opening tag
            int contentStart = modifiedXml.indexOf(">", startTag);
            if (contentStart == -1) break;
            contentStart++; // Move past the '>'

            // Find the closing tag
            int contentEnd = modifiedXml.indexOf("</w:t>", contentStart);
            if (contentEnd == -1) break;

            // Extract the content between the tags
            String content = modifiedXml.substring(contentStart, contentEnd);
            String originalContent = content; // Save original for length calculation later

            // Check if the content contains XML-encoded tags (potential placeholders)
            if (content.contains("&lt;") && content.contains("&gt;")) {
                int placeholderStart = content.indexOf("&lt;");
                int placeholderEnd = content.lastIndexOf("&gt;") + 4; // Include the "&gt;"

                if (placeholderStart >= 0 && placeholderEnd > placeholderStart) {
                    // Extract the encoded placeholder
                    String placeholder = content.substring(placeholderStart, placeholderEnd);
                    // Convert HTML entities to XML characters
                    String processedPlaceholder = placeholder.replace("&lt;", "<").replace("&gt;", ">");
                    // Extract the variable name from the placeholder
                    String variableName = extractVariableName(processedPlaceholder);

                    // If the variable exists in our input data, replace it
                    if (variableName != null && input.containsKey(variableName)) {
                        String replacementValue = input.get(variableName);
                        // Replace the placeholder with the actual value
                        content = content.replace(placeholder, replacementValue);

                        // Update the XML content with the replaced value
                        modifiedXml = modifiedXml.substring(0, contentStart) + content + modifiedXml.substring(contentEnd);
                        // Adjust the content end position based on the new content length
                        contentEnd += (content.length() - originalContent.length());
                    }
                }
            }

            // Move to the position after the closing tag
            pos = contentEnd + 6; // Length of "</w:t>"
        }

        // Remove any control tags that might remain
        return removeXmlTags(modifiedXml);
    }

    /**
     * Extracts variable name from a placeholder XML tag.
     * This method parses different formats of placeholders to extract the variable name.
     * It handles both the standard format &lt;Content Select="./VARIABLE"/&gt; and other formats.
     * 
     * @param placeholder The placeholder string to parse
     * @return The extracted variable name or null if no valid variable name is found
     */
    private static String extractVariableName(String placeholder) {
        // Look for the standard format: Select="./something" pattern
        if (placeholder.contains("Select=\"./")) {
            int selectIndex = placeholder.indexOf("Select=\"");
            int start = selectIndex + 8; // Length of "Select=\"" is 8
            int end = placeholder.indexOf("\"", start);

            if (end != -1) {
                String fullValue = placeholder.substring(start, end);
                if (fullValue.startsWith("./")) {
                    // Remove "./" prefix and trim spaces
                    return fullValue.substring(2).trim();
                }
            }
        } else {
            // Alternative format: look for "./" patterns directly
            int start = placeholder.indexOf("./");
            if (start != -1) {
                int end = placeholder.indexOf("/", start + 2);
                if (end != -1) {
                    // Extract the variable name between "./" and the next "/"
                    return placeholder.substring(start + 2, end).trim();
                }
            }
        }
        // No valid variable name found
        return null;
    }


    /**
     * Creates a new DOCX file with the replaced placeholders.
     * This method takes the original DOCX file as a template, replaces the document.xml
     * content with the modified XML, and creates a new DOCX file with the changes.
     *
     * @param originalFilePath Path to the original DOCX template file
     * @param modifiedXml The modified XML content with placeholders replaced by actual values
     * @throws IOException If there's an error reading the original file or writing the new file
     */
    private static void createDocxFile(String originalFilePath, String modifiedXml) throws IOException {
        // Get the current working directory
        String currentDir = System.getProperty("user.dir");
        // Ensure the src directory exists
        File srcDir = new File(currentDir, "src");
        if (!srcDir.exists()) {
            srcDir.mkdir();
        }

        // Generate the output filename
        String fileName = "output_all.docx";
        File outputFile = new File(srcDir, fileName);

        // Open streams for reading the original file and writing the new file
        try (FileInputStream fis = new FileInputStream(new File(originalFilePath));
             ZipInputStream zis = new ZipInputStream(fis);
             FileOutputStream fos = new FileOutputStream(outputFile);
             ZipOutputStream zos = new ZipOutputStream(fos)) {

            ZipEntry entry;
            byte[] buffer = new byte[1024];

            // Copy all entries from the original DOCX file to the new file
            while ((entry = zis.getNextEntry()) != null) {
                // Create a new entry in the output ZIP file
                ZipEntry newEntry = new ZipEntry(entry.getName());
                zos.putNextEntry(newEntry);

                // If this is the document.xml file, write our modified XML content
                if (entry.getName().equals("word/document.xml")) {
                    zos.write(modifiedXml.getBytes(StandardCharsets.UTF_8));
                } else {
                    // Otherwise, copy the original content
                    int len;
                    while ((len = zis.read(buffer)) > 0) {
                        zos.write(buffer, 0, len);
                    }
                }

                // Close the entries
                zos.closeEntry();
                zis.closeEntry();
            }
        }
        // Resources are automatically closed by try-with-resources
    }


}
