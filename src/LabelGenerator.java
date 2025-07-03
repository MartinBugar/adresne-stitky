import java.io.BufferedReader;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import org.w3c.dom.Document;
import org.xml.sax.SAXException;

/**
 * Utility class for processing DOCX files and replacing placeholders with actual values.
 * This utility reads a DOCX template file, identifies XML placeholders in the format &lt;Content Select="./KEY"/&gt;,
 * and replaces them with corresponding values from a data list. The program can process multiple data records
 * and generate a single output file with all records or individual files for each record.
 */
public class LabelGenerator {
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
     * Reads a DOCX file, replaces placeholders with values from the dataList, and returns a DocumentAssemblerResponse
     * object containing the processed document. This method handles the core document processing logic.
     *
     * @param document Document object representing the DOCX file (template)
     * @param dataList List of maps containing key-value pairs for placeholder replacement
     * @return DocumentAssemblerResponse containing the processed document
     * @throws IOException If there's an error reading or writing files
     */
    public static DocumentAssemblerResponse readDocxFile(Document document, List<Map<String, String>> dataList) throws IOException {
        // Convert Document to XML string
        String originalXml = documentToString(document);

        // Get the file path from the document's base URI or use a default path
        String filePath = document.getBaseURI();
        if (filePath == null || filePath.isEmpty() || filePath.equals("null")) {
            filePath = "input\\file"; // Default path if not available
        }

        try (FileInputStream fis = new FileInputStream(new File(filePath));
             ZipInputStream zis = new ZipInputStream(fis)) {

            // We still need to process the ZIP file for other entries if needed
            // But we already have the document.xml content from the Document object
            // This loop is kept for compatibility with the original method
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                // Skip document.xml as we already have it
                if (!entry.getName().equals("word/document.xml")) {
                    // Process other entries if needed
                }
            }

            // Create a DocumentAssemblerResponse object to return
            DocumentAssemblerResponse response = new DocumentAssemblerResponse();
            response.templateError = false; // Default to no error

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

                // Create the output DOCX file with all processed records and get the byte array
                byte[] docxData = createDocxFile(filePath, removeXmlTags(modifiedXml));

                // File saving removed as per requirements

                // Create a WmlDocument and set the byte array
                WmlDocument wmlDoc = new WmlDocument(null, null);
                wmlDoc.setByteArray(docxData);

                // Set the WmlDocument in the response
                response.wmlDocument = wmlDoc;
            } else {
                // If we couldn't find the document.xml, set templateError to true
                response.templateError = true;
            }

            return response;
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
     * Creates a new DOCX file with the replaced placeholders and returns it as a byte array.
     * This method takes the original DOCX file as a template, replaces the document.xml
     * content with the modified XML, and returns the resulting DOCX file as a byte array.
     *
     * @param originalFilePath Path to the original DOCX template file
     * @param modifiedXml The modified XML content with placeholders replaced by actual values
     * @return The byte array containing the generated DOCX file
     * @throws IOException If there's an error reading the original file or writing the new file
     */
    private static byte[] createDocxFile(String originalFilePath, String modifiedXml) throws IOException {
        // Create an in-memory output stream to hold the DOCX data
        try (ByteArrayOutputStream baos = new ByteArrayOutputStream();
             ZipOutputStream zos = new ZipOutputStream(baos);
             FileInputStream fis = new FileInputStream(new File(originalFilePath));
             ZipInputStream zis = new ZipInputStream(fis)) {

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

            // Close the ZipOutputStream to ensure all data is written
            zos.close();

            // Get the byte array from the ByteArrayOutputStream
            return baos.toByteArray();
        }
        // Resources are automatically closed by try-with-resources
    }

    /**
     * Creates a new DOCX file on disk from the WmlDocument.
     * This method saves the WmlDocument data to a file named "outputFinal.docx".
     *
     * @param wmlDocument The WmlDocument containing the DOCX file data
     * @throws IOException If there's an error writing the file
     */
    public static void saveResponseWmlDocument(WmlDocument wmlDocument) throws IOException {
        if (wmlDocument == null) {
            System.out.println("Error: WmlDocument is null");
            return;
        }

        // Get the byte array from the WmlDocument
        byte[] docxData = wmlDocument.DocumentByteArray();

        // Get the current working directory
        String currentDir = System.getProperty("user.dir");
        // Ensure the src directory exists
        File srcDir = new File(currentDir, "src");
        if (!srcDir.exists()) {
            srcDir.mkdir();
        }

        // Generate the output filename
        String fileName = "outputFinal.docx";
        File outputFile = new File(srcDir, fileName);

        // Write the byte array to the file
        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
            fos.write(docxData);
        }

        System.out.println("File saved successfully: " + outputFile.getAbsolutePath());
    }

    public static class DocumentAssemblerResponse {
        public WmlDocument wmlDocument;
        // out Boolean templateError
        public Boolean templateError;
    }

    // TODO konv toto je docasne, kym zistime, na co to vlastne je
    public static class WmlDocument {
        private byte[] data = new byte[0];
        public WmlDocument(Object a, Object b) {

        }

        public void setByteArray(byte[] new_data) {
            data = new_data;
        }

        public byte[] DocumentByteArray() {
            return data;
        }
    }

    /**
     * Creates an org.w3c.dom.Document from a file path.
     * This method reads the file at the given path and parses it into a DOM Document.
     * It handles different file types including XML files and DOCX files (which are ZIP archives containing XML).
     * It also handles potential BOM (Byte Order Mark), malformed XML tags, or other non-XML content.
     *
     * @param filePath The path to the file
     * @return The parsed DOM Document
     * @throws ParserConfigurationException If there's an error creating the DocumentBuilder
     * @throws SAXException If there's an error parsing the XML
     * @throws IOException If there's an error reading the file
     */
    public static Document createDocumentFromFilePath(String filePath) throws ParserConfigurationException, SAXException, IOException {
        // Create a DocumentBuilderFactory
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();

        // Create a DocumentBuilder
        DocumentBuilder builder = factory.newDocumentBuilder();

        // Read the file content
        File file = new File(filePath);
        String content = null;

        // Check if the file is a DOCX file (which is a ZIP archive)
        if (filePath.toLowerCase().endsWith(".docx") || isZipFile(file)) {
            // Extract the document.xml from the DOCX file
            try (FileInputStream fis = new FileInputStream(file);
                 ZipInputStream zis = new ZipInputStream(fis)) {

                ZipEntry entry;
                while ((entry = zis.getNextEntry()) != null) {
                    if (entry.getName().equals("word/document.xml")) {
                        // Found the main document XML content
                        content = readInputStream(zis);
                        break;
                    }
                }

                if (content == null) {
                    throw new IOException("Could not find word/document.xml in the DOCX file");
                }
            }
        } else {
            // Handle as a regular file
            try {
                // First try the standard way
                Document document = builder.parse(file);
                document.getDocumentElement().normalize();
                return document;
            } catch (SAXException e) {
                // Read the file into a string for further processing
                try (FileInputStream fis = new FileInputStream(file)) {
                    byte[] bytes = new byte[(int) file.length()];
                    fis.read(bytes);
                    content = new String(bytes, StandardCharsets.UTF_8);
                }
            }
        }

        // Process the content string
        if (content != null) {
            // Find the start of the XML content (skip BOM or other non-XML content)
            int xmlStart = content.indexOf("<?xml");
            if (xmlStart == -1) {
                xmlStart = content.indexOf("<");
            }

            if (xmlStart > 0) {
                // Remove content before the XML declaration
                content = content.substring(xmlStart);
            }

            // Check if the content is actually XML
            if (!content.trim().startsWith("<?xml") && !content.trim().startsWith("<")) {
                throw new SAXException("Not a valid XML content");
            }

            // Ensure proper XML prolog if needed
            if (!content.trim().startsWith("<?xml")) {
                // If there's no XML declaration but content starts with a tag,
                // add a proper XML declaration
                content = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" + content;
            }

            // Fix common XML issues
            content = fixMalformedXml(content);

            try {
                // Try to parse the cleaned content
                Document document = builder.parse(new org.xml.sax.InputSource(
                        new java.io.StringReader(content)));
                document.getDocumentElement().normalize();
                return document;
            } catch (SAXException innerEx) {
                // If still failing, try more aggressive cleaning
                content = removeProblematicTags(content);
                Document document = builder.parse(new org.xml.sax.InputSource(
                        new java.io.StringReader(content)));
                document.getDocumentElement().normalize();
                return document;
            }
        }

        throw new SAXException("Failed to parse file as XML: " + filePath);
    }

    /**
     * Checks if a file is a ZIP archive.
     * 
     * @param file The file to check
     * @return true if the file is a ZIP archive, false otherwise
     */
    private static boolean isZipFile(File file) {
        try (FileInputStream fis = new FileInputStream(file)) {
            byte[] signature = new byte[4];
            if (fis.read(signature) != 4) {
                return false;
            }
            // ZIP file signature: 50 4B 03 04
            return signature[0] == 0x50 && signature[1] == 0x4B && 
                   signature[2] == 0x03 && signature[3] == 0x04;
        } catch (IOException e) {
            return false;
        }
    }

    /**
     * Fixes common malformed XML issues, particularly focusing on unclosed tags
     * and invalid characters.
     * This method addresses issues with malformed tags and removes invalid XML characters.
     *
     * @param content The XML content to fix
     * @return The fixed XML content
     */
    private static String fixMalformedXml(String content) {
        // Remove invalid XML characters (including Unicode control characters)
        StringBuilder cleanContent = new StringBuilder();
        for (int i = 0; i < content.length(); i++) {
            char c = content.charAt(i);
            // XML 1.0 allows only these characters:
            // #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
            if (c == 0x9 || c == 0xA || c == 0xD || (c >= 0x20 && c <= 0xD7FF) ||
                (c >= 0xE000 && c <= 0xFFFD)) {
                cleanContent.append(c);
            }
            // Skip any other characters (including Unicode 0x10 mentioned in the error)
        }
        content = cleanContent.toString();

        // Fix malformed <u> tags (ensure they have proper closing)
        content = content.replaceAll("<u(?![>\\s/])", "<u>");
        content = content.replaceAll("<u\\s+(?!>|/|\\w+=)", "<u>");

        // Ensure all opening <u> tags have corresponding closing tags
        int openCount = countOccurrences(content, "<u>");
        int closeCount = countOccurrences(content, "</u>");

        if (openCount > closeCount) {
            // Add missing closing tags
            StringBuilder sb = new StringBuilder(content);
            for (int i = 0; i < openCount - closeCount; i++) {
                sb.append("</u>");
            }
            content = sb.toString();
        }

        return content;
    }

    /**
     * Removes problematic XML tags that might be causing parsing errors.
     * This is a more aggressive approach when other fixes fail.
     *
     * @param content The XML content to clean
     * @return The cleaned XML content
     */
    private static String removeProblematicTags(String content) {
        // Remove problematic <u> tags completely
        content = content.replaceAll("<u[^>]*>", "");
        content = content.replaceAll("</u>", "");

        // Remove other potentially problematic inline formatting tags
        String[] problematicTags = {"b", "i", "em", "strong"};
        for (String tag : problematicTags) {
            content = content.replaceAll("<" + tag + "[^>]*>", "");
            content = content.replaceAll("</" + tag + ">", "");
        }

        return content;
    }

    /**
     * Counts the number of occurrences of a substring within a string.
     *
     * @param content The string to search in
     * @param substring The substring to count
     * @return The number of occurrences
     */
    private static int countOccurrences(String content, String substring) {
        int count = 0;
        int index = 0;
        while ((index = content.indexOf(substring, index)) != -1) {
            count++;
            index += substring.length();
        }
        return count;
    }

    /**
     * Converts a Document object to its XML string representation.
     *
     * @param document The Document object to convert
     * @return The XML string representation of the document
     */
    private static String documentToString(Document document) {
        try {
            javax.xml.transform.TransformerFactory tf = javax.xml.transform.TransformerFactory.newInstance();
            javax.xml.transform.Transformer transformer = tf.newTransformer();
            transformer.setOutputProperty(javax.xml.transform.OutputKeys.OMIT_XML_DECLARATION, "no");
            transformer.setOutputProperty(javax.xml.transform.OutputKeys.METHOD, "xml");
            transformer.setOutputProperty(javax.xml.transform.OutputKeys.INDENT, "yes");
            transformer.setOutputProperty(javax.xml.transform.OutputKeys.ENCODING, "UTF-8");

            java.io.StringWriter writer = new java.io.StringWriter();
            transformer.transform(new javax.xml.transform.dom.DOMSource(document),
                                 new javax.xml.transform.stream.StreamResult(writer));
            return writer.getBuffer().toString();
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
}