import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class Main {
    private static final String[] XML_TAGS_TO_REMOVE = {
        "&lt;EndRepeat/&gt;", "<EndRepeat/>", 
        "&lt;Repeat Select=\"./Label \"/&gt;", "<Repeat Select=\"./Label \"/>"
    };

    public static void main(String[] args) {
        try {
            // Path to the input file
            String filePath = "input\\file";

            List<Map<String, String>> dataList = new ArrayList<>();

            Map<String,String> rec1 = new HashMap<>();
            rec1.put("Label",       "ID-12345");
            rec1.put("VS",          "2025");
            rec1.put("Oslovenie",   "Vážený pán Novák,");
            rec1.put("Adresat",     "Ján Novák");
            rec1.put("AdresnyRiadok1", "Hlavná 10");
            rec1.put("AdresnyRiadok2", "010 01 Žilina");
            rec1.put("Stat",        "Slovensko");

            Map<String,String> rec2 = new HashMap<>();
            rec2.put("Label",       "ID-12345");
            rec2.put("VS",          "2026");
            rec2.put("Oslovenie",   "Vážený pán Hurban,");
            rec2.put("Adresat",     "Ján Hurban");
            rec2.put("AdresnyRiadok1", "Hlavná 25");
            rec2.put("AdresnyRiadok2", "010 01 Martin");
            rec2.put("Stat",        "Madarsko");

            Map<String,String> rec3 = new HashMap<>();
            rec3.put("Label",       "ID-12345");
            rec3.put("VS",          "2026");
            rec3.put("Oslovenie",   "Vážený pán Hurban,");
            rec3.put("Adresat",     "Ján Hurban");
            rec3.put("AdresnyRiadok1", "Hlavná 25");
            rec3.put("AdresnyRiadok2", "010 01 Martin");
            rec3.put("Stat",        "Madarsko");
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
     * Helper method to remove XML tags from a string
     * 
     * @param content String to process
     * @return String with XML tags removed
     */
    private static String removeXmlTags(String content) {
        for (String tag : XML_TAGS_TO_REMOVE) {
            content = content.replace(tag, "");
        }
        return content;
    }

    /**
     * Reads a DOCX file, replaces placeholders with values from the dataList, and prints its content to the console.
     *
     * @param filePath Path to the DOCX file
     * @param dataList List of maps containing key-value pairs for placeholder replacement
     * @throws IOException If there's an error reading the file
     */
    private static void readDocxFile(String filePath, List<Map<String, String>> dataList) throws IOException {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             ZipInputStream zis = new ZipInputStream(fis)) {

            String originalXml = null;
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                if (entry.getName().equals("word/document.xml")) {
                    originalXml = readInputStream(zis);
                }
            }

            if (originalXml != null) {
                String modifiedXml = removeXmlTags(originalXml);
                Map<String, String> allReplacedValues = new LinkedHashMap<>();

                for (int i = 0; i < dataList.size(); i++) {
                    Map<String, String> datasetReplacedValues = new LinkedHashMap<>();
                    modifiedXml = replacePlaceholders(modifiedXml, dataList.get(i), i, datasetReplacedValues);

                    final int datasetIndex = i + 1;
                    datasetReplacedValues.forEach((key, value) -> 
                        allReplacedValues.put("Dataset " + datasetIndex + " - " + key, value));

                    if (i < dataList.size() - 1) {
                        int bodyEndPos = modifiedXml.lastIndexOf("</w:body>");
                        if (bodyEndPos != -1) {
                            String paragraphBreak = "<w:p><w:r><w:t xml:space=\"preserve\">&#10;</w:t></w:r></w:p>";
                            String templateContent = removeXmlTags(originalXml.substring(
                                originalXml.indexOf("<w:body>") + 8, 
                                originalXml.lastIndexOf("</w:body>")
                            ));

                            modifiedXml = modifiedXml.substring(0, bodyEndPos) + 
                                          paragraphBreak + 
                                          templateContent + 
                                          modifiedXml.substring(bodyEndPos);
                        }
                    }
                }

                createDocxFile(filePath, removeXmlTags(modifiedXml), -1);
            }
        }
    }

    /**
     * Reads an input stream into a string without closing the stream.
     */
    private static String readInputStream(InputStream is) throws IOException {
        StringBuilder sb = new StringBuilder();
        BufferedReader reader = new BufferedReader(new InputStreamReader(is, "UTF-8"));
        String line;
        while ((line = reader.readLine()) != null) {
            sb.append(line).append("\n");
        }
        return sb.toString();
    }

    /**
     * Extracts template placeholders from XML by looking for text between <w:t> tags
     * that contains XML-like content (e.g., &lt;Content Select="./VS"/&gt;).
     */
    private static List<String> extractTextFromXml(String xml) {
        List<String> lines = new ArrayList<>();
        int pos = 0;

        while (pos < xml.length()) {
            int startTag = xml.indexOf("<w:t", pos);
            if (startTag == -1) break;

            int contentStart = xml.indexOf(">", startTag);
            if (contentStart == -1) break;
            contentStart++; // Move past the '>'

            int contentEnd = xml.indexOf("</w:t>", contentStart);
            if (contentEnd == -1) break;

            String content = xml.substring(contentStart, contentEnd);

            if (content.contains("&lt;") && content.contains("&gt;")) {
                int placeholderStart = content.indexOf("&lt;");
                int placeholderEnd = content.lastIndexOf("&gt;") + 4;

                if (placeholderStart >= 0 && placeholderEnd > placeholderStart) {
                    String placeholder = content.substring(placeholderStart, placeholderEnd)
                                               .replace("&lt;", "<")
                                               .replace("&gt;", ">");
                    lines.add(placeholder);
                }
            }

            pos = contentEnd + 6; // Length of "</w:t>"
        }

        return lines;
    }

    /**
     * Replaces placeholders in the XML content with values from the input map.
     *
     * @param xmlContent The original XML content
     * @param input      Map containing key-value pairs for placeholder replacement
     * @param index      Index of the dataset being processed
     * @param outReplacedValues Optional map to collect replaced values (if null, values are exported to a file)
     * @return The modified XML content with placeholders replaced
     */
    private static String replacePlaceholders(String xmlContent, Map<String, String> input, int index, Map<String, String> outReplacedValues) {
        String modifiedXml = xmlContent;
        Map<String, String> replacedValues = outReplacedValues != null ? outReplacedValues : new LinkedHashMap<>();

        int pos = 0;
        while (pos < modifiedXml.length()) {
            int startTag = modifiedXml.indexOf("<w:t", pos);
            if (startTag == -1) break;

            int contentStart = modifiedXml.indexOf(">", startTag);
            if (contentStart == -1) break;
            contentStart++; // Move past the '>'

            int contentEnd = modifiedXml.indexOf("</w:t>", contentStart);
            if (contentEnd == -1) break;

            String content = modifiedXml.substring(contentStart, contentEnd);
            String originalContent = content;

            if (content.contains("&lt;") && content.contains("&gt;")) {
                int placeholderStart = content.indexOf("&lt;");
                int placeholderEnd = content.lastIndexOf("&gt;") + 4;

                if (placeholderStart >= 0 && placeholderEnd > placeholderStart) {
                    String placeholder = content.substring(placeholderStart, placeholderEnd);
                    String processedPlaceholder = placeholder.replace("&lt;", "<").replace("&gt;", ">");
                    String variableName = extractVariableName(processedPlaceholder);

                    if (variableName != null && input.containsKey(variableName)) {
                        String replacementValue = input.get(variableName);
                        content = content.replace(placeholder, replacementValue);
                        replacedValues.put(variableName, replacementValue);

                        modifiedXml = modifiedXml.substring(0, contentStart) + content + modifiedXml.substring(contentEnd);
                        contentEnd += (content.length() - originalContent.length());
                    }
                }
            }

            pos = contentEnd + 6; // Length of "</w:t>"
        }

        return removeXmlTags(modifiedXml);
    }

    /**
     * Extracts variable name from a placeholder
     * 
     * @param placeholder The placeholder string
     * @return The extracted variable name or null if not found
     */
    private static String extractVariableName(String placeholder) {
        // Look for Select="./something" pattern
        if (placeholder.contains("Select=\"./")) {
            int selectIndex = placeholder.indexOf("Select=\"");
            int start = selectIndex + 8; // Length of "Select=\"" is 8
            int end = placeholder.indexOf("\"", start);

            if (end != -1) {
                String fullValue = placeholder.substring(start, end);
                if (fullValue.startsWith("./")) {
                    return fullValue.substring(2).trim(); // Remove "./" prefix and trim spaces
                }
            }
        } else {
            // General case for finding "./" patterns
            int start = placeholder.indexOf("./");
            if (start != -1) {
                int end = placeholder.indexOf("/", start + 2);
                if (end != -1) {
                    return placeholder.substring(start + 2, end).trim(); // Remove "./" prefix and trim spaces
                }
            }
        }
        return null;
    }


    /**
     * Creates a DOCX file with the replaced placeholders.
     *
     * @param originalFilePath Path to the original DOCX file
     * @param modifiedXml      The modified XML content with replaced placeholders
     * @param index            Index of the dataset being processed (used to create unique filenames)
     * @throws IOException If there's an error creating the file
     */
    private static void createDocxFile(String originalFilePath, String modifiedXml, int index) throws IOException {
        String currentDir = System.getProperty("user.dir");
        File srcDir = new File(currentDir, "src");
        if (!srcDir.exists()) {
            srcDir.mkdir();
        }

        String fileName = index == -1 ? "output_all.docx" : "output_" + (index + 1) + ".docx";
        File outputFile = new File(srcDir, fileName);

        try (FileInputStream fis = new FileInputStream(new File(originalFilePath));
             ZipInputStream zis = new ZipInputStream(fis);
             FileOutputStream fos = new FileOutputStream(outputFile);
             ZipOutputStream zos = new ZipOutputStream(fos)) {

            ZipEntry entry;
            byte[] buffer = new byte[1024];

            while ((entry = zis.getNextEntry()) != null) {
                ZipEntry newEntry = new ZipEntry(entry.getName());
                zos.putNextEntry(newEntry);

                if (entry.getName().equals("word/document.xml")) {
                    zos.write(modifiedXml.getBytes("UTF-8"));
                } else {
                    int len;
                    while ((len = zis.read(buffer)) > 0) {
                        zos.write(buffer, 0, len);
                    }
                }

                zos.closeEntry();
                zis.closeEntry();
            }

        }
    }


}
