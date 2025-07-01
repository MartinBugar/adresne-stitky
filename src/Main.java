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

            dataList.add(rec1);
            dataList.add(rec2);


            // Read the DOCX file, replace placeholders, and export to XML file
            readDocxFile(filePath, dataList);


        } catch (Exception e) {
            System.err.println("Error processing file: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private static void transformFile() {


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

            System.out.println("Processing all datasets in a single document");
            System.out.println("DOCX file contents:");

            ZipEntry entry;
            String originalXml = null;
            while ((entry = zis.getNextEntry()) != null) {
                // Look for the main document content
                if (entry.getName().equals("word/document.xml")) {
                    // Read the XML content
                    originalXml = readInputStream(zis);

                    // Extract text from XML
                    List<String> textLines = extractTextFromXml(originalXml);

                    // Print all lines
                    for (String line : textLines) {
                        System.out.println(line);
                    }
                }
            }

            // Process all datasets and create a single document
            if (originalXml != null) {
                String modifiedXml = originalXml;
                Map<String, String> allReplacedValues = new LinkedHashMap<>();

                // Process each dataset in the dataList
                for (int i = 0; i < dataList.size(); i++) {
                    System.out.println("Processing dataset " + (i + 1) + " of " + dataList.size());

                    // Create a map to collect replaced values for this dataset
                    Map<String, String> datasetReplacedValues = new LinkedHashMap<>();

                    // Find and replace variable placeholders for the current dataset
                    modifiedXml = replacePlaceholders(modifiedXml, dataList.get(i), i, datasetReplacedValues);

                    // Add dataset number prefix to keys to avoid overwriting values from different datasets
                    for (Map.Entry<String, String> mapEntry : datasetReplacedValues.entrySet()) {
                        allReplacedValues.put("Dataset " + (i + 1) + " - " + mapEntry.getKey(), mapEntry.getValue());
                    }

                    // Print the modified XML content
                    System.out.println("\nModified XML content after dataset " + (i + 1) + ":");
                    System.out.println(modifiedXml);

                    // If this is not the last dataset, we need to duplicate the template for the next dataset
                    if (i < dataList.size() - 1) {
                        // Find the end of the document body to insert a vertical spacing and duplicate the template
                        int bodyEndPos = modifiedXml.lastIndexOf("</w:body>");
                        if (bodyEndPos != -1) {
                            // Insert a paragraph break instead of a page break to stack labels vertically
                            String paragraphBreak = "<w:p><w:r><w:t xml:space=\"preserve\">&#10;</w:t></w:r></w:p>";
                            String beforeBodyEnd = modifiedXml.substring(0, bodyEndPos);
                            String afterBodyEnd = modifiedXml.substring(bodyEndPos);

                            // Get the original template content (we'll use the original XML for the next dataset)
                            modifiedXml = beforeBodyEnd + paragraphBreak + originalXml.substring(
                                originalXml.indexOf("<w:body>") + 8, 
                                originalXml.lastIndexOf("</w:body>")
                            ) + afterBodyEnd;
                        }
                    }
                }

                // Export all replaced values to a single TXT file
                exportReplacedValuesToTxt(allReplacedValues, -1);

                // Create a single DOCX file with all datasets
                createDocxFile(filePath, modifiedXml, -1);
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
     * @param index      Index of the dataset being processed
     * @param outReplacedValues Optional map to collect replaced values (if null, values are exported to a file)
     * @return The modified XML content with placeholders replaced
     */
    private static String replacePlaceholders(String xmlContent, Map<String, String> input, int index, Map<String, String> outReplacedValues) {
        String modifiedXml = xmlContent;
        Map<String, String> replacedValues = outReplacedValues != null ? outReplacedValues : new LinkedHashMap<>();

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
                        // Store the replaced value for later export to TXT file
                        replacedValues.put(variableName, replacementValue);

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

        // Export the replaced values to a TXT file only if outReplacedValues is null
        if (outReplacedValues == null) {
            exportReplacedValuesToTxt(replacedValues, index);
        }

        return modifiedXml;
    }

    /**
     * Exports the replaced values to a TXT file in the src folder.
     *
     * @param replacedValues Map containing the replaced values
     * @param index Index of the dataset being processed (used to create unique filenames)
     */
    private static void exportReplacedValuesToTxt(Map<String, String> replacedValues, int index) {
        try {
            // Get the current working directory
            String currentDir = System.getProperty("user.dir");

            // Create the src directory if it doesn't exist
            File srcDir = new File(currentDir, "src");
            if (!srcDir.exists()) {
                srcDir.mkdir();
            }

            // Create the file path
            String filePath;
            if (index == -1) {
                // If index is -1, it means we're creating a single file with all datasets
                filePath = currentDir + "\\src\\replaced_values_all.txt";
            } else {
                // Otherwise, create a file with a unique name based on the index
                filePath = currentDir + "\\src\\replaced_values_" + (index + 1) + ".txt";
            }

            // Write the replaced values to the file
            try (FileWriter writer = new FileWriter(filePath);
                 BufferedWriter bufferedWriter = new BufferedWriter(writer)) {

                // Write a header
                bufferedWriter.write("Replaced Placeholders:\n");
                bufferedWriter.write("=====================\n\n");

                // Write each replaced value
                for (Map.Entry<String, String> entry : replacedValues.entrySet()) {
                    bufferedWriter.write(entry.getKey() + ": " + entry.getValue() + "\n");
                }

                System.out.println("TXT file with replaced values exported to: " + filePath);
            }
        } catch (IOException e) {
            System.err.println("Error exporting TXT file: " + e.getMessage());
            e.printStackTrace();
        }
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
        // Get the current working directory
        String currentDir = System.getProperty("user.dir");

        // Create the src directory if it doesn't exist
        File srcDir = new File(currentDir, "src");
        if (!srcDir.exists()) {
            srcDir.mkdir();
        }

        // Create the output file path
        String outputFilePath;
        if (index == -1) {
            // If index is -1, it means we're creating a single file with all datasets
            outputFilePath = currentDir + "\\src\\output_all.docx";
        } else {
            // Otherwise, create a file with a unique name based on the index
            outputFilePath = currentDir + "\\src\\output_" + (index + 1) + ".docx";
        }

        try (FileInputStream fis = new FileInputStream(new File(originalFilePath));
             ZipInputStream zis = new ZipInputStream(fis);
             FileOutputStream fos = new FileOutputStream(outputFilePath);
             ZipOutputStream zos = new ZipOutputStream(fos)) {

            ZipEntry entry;
            byte[] buffer = new byte[1024];

            while ((entry = zis.getNextEntry()) != null) {
                // Create a new entry in the output file
                ZipEntry newEntry = new ZipEntry(entry.getName());
                zos.putNextEntry(newEntry);

                // If this is the main document content, write the modified XML
                if (entry.getName().equals("word/document.xml")) {
                    // Write the modified XML content
                    zos.write(modifiedXml.getBytes("UTF-8"));
                } else {
                    // Copy the entry content from the original file
                    int len;
                    while ((len = zis.read(buffer)) > 0) {
                        zos.write(buffer, 0, len);
                    }
                }

                // Close the entry
                zos.closeEntry();
                zis.closeEntry();
            }

            System.out.println("DOCX file with replaced placeholders exported to: " + outputFilePath);
        }
    }


}
