import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import javax.xml.parsers.ParserConfigurationException;
import org.w3c.dom.Document;
import org.xml.sax.SAXException;

/**
 * Main application class that serves as the entry point for the address label generation program.
 * This class uses the LabelGenerator utility to process DOCX templates and generate output documents.
 */
public class Main {

    /**
     * Main entry point of the application.
     * Processes the template file with sample data records and generates output documents.
     *
     * @param args Command line arguments (not used)
     */
    public static void main(String[] args) {
        try {
            // Path to the input file (DOCX template)
            String filePath = "input\\file";
            Document documentFromFilePath = LabelGenerator.createDocumentFromFilePath(filePath);

            // Get sample data records
            List<Map<String, String>> dataList = createSampleRecords();

            // Read the DOCX file, replace placeholders, and get the DocumentAssemblerResponse
            LabelGenerator.DocumentAssemblerResponse response = LabelGenerator.readDocxFile(documentFromFilePath, dataList);

            // You can use the response object here if needed
            if (response.templateError) {
                System.out.println("Error: Template processing failed");
            } else {
                System.out.println("Document processed successfully");
                // Access the document bytes if needed
                // byte[] docBytes = response.wmlDocument.DocumentByteArray();
                // Save the response.wmlDocument to outputFinal.docx
                LabelGenerator.saveResponseWmlDocument(response.wmlDocument);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Creates sample data records for testing the label generation.
     * 
     * @return List of maps containing sample data records
     */
    private static List<Map<String, String>> createSampleRecords() {
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

        return dataList;
    }
}
