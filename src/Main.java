import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.StringReader;
import java.util.List;
import java.util.regex.Pattern;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.w3c.dom.Node;
import org.xml.sax.InputSource;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFSDT;
import org.apache.poi.xwpf.usermodel.ISDTContent;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class Main {
    // Pattern to match valid XML tag name characters
    private static final Pattern INVALID_XML_CHARS = Pattern.compile("[^a-zA-Z0-9._-]");

    /**
     * Sanitizes a string to be used as an XML tag name.
     * Replaces invalid characters with underscores and ensures it starts with a letter or underscore.
     * 
     * @param tagName the original tag name
     * @return a valid XML tag name
     */
    private static String sanitizeXmlTagName(String tagName) {
        if (tagName == null || tagName.isEmpty()) {
            return "_empty";
        }

        // Replace invalid characters with underscores
        String sanitized = INVALID_XML_CHARS.matcher(tagName).replaceAll("_");

        // Ensure the tag starts with a letter or underscore (XML requirement)
        char firstChar = sanitized.charAt(0);
        if (!Character.isLetter(firstChar) && firstChar != '_') {
            sanitized = "_" + sanitized;
        }

        return sanitized;
    }

    /**
     * Attempts to parse a string as XML and returns the root node.
     * 
     * @param xmlString the XML string to parse
     * @param dBuilder the DocumentBuilder to use
     * @return the parsed XML node, or null if parsing failed
     */
    private static Node parseXmlString(String xmlString, DocumentBuilder dBuilder) {
        try {
            // Parse the XML string
            Document doc = dBuilder.parse(new InputSource(new StringReader(xmlString)));
            return doc.getDocumentElement();
        } catch (Exception e) {
            // If parsing fails, return null
            return null;
        }
    }

    public static void main(String[] args) throws Exception {
        // Configure Log4j to use SimpleLogger to avoid the warning about missing log4j-core
        System.setProperty("org.apache.logging.log4j.simplelog.StatusLogger.level", "OFF");

        String inputFile = "src\\file.docx";
        String outputFile = "C:\\Users\\mbugar\\adresne-stitky\\output\\output.xml";

        if (args.length >= 2) {
            inputFile = args[0];
            outputFile = args[1];
        } else if (args.length == 1) {
            inputFile = args[0];
        }

        convert(inputFile, outputFile);
    }

    public static void convert(String docxPath, String xmlPath) throws Exception {
        try (FileInputStream fis = new FileInputStream(docxPath);
             XWPFDocument document = new XWPFDocument(fis)) {

            // Setup XML document
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document xmlDoc = dBuilder.newDocument();
            Element root = xmlDoc.createElement("Root");
            xmlDoc.appendChild(root);

            // Iterate through body elements to find content controls
            List<IBodyElement> elements = document.getBodyElements();
            for (IBodyElement element : elements) {
                if (element instanceof XWPFSDT) {
                    XWPFSDT sdt = (XWPFSDT) element;
                    String tag = sdt.getTag();
                    if (tag == null) continue;

                    // Handle Repeat blocks
                    if (tag.startsWith("Repeat:")) {
                        String repeatName = tag.substring("Repeat:".length());
                        String sanitizedRepeatName = sanitizeXmlTagName(repeatName);
                        Element repeatElem = xmlDoc.createElement(sanitizedRepeatName);
                        root.appendChild(repeatElem);

                        // Process nested content controls
                        ISDTContent content = sdt.getContent();
                        // In newer versions of Apache POI, we need to handle content differently
                        // Since we can't directly access body elements, we'll use the text content
                        String contentText = content.getText();
                        Element cont = xmlDoc.createElement("Content");
                        cont.setTextContent(contentText);
                        repeatElem.appendChild(cont);

                    } else {
                        // Simple content control
                        String sanitizedTag = sanitizeXmlTagName(tag);
                        Element elem = xmlDoc.createElement(sanitizedTag);

                        // Get the content text
                        String contentText = sdt.getContent().getText();

                        // Check if the content text looks like an XML tag
                        if (contentText.trim().startsWith("<") && contentText.trim().endsWith(">")) {
                            try {
                                // Try to parse it as XML
                                Node xmlNode = parseXmlString(contentText, dBuilder);
                                if (xmlNode != null) {
                                    // Successfully parsed as XML, import the node into our document
                                    Node importedNode = xmlDoc.importNode(xmlNode, true);
                                    elem.appendChild(importedNode);
                                } else {
                                    // Parsing failed, treat as regular text
                                    Element textHolder = xmlDoc.createElement("value");
                                    textHolder.setTextContent(contentText);
                                    elem.appendChild(textHolder);
                                }
                            } catch (Exception e) {
                                // Any error, treat as regular text
                                Element textHolder = xmlDoc.createElement("value");
                                textHolder.setTextContent(contentText);
                                elem.appendChild(textHolder);
                            }
                        } else {
                            // Regular text content
                            Element textHolder = xmlDoc.createElement("value");
                            textHolder.setTextContent(contentText);
                            elem.appendChild(textHolder);
                        }

                        root.appendChild(elem);
                    }
                }
            }

            // Write XML to file
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            DOMSource source = new DOMSource(xmlDoc);

            // Ensure output directory exists
            File outputFile = new File(xmlPath);
            File outputDir = outputFile.getParentFile();
            if (outputDir != null && !outputDir.exists()) {
                outputDir.mkdirs();
            }

            try (FileOutputStream fos = new FileOutputStream(xmlPath)) {
                transformer.transform(source, new StreamResult(fos));
            }

            System.out.println("Conversion completed: " + xmlPath);
        }
    }
}
