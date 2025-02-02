package com.xmlparser;
import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

public class XMLParser {
    public static void main(String[] args) {
        if (args.length != 2) {
            System.out.println("Usage: java XMLParser <input_xml_file> <output_folder>");
            System.exit(1);
        }

        String inputXmlFile = args[0];
        String outputFolder = args[1];

        try {
            // Create parsed folder if it doesn't exist
            File parsedFolder = new File(outputFolder);
            if (!parsedFolder.exists()) {
                parsedFolder.mkdirs();
            }

            // Generate output Excel filename with timestamp
            String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
            String excelFile = outputFolder + File.separator + "parsed_" + timestamp + ".xlsx";

            // Create new Excel workbook
            try (Workbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Parsed Data");
                
                // Create header row
                Row headerRow = sheet.createRow(0);
                headerRow.createCell(0).setCellValue("Title");
                headerRow.createCell(1).setCellValue("Number");
                headerRow.createCell(2).setCellValue("Video Link");

                // Parse XML file
                DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                DocumentBuilder builder = factory.newDocumentBuilder();
                Document doc = builder.parse(new File(inputXmlFile));
                doc.getDocumentElement().normalize();

                // Find all entry elements
                NodeList entryList = doc.getElementsByTagName("entry");
                int rowIndex = 1;

                for (int i = 0; i < entryList.getLength(); i++) {
                    Element entry = (Element) entryList.item(i);
                    
                    // Get title
                    String title = getElementText(entry, "title");
                    
                    // Get content and extract YouTube link
                    String content = getElementText(entry, "content");
                    String videoLink = extractYoutubeLink(content);

                    // Create row and add data
                    Row row = sheet.createRow(rowIndex++);
                    row.createCell(0).setCellValue(title);
                    
                    // Extract number from title if exists
                    String number = extractNumber(title);
                    if (number != null) {
                        row.createCell(1).setCellValue(number);
                    }
                    
                    if (videoLink != null) {
                        row.createCell(2).setCellValue(videoLink);
                    }
                }

                // Auto-size columns
                for (int i = 0; i < 3; i++) {
                    sheet.autoSizeColumn(i);
                }

                // Write to Excel file
                try (FileOutputStream outputStream = new FileOutputStream(excelFile)) {
                    workbook.write(outputStream);
                }

                System.out.println("Successfully parsed XML and created Excel file: " + excelFile);
            }

        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private static String getElementText(Element parent, String tagName) {
        NodeList nodeList = parent.getElementsByTagName(tagName);
        if (nodeList.getLength() > 0) {
            return nodeList.item(0).getTextContent();
        }
        return null;
    }

    private static String extractYoutubeLink(String content) {
        if (content == null) return null;
        
        Pattern pattern = Pattern.compile(
            "(?:https?://)?(?:www\\.)?(?:youtube\\.com/watch\\?v=|youtu\\.be/)([\\w-]+)"
        );
        Matcher matcher = pattern.matcher(content);
        
        if (matcher.find()) {
            return matcher.group(0);
        }
        return null;
    }

    private static String extractNumber(String title) {
        if (title == null) return null;
        
        // First try to get the character before first whitespace
        Pattern firstCharPattern = Pattern.compile("^(\\S)(?=\\s)");
        Matcher firstCharMatcher = firstCharPattern.matcher(title);
        
        if (firstCharMatcher.find()) {
            return firstCharMatcher.group(1);
        }
        
        // If no match found, fall back to finding any number
        Pattern pattern = Pattern.compile("\\d+");
        Matcher matcher = pattern.matcher(title);
        
        if (matcher.find()) {
            return matcher.group(0);
        }
        return null;
    }
} 