package DocToPDF123.DocxToPdf123;

import org.apache.fop.apps.FOUserAgent;
import org.apache.fop.apps.Fop;
import org.apache.fop.apps.FopFactory;
import org.apache.fop.apps.MimeConstants;
import org.apache.poi.xwpf.usermodel.*;
import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.xml.transform.Result;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.sax.SAXResult;
import javax.xml.transform.stream.StreamSource;
import java.io.*;

public class Demo {
    String PATH = "src/main/resources/temp/";

    public void convertToXslFo() throws Exception {

        File docxFile = new File(PATH + "test.docx");
        File foFile = new File(PATH + "out.fo");

        FileInputStream docxInputStream = new FileInputStream(docxFile);
        OutputStream foOutputStream = new FileOutputStream(foFile);

        XWPFDocument document = new XWPFDocument(docxInputStream);

        StringBuilder foContent = new StringBuilder();

        if (document.getParagraphs() != null) {
            // Process paragraphs
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                processParagraph(foContent, paragraph);

            }
        }


        if (document.getTables() != null) {
            // Process tables
            for (XWPFTable table : document.getTables()) {
                processTable(foContent, table);
            }
        }

        // Generate the final XSL-FO content
        String xslFoContent = generateXslFoContent(foContent);

        // Write the XSL-FO content to the output stream
        foOutputStream.write(xslFoContent.getBytes());
        System.out.println("Completed");

        hede();
    }

    private void processParagraph(StringBuilder foContent, XWPFParagraph paragraph) {

        if (paragraph.getText().equals("")) {

            // Handle empty paragraphs by adding a line break
            foContent.append("<fo:block linefeed-treatment=\"preserve\">&#xA;");
            for (XWPFRun run : paragraph.getRuns()) {
                processRun(foContent, run);
            }

            foContent.append("</fo:block>");
        } else {
            foContent.append("<fo:block");

            // Handle text alignment
            foContent.append(" text-align=\"").append(getTextAlignment(paragraph.getAlignment())).append("\"");

            applyParagraphProperties(foContent, paragraph);

            foContent.append(">");

            boolean isPreviousRunEmpty = false;
            for (XWPFRun run : paragraph.getRuns()) {

                boolean isCurrentRunEmpty = run.text().trim().isEmpty();

                // Check if the current run is empty and the previous run was also empty
                if (isCurrentRunEmpty && isPreviousRunEmpty) {
                    // Add a line break element
                    foContent.append("<fo:block/>");
                }

                processRun(foContent, run);
                isPreviousRunEmpty = isCurrentRunEmpty;

            }

            foContent.append("</fo:block>");


        }
    }

    private void processRun(StringBuilder foContent, XWPFRun run) {
        foContent.append("<fo:inline");

        // Retrieve and apply run properties
        applyRunProperties(foContent, run);

        foContent.append(">");


        String text = run.text();

        // Replace tab characters with XSL-FO representation
        text = text.replace("\t", "&#x9;");

        // Replace space characters with XSL-FO representation
        text = text.replace(" ", "&#x20;");

        foContent.append(text);

        foContent.append("</fo:inline>");
    }

    private void applyRunProperties(StringBuilder foContent, XWPFRun run) {
        CTRPr rPr = run.getCTR().getRPr();
        if (rPr != null) {
            CTFonts fonts = rPr.getRFonts();
            if (fonts != null && fonts.getAscii() != null) {
                foContent.append(" font-family=\"").append(fonts.getAscii()).append("\"");
            }
            CTHpsMeasure fontSize = rPr.getSz();
            if (fontSize != null) {
                foContent.append(" font-size=\"").append(fontSize.getVal().divide(new java.math.BigInteger("2"))).append("pt\"");
            }
            if (rPr.getB() != null) {
                foContent.append(" font-weight=\"bold\"");
            }
            if (rPr.getI() != null) {
                foContent.append(" font-style=\"italic\"");
            }
            if (rPr.getU() != null) {
                foContent.append(" text-decoration=\"underline\"");
            }
            if (rPr.isSetColor()){
                String s = convertHexColorToXslFo(run.getColor());
                foContent.append(" color=\"").append(s).append("\"");
            }

            if (rPr.isSetHighlight()) {
                STHighlightColor.Enum textHightlightColor = run.getTextHightlightColor();
                foContent.append(" background-color=\"").append(textHightlightColor).append("\"");

            }

        }
    }

    private String convertHexColorToXslFo(String hexColor) {
        if (hexColor.startsWith("#")) {
            hexColor = hexColor.substring(1); // Remove the leading #
        }

        // Convert the hex color to RGB format
        int red = Integer.parseInt(hexColor.substring(0, 2), 16);
        int green = Integer.parseInt(hexColor.substring(2, 4), 16);
        int blue = Integer.parseInt(hexColor.substring(4, 6), 16);

        // Format the RGB values as required by XSL-FO

        return String.format("#%02X%02X%02X", red, green, blue);
    }

    private void applyParagraphProperties(StringBuilder foContent, XWPFParagraph paragraph) {
        CTPPr pPr = paragraph.getCTP().getPPr();
        if (pPr != null) {
            CTPBdr bdr = pPr.getPBdr();
            if (bdr != null) {
                foContent.append(" border-width=\"").append(bdr.getTop().getSz().intValue() / 8).append("mm\"");
                foContent.append(" border-style=\"solid\"");
                foContent.append(" border-color=\"").append(bdr.getTop().getColor()).append("\"");
            }
        }
    }

    private void processTable(StringBuilder foContent, XWPFTable table) {
        foContent.append("<fo:table>");

        foContent.append("<fo:table-body>");
        for (XWPFTableRow row : table.getRows()) {
            processTableRow(foContent, row);
        }
        foContent.append("</fo:table-body>");

        foContent.append("</fo:table>");
    }

    private void processTableRow(StringBuilder foContent, XWPFTableRow row) {
        foContent.append("<fo:table-row>");

        for (XWPFTableCell cell : row.getTableCells()) {
            processTableCell(foContent, cell);
        }

        foContent.append("</fo:table-row>");
    }

    private void processTableCell(StringBuilder foContent, XWPFTableCell cell) {
        foContent.append("<fo:table-cell>");

        for (XWPFParagraph paragraph : cell.getParagraphs()) {
            processParagraph(foContent, paragraph);
        }

        foContent.append("</fo:table-cell>");
    }

    private String generateXslFoContent(StringBuilder foContent) {

        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                "<fo:root xmlns:fo=\"http://www.w3.org/1999/XSL/Format\">\n" +
                "<fo:layout-master-set>\n" +
                "<fo:simple-page-master master-name=\"A4\">\n" +
                "<fo:region-body margin=\"1in\"/>\n" +
                "</fo:simple-page-master>\n" +
                "</fo:layout-master-set>\n" +
                "<fo:page-sequence master-reference=\"A4\">\n" +
                "<fo:flow flow-name=\"xsl-region-body\">\n" +
                foContent + "\n" +
                "</fo:flow>\n" +
                "</fo:page-sequence>\n" +
                "</fo:root>";
    }


    public void hede() {
        try {
            // Configure FOP
            FopFactory fopFactory = FopFactory.newInstance(new File(".").toURI());
            FOUserAgent foUserAgent = fopFactory.newFOUserAgent();

            // Create output stream for PDF
            OutputStream out = new BufferedOutputStream(new FileOutputStream(new File(PATH + "out.pdf")));

            // Create FOP's PDF renderer
            Fop fop = fopFactory.newFop(MimeConstants.MIME_PDF, foUserAgent, out);

            // Load XSL-FO file
            File foFile = new File(PATH + "out.fo");
            FileInputStream foStream = new FileInputStream(foFile);

            // Transform XSL-FO to PDF
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();

            StreamSource foSource = new StreamSource(foStream);
            Result pdfResult = new SAXResult(fop.getDefaultHandler());
            transformer.transform(foSource, pdfResult);

            // Close resources
            out.close();
            foStream.close();

            System.out.println("PDF conversion completed successfully!");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private String getTextAlignment(ParagraphAlignment alignment) {
        switch (alignment) {
            case RIGHT:
                return "end";
            case CENTER:
                return "center";
            case BOTH:
                return "justify";
            default:
                return "start";
        }
    }

    public void test() throws Docx4JException, IOException {
        // Load the Word document
        WordprocessingMLPackage wordMLPackage = Docx4J.load(new File("path/to/word/document.docx"));

        // Configure the FO settings
        FOSettings foSettings = Docx4J.createFOSettings();
        foSettings.setWmlPackage(wordMLPackage);

        // Convert the Word document to XSL-FO
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        Docx4J.toFO(foSettings, outputStream, Docx4J.FLAG_EXPORT_PREFER_XSL);

        // Save the XSL-FO content to a file
        File outputFile = new File("path/to/output.xslfo");
        FileOutputStream fileOutputStream = new FileOutputStream(outputFile);
        outputStream.writeTo(fileOutputStream);
        fileOutputStream.close();

        System.out.println("Conversion completed. XSL-FO saved to: " + outputFile.getAbsolutePath());
    }
}
