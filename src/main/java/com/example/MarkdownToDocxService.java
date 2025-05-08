package com.example;

import com.vladsch.flexmark.docx.converter.DocxRenderer;
import com.vladsch.flexmark.parser.Parser;
import com.vladsch.flexmark.util.ast.Node;
import com.vladsch.flexmark.util.data.MutableDataSet;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.math.BigInteger;

import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

public class MarkdownToDocxService {

    public void convertMarkdownToDocx(String inputMarkdown, String outputDocxFile)
            throws IOException, org.docx4j.openpackaging.exceptions.InvalidFormatException,
            org.docx4j.openpackaging.exceptions.Docx4JException {
        // Configure the parser and renderer
        MutableDataSet options = new MutableDataSet();
        InputStream templateInputStream = new FileInputStream("src/main/resources/empty.xml");

        Parser parser = Parser.builder(options).build();
        DocxRenderer renderer = DocxRenderer.builder(options).build();

        // Parse and render the markdown to DOCX
        Node document = parser.parse(inputMarkdown);

        // Load the template file as a WordprocessingMLPackage
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(templateInputStream);

        // Ensure the template's styles are applied
        if (wordMLPackage.getMainDocumentPart().getStyleDefinitionsPart() != null) {
            org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart templateStyles = wordMLPackage
                    .getMainDocumentPart().getStyleDefinitionsPart();
            wordMLPackage.getMainDocumentPart().addTargetPart(templateStyles);
        } else {
            System.out.println("No styles found in the template.");
        }

        // Render the document to the WordprocessingMLPackage
        renderer.render(document, wordMLPackage);

        // Save the WordprocessingMLPackage to the output file
        wordMLPackage.save(new File(outputDocxFile));

        System.out.println("DOCX file generated successfully: " + outputDocxFile);
    }

    public void convertMarkdownToDocx(Path inputMarkdownFile, String outputDocxFile)
            throws IOException, org.docx4j.openpackaging.exceptions.InvalidFormatException,
            org.docx4j.openpackaging.exceptions.Docx4JException {
        // Read the markdown content
        String markdown = new String(Files.readAllBytes(inputMarkdownFile), StandardCharsets.UTF_8);

        // Delegate to the overridden method
        try {
            convertMarkdownToDocx(markdown, outputDocxFile);
        } catch (IOException | Docx4JException e) {
            e.printStackTrace();
        }
    }

    public void standardizeDocxFile(String inputDocxFile) throws IOException, Docx4JException {
        // Load the DOCX file
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(inputDocxFile));
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

        // Get the tables in the document
        List<Object> tables = null;
        try {
            tables = documentPart.getJAXBNodesViaXPath("//w:tbl", true);
        } catch (XPathBinderAssociationIsPartialException e) {
            System.out.println(e);
        } catch (javax.xml.bind.JAXBException e) {
            System.out.println(e);
        }

        if (tables == null || tables.isEmpty()) {
            System.out.println("No tables found in the document.");
            return;
        }

        try {

            // Extract the actual value from JAXBElement before casting
            Tbl table = (Tbl) ((javax.xml.bind.JAXBElement<?>) tables.get(0)).getValue();

            // Get the rows of the table
            List<Object> rows = table.getContent();

            if (rows.isEmpty()) {
                System.out.println("The table is empty.");
                return;
            }

            // Get the first row
            Tr firstRow = (Tr) rows.get(0);

            // Extract the actual value from JAXBElement before casting
            List<Object> cells = new java.util.ArrayList<>();
            for (Object cell : firstRow.getContent()) {
                if (cell instanceof javax.xml.bind.JAXBElement) {
                    Object value = ((javax.xml.bind.JAXBElement<?>) cell).getValue();
                    if (value instanceof Tc) {
                        cells.add(value);
                    }
                }
            }

            if (cells.size() < 2) {
                System.out.println("The first row does not have enough cells to combine.");
                return;
            }

            // Get the value of the first cell
            Tc firstCell = (Tc) cells.get(0);
            String firstCellValue = "";
            for (Object content : firstCell.getContent()) {
                firstCellValue += content.toString();
            }

            // Remove the second cell from the firstRow's content
            Object secondCell = firstRow.getContent().get(1);
            firstRow.getContent().remove(secondCell);

            // Merge the first cell to span two columns
            TcPr tcPr = firstCell.getTcPr();
            if (tcPr == null) {
                tcPr = new TcPr();
                firstCell.setTcPr(tcPr);
            }
            TcPrInner.GridSpan gridSpan = new TcPrInner.GridSpan();
            gridSpan.setVal(BigInteger.valueOf(2));
            tcPr.setGridSpan(gridSpan);

            // Set the value of the first cell
            firstCell.getContent().clear();
            firstCell.getContent().add(documentPart.createParagraphOfText(firstCellValue));

            // Set black background color for the first row
            for (Object cell : cells) {
                Tc tc = (Tc) cell;
                TcPr tcPrCell = tc.getTcPr();
                if (tcPrCell == null) {
                    tcPrCell = new TcPr();
                    tc.setTcPr(tcPrCell);
                }

                CTShd shd = new CTShd();
                shd.setFill("000000"); // Black color
                tcPrCell.setShd(shd);

                // Set font color to white and make text bold
                for (Object content : tc.getContent()) {
                    if (content instanceof P) {
                        P paragraph = (P) content;
                        List<Object> paragraphContents = paragraph.getContent();
                        for (Object paragraphContent : paragraphContents) {
                            if (paragraphContent instanceof R) {
                                R run = (R) paragraphContent;
                                RPr rPr = run.getRPr();
                                if (rPr == null) {
                                    rPr = new RPr();
                                    run.setRPr(rPr);
                                }

                                // Set font color to white
                                Color color = new Color();
                                color.setVal("FFFFFF"); // White color
                                rPr.setColor(color);

                                // Make text bold
                                BooleanDefaultTrue b = new BooleanDefaultTrue();
                                rPr.setB(b);
                            }
                        }
                    }
                }
            }

            // Set background color and bold font for the first cell of each row except the first row
            for (int i = 1; i < rows.size(); i++) { // Start from the second row
                Tr row = (Tr) rows.get(i);
                if (!row.getContent().isEmpty()) {
                    Object firstCellObject = row.getContent().get(0);
                    if (firstCellObject instanceof javax.xml.bind.JAXBElement) {
                        Object value = ((javax.xml.bind.JAXBElement<?>) firstCellObject).getValue();
                        if (value instanceof Tc) {
                            Tc firstCellInRow = (Tc) value; // Renamed to avoid conflict
                            TcPr firstCellTcPr = firstCellInRow.getTcPr(); // Renamed to avoid conflict
                            if (firstCellTcPr == null) {
                                firstCellTcPr = new TcPr();
                                firstCellInRow.setTcPr(firstCellTcPr);
                            }

                            // Set background color to light gray
                            CTShd shd = new CTShd();
                            shd.setFill("F2F2F2"); // Light gray color
                            firstCellTcPr.setShd(shd);

                            // Set bold font for the text in the first cell
                            for (Object content : firstCellInRow.getContent()) {
                                if (content instanceof P) {
                                    P paragraph = (P) content;
                                    List<Object> paragraphContents = paragraph.getContent();
                                    for (Object paragraphContent : paragraphContents) {
                                        if (paragraphContent instanceof R) {
                                            R run = (R) paragraphContent;
                                            RPr rPr = run.getRPr();
                                            if (rPr == null) {
                                                rPr = new RPr();
                                                run.setRPr(rPr);
                                            }

                                            // Make text bold
                                            BooleanDefaultTrue b = new BooleanDefaultTrue();
                                            rPr.setB(b);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Process the second and third tables
            for (int tableIndex = 1; tableIndex <= 2; tableIndex++) {
                if (tables.size() > tableIndex) {
                    Tbl currentTable = (Tbl) ((javax.xml.bind.JAXBElement<?>) tables.get(tableIndex)).getValue();

                    // Get the rows of the table
                    List<Object> tableRows = currentTable.getContent();

                    if (!tableRows.isEmpty()) {
                        // Get the first row of the table
                        Tr currentFirstRow = (Tr) tableRows.get(0);

                        // Set black background color for the first row of the table
                        for (Object cell : currentFirstRow.getContent()) {
                            if (cell instanceof javax.xml.bind.JAXBElement) {
                                Object value = ((javax.xml.bind.JAXBElement<?>) cell).getValue();
                                if (value instanceof Tc) {
                                    Tc tc = (Tc) value;
                                    TcPr tcPrCell = tc.getTcPr();
                                    if (tcPrCell == null) {
                                        tcPrCell = new TcPr();
                                        tc.setTcPr(tcPrCell);
                                    }

                                    CTShd shd = new CTShd();
                                    shd.setFill("000000"); // Black color
                                    tcPrCell.setShd(shd);

                                    // Set font color to white and make text bold
                                    for (Object content : tc.getContent()) {
                                        if (content instanceof P) {
                                            P paragraph = (P) content;
                                            List<Object> paragraphContents = paragraph.getContent();
                                            for (Object paragraphContent : paragraphContents) {
                                                if (paragraphContent instanceof R) {
                                                    R run = (R) paragraphContent;
                                                    RPr rPr = run.getRPr();
                                                    if (rPr == null) {
                                                        rPr = new RPr();
                                                        run.setRPr(rPr);
                                                    }

                                                    // Set font color to white
                                                    Color color = new Color();
                                                    color.setVal("FFFFFF"); // White color
                                                    rPr.setColor(color);

                                                    // Make text bold
                                                    BooleanDefaultTrue b = new BooleanDefaultTrue();
                                                    rPr.setB(b);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Merge the first row of the third table into a single cell
            List<Object> thirdTableRows = new java.util.ArrayList<>();
            if (tables.size() > 2) {
                Tbl thirdTable = (Tbl) ((javax.xml.bind.JAXBElement<?>) tables.get(2)).getValue();

                // Get the rows of the third table
                thirdTableRows = thirdTable.getContent();

                if (!thirdTableRows.isEmpty()) {
                    // Get the first row of the third table
                    Tr firstRowThirdTable = (Tr) thirdTableRows.get(0);

                    // Extract the actual value from JAXBElement before casting
                    List<Object> thirdTableCells = new java.util.ArrayList<>();
                    for (Object cell : firstRowThirdTable.getContent()) {
                        if (cell instanceof javax.xml.bind.JAXBElement) {
                            Object value = ((javax.xml.bind.JAXBElement<?>) cell).getValue();
                            if (value instanceof Tc) {
                                thirdTableCells.add(value);
                            }
                        }
                    }

                    if (thirdTableCells.size() < 4) {
                        System.out.println("The first row of the third table does not have enough cells to merge.");
                        return;
                    }

                    // Get the value of all cells in the first row
                    StringBuilder mergedCellValue = new StringBuilder();
                    for (Object cell : thirdTableCells) {
                        Tc tc = (Tc) cell;
                        for (Object content : tc.getContent()) {
                            mergedCellValue.append(content.toString());
                        }
                    }

                    // Remove all cells except the first one
                    for (int i = 1; i < thirdTableCells.size(); i++) {
                        firstRowThirdTable.getContent().remove(thirdTableCells.get(i));
                    }

                    // Merge the first cell to span four columns
                    Tc firstCellThirdTable = (Tc) thirdTableCells.get(0);
                    TcPr tcPrThirdTable = firstCellThirdTable.getTcPr();
                    if (tcPrThirdTable == null) {
                        tcPrThirdTable = new TcPr();
                        firstCellThirdTable.setTcPr(tcPrThirdTable);
                    }
                    TcPrInner.GridSpan gridSpanThirdTable = new TcPrInner.GridSpan();
                    gridSpanThirdTable.setVal(BigInteger.valueOf(4));
                    tcPrThirdTable.setGridSpan(gridSpanThirdTable);

                    // Set the value of the merged cell
                    firstCellThirdTable.getContent().clear();
                    firstCellThirdTable.getContent().add(documentPart.createParagraphOfText(mergedCellValue.toString()));
                }
            }

            // Set font to bold in the first and second rows of the third table
            for (int rowIndex = 0; rowIndex < Math.min(2, thirdTableRows.size()); rowIndex++) {
                Tr row = (Tr) thirdTableRows.get(rowIndex);
                for (Object cell : row.getContent()) {
                    if (cell instanceof javax.xml.bind.JAXBElement) {
                        Object value = ((javax.xml.bind.JAXBElement<?>) cell).getValue();
                        if (value instanceof Tc) {
                            Tc tc = (Tc) value;
                            for (Object content : tc.getContent()) {
                                if (content instanceof P) {
                                    P paragraph = (P) content;
                                    for (Object paragraphContent : paragraph.getContent()) {
                                        if (paragraphContent instanceof R) {
                                            R run = (R) paragraphContent;
                                            RPr rPr = run.getRPr();
                                            if (rPr == null) {
                                                rPr = new RPr();
                                                run.setRPr(rPr);
                                            }
                                            BooleanDefaultTrue b = new BooleanDefaultTrue();
                                            rPr.setB(b);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Save the changes
            wordMLPackage.save(new File(inputDocxFile));

            System.out.println("Standardized the DOCX file: " + inputDocxFile);
        } catch (Exception e) {
            // TODO: handle exception
            System.out.println("Error while standardizing the DOCX file: " + e.getMessage());
        }
    }
}
