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

    private void setCellBackgroundColor(Tc cell, String color) {
        TcPr tcPr = cell.getTcPr();
        if (tcPr == null) {
            tcPr = new TcPr();
            cell.setTcPr(tcPr);
        }

        CTShd shd = new CTShd();
        shd.setFill(color);
        tcPr.setShd(shd);
    }

    private void setFontColorAndBold(Tc cell, String color) {
        for (Object content : cell.getContent()) {
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

                        // Set font color
                        Color fontColor = new Color();
                        fontColor.setVal(color);
                        rPr.setColor(fontColor);

                        // Make text bold
                        BooleanDefaultTrue bold = new BooleanDefaultTrue();
                        rPr.setB(bold);
                    }
                }
            }
        }
    }

    private void styleTableFirstRow(Tbl table, String backgroundColor, String fontColor) {
        // Get the rows of the table
        List<Object> rows = table.getContent();

        if (!rows.isEmpty()) {
            // Get the first row of the table
            Tr firstRow = (Tr) rows.get(0);

            // Set background color and font color for the first row
            for (Object cell : firstRow.getContent()) {
                if (cell instanceof javax.xml.bind.JAXBElement) {
                    Object value = ((javax.xml.bind.JAXBElement<?>) cell).getValue();
                    if (value instanceof Tc) {
                        Tc tc = (Tc) value;
                        setCellBackgroundColor(tc, backgroundColor);
                        setFontColorAndBold(tc, fontColor);
                    }
                }
            }
        }
    }

    private void mergeAndStyleFirstRow(Tbl table, MainDocumentPart documentPart, String backgroundColor,
            String fontColor) {
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

        if (cells.size() >= 2) {
            // Get the value of the first cell
            Tc firstCell = (Tc) cells.get(0);
            StringBuilder firstCellValue = new StringBuilder();
            for (Object content : firstCell.getContent()) {
                firstCellValue.append(content.toString());
            }

            // Merge all cells into the first cell
            for (int i = 1; i < cells.size(); i++) {
                Tc cellToMerge = (Tc) cells.get(i);
                for (Object content : cellToMerge.getContent()) {
                    firstCellValue.append(content.toString());
                }
                firstRow.getContent().remove(cells.get(i));
            }

            // Update the first cell to span all columns
            TcPr tcPr = firstCell.getTcPr();
            if (tcPr == null) {
                tcPr = new TcPr();
                firstCell.setTcPr(tcPr);
            }
            TcPrInner.GridSpan gridSpan = new TcPrInner.GridSpan();
            gridSpan.setVal(BigInteger.valueOf(cells.size()));
            tcPr.setGridSpan(gridSpan);

            // Set the value of the first cell
            firstCell.getContent().clear();
            firstCell.getContent().add(documentPart.createParagraphOfText(firstCellValue.toString()));
        }

        // Apply background color and font color to all cells
        for (Object cell : cells) {
            Tc tc = (Tc) cell;
            setCellBackgroundColor(tc, backgroundColor);
            setFontColorAndBold(tc, fontColor);
        }
    }

    private void processFirstTable(Tbl firstTable, MainDocumentPart documentPart) {
        mergeAndStyleFirstRow(firstTable, documentPart, "000000", "FFFFFF"); // Black background, white font
        List<Object> rows = firstTable.getContent();
        for (int i = 1; i < rows.size(); i++) { // Start from the second row
            Tr row = (Tr) rows.get(i);
            if (!row.getContent().isEmpty()) {
                Object cellObj = row.getContent().get(0);
                if (cellObj instanceof javax.xml.bind.JAXBElement) {
                    Object value = ((javax.xml.bind.JAXBElement<?>) cellObj).getValue();
                    if (value instanceof Tc) {
                        Tc firstCell = (Tc) value;
                        setCellBackgroundColor(firstCell, "F2F2F2"); // Light gray background
                        setFontColorAndBold(firstCell, "000000"); // Black font color
                    }
                }
            }
        }
    }

    private void processSecondTable(Tbl secondTable) {
        styleTableFirstRow(secondTable, "000000", "FFFFFF"); // Black background, white font
    }

    private void processThirdTable(Tbl thirdTable, MainDocumentPart documentPart) {
        mergeAndStyleFirstRow(thirdTable, documentPart, "000000", "FFFFFF"); // Black background, white font
        List<Object> secondRows = thirdTable.getContent();
        if (secondRows.size() > 1) {
            Tr secondRow = (Tr) secondRows.get(1);
            for (Object cellObj : secondRow.getContent()) {
                if (cellObj instanceof javax.xml.bind.JAXBElement) {
                    Object value = ((javax.xml.bind.JAXBElement<?>) cellObj).getValue();
                    if (value instanceof Tc) {
                        Tc cell = (Tc) value;
                        setCellBackgroundColor(cell, "F2F2F2"); // Light gray background
                        setFontColorAndBold(cell, "000000"); // Black font color
                    }
                }
            }
        }
    }

    private void styleFirstParagraph(MainDocumentPart documentPart) {
        // Get the first paragraph in the document
        List<Object> content = documentPart.getContent();
        if (content.isEmpty()) {
            System.out.println("The document is empty.");
            return;
        }

        Object firstElement = content.get(0);
        if (firstElement instanceof P) {
            P firstParagraph = (P) firstElement;

            // Apply bold styling
            for (Object paragraphContent : firstParagraph.getContent()) {
                if (paragraphContent instanceof R) {
                    R run = (R) paragraphContent;
                    RPr rPr = run.getRPr();
                    if (rPr == null) {
                        rPr = new RPr();
                        run.setRPr(rPr);
                    }

                    BooleanDefaultTrue bold = new BooleanDefaultTrue();
                    rPr.setB(bold);
                }
            }

            // Center align the paragraph
            PPr pPr = firstParagraph.getPPr();
            if (pPr == null) {
                pPr = new PPr();
                firstParagraph.setPPr(pPr);
            }

            Jc justification = new Jc();
            justification.setVal(JcEnumeration.CENTER);
            pPr.setJc(justification);
        } else {
            System.out.println("The first element is not a paragraph.");
        }
    }

    public void standardizeDocxFile(String inputDocxFile) throws IOException, Docx4JException {
        // Load the DOCX file
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(inputDocxFile));
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

        // Style the first paragraph
        styleFirstParagraph(documentPart);

        // Get the tables in the document
        List<Object> tables = null;
        try {
            tables = documentPart.getJAXBNodesViaXPath("//w:tbl", true);
        } catch (XPathBinderAssociationIsPartialException | javax.xml.bind.JAXBException e) {
            System.out.println(e);
        }

        if (tables == null || tables.isEmpty() || tables.size() < 3) {
            System.out.println("Incorrect number of tables found in the document.");
            return;
        } else {
            try {
                Tbl firstTable = (Tbl) ((javax.xml.bind.JAXBElement<?>) tables.get(0)).getValue();
                processFirstTable(firstTable, documentPart);

                Tbl secondTable = (Tbl) ((javax.xml.bind.JAXBElement<?>) tables.get(1)).getValue();
                processSecondTable(secondTable);

                Tbl thirdTable = (Tbl) ((javax.xml.bind.JAXBElement<?>) tables.get(2)).getValue();
                processThirdTable(thirdTable, documentPart);

                System.out.println("Standardized the DOCX file: " + inputDocxFile);
            } catch (Exception e) {
                System.out.println("Error while standardizing the DOCX file: " + e.getMessage());
            }
        }
        // Save the changes
        wordMLPackage.save(new File(inputDocxFile));
    }
}
