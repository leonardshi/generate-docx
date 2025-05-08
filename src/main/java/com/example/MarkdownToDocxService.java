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

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

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
}
