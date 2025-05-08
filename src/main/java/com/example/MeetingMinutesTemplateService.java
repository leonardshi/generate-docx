package com.example;

import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.util.List;

public class MeetingMinutesTemplateService {

    private static final Logger logger = LoggerFactory.getLogger(MeetingMinutesTemplateService.class);

    private final DocxFileHandler docxFileHandler;
    private final ParagraphStyler paragraphStyler;
    private final TableProcessor tableProcessor;

    // Added constructor for dependency injection
    public MeetingMinutesTemplateService(DocxFileHandler docxFileHandler, ParagraphStyler paragraphStyler, TableProcessor tableProcessor) {
        this.docxFileHandler = docxFileHandler;
        this.paragraphStyler = paragraphStyler;
        this.tableProcessor = tableProcessor;
    }

    public void standardizeDocxFile(String inputDocxFilePath) throws IOException, Docx4JException {
        WordprocessingMLPackage wordMLPackage = docxFileHandler.loadDocxFile(inputDocxFilePath);
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

        paragraphStyler.styleFirstParagraph(documentPart);

        List<Object> tables = null;
        try {
            tables = documentPart.getJAXBNodesViaXPath("//w:tbl", true);
        } catch (XPathBinderAssociationIsPartialException | javax.xml.bind.JAXBException e) {
            logger.error("Error while fetching tables via XPath", e);
        }

        if (tables == null || tables.isEmpty() || tables.size() < 3) {
            logger.warn("Incorrect number of tables found in the document.");
            return;
        } else {
            try {
                Tbl firstTable = (Tbl) ((javax.xml.bind.JAXBElement<?>) tables.get(0)).getValue();
                tableProcessor.processFirstTable(firstTable, documentPart);

                Tbl secondTable = (Tbl) ((javax.xml.bind.JAXBElement<?>) tables.get(1)).getValue();
                tableProcessor.processSecondTable(secondTable);

                Tbl thirdTable = (Tbl) ((javax.xml.bind.JAXBElement<?>) tables.get(2)).getValue();
                tableProcessor.processThirdTable(thirdTable, documentPart);

                logger.info("Standardized the DOCX file: {}", inputDocxFilePath);
            } catch (Exception e) {
                logger.error("Error while standardizing the DOCX file", e);
            }
        }

        docxFileHandler.saveDocxFile(wordMLPackage, inputDocxFilePath);
    }
}
