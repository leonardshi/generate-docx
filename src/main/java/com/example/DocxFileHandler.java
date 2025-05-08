package com.example;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

public class DocxFileHandler {

    private static final Logger logger = LoggerFactory.getLogger(DocxFileHandler.class);

    public WordprocessingMLPackage loadDocxFile(String filePath) throws IOException, Docx4JException {
        logger.info("Loading DOCX file from path: {}", filePath);
        return WordprocessingMLPackage.load(new File(filePath));
    }

    public void saveDocxFile(WordprocessingMLPackage wordMLPackage, String filePath) throws IOException, Docx4JException {
        logger.info("Saving DOCX file to path: {}", filePath);
        wordMLPackage.save(new File(filePath));
    }

    public WordprocessingMLPackage loadDocxFile(InputStream inputStream) throws IOException, Docx4JException {
        logger.info("Loading DOCX file from InputStream");
        return WordprocessingMLPackage.load(inputStream);
    }
}
