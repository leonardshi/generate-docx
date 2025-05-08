package com.example;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.mockito.Mockito;

import java.io.IOException;

import static org.mockito.Mockito.*;

class MeetingMinutesTemplateServiceTest {

    private MeetingMinutesTemplateService service;
    private DocxFileHandler mockDocxFileHandler;
    private ParagraphStyler mockParagraphStyler;
    private TableProcessor mockTableProcessor;
    private WordprocessingMLPackage mockWordMLPackage;
    private MainDocumentPart mockDocumentPart;

    @BeforeEach
    void setUp() {
        mockDocxFileHandler = mock(DocxFileHandler.class);
        mockParagraphStyler = mock(ParagraphStyler.class);
        mockTableProcessor = mock(TableProcessor.class);
        mockWordMLPackage = mock(WordprocessingMLPackage.class);
        mockDocumentPart = mock(MainDocumentPart.class);

        service = new MeetingMinutesTemplateService(mockDocxFileHandler, mockParagraphStyler, mockTableProcessor);
    }

    @Test
    void testStandardizeDocxFile() throws IOException, Docx4JException {
        when(mockDocxFileHandler.loadDocxFile(anyString())).thenReturn(mockWordMLPackage);
        when(mockWordMLPackage.getMainDocumentPart()).thenReturn(mockDocumentPart);

        service.standardizeDocxFile("test.docx");

        verify(mockDocxFileHandler, times(1)).loadDocxFile(anyString());
        verify(mockParagraphStyler, times(1)).styleFirstParagraph(mockDocumentPart);
    }
}
