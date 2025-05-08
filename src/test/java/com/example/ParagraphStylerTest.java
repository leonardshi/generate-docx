package com.example;

import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.P;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.mockito.Mockito;

import java.util.ArrayList;
import java.util.List;

import static org.mockito.Mockito.*;

class ParagraphStylerTest {

    private ParagraphStyler paragraphStyler;
    private MainDocumentPart mockDocumentPart;

    @BeforeEach
    void setUp() {
        paragraphStyler = new ParagraphStyler();
        mockDocumentPart = mock(MainDocumentPart.class);
    }

    @Test
    void testStyleFirstParagraph() {
        List<Object> content = new ArrayList<>();
        P mockParagraph = mock(P.class);
        content.add(mockParagraph);
        when(mockDocumentPart.getContent()).thenReturn(content);

        paragraphStyler.styleFirstParagraph(mockDocumentPart);

        verify(mockDocumentPart, times(1)).getContent();
    }
}
