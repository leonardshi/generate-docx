package com.example;

import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.mockito.Mockito;

import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

class TableProcessorTest {

    private TableProcessor tableProcessor;
    private MainDocumentPart mockDocumentPart;
    private Tbl mockTable;

    @BeforeEach
    void setUp() {
        tableProcessor = new TableProcessor();
        mockDocumentPart = mock(MainDocumentPart.class);
        mockTable = mock(Tbl.class);
    }

    @Test
    void testProcessFirstTable() {
        List<Object> rows = new ArrayList<>();
        Tr mockRow = mock(Tr.class);
        rows.add(mockRow);
        when(mockTable.getContent()).thenReturn(rows);

        tableProcessor.processFirstTable(mockTable, mockDocumentPart);

        verify(mockTable, times(2)).getContent();
    }

    @Test
    void testProcessSecondTable() {
        List<Object> rows = new ArrayList<>();
        Tr mockRow = mock(Tr.class);
        rows.add(mockRow);
        when(mockTable.getContent()).thenReturn(rows);

        tableProcessor.processSecondTable(mockTable);

        verify(mockTable, times(1)).getContent();
    }

    @Test
    void testProcessThirdTable() {
        List<Object> rows = new ArrayList<>();
        Tr mockRow = mock(Tr.class);
        rows.add(mockRow);
        when(mockTable.getContent()).thenReturn(rows);

        tableProcessor.processThirdTable(mockTable, mockDocumentPart);

        verify(mockTable, times(2)).getContent();
    }
}
