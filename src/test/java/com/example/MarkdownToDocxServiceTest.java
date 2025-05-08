package com.example;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

public class MarkdownToDocxServiceTest {

    @Test
    public void testConvertMarkdownToDocxWithStringInput(@TempDir Path tempDir) throws Exception {
        // Arrange
        String markdownContent = "# Hello World\nThis is a test markdown file.";
        String outputDocxFile = tempDir.resolve("output.docx").toString();
        MarkdownToDocxService service = new MarkdownToDocxService();

        // Act
        service.convertMarkdownToDocx(markdownContent, outputDocxFile);

        // Assert
        File outputFile = new File(outputDocxFile);
        assertTrue(outputFile.exists(), "The output DOCX file should exist.");
        assertTrue(outputFile.length() > 0, "The output DOCX file should not be empty.");

        // Copy the output file to the test resources tmp folder
        Path testResourcesTmpDir = Path.of("src/test/resources/tmp");
        Files.createDirectories(testResourcesTmpDir);
        Files.copy(outputFile.toPath(), testResourcesTmpDir.resolve(outputFile.getName()), java.nio.file.StandardCopyOption.REPLACE_EXISTING);
    }

    @Test
    public void testConvertMarkdownToDocxWithFileInput(@TempDir Path tempDir) throws Exception {
        // Arrange
        String markdownContent = "# Title\n" +
                                 "## Subtitle\n" +
                                 "- Item 1\n" +
                                 "  - Subitem 1.1\n" +
                                 "  - Subitem 1.2\n" +
                                 "- Item 2\n\n" +
                                 "**Bold Text with *italic inside***\n" +
                                 "*Italic Text with **bold inside***\n\n" +
                                 "[Link with **bold text**](https://example.com)\n\n" +
                                 "```java\n" +
                                 "// Nested code block\n" +
                                 "public class Example {\n" +
                                 "    public static void main(String[] args) {\n" +
                                 "        System.out.println(\"Hello, World!\");\n" +
                                 "    }\n" +
                                 "}\n" +
                                 "```\n";

        String outputDocxFile = tempDir.resolve("output-2.docx").toString();
        MarkdownToDocxService service = new MarkdownToDocxService();

        // Act
        service.convertMarkdownToDocx(markdownContent, outputDocxFile);

        // Assert
        File outputFile2 = new File(outputDocxFile);
        assertTrue(outputFile2.exists(), "The output DOCX file should exist.");
        assertTrue(outputFile2.length() > 0, "The output DOCX file should not be empty.");

        // Copy the output file to the test resources tmp folder
        //Path testResourcesTmpDir = Path.of("src/test/resources/tmp");
        Path testResourcesTmpDir = Path.of("src/test/resources/tmp");
        Files.createDirectories(testResourcesTmpDir);
        Files.copy(outputFile2.toPath(), testResourcesTmpDir.resolve(outputFile2.getName()), java.nio.file.StandardCopyOption.REPLACE_EXISTING);

    }

    @Test
    public void testMockedStandardizeDocxFile() throws Exception {
        // Arrange
        DocxFileHandler mockDocxFileHandler = mock(DocxFileHandler.class);
        ParagraphStyler mockParagraphStyler = mock(ParagraphStyler.class);
        TableProcessor mockTableProcessor = mock(TableProcessor.class);
        MeetingMinutesTemplateService service = new MeetingMinutesTemplateService(mockDocxFileHandler, mockParagraphStyler, mockTableProcessor);
        Path tempFile = Files.createTempFile("meeting-minutes", ".docx");
        Files.copy(new File("src/main/resources/meeting-minutes.docx").toPath(), tempFile, java.nio.file.StandardCopyOption.REPLACE_EXISTING);

        // Mock the behavior of DocxFileHandler to return a valid WordprocessingMLPackage
        WordprocessingMLPackage mockWordMLPackage = mock(WordprocessingMLPackage.class);
        MainDocumentPart mockDocumentPart = mock(MainDocumentPart.class);
        when(mockDocxFileHandler.loadDocxFile(anyString())).thenReturn(mockWordMLPackage);
        when(mockWordMLPackage.getMainDocumentPart()).thenReturn(mockDocumentPart);

        // Act
        service.standardizeDocxFile(tempFile.toString());

        // Clean up
        Files.copy(tempFile, Path.of("/mnt/c/temp/output.docx"), java.nio.file.StandardCopyOption.REPLACE_EXISTING);
    }

    @Test
    public void testStandardizeDocxFile() throws Exception {
        // Arrange
        DocxFileHandler docxFileHandler = new DocxFileHandler();
        ParagraphStyler paragraphStyler = new ParagraphStyler();
        TableProcessor tableProcessor = new TableProcessor();
        MeetingMinutesTemplateService service = new MeetingMinutesTemplateService(docxFileHandler, paragraphStyler, tableProcessor);
        Path tempFile = Files.createTempFile("meeting-minutes", ".docx");
        Files.copy(new File("src/main/resources/meeting-minutes.docx").toPath(), tempFile, java.nio.file.StandardCopyOption.REPLACE_EXISTING);

        // Act
        service.standardizeDocxFile(tempFile.toString());

        // Clean up
        // Files.copy(tempFile, Path.of("/mnt/c/temp/output.docx"), java.nio.file.StandardCopyOption.REPLACE_EXISTING);
    }
}
