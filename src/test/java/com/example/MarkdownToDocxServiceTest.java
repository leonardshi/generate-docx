package com.example;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertTrue;

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
        Files.copy(outputFile.toPath(), testResourcesTmpDir.resolve(outputFile.getName()));
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
        File outputFile = new File(outputDocxFile);
        assertTrue(outputFile.exists(), "The output DOCX file should exist.");
        assertTrue(outputFile.length() > 0, "The output DOCX file should not be empty.");

        // Copy the output file to the test resources tmp folder
        //Path testResourcesTmpDir = Path.of("src/test/resources/tmp");
        Path testResourcesTmpDir = Path.of("/mnt/c/temp");
        Files.createDirectories(testResourcesTmpDir);
        Files.copy(outputFile.toPath(), testResourcesTmpDir.resolve(outputFile.getName()), java.nio.file.StandardCopyOption.REPLACE_EXISTING);

    }
}
