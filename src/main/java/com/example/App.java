package com.example;

import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;

public class App {

    public static void main(String[] args) throws org.docx4j.openpackaging.exceptions.InvalidFormatException, org.docx4j.openpackaging.exceptions.Docx4JException {
        if (args.length < 2) {
            System.out.println("Usage: java App <input-markdown-file> <output-docx-file>");
            return;
        }

        Path inputMarkdownFile = Paths.get(args[0]);
        String outputDocxFile = args[1];

        MarkdownToDocxService service = new MarkdownToDocxService();
        try {
            service.convertMarkdownToDocx(inputMarkdownFile, outputDocxFile);
        } catch (IOException e) {
            System.err.println("Error: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
