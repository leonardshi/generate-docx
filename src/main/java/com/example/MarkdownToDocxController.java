package com.example;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.io.IOException;

@RestController
@RequestMapping("/api/v1/docx")
public class MarkdownToDocxController {

    private final MarkdownToDocxService markdownToDocxService;

    @Autowired
    public MarkdownToDocxController(MarkdownToDocxService markdownToDocxService) {
        this.markdownToDocxService = markdownToDocxService;
    }

    @PostMapping("/convert")
    public String convertMarkdownToDocx(@RequestBody String markdownContent) {
        try {
            markdownToDocxService.convertMarkdownToDocx(markdownContent, "/mnt/c/temp/converted.docx");
            return "Markdown converted to DOCX successfully.";
        } catch (IOException e) {
            return "Error during conversion: " + e.getMessage();
        } catch (InvalidFormatException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
            return "Error during conversion: " + e.getMessage();
        } catch (Docx4JException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
            return "Error during conversion: " + e.getMessage();
        }
    }
}
