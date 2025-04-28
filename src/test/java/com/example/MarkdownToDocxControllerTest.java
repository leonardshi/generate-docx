package com.example;

import org.junit.jupiter.api.Test;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;
import org.springframework.http.MediaType;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.test.web.servlet.setup.MockMvcBuilders;

import java.io.IOException;

import static org.mockito.Mockito.doThrow;
import static org.mockito.Mockito.verify;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.post;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

public class MarkdownToDocxControllerTest {

    private MockMvc mockMvc;

    @Mock
    private MarkdownToDocxService markdownToDocxService;

    @InjectMocks
    private MarkdownToDocxController markdownToDocxController;

    public MarkdownToDocxControllerTest() {
        MockitoAnnotations.openMocks(this);
        this.mockMvc = MockMvcBuilders.standaloneSetup(markdownToDocxController).build();
    }

    @Test
    void testConvertMarkdownToDocx_Success() throws Exception {
        String markdownContent = "# Sample Markdown";

        mockMvc.perform(post("/api/v1/docx/convert")
                .content(markdownContent)
                .contentType(MediaType.TEXT_PLAIN))
                .andExpect(status().isOk());

        verify(markdownToDocxService).convertMarkdownToDocx(markdownContent, "/mnt/c/temp/converted.docx");
    }

    @Test
    void testConvertMarkdownToDocx_IOException() throws Exception {
        String markdownContent = "# Sample Markdown";

        doThrow(new IOException("IO Error")).when(markdownToDocxService).convertMarkdownToDocx(markdownContent, "");

        mockMvc.perform(post("/api/v1/docx/convert")
                .content(markdownContent)
                .contentType(MediaType.TEXT_PLAIN))
                .andExpect(status().isOk());
    }

    @Test
    void testConvertMarkdownToDocx_InvalidFormatException() throws Exception {
        String markdownContent = "# Sample Markdown";

        doThrow(new org.docx4j.openpackaging.exceptions.InvalidFormatException("Invalid Format"))
                .when(markdownToDocxService).convertMarkdownToDocx(markdownContent, "");

        mockMvc.perform(post("/api/v1/docx/convert")
                .content(markdownContent)
                .contentType(MediaType.TEXT_PLAIN))
                .andExpect(status().isOk());
    }

    @Test
    void testConvertMarkdownToDocx_Docx4JException() throws Exception {
        String markdownContent = "# Sample Markdown";

        doThrow(new org.docx4j.openpackaging.exceptions.Docx4JException("Docx4J Error"))
                .when(markdownToDocxService).convertMarkdownToDocx(markdownContent, "");

        mockMvc.perform(post("/api/v1/docx/convert")
                .content(markdownContent)
                .contentType(MediaType.TEXT_PLAIN))
                .andExpect(status().isOk());
    }
}
