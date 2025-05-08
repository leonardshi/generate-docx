package com.example;

import org.docx4j.wml.*;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

public class ParagraphStyler {

    private static final Logger logger = LoggerFactory.getLogger(ParagraphStyler.class);

    public void styleFirstParagraph(MainDocumentPart documentPart) {
        List<Object> content = documentPart.getContent();
        if (content.isEmpty()) {
            logger.warn("The document is empty.");
            return;
        }

        Object firstElement = content.get(0);
        if (firstElement instanceof P) {
            P firstParagraph = (P) firstElement;

            // Apply bold styling
            for (Object paragraphContent : firstParagraph.getContent()) {
                if (paragraphContent instanceof R) {
                    R run = (R) paragraphContent;
                    RPr rPr = run.getRPr();
                    if (rPr == null) {
                        rPr = new RPr();
                        run.setRPr(rPr);
                    }

                    BooleanDefaultTrue bold = new BooleanDefaultTrue();
                    rPr.setB(bold);
                }
            }

            // Center align the paragraph
            PPr pPr = firstParagraph.getPPr();
            if (pPr == null) {
                pPr = new PPr();
                firstParagraph.setPPr(pPr);
            }

            Jc justification = new Jc();
            justification.setVal(JcEnumeration.CENTER);
            pPr.setJc(justification);
        } else {
            logger.warn("The first element is not a paragraph.");
        }
    }
}
