package com.example;

import org.docx4j.wml.*;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.math.BigInteger;
import java.util.List;

public class TableProcessor {

    private static final Logger logger = LoggerFactory.getLogger(TableProcessor.class);

    public void processFirstTable(Tbl firstTable, MainDocumentPart documentPart) {
        mergeAndStyleFirstRow(firstTable, documentPart, "000000", "FFFFFF"); // Black background, white font
        List<Object> rows = firstTable.getContent();
        for (int i = 1; i < rows.size(); i++) { // Start from the second row
            Tr row = (Tr) rows.get(i);
            if (!row.getContent().isEmpty()) {
                Object cellObj = row.getContent().get(0);
                Tc firstCell = extractTableCell(cellObj);
                if (firstCell != null) {
                    setCellBackgroundColor(firstCell, "F2F2F2"); // Light gray background
                    setFontColorAndBold(firstCell, "000000"); // Black font color
                }
            }
        }
    }

    public void processSecondTable(Tbl secondTable) {
        styleTableFirstRow(secondTable, "000000", "FFFFFF"); // Black background, white font
    }

    public void processThirdTable(Tbl thirdTable, MainDocumentPart documentPart) {
        mergeAndStyleFirstRow(thirdTable, documentPart, "000000", "FFFFFF"); // Black background, white font
        List<Object> secondRows = thirdTable.getContent();
        if (secondRows.size() > 1) {
            Tr secondRow = (Tr) secondRows.get(1);
            for (Object cellObj : secondRow.getContent()) {
                Tc cell = extractTableCell(cellObj);
                if (cell != null) {
                    setCellBackgroundColor(cell, "F2F2F2"); // Light gray background
                    setFontColorAndBold(cell, "000000"); // Black font color
                }
            }
        } else {
            logger.warn("The third table does not have enough rows.");
        }
    }

    private void setCellBackgroundColor(Tc cell, String color) {
        TcPr tcPr = cell.getTcPr();
        if (tcPr == null) {
            tcPr = new TcPr();
            cell.setTcPr(tcPr);
        }

        CTShd shd = new CTShd();
        shd.setFill(color);
        tcPr.setShd(shd);
    }

    private void setFontColorAndBold(Tc cell, String color) {
        for (Object content : cell.getContent()) {
            if (content instanceof P) {
                P paragraph = (P) content;
                List<Object> paragraphContents = paragraph.getContent();
                for (Object paragraphContent : paragraphContents) {
                    if (paragraphContent instanceof R) {
                        R run = (R) paragraphContent;
                        RPr rPr = run.getRPr();
                        if (rPr == null) {
                            rPr = new RPr();
                            run.setRPr(rPr);
                        }

                        // Set font color
                        Color fontColor = new Color();
                        fontColor.setVal(color);
                        rPr.setColor(fontColor);

                        // Make text bold
                        BooleanDefaultTrue bold = new BooleanDefaultTrue();
                        rPr.setB(bold);
                    }
                }
            }
        }
    }

    private void styleTableFirstRow(Tbl table, String backgroundColor, String fontColor) {
        List<Object> rows = table.getContent();

        if (!rows.isEmpty()) {
            Tr firstRow = (Tr) rows.get(0);

            for (Object cell : firstRow.getContent()) {
                Tc tc = extractTableCell(cell);
                if (tc != null) {
                    setCellBackgroundColor(tc, backgroundColor);
                    setFontColorAndBold(tc, fontColor);
                }
            }
        }
    }

    private void mergeAndStyleFirstRow(Tbl table, MainDocumentPart documentPart, String backgroundColor, String fontColor) {
        List<Object> rows = table.getContent();

        if (rows.isEmpty()) {
            logger.warn("The table is empty.");
            return;
        }

        Tr firstRow = (Tr) rows.get(0);

        List<Tc> cells = new java.util.ArrayList<>();
        for (Object cell : firstRow.getContent()) {
            Tc tc = extractTableCell(cell);
            if (tc != null) {
                cells.add(tc);
            }
        }

        if (cells.size() >= 2) {
            Tc firstCell = cells.get(0);
            StringBuilder firstCellValue = new StringBuilder();
            for (Object content : firstCell.getContent()) {
                firstCellValue.append(content.toString());
            }

            for (int i = 1; i < cells.size(); i++) {
                Tc cellToMerge = cells.get(i);
                for (Object content : cellToMerge.getContent()) {
                    firstCellValue.append(content.toString());
                }
                firstRow.getContent().remove(cells.get(i));
            }

            TcPr tcPr = firstCell.getTcPr();
            if (tcPr == null) {
                tcPr = new TcPr();
                firstCell.setTcPr(tcPr);
            }
            TcPrInner.GridSpan gridSpan = new TcPrInner.GridSpan();
            gridSpan.setVal(BigInteger.valueOf(cells.size()));
            tcPr.setGridSpan(gridSpan);

            firstCell.getContent().clear();
            firstCell.getContent().add(documentPart.createParagraphOfText(firstCellValue.toString()));
        }

        for (Tc tc : cells) {
            setCellBackgroundColor(tc, backgroundColor);
            setFontColorAndBold(tc, fontColor);
        }
    }

    private Tc extractTableCell(Object cellObj) {
        if (cellObj instanceof Tc) {
            return (Tc) cellObj;
        } else if (cellObj instanceof jakarta.xml.bind.JAXBElement) {
            Object value = ((jakarta.xml.bind.JAXBElement<?>) cellObj).getValue();
            if (value instanceof Tc) {
                return (Tc) value;
            }
        }
        return null;
    }

    public Tbl extractTable(Object tableObj) {
        if (tableObj instanceof Tbl) {
            return (Tbl) tableObj;
        } else if (tableObj instanceof jakarta.xml.bind.JAXBElement) {
            Object value = ((jakarta.xml.bind.JAXBElement<?>) tableObj).getValue();
            if (value instanceof Tbl) {
                return (Tbl) value;
            }
        }
        return null;
    }
}
