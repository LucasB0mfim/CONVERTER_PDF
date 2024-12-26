package br.com.converter.converter.service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;

import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.properties.UnitValue;
import com.itextpdf.io.font.constants.StandardFonts;

@Service
public class DocxToPdfService {

    public byte[] convertDocxToPdfBytes(InputStream docxInputStream) throws IOException {
        try (docxInputStream; XWPFDocument docx = new XWPFDocument(docxInputStream)) {

            ByteArrayOutputStream pdfOutputStream = new ByteArrayOutputStream();
            PdfDocument pdfDoc = new PdfDocument(new PdfWriter(pdfOutputStream));
            Document document = new Document(pdfDoc);

            // Processar parágrafos
            for (XWPFParagraph docxParagraph : docx.getParagraphs()) {
                Paragraph pdfParagraph = new Paragraph();

                // Definir alinhamento
                ParagraphAlignment alignment = docxParagraph.getAlignment();
                if (alignment != null) {
                    switch (alignment) {
                        case CENTER:
                            pdfParagraph.setTextAlignment(TextAlignment.CENTER);
                            break;
                        case RIGHT:
                            pdfParagraph.setTextAlignment(TextAlignment.RIGHT);
                            break;
                        case BOTH:
                            pdfParagraph.setTextAlignment(TextAlignment.JUSTIFIED);
                            break;
                        default:
                            pdfParagraph.setTextAlignment(TextAlignment.LEFT);
                    }
                }

                // Processar runs (partes do texto com formatação específica)
                for (XWPFRun run : docxParagraph.getRuns()) {
                    String text = run.getText(0);
                    if (text != null) {
                        com.itextpdf.layout.element.Text pdfText = new com.itextpdf.layout.element.Text(text);

                        // Aplicar estilos (negrito e itálico)
                        PdfFont font;
                        try {
                            if (run.isBold() && run.isItalic()) {
                                font = PdfFontFactory.createFont(StandardFonts.HELVETICA_BOLDOBLIQUE);
                            } else if (run.isBold()) {
                                font = PdfFontFactory.createFont(StandardFonts.HELVETICA_BOLD);
                            } else if (run.isItalic()) {
                                font = PdfFontFactory.createFont(StandardFonts.HELVETICA_OBLIQUE);
                            } else {
                                font = PdfFontFactory.createFont(StandardFonts.HELVETICA);
                            }
                        } catch (IOException e) {
                            font = PdfFontFactory.createFont(StandardFonts.HELVETICA);
                        }
                        pdfText.setFont(font);

                        // Aplicar tamanho da fonte
                        Double fontSize = run.getFontSizeAsDouble();
                        if (fontSize != null && fontSize > 0) {
                            pdfText.setFontSize(fontSize.floatValue());
                        }

                        // Aplicar cor
                        String color = run.getColor();
                        if (color != null) {
                            pdfText.setFontColor(convertColor(color));
                        }

                        pdfParagraph.add(pdfText);
                    }
                }

                document.add(pdfParagraph);
            }

            // Processar tabelas
            for (XWPFTable docxTable : docx.getTables()) {
                float[] columnWidths = new float[docxTable.getRow(0).getTableCells().size()];
                for (int i = 0; i < columnWidths.length; i++) {
                    columnWidths[i] = 1; // Distribuição uniforme
                }

                Table pdfTable = new Table(UnitValue.createPercentArray(columnWidths));

                for (XWPFTableRow row : docxTable.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        Cell pdfCell = new Cell();

                        // Processar parágrafos dentro da célula
                        for (XWPFParagraph cellParagraph : cell.getParagraphs()) {
                            pdfCell.add(new Paragraph(cellParagraph.getText()));
                        }

                        pdfTable.addCell(pdfCell);
                    }
                }

                document.add(pdfTable);
            }

            document.close();
            return pdfOutputStream.toByteArray();
        }
    }

    // Método auxiliar para converter cor hexadecimal para DeviceRgb
    private DeviceRgb convertColor(String color) {
        if (color != null && color.length() == 6) {
            // A cor fornecida é uma string hexadecimal (ex: "FF5733")
            int r = Integer.parseInt(color.substring(0, 2), 16);
            int g = Integer.parseInt(color.substring(2, 4), 16);
            int b = Integer.parseInt(color.substring(4, 6), 16);
            return new DeviceRgb(r, g, b);
        }
        return (DeviceRgb) DeviceRgb.BLACK; // Retorna preto caso a cor seja inválida
    }
}
