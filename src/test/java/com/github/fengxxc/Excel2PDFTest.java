package com.github.fengxxc;

import com.github.fengxxc.util.ITextUtil;
import com.itextpdf.kernel.colors.ColorConstants;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.geom.Rectangle;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfPage;
import com.itextpdf.kernel.pdf.canvas.PdfCanvas;
import com.itextpdf.layout.Canvas;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.properties.VerticalAlignment;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import static org.junit.Assert.*;

/**
 * @author fengxxc
 * @date 2022-12-08
 */
public class Excel2PDFTest {

    @Before
    public void setUp() throws Exception {
    }

    @After
    public void tearDown() throws Exception {
    }

    @Test
    public void process() {
        try (InputStream is = Excel2PDF.class.getResourceAsStream("/mayilong_resume.xlsx");
             OutputStream os = new FileOutputStream("F:\\temp\\excel2pdf\\mayilong_resume.pdf")
        ) {
            Excel2PDF.process(is, os, document -> {

            });
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void process2() {
        try (InputStream is = Excel2PDF.class.getResourceAsStream("/mayilong_resume2.xlsx");
             OutputStream os = new FileOutputStream("F:\\temp\\excel2pdf\\mayilong_resume2.pdf")
        ) {
            Excel2PDF.process(is, os, document -> {

            });
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    @Test
    public void process3() {
        try (InputStream is = Excel2PDF.class.getResourceAsStream("/large_page.xlsx");
             OutputStream os = new FileOutputStream("C:\\temp\\excel2pdf\\large_page.pdf");
        ) {
            PdfFont font = ITextUtil.createFont(ITextUtil.FONT_SIMHEI);
            Excel2PDF.process(is, os, document -> {
                // document callback...
            }, e2pPageEvent -> {
                // page started callback
                PdfDocument pdfDocument = e2pPageEvent.getPdfDocument();
                PdfPage pdfPage = e2pPageEvent.getPdfPage();
                String sheetName = e2pPageEvent.getSheetName();
                PdfCanvas pdfCanvas = new PdfCanvas(pdfPage.newContentStreamBefore(), pdfPage.getResources(), pdfDocument);
                Rectangle pageSize = pdfPage.getPageSize();
                // 绘制水印
                new Canvas(pdfCanvas, pdfPage.getPageSize())
                        .setFontColor(ColorConstants.LIGHT_GRAY).setFontSize(60)
                        .setFont(font)
                        .showTextAligned(
                                new Paragraph(sheetName)
                                , pageSize.getWidth() / 2, pageSize.getHeight() / 2, pdfDocument.getPageNumber(pdfPage), TextAlignment.CENTER
                                , VerticalAlignment.MIDDLE, 45
                        );

                pdfCanvas.release();
            });
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}