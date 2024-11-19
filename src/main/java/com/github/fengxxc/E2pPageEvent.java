package com.github.fengxxc;

import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfPage;

/**
 * @author fengxxc
 * @date 2024-11-19
 */
public class E2pPageEvent {
    private int sheetAt;
    private String sheetName;
    private PdfDocument pdfDocument;
    private PdfPage pdfPage;

    public E2pPageEvent(int sheetAt, String sheetName, PdfDocument pdfDocument, PdfPage pdfPage) {
        this.sheetAt = sheetAt;
        this.sheetName = sheetName;
        this.pdfDocument = pdfDocument;
        this.pdfPage = pdfPage;
    }

    public int getSheetAt() {
        return sheetAt;
    }

    public String getSheetName() {
        return sheetName;
    }

    public PdfDocument getPdfDocument() {
        return pdfDocument;
    }

    public PdfPage getPdfPage() {
        return pdfPage;
    }

    public int getPageNumber() {
        return pdfDocument.getPageNumber(pdfPage);
    }
}
