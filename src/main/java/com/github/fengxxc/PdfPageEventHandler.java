package com.github.fengxxc;

import com.itextpdf.kernel.events.Event;
import com.itextpdf.kernel.events.IEventHandler;
import com.itextpdf.kernel.events.PdfDocumentEvent;
import com.itextpdf.kernel.geom.Rectangle;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfPage;
import com.itextpdf.kernel.pdf.canvas.PdfCanvas;

import java.util.function.Consumer;

/**
 * @author fengxxc
 * @date 2024-11-19
 */
public class PdfPageEventHandler implements IEventHandler {
    // 页码从第几页开始（第一页是1）
    private static int START_PAGE = 1;

    int sheetAt = 0;
    String sheetName = "";

    /**
     * 水印 function (pageNumber) => watermarkText
     * 注意：pageNumber 从1开始
     */
    Consumer<E2pPageEvent> callback;

    public PdfPageEventHandler(Consumer<E2pPageEvent> callback) {
        this.callback = callback;
    }

    public PdfPageEventHandler setSheetAt(int sheetAt) {
        this.sheetAt = sheetAt;
        return this;
    }

    public PdfPageEventHandler setSheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    @Override
    public void handleEvent(Event event) {
        PdfDocumentEvent docEvent = (PdfDocumentEvent) event;
        PdfDocument pdfDoc = docEvent.getDocument();
        PdfPage page = docEvent.getPage();
        int pageNumber = pdfDoc.getPageNumber(page);
        Rectangle pageSize = page.getPageSize();
        PdfCanvas pdfCanvas = new PdfCanvas(page.newContentStreamBefore(), page.getResources(), pdfDoc);
        // PdfFont font = this.font != null ? this.font : pdfDoc.getDefaultFont();
        // // 绘制页码
        // // 实际页码
        // String _pageNumber = pageNumber < START_PAGE ? "" : pageNumber - START_PAGE + 1 + "";
        // pdfCanvas.beginText().setFontAndSize(font, 9)
        //         .setFillColorGray(0.5f) // 设置为灰色
        //         .moveText(pageSize.getWidth() - 40, 30)
        //         .showText(_pageNumber)
        //         .endText()
        //         .resetFillColorRgb() // 还原颜色
        // ;

        if (callback != null) {
            callback.accept(new E2pPageEvent(sheetAt, sheetName, pdfDoc, page));
        }
        // System.out.println(event.getType() + ": " + pageNumber + " " + watermarkText);
        // // 绘制水印
        // new Canvas(pdfCanvas, page.getPageSize())
        //         .setFontColor(ColorConstants.LIGHT_GRAY).setFontSize(60)
        //         .setFont(font)
        //         .showTextAligned(
        //                 new Paragraph(watermarkText)
        //                 , pageSize.getWidth() / 2, pageSize.getHeight() / 2, pdfDoc.getPageNumber(page), TextAlignment.CENTER
        //                 , VerticalAlignment.MIDDLE, 45
        //         );
        //
        // pdfCanvas.release();
    }
}
