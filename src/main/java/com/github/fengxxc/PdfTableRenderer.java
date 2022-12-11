package com.github.fengxxc;

import com.github.fengxxc.util.Position;
import com.itextpdf.io.image.ImageData;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.geom.Rectangle;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.renderer.CellRenderer;
import com.itextpdf.layout.renderer.DrawContext;
import com.itextpdf.layout.renderer.IRenderer;
import com.itextpdf.layout.renderer.TableRenderer;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.util.ImageUtils;
import org.apache.poi.util.Units;

import java.awt.*;
import java.util.Map;

/**
 * @author fengxxc
 * @date 2022-12-09
 */
public class PdfTableRenderer<T extends Picture> extends TableRenderer {
    private Map<Position, T> pos2picture;
    public PdfTableRenderer(Table modelElement, Map<Position, T> pos2picture) {
        super(modelElement);
        this.pos2picture = pos2picture;
    }

    @Override
    public void drawChildren(DrawContext drawContext) {
        super.drawChildren(drawContext);
        if (pos2picture == null) {
            return;
        }
        pos2picture.forEach((pos, picture) -> {
            if (pos.getRowIndex() >= super.rows.size()) {
                return;
            }
            final CellRenderer[] cellRenderers = super.rows.get(pos.getRowIndex());
            final CellRenderer cellRenderer = cellRenderers[pos.getColIndex()];
            final Rectangle areaBBox = cellRenderer.getOccupiedAreaBBox();
            final ImageData imageData = ImageDataFactory.create(picture.getPictureData().getData());
            final Dimension imageDimension = ImageUtils.getDimensionFromAnchor(picture);
            final double width = Units.toPoints((long) (imageDimension.getWidth()));
            final double height = Units.toPoints((long) (imageDimension.getHeight()));
            final Rectangle rectangle = areaBBox.clone();
            // test
            rectangle.setWidth((float) width);
            rectangle.setHeight((float) height);
            drawContext.getCanvas().addImageFittedIntoRectangle(imageData, rectangle, false);
        });
    }

    @Override
    public IRenderer getNextRenderer() {
        return new PdfTableRenderer<T>((Table) modelElement, pos2picture);
    }
}
