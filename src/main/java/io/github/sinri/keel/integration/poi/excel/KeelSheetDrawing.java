package io.github.sinri.keel.integration.poi.excel;

import io.github.sinri.keel.core.TechnicalPreview;
import io.github.sinri.keel.core.ValueBox;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @since 4.1.1
 */
@TechnicalPreview(since = "4.1.1")
class KeelSheetDrawing {
    @Nonnull
    private final ValueBox<XSSFDrawing> drawingForXlsxValueBox = new ValueBox<>();
    @Nonnull
    private final ValueBox<HSSFPatriarch> drawingForXlsValueBox = new ValueBox<>();


    public KeelSheetDrawing(@Nonnull KeelSheet keelSheet) {
        drawingForXlsxValueBox.setValue(null);
        drawingForXlsValueBox.setValue(null);
        if (keelSheet.getSheetsReaderType() == KeelSheetsReaderType.XLSX) {
            var x = keelSheet.getSheet().getDrawingPatriarch();
            if (x == null) {
                drawingForXlsxValueBox.setValue(null);
            } else {
                drawingForXlsxValueBox.setValue((XSSFDrawing) x);
            }
        } else if (keelSheet.getSheetsReaderType() == KeelSheetsReaderType.XLS) {
            var x = keelSheet.getSheet().getDrawingPatriarch();
            if (x == null) {
                drawingForXlsValueBox.setValue(null);
            } else {
                drawingForXlsValueBox.setValue((HSSFPatriarch) x);
            }
        }
    }

    @Nonnull
    public List<KeelPictureInSheet> getPictures() {
        List<KeelPictureInSheet> list = new ArrayList<>();

        List<XSSFPicture> picturesForXlsx = getPicturesForXlsx();
        if (picturesForXlsx != null) {
            picturesForXlsx.forEach(picture -> {
                list.add(new KeelPictureInSheet(picture));
            });
        }

        List<HSSFPicture> picturesForXls = getPicturesForXls();
        if (picturesForXls != null) {
            picturesForXls.forEach(picture -> {
                list.add(new KeelPictureInSheet(picture));
            });
        }

        return list;
    }

    @Nullable
    private XSSFDrawing getDrawingForXlsx() {
        return drawingForXlsxValueBox.getValue();
    }

    @Nullable
    private List<XSSFShape> getShapesForXlsx() {
        XSSFDrawing xlsxDrawing = getDrawingForXlsx();
        if (xlsxDrawing == null) return null;
        return xlsxDrawing.getShapes();
    }

    @Nullable
    private List<XSSFPicture> getPicturesForXlsx() {
        List<XSSFShape> xlsxShapes = getShapesForXlsx();
        if (xlsxShapes == null) return null;
        return xlsxShapes.stream()
                         .filter(shape -> shape instanceof XSSFPicture)
                         .map(x -> (XSSFPicture) x)
                         .collect(Collectors.toList());
    }

    @Nullable
    private HSSFPatriarch getDrawingForXls() {
        return drawingForXlsValueBox.getValue();
    }

    @Nullable
    private List<org.apache.poi.hssf.usermodel.HSSFShape> getShapesForXls() {
        HSSFPatriarch drawing = getDrawingForXls();
        if (drawing == null) return null;
        return drawing.getChildren();
    }

    @Nullable
    private List<HSSFPicture> getPicturesForXls() {
        List<HSSFShape> shapesForXls = getShapesForXls();
        if (shapesForXls == null) return null;
        return shapesForXls.stream()
                           .filter(x -> x instanceof HSSFPicture)
                           .map(x -> (HSSFPicture) x)
                           .collect(Collectors.toList());
    }

}
