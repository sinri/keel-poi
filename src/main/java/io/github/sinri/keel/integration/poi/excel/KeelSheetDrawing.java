package io.github.sinri.keel.integration.poi.excel;

import io.github.sinri.keel.core.utils.value.ValueBox;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.jspecify.annotations.NullMarked;
import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Excel 工作表绘制类，用于处理 Excel 工作表中的图形元素。
 * <p>
 * 该类封装了对工作表中图片等绘制元素的访问。
 *
 * @since 5.0.0
 */
@NullMarked
class KeelSheetDrawing {
    private final ValueBox<XSSFDrawing> drawingForXlsxValueBox = new ValueBox<>();
    private final ValueBox<HSSFPatriarch> drawingForXlsValueBox = new ValueBox<>();


    /**
     * 构造函数，使用指定的工作表创建绘制类实例。
     * 该构造函数会根据工作表的类型（XLSX 或 XLS）初始化相应的绘图对象。
     *
     * @param keelSheet 工作表实例
     */
    public KeelSheetDrawing(KeelSheet keelSheet) {
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

    /**
     * 获取工作表中的所有图片列表。
     * 该方法支持 XLSX 和 XLS 格式的工作表，会返回所有嵌入的图片。
     *
     * @return 工作表中的图片列表
     */
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


    private @Nullable XSSFDrawing getDrawingForXlsx() {
        return drawingForXlsxValueBox.getValue();
    }


    private @Nullable List<XSSFShape> getShapesForXlsx() {
        XSSFDrawing xlsxDrawing = getDrawingForXlsx();
        if (xlsxDrawing == null) return null;
        return xlsxDrawing.getShapes();
    }


    private @Nullable List<XSSFPicture> getPicturesForXlsx() {
        List<XSSFShape> xlsxShapes = getShapesForXlsx();
        if (xlsxShapes == null) return null;
        return xlsxShapes.stream()
                         .filter(shape -> shape instanceof XSSFPicture)
                         .map(x -> (XSSFPicture) x)
                         .collect(Collectors.toList());
    }


    private @Nullable HSSFPatriarch getDrawingForXls() {
        return drawingForXlsValueBox.getValue();
    }


    private @Nullable List<org.apache.poi.hssf.usermodel.HSSFShape> getShapesForXls() {
        HSSFPatriarch drawing = getDrawingForXls();
        if (drawing == null) return null;
        return drawing.getChildren();
    }


    private @Nullable List<HSSFPicture> getPicturesForXls() {
        List<HSSFShape> shapesForXls = getShapesForXls();
        if (shapesForXls == null) return null;
        return shapesForXls.stream()
                           .filter(x -> x instanceof HSSFPicture)
                           .map(x -> (HSSFPicture) x)
                           .collect(Collectors.toList());
    }

}
