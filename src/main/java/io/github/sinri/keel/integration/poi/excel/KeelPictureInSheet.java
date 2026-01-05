package io.github.sinri.keel.integration.poi.excel;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.jspecify.annotations.NullMarked;


/**
 * Excel 工作表中的图片类，表示在 Excel 工作表中嵌入的图片。
 * <p>
 * 该类封装了图片的位置、尺寸、数据等信息。
 *
 * @since 5.0.0
 */
@NullMarked
public class KeelPictureInSheet {
    private final int atRow;
    private final int atCol;
    private final int width;
    private final int height;
    private final String suggestFileExtension;
    private final String mimeType;
    private final byte[] data;
    private final int workbookDefinedPictureType;

    /**
     * 构造函数，使用指定的 XSSF 图片对象创建 Excel 工作表中的图片实例。
     *
     * @param xssfPicture XSSF 图片对象
     */
    KeelPictureInSheet(XSSFPicture xssfPicture) {
        if (xssfPicture.getClientAnchor() != null) {
            var anchor = xssfPicture.getClientAnchor();
            atRow = anchor.getRow1() + 1; // 转换为Excel行号（从1开始）
            atCol = anchor.getCol1() + 1; // 转换为Excel列号（从1开始）
        } else {
            atRow = -1;
            atCol = -1;
        }

        var dimension = xssfPicture.getImageDimension();
        width = dimension.width;
        height = dimension.height;

        XSSFPictureData pictureData = xssfPicture.getPictureData();
        this.workbookDefinedPictureType = pictureData.getPictureType();
        this.data = pictureData.getData();
        this.suggestFileExtension = pictureData.suggestFileExtension();
        this.mimeType = pictureData.getMimeType();

    }

    /**
     * 构造函数，使用指定的 HSSF 图片对象创建 Excel 工作表中的图片实例。
     *
     * @param hssfPicture HSSF 图片对象
     */
    KeelPictureInSheet(HSSFPicture hssfPicture) {
        HSSFClientAnchor anchor = hssfPicture.getClientAnchor();
        if (anchor != null) {
            atRow = anchor.getRow1() + 1; // 转换为Excel行号（从1开始）
            atCol = anchor.getCol1() + 1; // 转换为Excel列号（从1开始）

            // 使用anchor的尺寸信息 - POI 5.4.1兼容方式
            // 转换为像素（近似值）
            width = Math.max(1, anchor.getDx1() / 9525); // 9525 EMUs per pixel
            height = Math.max(1, anchor.getDy1() / 9525);
        } else {
            atRow = -1;
            atCol = -1;

            width = -1;
            height = -1;
        }

        HSSFPictureData pictureData = hssfPicture.getPictureData();
        this.workbookDefinedPictureType = pictureData.getPictureType();
        // int format = pictureData.getFormat();// HSSF specific
        this.data = pictureData.getData();
        this.suggestFileExtension = pictureData.suggestFileExtension();
        this.mimeType = pictureData.getMimeType();
    }

    /**
     * 获取图片所在行号。
     *
     * @return 图片所在行号
     */
    public int getAtRow() {
        return atRow;
    }

    /**
     * 获取图片所在列号。
     *
     * @return 图片所在列号
     */
    public int getAtCol() {
        return atCol;
    }

    /**
     * 获取图片宽度。
     *
     * @return 图片宽度
     */
    public int getWidth() {
        return width;
    }

    /**
     * 获取图片高度。
     *
     * @return 图片高度
     */
    public int getHeight() {
        return height;
    }

    /**
     * 获取建议的文件扩展名。
     *
     * @return 建议的文件扩展名
     */
    public String getSuggestFileExtension() {
        return suggestFileExtension;
    }

    /**
     * 获取 MIME 类型。
     *
     * @return MIME 类型
     */
    public String getMimeType() {
        return mimeType;
    }

    /**
     * 获取图片数据。
     *
     * @return 图片数据字节数组
     */
    public byte[] getData() {
        return data;
    }

    /**
     * 获取工作簿定义的图片类型。
     *
     * @return 工作簿定义的图片类型
     */
    public int getWorkbookDefinedPictureType() {
        return workbookDefinedPictureType;
    }
}
