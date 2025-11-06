package io.github.sinri.keel.integration.poi.excel;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPictureData;

import javax.annotation.Nonnull;

public class KeelPictureInSheet {
    private final int atRow;
    private final int atCol;
    private final int width;
    private final int height;
    private final String suggestFileExtension;
    private final String mimeType;
    private final byte[] data;
    private final int workbookDefinedPictureType;

    public KeelPictureInSheet(@Nonnull XSSFPicture xssfPicture) {
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

    public KeelPictureInSheet(@Nonnull HSSFPicture hssfPicture) {
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

    public int getAtRow() {
        return atRow;
    }

    public int getAtCol() {
        return atCol;
    }

    public int getWidth() {
        return width;
    }

    public int getHeight() {
        return height;
    }

    public String getSuggestFileExtension() {
        return suggestFileExtension;
    }

    public String getMimeType() {
        return mimeType;
    }

    public byte[] getData() {
        return data;
    }

    public int getWorkbookDefinedPictureType() {
        return workbookDefinedPictureType;
    }
}
