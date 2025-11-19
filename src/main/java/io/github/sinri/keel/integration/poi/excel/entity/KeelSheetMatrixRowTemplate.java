package io.github.sinri.keel.integration.poi.excel.entity;

import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.util.List;

/**
 * Excel 表格矩阵行模板接口，定义模板化矩阵中行的列结构。
 * 该接口用于定义表格中的列名和列索引的映射关系。
 *
 * @since 5.0.0
 */
public interface KeelSheetMatrixRowTemplate {
    static KeelSheetMatrixRowTemplate create(@NotNull List<String> headerRow) {
        return new KeelSheetMatrixRowTemplateImpl(headerRow);
    }

    /**
     * @param i Column index start from 0.
     * @return Column name at index.
     * @throws RuntimeException if index is out of bound
     */
    @NotNull
    String getColumnName(int i);

    /**
     * @param name the column name to seek.
     * @return The first met (or by customized logic) index of the given column name, and null if not found.
     */
    @Nullable
    Integer getColumnIndex(String name);

    @NotNull
    List<String> getColumnNames();
}
