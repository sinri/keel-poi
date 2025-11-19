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
    /**
     * 创建行模板实例。
     *
     * @param headerRow 表头行数据列表
     * @return 行模板实例
     * @since 5.0.0
     */
    static KeelSheetMatrixRowTemplate create(@NotNull List<String> headerRow) {
        return new KeelSheetMatrixRowTemplateImpl(headerRow);
    }

    /**
     * 获取指定索引处的列名。
     *
     * @param i 列索引，从 0 开始
     * @return 指定索引处的列名
     * @throws RuntimeException 如果索引超出范围
     * @since 5.0.0
     */
    @NotNull
    String getColumnName(int i);

    /**
     * 获取指定列名的索引。
     *
     * @param name 要查找的列名
     * @return 给定列名的第一个匹配索引（或按自定义逻辑），如果未找到则返回 null
     * @since 5.0.0
     */
    @Nullable
    Integer getColumnIndex(String name);

    /**
     * 获取所有列名列表。
     *
     * @return 列名列表
     * @since 5.0.0
     */
    @NotNull
    List<String> getColumnNames();
}
