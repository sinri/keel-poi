package io.github.sinri.keel.integration.poi.excel.entity;

import org.jetbrains.annotations.NotNull;

import java.util.List;

/**
 * Excel 表格模板化矩阵接口，表示带有模板结构的表格矩阵。
 * 该接口提供了基于模板访问表格矩阵数据的方法。
 *
 * @since 5.0.0
 */
public interface KeelSheetTemplatedMatrix {
    /**
     * 创建模板化矩阵实例。
     *
     * @param template 行模板
     * @return 模板化矩阵实例
     * @since 5.0.0
     */
    static KeelSheetTemplatedMatrix create(@NotNull KeelSheetMatrixRowTemplate template) {
        return new KeelSheetTemplatedMatrixImpl(template);
    }

    /**
     * 获取矩阵模板。
     *
     * @return 矩阵模板
     * @since 5.0.0
     */
    KeelSheetMatrixRowTemplate getTemplate();

    /**
     * 获取指定索引处的模板化行。
     *
     * @param index 行索引
     * @return 指定索引处的模板化行
     * @since 5.0.0
     */
    KeelSheetMatrixTemplatedRow getRow(int index);

    /**
     * 获取所有模板化行列表。
     *
     * @return 模板化行列表
     * @since 5.0.0
     */
    List<KeelSheetMatrixTemplatedRow> getRows();

    /**
     * 获取所有原始行数据列表。
     *
     * @return 原始行数据列表
     * @since 5.0.0
     */
    List<List<String>> getRawRows();

    /**
     * 添加原始行数据。
     *
     * @param rawRow 原始行数据
     * @return 当前模板化矩阵实例，支持链式调用
     * @since 5.0.0
     */
    KeelSheetTemplatedMatrix addRawRow(@NotNull List<String> rawRow);

    /**
     * 添加多个原始行数据。
     *
     * @param rawRows 原始行数据列表
     * @return 当前模板化矩阵实例，支持链式调用
     * @since 5.0.0
     */
    default KeelSheetTemplatedMatrix addRawRows(@NotNull List<List<String>> rawRows) {
        rawRows.forEach(this::addRawRow);
        return this;
    }

    /**
     * 将模板化矩阵转换为普通矩阵。
     *
     * @return 转换后的普通矩阵
     * @since 5.0.0
     */
    default KeelSheetMatrix transformToMatrix() {
        KeelSheetMatrix keelSheetMatrix = new KeelSheetMatrix();
        keelSheetMatrix.setHeaderRow(getTemplate().getColumnNames());
        keelSheetMatrix.addRows(getRawRows());
        return keelSheetMatrix;
    }

}
