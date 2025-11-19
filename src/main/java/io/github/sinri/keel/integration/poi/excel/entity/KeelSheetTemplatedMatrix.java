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
    static KeelSheetTemplatedMatrix create(@NotNull KeelSheetMatrixRowTemplate template) {
        return new KeelSheetTemplatedMatrixImpl(template);
    }

    KeelSheetMatrixRowTemplate getTemplate();

    KeelSheetMatrixTemplatedRow getRow(int index);

    List<KeelSheetMatrixTemplatedRow> getRows();

    List<List<String>> getRawRows();

    KeelSheetTemplatedMatrix addRawRow(@NotNull List<String> rawRow);

    default KeelSheetTemplatedMatrix addRawRows(@NotNull List<List<String>> rawRows) {
        rawRows.forEach(this::addRawRow);
        return this;
    }

    default KeelSheetMatrix transformToMatrix() {
        KeelSheetMatrix keelSheetMatrix = new KeelSheetMatrix();
        keelSheetMatrix.setHeaderRow(getTemplate().getColumnNames());
        keelSheetMatrix.addRows(getRawRows());
        return keelSheetMatrix;
    }

}
