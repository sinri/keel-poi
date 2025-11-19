package io.github.sinri.keel.integration.poi.excel.entity;

import org.jetbrains.annotations.NotNull;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel 表格模板化矩阵实现类，实现模板化矩阵接口。
 * 该类提供了基于模板结构管理表格矩阵数据的具体实现。
 *
 * @since 5.0.0
 */
public class KeelSheetTemplatedMatrixImpl implements KeelSheetTemplatedMatrix {
    private final KeelSheetMatrixRowTemplate template;
    private final List<List<String>> rawRows;
    //private final List<KeelSheetMatrixTemplatedRow> templatedRows;

    KeelSheetTemplatedMatrixImpl(@NotNull KeelSheetMatrixRowTemplate template) {
        this.template = template;
        this.rawRows = new ArrayList<>();
//        this.templatedRows = new ArrayList<>();
    }

    @Override
    public List<List<String>> getRawRows() {
        return rawRows;
    }

    @Override
    public KeelSheetMatrixRowTemplate getTemplate() {
        return template;
    }

    @Override
    public KeelSheetMatrixTemplatedRow getRow(int index) {
        return KeelSheetMatrixTemplatedRow.create(getTemplate(), this.rawRows.get(index));
//        return this.templatedRows.get(index);
    }

    @Override
    public List<KeelSheetMatrixTemplatedRow> getRows() {
        List<KeelSheetMatrixTemplatedRow> templatedRows = new ArrayList<>();
        this.rawRows.forEach(rawRow -> templatedRows.add(KeelSheetMatrixTemplatedRow.create(getTemplate(), rawRow)));
        return templatedRows;
    }

    @Override
    public KeelSheetTemplatedMatrix addRawRow(@NotNull List<String> rawRow) {
        this.rawRows.add(rawRow);
        //this.templatedRows.add(KeelSheetMatrixTemplatedRow.create(getTemplate(), rawRow));
        return this;
    }
}
