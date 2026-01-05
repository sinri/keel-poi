package io.github.sinri.keel.integration.poi.excel.entity;


import org.jspecify.annotations.NullMarked;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel 表格模板化矩阵实现类，实现模板化矩阵接口。
 * 该类提供了基于模板结构管理表格矩阵数据的具体实现。
 *
 * @since 5.0.0
 */
@NullMarked
public class KeelSheetTemplatedMatrixImpl implements KeelSheetTemplatedMatrix {
    private final KeelSheetMatrixRowTemplate template;
    private final List<List<String>> rawRows;
    //private final List<KeelSheetMatrixTemplatedRow> templatedRows;

    /**
     * 构造函数，使用指定的模板创建模板化矩阵实现。
     * 该构造函数为包级访问权限。
     *
     * @param template 行模板
     * @since 5.0.0
     */
    KeelSheetTemplatedMatrixImpl(KeelSheetMatrixRowTemplate template) {
        this.template = template;
        this.rawRows = new ArrayList<>();
        //        this.templatedRows = new ArrayList<>();
    }

    /**
     * 获取所有原始行数据列表。
     *
     * @return 原始行数据列表
     * @since 5.0.0
     */
    @Override
    public List<List<String>> getRawRows() {
        return rawRows;
    }

    /**
     * 获取矩阵模板。
     *
     * @return 矩阵模板
     * @since 5.0.0
     */
    @Override
    public KeelSheetMatrixRowTemplate getTemplate() {
        return template;
    }

    /**
     * 获取指定索引处的模板化行。
     *
     * @param index 行索引
     * @return 指定索引处的模板化行
     * @since 5.0.0
     */
    @Override
    public KeelSheetMatrixTemplatedRow getRow(int index) {
        return KeelSheetMatrixTemplatedRow.create(getTemplate(), this.rawRows.get(index));
        //        return this.templatedRows.get(index);
    }

    /**
     * 获取所有模板化行列表。
     *
     * @return 模板化行列表
     * @since 5.0.0
     */
    @Override
    public List<KeelSheetMatrixTemplatedRow> getRows() {
        List<KeelSheetMatrixTemplatedRow> templatedRows = new ArrayList<>();
        this.rawRows.forEach(rawRow -> templatedRows.add(KeelSheetMatrixTemplatedRow.create(getTemplate(), rawRow)));
        return templatedRows;
    }

    /**
     * 添加原始行数据。
     *
     * @param rawRow 原始行数据
     * @return 当前模板化矩阵实例，支持链式调用
     * @since 5.0.0
     */
    @Override
    public KeelSheetTemplatedMatrix addRawRow(List<String> rawRow) {
        this.rawRows.add(rawRow);
        //this.templatedRows.add(KeelSheetMatrixTemplatedRow.create(getTemplate(), rawRow));
        return this;
    }
}
