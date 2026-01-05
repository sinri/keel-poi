package io.github.sinri.keel.integration.poi.excel.entity;

import io.vertx.core.json.JsonObject;
import org.jspecify.annotations.NullMarked;

import java.util.List;

/**
 * Excel 表格模板化行接口，表示带有模板结构的表格行。
 * 该接口提供了基于列索引或列名获取列值的方法，以及将行数据转换为 JSON 对象的功能。
 *
 * @since 5.0.0
 */
@NullMarked
public interface KeelSheetMatrixTemplatedRow {
    /**
     * 创建模板化行实例。
     *
     * @param template 行模板
     * @param rawRow   原始行数据
     * @return 模板化行实例
     * @since 5.0.0
     */
    static KeelSheetMatrixTemplatedRow create(KeelSheetMatrixRowTemplate template, List<String> rawRow) {
        return new KeelSheetMatrixTemplatedRowImpl(template, rawRow);
    }

    /**
     * 获取行模板。
     *
     * @return 行模板
     * @since 5.0.0
     */
    KeelSheetMatrixRowTemplate getTemplate();

    /**
     * 获取指定索引处的列值。
     *
     * @param i 列索引
     * @return 指定索引处的列值
     * @since 5.0.0
     */
    String getColumnValue(int i);

    /**
     * 获取指定列名的列值。
     *
     * @param name 列名
     * @return 指定列名的列值
     * @since 5.0.0
     */
    String getColumnValue(String name);

    /**
     * 获取原始行数据。
     *
     * @return 原始行数据列表
     * @since 5.0.0
     */
    List<String> getRawRow();

    /**
     * 将行数据转换为 JSON 对象。
     *
     * @return 包含行数据的 JSON 对象
     * @since 5.0.0
     */
    JsonObject toJsonObject();
}
