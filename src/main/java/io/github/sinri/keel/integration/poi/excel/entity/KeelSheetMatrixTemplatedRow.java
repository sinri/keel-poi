package io.github.sinri.keel.integration.poi.excel.entity;

import io.vertx.core.json.JsonObject;
import org.jetbrains.annotations.NotNull;

import java.util.List;

/**
 * Excel 表格模板化行接口，表示带有模板结构的表格行。
 * 该接口提供了基于列索引或列名获取列值的方法，以及将行数据转换为 JSON 对象的功能。
 *
 * @since 5.0.0
 */
public interface KeelSheetMatrixTemplatedRow {
    static KeelSheetMatrixTemplatedRow create(@NotNull KeelSheetMatrixRowTemplate template, @NotNull List<String> rawRow) {
        return new KeelSheetMatrixTemplatedRowImpl(template, rawRow);
    }

    KeelSheetMatrixRowTemplate getTemplate();

    String getColumnValue(int i);

    String getColumnValue(String name);

    List<String> getRawRow();

    JsonObject toJsonObject();
}
