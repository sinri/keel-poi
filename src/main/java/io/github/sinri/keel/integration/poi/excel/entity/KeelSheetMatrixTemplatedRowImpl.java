package io.github.sinri.keel.integration.poi.excel.entity;

import io.vertx.core.json.JsonObject;
import org.jetbrains.annotations.NotNull;

import java.util.List;
import java.util.Objects;

/**
 * Excel 表格模板化行实现类，实现模板化行接口。
 * 该类提供了基于模板结构访问表格行数据的具体实现。
 *
 * @since 5.0.0
 */
public class KeelSheetMatrixTemplatedRowImpl implements KeelSheetMatrixTemplatedRow {
    private final KeelSheetMatrixRowTemplate template;
    private final List<String> rawRow;

    KeelSheetMatrixTemplatedRowImpl(@NotNull KeelSheetMatrixRowTemplate template, @NotNull List<String> rawRow) {
        this.template = template;
        this.rawRow = rawRow;
    }

    @Override
    public KeelSheetMatrixRowTemplate getTemplate() {
        return template;
    }

    @Override
    public String getColumnValue(int i) {
        return this.rawRow.get(i);
    }

    @Override
    public String getColumnValue(String name) {
        Integer columnIndex = getTemplate().getColumnIndex(name);
        return this.rawRow.get(Objects.requireNonNull(columnIndex));
    }

    /**
     * @since 3.2.16 Fix bug
     */
    @Override
    public List<String> getRawRow() {
        return rawRow;
    }

    @Override
    public JsonObject toJsonObject() {
        var x = new JsonObject();
        List<String> columnNames = this.template.getColumnNames();
        for (int i = 0; i < columnNames.size(); i++) {
            x.put(columnNames.get(i), getColumnValue(i));
        }
        return x;
    }
}
