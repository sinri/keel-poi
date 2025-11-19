package io.github.sinri.keel.integration.poi.excel.entity;

import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * Excel 表格矩阵行模板实现类，实现列结构定义的接口。
 * 该类提供了列名到列索引的映射功能。
 *
 * @since 5.0.0
 */
public class KeelSheetMatrixRowTemplateImpl implements KeelSheetMatrixRowTemplate {
    private final List<String> headerRow;
    private final Map<String, Integer> headerMap;

    KeelSheetMatrixRowTemplateImpl(@NotNull List<String> headerRow) {
        this.headerRow = headerRow;
        this.headerMap = new LinkedHashMap<>();
        for (int i = 0; i < headerRow.size(); i++) {
            this.headerMap.put(Objects.requireNonNullElse(headerRow.get(i), ""), i);
        }
    }

    @NotNull
    @Override
    public String getColumnName(int i) {
        return this.headerRow.get(i);
    }

    @Nullable
    @Override
    public Integer getColumnIndex(String name) {
        return this.headerMap.get(name);
    }

    @NotNull
    @Override
    public List<String> getColumnNames() {
        return headerRow;
    }
}
