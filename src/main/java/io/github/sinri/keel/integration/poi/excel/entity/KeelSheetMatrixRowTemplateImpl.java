package io.github.sinri.keel.integration.poi.excel.entity;

import org.jspecify.annotations.NullMarked;
import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.Collections;
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
@NullMarked
public class KeelSheetMatrixRowTemplateImpl implements KeelSheetMatrixRowTemplate {
    private final List<String> headerRow;
    private final Map<String, Integer> headerMap;

    /**
     * 构造函数，使用指定的表头行数据创建行模板实现。
     * 该构造函数为包级访问权限。
     * <p>
     * 构造时会复制并归一化表头行；调用方后续修改原始列表不会影响模板内部状态。
     *
     * @param headerRow 表头行数据列表
     * @since 5.0.0
     */
    KeelSheetMatrixRowTemplateImpl(List<String> headerRow) {
        this.headerRow = new ArrayList<>(headerRow.size());
        this.headerMap = new LinkedHashMap<>();
        for (int i = 0; i < headerRow.size(); i++) {
            String key = Objects.requireNonNullElse(headerRow.get(i), "");
            this.headerRow.add(key);
            if (headerMap.containsKey(key)) {
                throw new IllegalArgumentException(
                        "Duplicate column name in header: \"" + key + "\" at index " + i
                );
            }
            headerMap.put(key, i);
        }
    }

    /**
     * 获取指定索引处的列名。
     *
     * @param i 列索引
     * @return 指定索引处的列名
     * @since 5.0.0
     */
    @Override
    public String getColumnName(int i) {
        return this.headerRow.get(i);
    }

    /**
     * 获取指定列名的索引。
     *
     * @param name 列名
     * @return 列索引，如果未找到则返回 null
     * @since 5.0.0
     */

    @Override
    public @Nullable Integer getColumnIndex(String name) {
        return this.headerMap.get(name);
    }

    /**
     * 获取所有列名列表。
     * <p>
     * 返回的列名列表为只读视图，不允许调用方通过该列表修改模板内部状态。
     *
     * @return 列名列表
     * @since 5.0.0
     */
    @Override
    public List<String> getColumnNames() {
        return Collections.unmodifiableList(headerRow);
    }
}
