package io.github.sinri.keel.integration.poi.excel.entity;

import io.vertx.core.json.JsonObject;
import org.jspecify.annotations.NullMarked;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Excel 表格模板化行实现类，实现模板化行接口。
 * 该类提供了基于模板结构访问表格行数据的具体实现。
 *
 * @since 5.0.0
 */
@NullMarked
public class KeelSheetMatrixTemplatedRowImpl implements KeelSheetMatrixTemplatedRow {
    private final KeelSheetMatrixRowTemplate template;
    private final List<String> rawRow;

    /**
     * 构造函数，使用指定的模板和原始行数据创建模板化行实现。
     * 该构造函数为包级访问权限。
     * <p>
     * 构造时会保存传入行的副本；调用方后续修改原始列表不会影响本行对象。
     *
     * @param template 模板
     * @param rawRow   原始行数据
     * @since 5.0.0
     */
    KeelSheetMatrixTemplatedRowImpl(KeelSheetMatrixRowTemplate template, List<String> rawRow) {
        this.template = template;
        this.rawRow = new ArrayList<>(rawRow);
    }

    /**
     * 获取行模板。
     *
     * @return 行模板
     * @since 5.0.0
     */
    @Override
    public KeelSheetMatrixRowTemplate getTemplate() {
        return template;
    }

    /**
     * 获取指定索引处的列值。
     *
     * @param i 列索引
     * @return 指定索引处的列值
     * @since 5.0.0
     */
    @Override
    public String getColumnValue(int i) {
        if (i >= this.rawRow.size()) {
            return "";
        }
        return this.rawRow.get(i);
    }

    /**
     * 获取指定列名的列值。
     * 如果找不到指定的列名，则抛出运行时异常。
     *
     * @param name 列名
     * @return 指定列名的列值
     * @throws RuntimeException 如果找不到指定的列名
     * @since 5.0.0
     */
    @Override
    public String getColumnValue(String name) {
        Integer columnIndex = getTemplate().getColumnIndex(name);
        if (columnIndex == null) {
            throw new IllegalArgumentException("Unknown column name: " + name);
        }
        return getColumnValue(columnIndex);
    }

    /**
     * 获取原始行数据。
     * <p>
     * 返回的行列表为只读视图，不允许调用方通过该列表修改本行对象的内部状态。
     *
     * @return 原始行数据列表
     * @since 5.0.0
     */
    @Override
    public List<String> getRawRow() {
        return Collections.unmodifiableList(rawRow);
    }

    /**
     * 将行数据转换为 JSON 对象。
     * 该方法将使用模板中的列名作为键，行中的对应值作为值创建 JSON 对象。
     *
     * @return 包含行数据的 JSON 对象
     * @since 5.0.0
     */
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
