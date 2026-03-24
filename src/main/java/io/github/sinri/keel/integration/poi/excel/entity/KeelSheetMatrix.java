package io.github.sinri.keel.integration.poi.excel.entity;

import org.jspecify.annotations.NullMarked;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.function.Function;

/**
 * Excel 表格矩阵类，将 Excel 表格数据转换为单元格字符串值的矩阵，支持自定义行读取器。
 * 该类用于表示 Excel 表格中的数据结构，包含表头行和数据行。
 *
 * @since 5.0.0
 */
@NullMarked
public class KeelSheetMatrix {
    private final List<String> headerRow;
    private final List<List<String>> rows;

    /**
     * 构造函数，创建一个空的 Excel 表格矩阵实例。
     * 初始化表头行和数据行列表为空列表。
     *
     * @since 5.0.0
     */
    public KeelSheetMatrix() {
        this.headerRow = new ArrayList<>();
        this.rows = new ArrayList<>();
    }

    /**
     * 向矩阵中添加一行数据。
     *
     * @param row 要添加的数据行
     * @return 当前矩阵实例，支持链式调用
     * @since 5.0.0
     */
    public KeelSheetMatrix addRow(List<String> row) {
        this.rows.add(row);
        return this;
    }

    /**
     * 向矩阵中添加多行数据。
     *
     * @param rows 要添加的数据行列表
     * @return 当前矩阵实例，支持链式调用
     * @since 5.0.0
     */
    public KeelSheetMatrix addRows(List<List<String>> rows) {
        this.rows.addAll(rows);
        return this;
    }

    /**
     * 获取表头行数据。
     *
     * @return 表头行数据列表
     * @since 5.0.0
     */
    public List<String> getHeaderRow() {
        return Collections.unmodifiableList(headerRow);
    }

    /**
     * 设置表头行数据。
     * 该方法会清空当前的表头行数据，然后添加新的表头行数据。
     *
     * @param headerRow 新的表头行数据列表
     * @return 当前矩阵实例，支持链式调用
     * @since 5.0.0
     */
    public KeelSheetMatrix setHeaderRow(List<String> headerRow) {
        this.headerRow.clear();
        this.headerRow.addAll(headerRow);
        return this;
    }

    /**
     * 获取指定索引的原始行数据。
     *
     * @param i 行数据的索引
     * @return 指定索引的原始行数据
     * @since 5.0.0
     */
    public List<String> getRawRow(int i) {
        return this.rows.get(i);
    }

    /**
     * 获取所有原始行数据列表。
     *
     * @return 原始行数据列表
     * @since 5.0.0
     */
    public List<List<String>> getRawRowList() {
        return Collections.unmodifiableList(rows);
    }

    /**
     * 将当前矩阵转换为模板化矩阵。
     * 如果表头行为空，则抛出运行时异常。
     *
     * @return 转换后的模板化矩阵
     * @throws RuntimeException 如果表头行为空（列未定义）
     * @since 5.0.0
     */
    public KeelSheetTemplatedMatrix transformToTemplatedMatrix() {
        List<String> x = getHeaderRow();
        if (x.isEmpty()) throw new RuntimeException("Columns not defined");
        KeelSheetTemplatedMatrix templatedMatrix = KeelSheetTemplatedMatrix.create(KeelSheetMatrixRowTemplate.create(x));
        templatedMatrix.addRawRows(getRawRowList());
        return templatedMatrix;
    }

    /**
     * 获取行迭代器，用于遍历矩阵中的所有行。
     * 该方法使用指定的工厂函数将原始行数据转换为自定义行对象。
     *
     * @param rowFactory 将原始行数据转换为行对象的工厂函数
     * @return 行迭代器
     * @since 5.0.1
     */
    public <R extends KeelSheetMatrixRow> Iterator<R> getRowIterator(Function<List<String>, R> rowFactory) {
        return new RowReaderIterator<>(rowFactory, rows);
    }

    /**
     * 获取行迭代器，用于遍历矩阵中的所有行。
     * 该方法使用默认的 KeelSheetMatrixRow 类来创建迭代器。
     *
     * @return 行迭代器
     * @since 5.0.0
     */
    public Iterator<KeelSheetMatrixRow> getRowIterator() {
        return new RowReaderIterator<>(KeelSheetMatrixRow::new, rows);
    }

    /**
     * 行读取器迭代器类，用于遍历表格矩阵中的行。
     * 该类实现了迭代器接口，允许逐行访问矩阵数据。
     * <p>
     * 与 {@link java.util.ArrayList} 一样，本迭代器非线程安全；不得在并发环境下共享同一实例。
     *
     * @since 5.0.0
     */
    public static class RowReaderIterator<R extends KeelSheetMatrixRow> implements Iterator<R> {
        private final List<List<String>> rows;
        private int nextIndex;
        private final Function<List<String>, R> rawRow2row;

        /**
         * 构造函数，使用指定的转换函数和行数据列表创建行读取器迭代器。
         *
         * @param rawRow2row 将原始行转换为 KeelSheetMatrixRow 实例的函数
         * @param rows       原始行数据列表
         * @since 5.0.0
         */
        public RowReaderIterator(Function<List<String>, R> rawRow2row, List<List<String>> rows) {
            this.rawRow2row = rawRow2row;
            this.rows = rows;
        }


        /**
         * 检查迭代器是否还有下一个元素。
         *
         * @return 如果还有下一个元素则返回 true，否则返回 false
         * @since 5.0.0
         */
        @Override
        public boolean hasNext() {
            return this.rows.size() > nextIndex;
        }


        /**
         * 获取迭代器的下一个元素。
         *
         * @return 下一个元素
         * @since 5.0.0
         */
        @Override
        public R next() {
            if (!hasNext()) {
                throw new NoSuchElementException("No more rows in matrix (size=" + rows.size() + ")");
            }
            List<String> rawRow = this.rows.get(nextIndex);
            nextIndex++;
            return this.rawRow2row.apply(rawRow);
        }
    }
}
