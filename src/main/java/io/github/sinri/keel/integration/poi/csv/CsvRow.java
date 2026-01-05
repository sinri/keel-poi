package io.github.sinri.keel.integration.poi.csv;

import org.jspecify.annotations.NullMarked;

import java.util.ArrayList;
import java.util.List;

/**
 * CSV 行类，表示 CSV 文件中的一行数据。
 *
 * @since 5.0.0
 */
@NullMarked
public class CsvRow {
    private final List<CsvCell> cells = new ArrayList<>();

    /**
     * 向此 CSV 行中添加一个单元格。
     *
     * @param cell 要添加的 CSV 单元格
     * @return 当前 CSV 行对象，支持链式调用
     */
    public CsvRow addCell(CsvCell cell) {
        this.cells.add(cell);
        return this;
    }

    /**
     * 获取指定索引处的 CSV 单元格。
     *
     * @param i 单元格的索引
     * @return 指定索引处的 CSV 单元格
     */
    public CsvCell getCell(int i) {
        return cells.get(i);
    }

    /**
     * 获取此 CSV 行中的单元格数量。
     *
     * @return 此 CSV 行中的单元格数量
     */
    public int size() {
        return this.cells.size();
    }
}
