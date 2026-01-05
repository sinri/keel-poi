package io.github.sinri.keel.integration.poi.excel;

import org.jspecify.annotations.NullMarked;

import java.util.List;

/**
 * Excel 工作表行过滤器接口，用于过滤 Excel 工作表中的行。
 * <p>
 * 该接口允许用户定义自定义的行过滤逻辑。
 *
 * @since 5.0.0
 */
@NullMarked
public interface SheetRowFilter {
    /**
     * 创建一个过滤器，用于过滤掉所有单元格都为空的行。
     *
     * @return 一个过滤器实例，该过滤器会判断一行是否所有单元格都为空
     */
    static SheetRowFilter toThrowEmptyRows() {
        return rawRow -> {
            boolean allEmpty = true;
            for (String cell : rawRow) {
                if (!cell.isEmpty()) {
                    allEmpty = false;
                    break;
                }
            }
            return allEmpty;
        };
    }

    /**
     * 判断是否应该过滤掉当前行。
     *
     * @param rawRow 原始行数据，包含该行所有单元格的内容
     * @return 如果应该过滤掉此行则返回 true，否则返回 false
     */
    boolean shouldThrowThisRawRow(List<String> rawRow);


}
