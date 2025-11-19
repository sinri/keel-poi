package io.github.sinri.keel.integration.poi.excel;

import org.jetbrains.annotations.NotNull;

import java.util.List;

/**
 * Excel 工作表行过滤器接口，用于过滤 Excel 工作表中的行。
 * <p>
 * 该接口允许用户定义自定义的行过滤逻辑。
 *
 * @since 5.0.0
 */
public interface SheetRowFilter {
    static SheetRowFilter toThrowEmptyRows() {
        return rawRow -> {
            boolean allEmpty = true;
            for (String cell : rawRow) {
                if (cell != null && !cell.isEmpty()) {
                    allEmpty = false;
                    break;
                }
            }
            return allEmpty;
        };
    }

    boolean shouldThrowThisRawRow(@NotNull List<String> rawRow);


}
