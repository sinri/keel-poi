package io.github.sinri.keel.integration.poi.csv;

import org.jetbrains.annotations.NotNull;

import java.util.ArrayList;
import java.util.List;

/**
 * @since 3.1.1 Technical Preview
 */
public class CsvRow {
    private final List<CsvCell> cells = new ArrayList<>();

    public CsvRow addCell(@NotNull CsvCell cell) {
        this.cells.add(cell);
        return this;
    }

    @NotNull
    public CsvCell getCell(int i) {
        return cells.get(i);
    }

    public int size() {
        return this.cells.size();
    }
}
