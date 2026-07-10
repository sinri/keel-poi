package io.github.sinri.keel.integration.poi.excel.entity;

import io.github.sinri.keel.tesuto.KeelJUnit5Test;
import org.jspecify.annotations.NullMarked;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

@NullMarked
class KeelSheetMatrixRowTemplateTest extends KeelJUnit5Test {

    KeelSheetMatrixRowTemplateTest() {
        super();
    }

    @Test
    void duplicateColumnNameRejected() {
        assertThrows(
                IllegalArgumentException.class,
                () -> KeelSheetMatrixRowTemplate.create(Arrays.asList("a", "b", "a"))
        );
    }

    @Test
    void templateDefensivelyCopiesHeaderRowAndExposesReadonlyColumnNames() {
        List<String> headerRow = new ArrayList<>(List.of("A", "B"));

        KeelSheetMatrixRowTemplate template = KeelSheetMatrixRowTemplate.create(headerRow);
        headerRow.set(0, "changed");

        assertEquals("A", template.getColumnName(0));
        assertEquals(0, template.getColumnIndex("A"));
        assertThrows(UnsupportedOperationException.class, () -> template.getColumnNames().set(0, "x"));
    }

    @Test
    void templateNormalizesNullHeaderItemsConsistently() {
        KeelSheetMatrixRowTemplate template = KeelSheetMatrixRowTemplate.create(Arrays.asList("A", null));

        assertEquals("", template.getColumnName(1));
        assertEquals(1, template.getColumnIndex(""));
        assertEquals(List.of("A", ""), template.getColumnNames());
    }

    @Test
    void templatedRowReturnsBlankForTrailingMissingColumnByName() {
        KeelSheetMatrixRowTemplate template = KeelSheetMatrixRowTemplate.create(List.of("A", "B", "C"));
        KeelSheetMatrixTemplatedRow row = KeelSheetMatrixTemplatedRow.create(template, List.of("a", "b"));

        assertEquals("", row.getColumnValue(2));
        assertEquals("", row.getColumnValue("C"));

        IllegalArgumentException exception = assertThrows(
                IllegalArgumentException.class,
                () -> row.getColumnValue("D")
        );
        assertTrue(exception.getMessage().contains("D"));
    }

    @Test
    void matrixDefensivelyCopiesAddedRowsAndExposesReadonlyRows() {
        KeelSheetMatrix matrix = new KeelSheetMatrix();
        List<String> row = new ArrayList<>(List.of("a", "b"));

        matrix.addRow(row);
        row.set(0, "changed");

        assertEquals("a", matrix.getRawRow(0).get(0));
        assertThrows(UnsupportedOperationException.class, () -> matrix.getRawRow(0).set(0, "x"));
        assertThrows(UnsupportedOperationException.class, () -> matrix.getRawRowList().get(0).set(0, "x"));
    }

    @Test
    void matrixDefensivelyCopiesAddedRowGroups() {
        KeelSheetMatrix matrix = new KeelSheetMatrix();
        List<String> row = new ArrayList<>(List.of("a", "b"));
        List<List<String>> rows = new ArrayList<>();
        rows.add(row);

        matrix.addRows(rows);
        row.set(0, "changed");
        rows.clear();

        assertEquals(1, matrix.getRawRowList().size());
        assertEquals("a", matrix.getRawRow(0).get(0));
    }

    @Test
    void templatedMatrixDefensivelyCopiesRawRowsAndExposesReadonlyRows() {
        KeelSheetMatrixRowTemplate template = KeelSheetMatrixRowTemplate.create(List.of("A", "B"));
        KeelSheetTemplatedMatrix matrix = KeelSheetTemplatedMatrix.create(template);
        List<String> row = new ArrayList<>(List.of("a", "b"));

        matrix.addRawRow(row);
        row.set(0, "changed");

        assertEquals("a", matrix.getRawRows().get(0).get(0));
        assertThrows(UnsupportedOperationException.class, () -> matrix.getRawRows().get(0).set(0, "x"));
    }

    @Test
    void templatedRowDefensivelyCopiesRawRowAndExposesReadonlyRow() {
        KeelSheetMatrixRowTemplate template = KeelSheetMatrixRowTemplate.create(List.of("A", "B"));
        List<String> rawRow = new ArrayList<>(List.of("a", "b"));

        KeelSheetMatrixTemplatedRow row = KeelSheetMatrixTemplatedRow.create(template, rawRow);
        rawRow.set(0, "changed");

        assertEquals("a", row.getColumnValue("A"));
        assertThrows(UnsupportedOperationException.class, () -> row.getRawRow().set(0, "x"));
    }
}
