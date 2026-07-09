package io.github.sinri.keel.integration.poi.excel.entity;

import io.github.sinri.keel.tesuto.KeelJUnit5Test;
import org.jspecify.annotations.NullMarked;
import org.junit.jupiter.api.Test;

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
}
