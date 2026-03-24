package io.github.sinri.keel.integration.poi.excel.entity;

import io.github.sinri.keel.tesuto.KeelJUnit5Test;
import org.jspecify.annotations.NullMarked;
import org.junit.jupiter.api.Test;

import java.util.Arrays;

import static org.junit.jupiter.api.Assertions.assertThrows;

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
}
