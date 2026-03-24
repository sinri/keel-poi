package io.github.sinri.keel.integration.poi.excel;

import io.github.sinri.keel.core.utils.value.ValueBox;
import io.github.sinri.keel.integration.poi.excel.entity.KeelSheetMatrix;
import io.github.sinri.keel.tesuto.KeelJUnit5Test;
import io.vertx.core.Future;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.NullMarked;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.TimeUnit;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Excel 集成模块基础回归测试（读写、单元格字符串化、矩阵写入）。
 */
@NullMarked
class KeelExcelCoreTest extends KeelJUnit5Test {

    KeelExcelCoreTest() {
        super();
    }

    @Test
    void dumpCellToStringBooleanBlankAndNumeric() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Row row = wb.createSheet("s").createRow(0);
            ValueBox<FormulaEvaluator> box = new ValueBox<>();

            Cell boolCell = row.createCell(0);
            boolCell.setCellValue(true);
            assertEquals("true", KeelSheet.dumpCellToString(boolCell, box));

            Cell blankCell = row.createCell(1);
            blankCell.setBlank();
            assertEquals("", KeelSheet.dumpCellToString(blankCell, box));

            Cell numCell = row.createCell(2);
            numCell.setCellValue(3.5);
            assertEquals("3.5", KeelSheet.dumpCellToString(numCell, box));
        }
    }

    @Test
    void dumpCellToStringErrorCell() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Row row = wb.createSheet("s").createRow(0);
            Cell c = row.createCell(0);
            c.setCellErrorValue(FormulaError.NA.getCode());
            ValueBox<FormulaEvaluator> box = new ValueBox<>();
            String s = KeelSheet.dumpCellToString(c, box);
            assertFalse(s.isEmpty());
        }
    }

    @Test
    void createWriteMatrixAndSaveFile() throws Exception {
        Path tmp = Files.createTempFile("keel-poi-excel-", ".xlsx");
        tmp.toFile().deleteOnExit();

        SheetsCreateOptions opts = new SheetsCreateOptions()
                .setUseStreamWriting(false)
                .setWithFormulaEvaluator(false);

        Future<Void> done = KeelSheets.useSheets(opts, keelSheets -> {
            KeelSheet sheet = keelSheets.generateWriterForSheet("Data");
            KeelSheetMatrix matrix = new KeelSheetMatrix();
            matrix.setHeaderRow(List.of("H1", "H2"));
            matrix.addRow(List.of("a", "b"));
            matrix.addRow(List.of("c", "d"));
            sheet.writeMatrix(matrix);
            keelSheets.save(tmp.toFile());
            return Future.succeededFuture();
        });

        done.toCompletionStage().toCompletableFuture().get(30, TimeUnit.SECONDS);
        assertTrue(Files.size(tmp) > 50, "saved xlsx should have non-trivial size");

        try (XSSFWorkbook wb = new XSSFWorkbook(tmp.toFile())) {
            Sheet s = wb.getSheet("Data");
            assertNotNull(s);
            assertEquals("H1", s.getRow(0).getCell(0).getStringCellValue());
            assertEquals("d", s.getRow(2).getCell(1).getStringCellValue());
        }
    }
}
