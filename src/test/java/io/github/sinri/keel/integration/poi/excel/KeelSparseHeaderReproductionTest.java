package io.github.sinri.keel.integration.poi.excel;

import io.github.sinri.keel.base.async.Keel;
import io.github.sinri.keel.core.utils.value.ValueBox;
import io.github.sinri.keel.integration.poi.excel.entity.KeelSheetMatrix;
import io.vertx.core.Vertx;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.DynamicTest;
import org.junit.jupiter.api.TestFactory;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeUnit;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Issue #9 的复现测试：断言正确结果，稀疏行场景在修复前应失败。
 * 每次运行生成可独立打开的 XLSX 物料到 build/repro/issue-9。
 */
class KeelSparseHeaderReproductionTest {
    private static Vertx vertx;
    private static Keel keel;
    private static final Path MATERIALS = Path.of("build", "repro", "issue-9");

    @BeforeAll
    static void prepare() throws Exception {
        vertx = Vertx.vertx();
        keel = Keel.create(vertx);
        Files.createDirectories(MATERIALS);
        for (String scenario : List.of("continuous", "leading-missing", "middle-missing")) {
            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("S");
                if (!scenario.equals("leading-missing")) {
                    sheet.createRow(0).createCell(0).setCellValue("preamble");
                }
                if (scenario.equals("continuous")) {
                    sheet.createRow(1).createCell(0).setCellValue("preamble-2");
                }
                sheet.createRow(2).createCell(0).setCellValue("header");
                sheet.createRow(3).createCell(0).setCellValue("data");
                try (var output = Files.newOutputStream(MATERIALS.resolve(scenario + ".xlsx"))) {
                    workbook.write(output);
                }
            }
        }
    }

    @AfterAll
    static void closeVertx() throws Exception {
        if (vertx != null) {
            vertx.close().toCompletionStage().toCompletableFuture().get(30, TimeUnit.SECONDS);
        }
    }

    @TestFactory
    List<DynamicTest> physicalHeaderRowMustBeUsed() {
        List<DynamicTest> tests = new ArrayList<>();
        for (String scenario : List.of("continuous", "leading-missing", "middle-missing")) {
            for (String mode : List.of("matrix-sync", "template-sync", "matrix-async", "template-async")) {
                tests.add(DynamicTest.dynamicTest(scenario + " / " + mode, () -> {
                    try (XSSFWorkbook workbook = new XSSFWorkbook(MATERIALS.resolve(scenario + ".xlsx").toFile())) {
                        Sheet sheet = workbook.getSheet("S");
                        // 验证落盘后仍是未创建的行，而不是已创建的空行。
                        if (!scenario.equals("continuous")) assertNull(sheet.getRow(1));
                        if (scenario.equals("leading-missing")) assertNull(sheet.getRow(0));
                        assertEquals("header", sheet.getRow(2).getCell(0).getStringCellValue());
                        KeelSheet reader = new KeelSheet(null, sheet, new ValueBox<>());
                        KeelSheetMatrix matrix = readMatrix(reader, mode, 1);
                        System.out.println(scenario + " / " + mode + ": header=" + matrix.getHeaderRow()
                                + ", rows=" + matrix.getRawRowList());
                        assertAll(
                                () -> assertEquals(List.of("header"), matrix.getHeaderRow()),
                                () -> assertEquals(List.of(List.of("data")), matrix.getRawRowList())
                        );
                    }
                }));
            }
        }
        return tests;
    }

    @TestFactory
    List<DynamicTest> missingHeaderMustFailClearly() {
        List<DynamicTest> tests = new ArrayList<>();
        for (String scenario : List.of("empty", "beyond-last-row", "missing-between-rows")) {
            for (String mode : List.of("matrix-sync", "template-sync", "matrix-async", "template-async")) {
                tests.add(DynamicTest.dynamicTest(scenario + " / " + mode, () -> {
                    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                        Sheet sheet = workbook.createSheet("MissingHeader");
                        if (!scenario.equals("empty")) {
                            sheet.createRow(0).createCell(0).setCellValue("preamble");
                        }
                        if (scenario.equals("missing-between-rows")) {
                            sheet.createRow(3).createCell(0).setCellValue("data");
                        }
                        KeelSheet reader = new KeelSheet(null, sheet, new ValueBox<>());
                        Throwable failure;
                        if (mode.endsWith("async")) {
                            failure = assertThrows(ExecutionException.class,
                                    () -> readMatrix(reader, mode, 0)).getCause();
                        } else {
                            failure = assertThrows(IllegalArgumentException.class,
                                    () -> readMatrix(reader, mode, 0));
                        }
                        assertInstanceOf(IllegalArgumentException.class, failure);
                        assertEquals("Header row not found: sheet=MissingHeader, headerRowIndex=2 (0-based)",
                                failure.getMessage());
                    }
                }));
            }
        }
        return tests;
    }

    @TestFactory
    List<DynamicTest> streamingReaderMustUsePhysicalHeaderAndDetectColumns() {
        List<DynamicTest> tests = new ArrayList<>();
        for (String scenario : List.of("continuous", "leading-missing", "middle-missing")) {
            for (String mode : List.of("matrix-sync", "template-sync", "matrix-async", "template-async")) {
                tests.add(DynamicTest.dynamicTest("streaming / " + scenario + " / " + mode, () -> {
                    // 使用同一份落盘物料，验证流式读取不依赖随机访问表头。
                    try (var workbook = com.github.pjfanning.xlsx.StreamingReader.builder()
                            .rowCacheSize(2).open(MATERIALS.resolve(scenario + ".xlsx").toFile())) {
                        KeelSheet reader = new KeelSheet(null, workbook.getSheet("S"), new ValueBox<>());
                        KeelSheetMatrix matrix = readMatrix(reader, mode, 0);
                        assertEquals(List.of("header"), matrix.getHeaderRow());
                        assertEquals(List.of(List.of("data")), matrix.getRawRowList());
                    }
                }));
            }
        }
        return tests;
    }

    private static KeelSheetMatrix readMatrix(KeelSheet reader, String mode, int maxColumns) throws Exception {
        return switch (mode) {
            case "matrix-sync" -> reader.readAllRowsToMatrix(2, maxColumns, null);
            case "template-sync" -> reader.readAllRowsToTemplatedMatrix(2, maxColumns, null).transformToMatrix();
            case "matrix-async" -> reader.readAllRowsToMatrixAsync(keel, 2, maxColumns, null)
                    .toCompletionStage().toCompletableFuture().get(30, TimeUnit.SECONDS);
            case "template-async" -> reader.readAllRowsToTemplatedMatrixAsync(keel, 2, maxColumns, null)
                    .toCompletionStage().toCompletableFuture().get(30, TimeUnit.SECONDS).transformToMatrix();
            default -> throw new IllegalArgumentException(mode);
        };
    }

}
