package io.github.sinri.keel.integration.poi.excel;

import io.github.sinri.keel.core.utils.value.ValueBox;
import io.github.sinri.keel.tesuto.KeelJUnit5Test;
import io.vertx.core.Future;
import io.vertx.core.Vertx;
import io.vertx.junit5.VertxExtension;
import io.vertx.junit5.VertxTestContext;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;

import java.io.File;
import java.io.IOException;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

@ExtendWith(VertxExtension.class)
class KeelSheetsTest extends KeelJUnit5Test {
    private static final String TEST_XLSX_FILE_PATH = "/Users/sinri/code/keel/src/test/resources/runtime/with-pic.xlsx";
    private static final String TEST_XLS_FILE_PATH = "/Users/sinri/code/keel/src/test/resources/runtime/with-pic.xls";
    private static final String TEST_OUTPUT_DIR = "/Users/sinri/code/keel/src/test/resources/runtime/output";

    private File testOutputDir;

    KeelSheetsTest(Vertx vertx) {
        super(vertx);
    }

    @Override
    protected void test(VertxTestContext testContext) {
        testContext.completeNow();
    }

    @BeforeEach
    void setUp() {
        testOutputDir = new File(TEST_OUTPUT_DIR);
        if (!testOutputDir.exists()) {
            testOutputDir.mkdirs();
        }
    }

    @AfterEach
    void tearDown() {
        if (testOutputDir.exists()) {
            deleteDirectory(testOutputDir);
        }
    }

    private void deleteDirectory(File dir) {
        File[] files = dir.listFiles();
        if (files != null) {
            for (File file : files) {
                if (file.isDirectory()) {
                    deleteDirectory(file);
                } else {
                    file.delete();
                }
            }
        }
        dir.delete();
    }


    @Test
    @DisplayName("测试文件操作 - 打开、读取、保存")
    void testFileOperations() {
        // 测试打开XLSX文件
        File testFile = new File(TEST_XLSX_FILE_PATH);
        assertTrue(testFile.exists(), "测试文件应该存在");

        KeelSheets.useSheets(
                new SheetsOpenOptions().setFile(testFile),
                keelSheets -> {
                    assertNotNull(keelSheets);
                    assertNotNull(keelSheets.getWorkbook());
                    assertEquals(1, keelSheets.getSheetCount());

                    // 读取工作表数据
                    KeelSheet sheet = keelSheets.generateReaderForSheet(0);
                    assertNotNull(sheet);

                    return sheet.readAllRowsToMatrix()
                                .compose(matrix -> {
                                    assertNotNull(matrix);
                                    List<List<String>> rows = matrix.getRawRowList();
                                    assertFalse(rows.isEmpty(), "应该能读取到行数据");

                                    getUnitTestLogger().info(r -> r.message("成功读取工作表数据，行数: " + rows.size()));

                                    // 读取第一行详细信息
                                    Row firstRow = sheet.readRow(0);
                                    if (firstRow != null) {
                                        getUnitTestLogger().info(r -> r.message("第一行单元格数: " + firstRow.getLastCellNum()));

                                        for (int colIndex = 0; colIndex < firstRow.getLastCellNum(); colIndex++) {
                                            final int index = colIndex;
                                            Cell cell = firstRow.getCell(index);
                                            if (cell != null) {
                                                // 使用KeelSheet的公共方法获取单元格值
                                                String cellValue = KeelSheet.dumpCellToString(cell, new ValueBox<>(sheet.getFormulaEvaluator()));
                                                getUnitTestLogger().info(r -> r.message("单元格[" + index + "]: " + cellValue));
                                            }
                                        }
                                    }

                                    return Future.succeededFuture();
                                });
                }
        );
    }

    @Test
    @DisplayName("测试XLS格式文件")
    void testXlsFile() {
        File testFile = new File(TEST_XLS_FILE_PATH);
        assertTrue(testFile.exists(), "测试文件应该存在");

        KeelSheets.useSheets(
                new SheetsOpenOptions().setFile(testFile),
                keelSheets -> {
                    assertNotNull(keelSheets);
                    assertNotNull(keelSheets.getWorkbook());
                    assertEquals(1, keelSheets.getSheetCount());

                    getUnitTestLogger().info(r -> r.message("成功打开XLS文件，工作表数量: " + keelSheets.getSheetCount()));
                    return Future.succeededFuture();
                }
        );
    }

    @Test
    @DisplayName("测试工作簿创建和操作")
    void testWorkbookCreationAndOperations() {
        KeelSheets.useSheets(
                new SheetsCreateOptions().setUseXlsx(true),
                keelSheets -> {
                    assertNotNull(keelSheets);
                    assertEquals(0, keelSheets.getSheetCount());

                    // 添加工作表
                    KeelSheet sheet1 = keelSheets.generateWriterForSheet("Sheet1");
                    KeelSheet sheet2 = keelSheets.generateWriterForSheet("Sheet2", 0);
                    KeelSheet sheet3 = keelSheets.generateWriterForSheet("ThirdSheet");

                    assertEquals(3, keelSheets.getSheetCount());
                    assertEquals("Sheet2", keelSheets.getWorkbook().getSheetName(0));
                    assertEquals("Sheet1", keelSheets.getWorkbook().getSheetName(1));
                    assertEquals("ThirdSheet", keelSheets.getWorkbook().getSheetName(2));

                    // 测试通过名称获取工作表
                    KeelSheet reader1 = keelSheets.generateReaderForSheet("Sheet1");
                    KeelSheet reader2 = keelSheets.generateReaderForSheet("Sheet2");

                    assertNotNull(reader1);
                    assertNotNull(reader2);
                    assertEquals("Sheet1", reader1.getName());
                    assertEquals("Sheet2", reader2.getName());

                    getUnitTestLogger().info(r -> r.message("成功测试工作表操作"));
                    return Future.succeededFuture();
                }
        );
    }

    @Test
    @DisplayName("测试文件保存和图片读取")
    void testFileSaveAndPictureReading() {
        String outputPath = TEST_OUTPUT_DIR + "/test_output.xlsx";
        File outputFile = new File(outputPath);

        // 先创建并保存文件
        KeelSheets.useSheets(
                new SheetsCreateOptions().setUseXlsx(true),
                keelSheets -> {
                    KeelSheet sheet = keelSheets.generateWriterForSheet("TestSheet");

                    // 添加测试数据
                    Row row1 = sheet.getSheet().createRow(0);
                    row1.createCell(0).setCellValue("Name");
                    row1.createCell(1).setCellValue("Age");

                    Row row2 = sheet.getSheet().createRow(1);
                    row2.createCell(0).setCellValue("John");
                    row2.createCell(1).setCellValue(25);

                    keelSheets.save(outputPath);
                    assertTrue(outputFile.exists(), "输出文件应该被创建");

                    getUnitTestLogger().info(r -> r.message("成功保存工作簿到文件: " + outputPath));
                    return Future.succeededFuture();
                }
        );

        // 然后读取原文件测试图片功能
        KeelSheets.useSheets(
                new SheetsOpenOptions().setFile(TEST_XLSX_FILE_PATH),
                keelSheets -> {
                    KeelSheet sheet = keelSheets.generateReaderForSheet(0);
                    List<KeelPictureInSheet> pictures = sheet.getPictures();
                    assertNotNull(pictures);

                    getUnitTestLogger().info(r -> r.message("工作表包含图片数量: " + pictures.size()));

                    // 验证图片数据
                    for (int picIndex = 0; picIndex < pictures.size(); picIndex++) {
                        final int index = picIndex;
                        KeelPictureInSheet picture = pictures.get(index);
                        getUnitTestLogger().info(r -> r.message("图片[" + index + "]: " + picture.toString()));

                        assertNotNull(picture.getData());
                        assertTrue(picture.getData().length > 0, "图片数据应该不为空");
                        assertNotNull(picture.getMimeType());
                        assertNotNull(picture.getSuggestFileExtension());
                    }

                    return Future.succeededFuture();
                }
        );
    }

    @Test
    @DisplayName("测试高级功能 - 公式求值器和流式写入")
    void testAdvancedFeatures() {
        // 测试公式求值器
        KeelSheets.useSheets(
                new SheetsCreateOptions().setUseXlsx(true).setWithFormulaEvaluator(true),
                keelSheets -> {
                    KeelSheet sheet = keelSheets.generateWriterForSheet("FormulaTest");

                    Row row1 = sheet.getSheet().createRow(0);
                    row1.createCell(0).setCellValue("A");
                    row1.createCell(1).setCellValue("B");
                    row1.createCell(2).setCellValue("Sum");

                    Row row2 = sheet.getSheet().createRow(1);
                    row2.createCell(0).setCellValue(10);
                    row2.createCell(1).setCellValue(20);
                    row2.createCell(2).setCellFormula("A2+B2");

                    assertNotNull(keelSheets.getWorkbook().getCreationHelper().createFormulaEvaluator());
                    getUnitTestLogger().info(r -> r.message("成功创建带公式求值器的工作簿"));
                    return Future.succeededFuture();
                }
        );

        // 测试流式写入模式
        KeelSheets.useSheets(
                new SheetsCreateOptions().setUseXlsx(true).setUseStreamWriting(true),
                keelSheets -> {
                    KeelSheet sheet = keelSheets.generateWriterForSheet("StreamTest");

                    // 添加大量数据测试流式写入
                    for (int i = 0; i < 100; i++) {
                        Row row = sheet.getSheet().createRow(i);
                        for (int j = 0; j < 10; j++) {
                            row.createCell(j).setCellValue("Cell_" + i + "_" + j);
                        }
                    }

                    getUnitTestLogger().info(r -> r.message("成功使用流式写入模式创建大量数据"));
                    return Future.succeededFuture();
                }
        );
    }

    @Test
    @DisplayName("测试错误处理和资源管理")
    void testErrorHandlingAndResourceManagement() {
        // 测试错误处理
        File nonExistentFile = new File("/non/existent/file.xlsx");

        KeelSheets.useSheets(
                new SheetsOpenOptions().setFile(nonExistentFile),
                keelSheets -> {
                    fail("应该抛出异常，因为文件不存在");
                    return Future.succeededFuture();
                }
        ).onComplete(ar -> {
            if (ar.failed()) {
                getUnitTestLogger().info(r -> r.message("预期的错误: " + ar.cause().getMessage()));
                assertTrue(ar.cause() instanceof IOException || ar.cause().getCause() instanceof IOException);
            } else {
                fail("应该失败但没有失败");
            }
        });

        // 测试资源自动关闭
        KeelSheets.useSheets(
                new SheetsCreateOptions().setUseXlsx(true),
                keelSheets -> {
                    KeelSheet sheet = keelSheets.generateWriterForSheet("AutoCloseTest");
                    Row row = sheet.getSheet().createRow(0);
                    row.createCell(0).setCellValue("Test");

                    getUnitTestLogger().info(r -> r.message("在回调中成功操作工作簿"));
                    return Future.succeededFuture();
                }
        ).onComplete(ar -> {
            if (ar.succeeded()) {
                getUnitTestLogger().info(r -> r.message("资源自动关闭测试成功"));
            } else {
                getUnitTestLogger().error(r -> r.message("资源自动关闭测试失败: " + ar.cause().getMessage()));
            }
        });
    }
}