package io.github.sinri.keel.integration.poi.excel;

import io.github.sinri.keel.integration.poi.excel.entity.KeelSheetMatrix;
import io.github.sinri.keel.integration.poi.excel.entity.KeelSheetMatrixRow;
import io.github.sinri.keel.tesuto.KeelJUnit5Test;
import io.vertx.core.Future;
import io.vertx.core.Vertx;
import io.vertx.junit5.VertxExtension;
import io.vertx.junit5.VertxTestContext;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

@ExtendWith(VertxExtension.class)
class KeelSheetTest extends KeelJUnit5Test {
    private static final String TEST_OUTPUT_DIR = "src/test/resources/runtime/test_output";
    private File testOutputDir;
    private KeelSheet testSheet;
    private KeelSheets testSheets;
    private Workbook testWorkbook;

    KeelSheetTest(Vertx vertx) {
        super(vertx);
    }

    protected void test(VertxTestContext testContext) {
        testContext.completeNow();
    }

    @BeforeEach
    void setUp() throws IOException {
        testOutputDir = new File(TEST_OUTPUT_DIR);
        if (!testOutputDir.exists()) {
            testOutputDir.mkdirs();
        }

        // 创建测试工作簿和工作表
        testWorkbook = new XSSFWorkbook();
        Sheet sheet = testWorkbook.createSheet("TestSheet");
        
        // 创建测试数据
        createTestData(sheet);
        
        testSheets = new KeelSheets(null, testWorkbook);
        testSheet = testSheets.generateReaderForSheet(0);
    }

    @AfterEach
    void tearDown() throws IOException {
        if (testWorkbook != null) {
            testWorkbook.close();
        }
        if (testSheets != null) {
            try {
                testSheets.close(null);
            } catch (Exception e) {
                // 忽略关闭时的错误
            }
        }
        
        // 清理测试输出目录
        if (testOutputDir.exists()) {
            deleteDirectory(testOutputDir);
        }
    }

    private void createTestData(Sheet sheet) {
        // 创建表头行
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Name");
        headerRow.createCell(1).setCellValue("Age");
        headerRow.createCell(2).setCellValue("City");
        headerRow.createCell(3).setCellValue("Salary");

        // 创建数据行
        Row dataRow1 = sheet.createRow(1);
        dataRow1.createCell(0).setCellValue("Alice");
        dataRow1.createCell(1).setCellValue(25);
        dataRow1.createCell(2).setCellValue("Beijing");
        dataRow1.createCell(3).setCellValue(5000.0);

        Row dataRow2 = sheet.createRow(2);
        dataRow2.createCell(0).setCellValue("Bob");
        dataRow2.createCell(1).setCellValue(30);
        dataRow2.createCell(2).setCellValue("Shanghai");
        dataRow2.createCell(3).setCellValue(6000.0);

        Row dataRow3 = sheet.createRow(3);
        dataRow3.createCell(0).setCellValue("Charlie");
        dataRow3.createCell(1).setCellValue(35);
        dataRow3.createCell(2).setCellValue("Guangzhou");
        dataRow3.createCell(3).setCellValue(7000.0);

        // 创建空行（用于测试过滤）
        Row emptyRow = sheet.createRow(4);
        // 不添加任何单元格，创建空行

        // 创建公式行
        Row formulaRow = sheet.createRow(5);
        formulaRow.createCell(0).setCellValue("Total");
        formulaRow.createCell(1).setCellValue(""); // 空单元格
        formulaRow.createCell(2).setCellValue(""); // 空单元格
        formulaRow.createCell(3).setCellFormula("SUM(D2:D4)");
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
    @DisplayName("测试工作表基本信息")
    void testSheetBasicInfo() {
        assertNotNull(testSheet);
        assertEquals("TestSheet", testSheet.getName());
        assertNotNull(testSheet.getSheet());
        assertEquals(6, testSheet.getSheet().getLastRowNum() + 1); // 行数从0开始
    }

    @Test
    @DisplayName("测试读取单行")
    void testReadRow() {
        Row row = testSheet.readRow(0);
        assertNotNull(row);
        assertEquals(4, row.getLastCellNum());
        assertEquals("Name", row.getCell(0).getStringCellValue());
        assertEquals("Age", row.getCell(1).getStringCellValue());
        assertEquals("City", row.getCell(2).getStringCellValue());
        assertEquals("Salary", row.getCell(3).getStringCellValue());
    }

    @Test
    @DisplayName("测试读取不存在的行")
    void testReadNonExistentRow() {
        Row row = testSheet.readRow(999);
        assertNull(row);
    }

    @Test
    @DisplayName("测试行迭代器")
    void testRowIterator() {
        Iterator<Row> rowIterator = testSheet.getRowIterator();
        assertNotNull(rowIterator);
        
        int rowCount = 0;
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            assertNotNull(row);
            rowCount++;
        }
        assertEquals(6, rowCount);
    }

    @Test
    @DisplayName("测试读取原始行数据")
    void testReadRawRow() {
        List<String> rawRow = testSheet.readRawRow(0, 4, null);
        assertNotNull(rawRow);
        assertEquals(4, rawRow.size());
        assertEquals("Name", rawRow.get(0));
        assertEquals("Age", rawRow.get(1));
        assertEquals("City", rawRow.get(2));
        assertEquals("Salary", rawRow.get(3));
    }

    @Test
    @DisplayName("测试读取原始行数据（指定最大列数）")
    void testReadRawRowWithMaxColumns() {
        List<String> rawRow = testSheet.readRawRow(0, 2, null);
        assertNotNull(rawRow);
        assertEquals(2, rawRow.size());
        assertEquals("Name", rawRow.get(0));
        assertEquals("Age", rawRow.get(1));
    }

    @Test
    @DisplayName("测试原始行迭代器")
    void testRawRowIterator() {
        Iterator<List<String>> rawRowIterator = testSheet.getRawRowIterator(4, null);
        assertNotNull(rawRowIterator);
        
        int rowCount = 0;
        while (rawRowIterator.hasNext()) {
            List<String> rawRow = rawRowIterator.next();
            assertNotNull(rawRow);
            assertEquals(4, rawRow.size());
            rowCount++;
        }
        assertEquals(6, rowCount);
    }

    @Test
    @DisplayName("测试阻塞读取所有行")
    void testBlockReadAllRows() {
        List<Row> rows = new ArrayList<>();
        testSheet.blockReadAllRows(rows::add);
        
        assertEquals(6, rows.size());
        assertEquals("Name", rows.get(0).getCell(0).getStringCellValue());
        assertEquals("Alice", rows.get(1).getCell(0).getStringCellValue());
        assertEquals("Bob", rows.get(2).getCell(0).getStringCellValue());
    }

    @Test
    @DisplayName("测试读取所有行到矩阵")
    void testBlockReadAllRowsToMatrix() {
        KeelSheetMatrix matrix = testSheet.blockReadAllRowsToMatrix();
        assertNotNull(matrix);
        
        List<String> headers = matrix.getHeaderRow();
        assertEquals(4, headers.size());
        assertEquals("Name", headers.get(0));
        assertEquals("Age", headers.get(1));
        assertEquals("City", headers.get(2));
        assertEquals("Salary", headers.get(3));
        
        List<List<String>> dataRows = matrix.getRawRowList();
        assertEquals(4, dataRows.size()); // 减去表头行，空行被过滤
        
        // 验证第一行数据
        List<String> firstDataRow = dataRows.get(0);
        assertEquals("Alice", firstDataRow.get(0));
        assertEquals("25.0", firstDataRow.get(1)); // 数字单元格转换为字符串时保留小数点
        assertEquals("Beijing", firstDataRow.get(2));
        assertEquals("5000.0", firstDataRow.get(3));
    }

    @Test
    @DisplayName("测试读取所有行到矩阵（指定参数）")
    void testBlockReadAllRowsToMatrixWithParams() {
        KeelSheetMatrix matrix = testSheet.blockReadAllRowsToMatrix(1, 3, SheetRowFilter.toThrowEmptyRows());
        assertNotNull(matrix);
        
        List<String> headers = matrix.getHeaderRow();
        assertEquals(3, headers.size());
        assertEquals("Alice", headers.get(0)); // 第1行第1列的值
        assertEquals("25.0", headers.get(1)); // 第1行第2列的值
        assertEquals("Beijing", headers.get(2)); // 第1行第3列的值
        
        List<List<String>> dataRows = matrix.getRawRowList();
        assertEquals(3, dataRows.size()); // 从第2行开始，减去表头行，空行被过滤
    }

    @Test
    @DisplayName("测试异步读取所有行到矩阵")
    void testReadAllRowsToMatrix() {
        testSheet.readAllRowsToMatrix()
                .onComplete(ar -> {
                    assertTrue(ar.succeeded());
                    KeelSheetMatrix matrix = ar.result();
                    assertNotNull(matrix);
                    
                    List<String> headers = matrix.getHeaderRow();
                    assertEquals(4, headers.size());
                    assertEquals("Name", headers.get(0));
                    
                    List<List<String>> dataRows = matrix.getRawRowList();
                    assertEquals(4, dataRows.size());
                });
    }

    @Test
    @DisplayName("测试异步读取所有行到矩阵（指定参数）")
    void testReadAllRowsToMatrixWithParams() {
        testSheet.readAllRowsToMatrix(1, 2, null)
                .onComplete(ar -> {
                    assertTrue(ar.succeeded());
                    KeelSheetMatrix matrix = ar.result();
                    assertNotNull(matrix);
                    
                    List<String> headers = matrix.getHeaderRow();
                    assertEquals(2, headers.size());
                    assertEquals("Alice", headers.get(0)); // 第1行第1列的值
                    assertEquals("25.0", headers.get(1)); // 第1行第2列的值
                });
    }

    @Test
    @DisplayName("测试矩阵行迭代器")
    void testMatrixRowIterator() {
        KeelSheetMatrix matrix = testSheet.blockReadAllRowsToMatrix();
        Iterator<KeelSheetMatrixRow> rowIterator = matrix.getRowIterator();
        
        assertNotNull(rowIterator);
        int rowCount = 0;
        while (rowIterator.hasNext()) {
            KeelSheetMatrixRow row = rowIterator.next();
            assertNotNull(row);
            rowCount++;
        }
        assertEquals(4, rowCount); // 减去表头行，空行被过滤
    }

    @Test
    @DisplayName("测试矩阵行数据读取")
    void testMatrixRowDataReading() {
        KeelSheetMatrix matrix = testSheet.blockReadAllRowsToMatrix();
        Iterator<KeelSheetMatrixRow> rowIterator = matrix.getRowIterator();
        
        assertTrue(rowIterator.hasNext());
        KeelSheetMatrixRow firstRow = rowIterator.next();
        
        assertEquals("Alice", firstRow.readValue(0));
        assertEquals(25, firstRow.readValueToInteger(1));
        assertEquals(25.0, firstRow.readValueToDouble(1));
        assertEquals(25L, firstRow.readValueToLong(1));
        assertEquals(5000.0, firstRow.readValueToDouble(3));
    }

    @Test
    @DisplayName("测试工作表类型设置和获取")
    void testSheetTypeOperations() {
        assertNull(testSheet.getSheetsReaderType());
        
        testSheet.setSheetsReaderType(KeelSheetsReaderType.XLSX);
        assertEquals(KeelSheetsReaderType.XLSX, testSheet.getSheetsReaderType());
        
        testSheet.setSheetsReaderType(KeelSheetsReaderType.XLS);
        assertEquals(KeelSheetsReaderType.XLS, testSheet.getSheetsReaderType());
    }

    @Test
    @DisplayName("测试写入功能")
    void testWriteOperations() throws IOException {
        // 创建新的写入工作表
        KeelSheet writerSheet = testSheets.generateWriterForSheet("WriteTest");
        
        // 准备测试数据
        List<List<String>> testData = List.of(
                List.of("Product", "Price", "Quantity"),
                List.of("Apple", "2.5", "100"),
                List.of("Banana", "1.8", "150"),
                List.of("Orange", "3.2", "80")
        );
        
        // 写入数据
        writerSheet.blockWriteAllRows(testData);
        
        // 保存到文件
        File outputFile = new File(testOutputDir, "write_test.xlsx");
        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
            testSheets.save(fos);
        }
        
        // 验证文件已创建
        assertTrue(outputFile.exists());
        assertTrue(outputFile.length() > 0);
        
        // 清理
        outputFile.delete();
    }

    @Test
    @DisplayName("测试图片功能（无图片的工作表）")
    void testPictureOperations() {
        List<KeelPictureInSheet> pictures = testSheet.getPictures();
        assertNotNull(pictures);
        assertTrue(pictures.isEmpty(), "测试工作表应该没有图片");
    }

    @Test
    @DisplayName("测试公式单元格处理")
    void testFormulaCellHandling() {
        // 测试读取包含公式的行
        List<String> formulaRow = testSheet.readRawRow(5, 4, null);
        assertNotNull(formulaRow);
        assertEquals("Total", formulaRow.get(0));
        assertEquals("", formulaRow.get(1)); // 空单元格
        assertEquals("", formulaRow.get(2)); // 空单元格
        // 公式单元格的值取决于是否有公式求值器
        assertNotNull(formulaRow.get(3));
    }

    @Test
    @DisplayName("测试空行处理")
    void testEmptyRowHandling() {
        // 第4行是空行
        List<String> emptyRow = testSheet.readRawRow(4, 4, null);
        assertNotNull(emptyRow);
        assertEquals(4, emptyRow.size());
        
        // 所有单元格都应该是空字符串
        for (String cellValue : emptyRow) {
            assertEquals("", cellValue);
        }
    }

    @Test
    @DisplayName("测试列数自动检测")
    void testColumnCountAutoDetection() {
        // 测试不同行的列数
        Row headerRow = testSheet.readRow(0);
        assertEquals(4, headerRow.getLastCellNum());
        
        Row dataRow = testSheet.readRow(1);
        assertEquals(4, dataRow.getLastCellNum());
        
        // 创建一个列数不同的行
        Sheet sheet = testSheet.getSheet();
        Row customRow = sheet.createRow(10);
        customRow.createCell(0).setCellValue("Col1");
        customRow.createCell(2).setCellValue("Col3"); // 跳过第1列
        
        assertEquals(3, customRow.getLastCellNum());
    }

    @Test
    @DisplayName("测试单元格转字符串功能")
    void testCellToStringConversion() {
        // 测试不同类型的单元格
        Row testRow = testSheet.readRow(1);
        
        // 字符串单元格
        Cell stringCell = testRow.getCell(0);
        String stringValue = stringCell.getStringCellValue();
        assertEquals("Alice", stringValue);
        
        // 数字单元格
        Cell numberCell = testRow.getCell(1);
        String numberValue = String.valueOf((int) numberCell.getNumericCellValue());
        assertEquals("25", numberValue);
        
        // 双精度单元格
        Cell doubleCell = testRow.getCell(3);
        String doubleValue = String.valueOf(doubleCell.getNumericCellValue());
        assertEquals("5000.0", doubleValue);
    }

    @Test
    @DisplayName("测试行过滤功能")
    void testRowFiltering() {
        // 创建一个简单的行过滤器，过滤掉空行
        SheetRowFilter filter = new SheetRowFilter() {
            @Override
            public boolean shouldThrowThisRawRow(@NotNull List<String> rawRow) {
                return rawRow.stream().allMatch(String::isEmpty);
            }
        };
        
        // 使用过滤器读取行
        List<String> filteredRow = testSheet.readRawRow(4, 4, filter);
        // 第4行是空行，应该被过滤掉
        assertNull(filteredRow);
        
        // 非空行应该正常返回
        List<String> nonEmptyRow = testSheet.readRawRow(1, 4, filter);
        assertNotNull(nonEmptyRow);
        assertEquals("Alice", nonEmptyRow.get(0));
    }

    @Test
    @DisplayName("测试工作簿保存和加载")
    void testWorkbookSaveAndLoad() throws IOException {
        // 保存当前工作簿
        File outputFile = new File(testOutputDir, "save_load_test.xlsx");
        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
            testSheets.save(fos);
        }
        
        // 验证文件已创建
        assertTrue(outputFile.exists());
        assertTrue(outputFile.length() > 0);
        
        // 重新加载工作簿
        KeelSheets.useSheets(new SheetsOpenOptions().setFile(outputFile), loadedSheets -> {
            assertNotNull(loadedSheets);
            
            // 验证工作表数量
            assertEquals(1, loadedSheets.getSheetCount());
            
            // 验证工作表名称
            KeelSheet loadedSheet = loadedSheets.generateReaderForSheet(0);
            assertEquals("TestSheet", loadedSheet.getName());
            
            return Future.succeededFuture();
        });
        
        // 清理
        outputFile.delete();
    }
}