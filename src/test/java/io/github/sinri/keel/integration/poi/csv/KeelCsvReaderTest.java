package io.github.sinri.keel.integration.poi.csv;

import io.github.sinri.keel.tesuto.KeelJUnit5Test;
import io.vertx.core.Future;
import io.vertx.core.Vertx;
import org.jspecify.annotations.NullMarked;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.*;

@NullMarked
class KeelCsvReaderTest extends KeelJUnit5Test {

    public KeelCsvReaderTest(Vertx vertx) {
        super(vertx);
    }

    @Test
    void testReadCsvFile() throws IOException {
        String csvFile = "src/test/resources/runtime/test-csv-1.csv";
        try (FileInputStream fis = new FileInputStream(csvFile);
             KeelCsvReader reader = new KeelCsvReader(fis, StandardCharsets.UTF_8, ",")) {

            // 读取表头
            CsvRow header = reader.next();
            assertNotNull(header);
            assertEquals(4, header.size());
            assertEquals("id", header.getCell(0).getString());
            assertEquals("name", header.getCell(1).getString());
            assertEquals("category", header.getCell(2).getString());
            assertEquals("price", header.getCell(3).getString());

            // 读取第一行数据：1,apple,fruit,5
            CsvRow row1 = reader.next();
            assertNotNull(row1);
            assertEquals(4, row1.size());
            assertEquals("1", row1.getCell(0).getString());
            assertTrue(row1.getCell(0).isNumber());
            assertEquals(new BigDecimal("1"), row1.getCell(0).getNumber());
            assertEquals("apple", row1.getCell(1).getString());
            assertFalse(row1.getCell(1).isNumber());
            assertEquals("fruit", row1.getCell(2).getString());
            assertEquals("5", row1.getCell(3).getString());
            assertTrue(row1.getCell(3).isNumber());
            assertEquals(new BigDecimal("5"), row1.getCell(3).getNumber());

            // 读取第二行数据：2,bag,tool,24.5
            CsvRow row2 = reader.next();
            assertNotNull(row2);
            assertEquals(4, row2.size());
            assertEquals("2", row2.getCell(0).getString());
            assertTrue(row2.getCell(0).isNumber());
            assertEquals(new BigDecimal("2"), row2.getCell(0).getNumber());
            assertEquals("bag", row2.getCell(1).getString());
            assertEquals("tool", row2.getCell(2).getString());
            assertEquals("24.5", row2.getCell(3).getString());
            assertTrue(row2.getCell(3).isNumber());
            assertEquals(new BigDecimal("24.5"), row2.getCell(3).getNumber());

            // 读取第三行数据：3,car,vehicle,200000
            CsvRow row3 = reader.next();
            assertNotNull(row3);
            assertEquals(4, row3.size());
            assertEquals("3", row3.getCell(0).getString());
            assertTrue(row3.getCell(0).isNumber());
            assertEquals(new BigDecimal("3"), row3.getCell(0).getNumber());
            assertEquals("car", row3.getCell(1).getString());
            assertEquals("vehicle", row3.getCell(2).getString());
            assertEquals("200000", row3.getCell(3).getString());
            assertTrue(row3.getCell(3).isNumber());
            assertEquals(new BigDecimal("200000"), row3.getCell(3).getNumber());

            // 读取第四行数据：4,detailed paint,art,3999999.9
            CsvRow row4 = reader.next();
            assertNotNull(row4);
            assertEquals(4, row4.size());
            assertEquals("4", row4.getCell(0).getString());
            assertTrue(row4.getCell(0).isNumber());
            assertEquals(new BigDecimal("4"), row4.getCell(0).getNumber());
            assertEquals("detailed paint", row4.getCell(1).getString());
            assertEquals("art", row4.getCell(2).getString());
            assertEquals("3999999.9", row4.getCell(3).getString());
            assertTrue(row4.getCell(3).isNumber());
            assertEquals(new BigDecimal("3999999.9"), row4.getCell(3).getNumber());

            // 读取第五行数据：5,e-mail address 'vip@keel.com',virutal,8888
            CsvRow row5 = reader.next();
            assertNotNull(row5);
            assertEquals(4, row5.size());
            assertEquals("5", row5.getCell(0).getString());
            assertTrue(row5.getCell(0).isNumber());
            assertEquals(new BigDecimal("5"), row5.getCell(0).getNumber());
            assertEquals("e-mail address 'vip@keel.com'", row5.getCell(1).getString());
            assertFalse(row5.getCell(1).isNumber());
            assertEquals("virutal", row5.getCell(2).getString());
            assertEquals("8888", row5.getCell(3).getString());
            assertTrue(row5.getCell(3).isNumber());
            assertEquals(new BigDecimal("8888"), row5.getCell(3).getNumber());

            // 验证没有更多行了
            assertNull(reader.next());
        }
    }

    @Test
    void testReadCsvFileWithStaticMethod() {
        String csvFile = "src/test/resources/runtime/test-csv-1.csv";
        try (FileInputStream fis = new FileInputStream(csvFile)) {
            Future<Void> future = KeelCsvReader.read(
                    fis,
                    StandardCharsets.UTF_8,
                    ",",
                    reader -> {
                        try {
                            // 读取表头
                            CsvRow header = reader.next();
                            assertNotNull(header);
                            assertEquals(4, header.size());
                            assertEquals("id", header.getCell(0).getString());

                            // 读取第一行数据
                            CsvRow row1 = reader.next();
                            assertNotNull(row1);
                            assertEquals("1", row1.getCell(0).getString());
                            assertEquals("apple", row1.getCell(1).getString());
                            assertTrue(row1.getCell(0).isNumber());
                            assertEquals(new BigDecimal("1"), row1.getCell(0).getNumber());

                            // 验证还有更多行
                            CsvRow row2 = reader.next();
                            assertNotNull(row2);
                            assertEquals("2", row2.getCell(0).getString());

                            return Future.succeededFuture();
                        } catch (IOException e) {
                            return Future.failedFuture(e);
                        }
                    }
            );

            assertTrue(future.succeeded());
        } catch (IOException e) {
            fail("Failed to read CSV file: " + e.getMessage());
        }
    }

    @Test
    void testNumberParsing() throws IOException {
        String csvFile = "src/test/resources/runtime/test-csv-1.csv";
        try (FileInputStream fis = new FileInputStream(csvFile);
             KeelCsvReader reader = new KeelCsvReader(fis, StandardCharsets.UTF_8, ",")) {

            // 跳过表头
            reader.next();

            // 测试整数
            CsvRow row1 = reader.next();
            assertTrue(row1.getCell(0).isNumber());
            assertEquals(new BigDecimal("1"), row1.getCell(0).getNumber());
            assertEquals(new BigDecimal("1"), row1.getCell(0).getNumberOrElse(0));

            // 测试小数
            CsvRow row2 = reader.next();
            assertTrue(row2.getCell(3).isNumber());
            assertEquals(new BigDecimal("24.5"), row2.getCell(3).getNumber());
            assertEquals(new BigDecimal("24.5"), row2.getCell(3).getNumberOrElse(0.0));

            // 跳过 row3
            reader.next();

            // 测试大数
            CsvRow row4 = reader.next();
            assertTrue(row4.getCell(3).isNumber());
            assertEquals(new BigDecimal("3999999.9"), row4.getCell(3).getNumber());

            // 测试非数字字符串
            CsvRow row5 = reader.next();
            assertFalse(row5.getCell(1).isNumber());
            assertThrows(NumberFormatException.class, () -> row5.getCell(1).getNumber());
            assertEquals(new BigDecimal("0"), row5.getCell(1).getNumberOrElse(0));
        }
    }

    @Test
    void testCellMethods() throws IOException {
        String csvFile = "src/test/resources/runtime/test-csv-1.csv";
        try (FileInputStream fis = new FileInputStream(csvFile);
             KeelCsvReader reader = new KeelCsvReader(fis, StandardCharsets.UTF_8, ",")) {

            // 跳过表头
            CsvRow header = reader.next();
            assertNotNull(header);

            // 读取第一行数据
            CsvRow row1 = reader.next();
            assertNotNull(row1);

            // 测试 isNull
            assertFalse(row1.getCell(0).isNull());
            assertFalse(row1.getCell(1).isNull());

            // 测试 isEmpty
            assertFalse(row1.getCell(0).isEmpty());
            assertFalse(row1.getCell(1).isEmpty());

            // 测试 getString
            assertEquals("1", row1.getCell(0).getString());
            assertEquals("apple", row1.getCell(1).getString());
        }
    }

    @Test
    void testReadAllRows() throws IOException {
        String csvFile = "src/test/resources/runtime/test-csv-1.csv";
        try (FileInputStream fis = new FileInputStream(csvFile);
             KeelCsvReader reader = new KeelCsvReader(fis, StandardCharsets.UTF_8, ",")) {

            int rowCount = 0;
            CsvRow row;
            while ((row = reader.next()) != null) {
                rowCount++;
                assertNotNull(row);
                assertEquals(4, row.size());
            }

            // 应该有1个表头 + 5行数据 = 6行
            assertEquals(6, rowCount);
        }
    }
}