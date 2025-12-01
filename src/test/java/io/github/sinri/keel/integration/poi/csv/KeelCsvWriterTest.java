package io.github.sinri.keel.integration.poi.csv;

import io.github.sinri.keel.tesuto.KeelJUnit5Test;
import io.vertx.core.Future;
import io.vertx.core.Vertx;
import org.jetbrains.annotations.NotNull;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class KeelCsvWriterTest extends KeelJUnit5Test {

    public KeelCsvWriterTest(@NotNull Vertx vertx) {
        super(vertx);
    }

    @Test
    void testBlockWriteRow() throws IOException {
        Path outputFile = Paths.get("src/test/resources/runtime/test-block-write-row.csv");
        Files.createDirectories(outputFile.getParent());

        try (FileOutputStream fos = new FileOutputStream(outputFile.toFile());
             KeelCsvWriter writer = new KeelCsvWriter(fos)) {

            writer.blockWriteRow(Arrays.asList("Name", "Age", "City"));
            writer.blockWriteRow(Arrays.asList("Alice", "30", "New York"));
            writer.blockWriteRow(Arrays.asList("Bob", "25", "London"));
        }

        // 验证文件内容
        List<String> lines = Files.readAllLines(outputFile, StandardCharsets.UTF_8);
        assertEquals(3, lines.size());
        assertEquals("Name,Age,City", lines.get(0));
        assertEquals("Alice,30,New York", lines.get(1));
        assertEquals("Bob,25,London", lines.get(2));

        // 使用KeelCsvReader验证
        try (FileInputStream fis = new FileInputStream(outputFile.toFile());
             KeelCsvReader reader = new KeelCsvReader(fis, StandardCharsets.UTF_8, ",")) {

            CsvRow row1 = reader.next();
            assertNotNull(row1);
            assertEquals(3, row1.size());
            assertEquals("Name", row1.getCell(0).getString());
            assertEquals("Age", row1.getCell(1).getString());
            assertEquals("City", row1.getCell(2).getString());

            CsvRow row2 = reader.next();
            assertNotNull(row2);
            assertEquals("Alice", row2.getCell(0).getString());
            assertEquals("30", row2.getCell(1).getString());
            assertEquals("New York", row2.getCell(2).getString());

            CsvRow row3 = reader.next();
            assertNotNull(row3);
            assertEquals("Bob", row3.getCell(0).getString());
            assertEquals("25", row3.getCell(1).getString());
            assertEquals("London", row3.getCell(2).getString());

            assertNull(reader.next());
        }
    }

    @Test
    void testWriteCellAndRowEnding() throws IOException {
        Path outputFile = Paths.get("src/test/resources/runtime/test-write-cell-row-ending.csv");
        Files.createDirectories(outputFile.getParent());

        try (FileOutputStream fos = new FileOutputStream(outputFile.toFile());
             KeelCsvWriter writer = new KeelCsvWriter(fos)) {

            writer.writeCell("Name");
            writer.writeCell("Age");
            writer.writeCell("City");
            writer.writeRowEnding();

            writer.writeCell("Alice");
            writer.writeCell("30");
            writer.writeCell("New York");
            writer.writeRowEnding();
        }

        // 验证文件内容
        List<String> lines = Files.readAllLines(outputFile, StandardCharsets.UTF_8);
        assertEquals(2, lines.size());
        assertTrue(lines.get(0).startsWith("Name,"));
        assertTrue(lines.get(0).endsWith(",City,"));
        assertTrue(lines.get(1).startsWith("Alice,"));
    }

    @Test
    void testSpecialCharacters() throws IOException {
        Path outputFile = Paths.get("src/test/resources/runtime/test-special-characters.csv");
        Files.createDirectories(outputFile.getParent());

        try (FileOutputStream fos = new FileOutputStream(outputFile.toFile());
             KeelCsvWriter writer = new KeelCsvWriter(fos)) {

            // 测试包含引号的单元格
            writer.blockWriteRow(Arrays.asList("Say \"Hello\"", "Normal", "Test"));
            // 测试包含逗号的单元格
            writer.blockWriteRow(Arrays.asList("Item, with comma", "Price", "100"));
            // 测试包含换行符的单元格
            writer.blockWriteRow(Arrays.asList("Line1\nLine2", "Multi", "Line"));
            // 测试包含引号和逗号的单元格
            writer.blockWriteRow(Arrays.asList("Say \"Hello\", world", "Complex", "Cell"));
        }

        // 使用KeelCsvReader验证
        try (FileInputStream fis = new FileInputStream(outputFile.toFile());
             KeelCsvReader reader = new KeelCsvReader(fis, StandardCharsets.UTF_8, ",")) {

            CsvRow row1 = reader.next();
            assertNotNull(row1);
            assertEquals("Say \"Hello\"", row1.getCell(0).getString());
            assertEquals("Normal", row1.getCell(1).getString());

            CsvRow row2 = reader.next();
            assertNotNull(row2);
            assertEquals("Item, with comma", row2.getCell(0).getString());

            CsvRow row3 = reader.next();
            assertNotNull(row3);
            String cell0 = row3.getCell(0).getString();
            assertTrue(cell0.contains("Line1") && cell0.contains("Line2"));

            CsvRow row4 = reader.next();
            assertNotNull(row4);
            assertEquals("Say \"Hello\", world", row4.getCell(0).getString());
        }
    }

    @Test
    void testCustomSeparator() throws IOException {
        Path outputFile = Paths.get("src/test/resources/runtime/test-custom-separator.csv");
        Files.createDirectories(outputFile.getParent());

        try (FileOutputStream fos = new FileOutputStream(outputFile.toFile());
             KeelCsvWriter writer = new KeelCsvWriter(fos, ";", StandardCharsets.UTF_8)) {

            writer.blockWriteRow(Arrays.asList("Name", "Age", "City"));
            writer.blockWriteRow(Arrays.asList("Alice", "30", "New York"));
        }

        // 验证文件内容
        List<String> lines = Files.readAllLines(outputFile, StandardCharsets.UTF_8);
        assertEquals(2, lines.size());
        assertEquals("Name;Age;City", lines.get(0));
        assertEquals("Alice;30;New York", lines.get(1));

        // 使用KeelCsvReader验证
        try (FileInputStream fis = new FileInputStream(outputFile.toFile());
             KeelCsvReader reader = new KeelCsvReader(fis, StandardCharsets.UTF_8, ";")) {

            CsvRow row1 = reader.next();
            assertNotNull(row1);
            assertEquals(3, row1.size());
            assertEquals("Name", row1.getCell(0).getString());

            CsvRow row2 = reader.next();
            assertNotNull(row2);
            assertEquals("Alice", row2.getCell(0).getString());
        }
    }

    @Test
    void testCustomCharset() throws IOException {
        Path outputFile = Paths.get("src/test/resources/runtime/test-custom-charset.csv");
        Files.createDirectories(outputFile.getParent());

        Charset gbk = Charset.forName("GBK");
        try (FileOutputStream fos = new FileOutputStream(outputFile.toFile());
             KeelCsvWriter writer = new KeelCsvWriter(fos, ",", gbk)) {

            writer.blockWriteRow(Arrays.asList("姓名", "年龄", "城市"));
            writer.blockWriteRow(Arrays.asList("张三", "25", "北京"));
        }

        // 使用KeelCsvReader验证
        try (FileInputStream fis = new FileInputStream(outputFile.toFile());
             KeelCsvReader reader = new KeelCsvReader(fis, gbk, ",")) {

            CsvRow row1 = reader.next();
            assertNotNull(row1);
            assertEquals("姓名", row1.getCell(0).getString());
            assertEquals("年龄", row1.getCell(1).getString());
            assertEquals("城市", row1.getCell(2).getString());

            CsvRow row2 = reader.next();
            assertNotNull(row2);
            assertEquals("张三", row2.getCell(0).getString());
            assertEquals("25", row2.getCell(1).getString());
            assertEquals("北京", row2.getCell(2).getString());
        }
    }

    @Test
    void testNullAndEmptyCells() throws IOException {
        Path outputFile = Paths.get("src/test/resources/runtime/test-null-empty-cells.csv");
        Files.createDirectories(outputFile.getParent());

        try (FileOutputStream fos = new FileOutputStream(outputFile.toFile());
             KeelCsvWriter writer = new KeelCsvWriter(fos)) {

            writer.blockWriteRow(Arrays.asList("Name", "", null, "City"));
            writer.blockWriteRow(Arrays.asList("Alice", "30", "", null));
        }

        // 使用KeelCsvReader验证
        try (FileInputStream fis = new FileInputStream(outputFile.toFile());
             KeelCsvReader reader = new KeelCsvReader(fis, StandardCharsets.UTF_8, ",")) {

            CsvRow row1 = reader.next();
            assertNotNull(row1);
            assertEquals(4, row1.size());
            assertEquals("Name", row1.getCell(0).getString());
            assertTrue(row1.getCell(1).isEmpty());
            assertTrue(row1.getCell(2).isNull() || row1.getCell(2).isEmpty());
            assertEquals("City", row1.getCell(3).getString());

            CsvRow row2 = reader.next();
            assertNotNull(row2);
            assertEquals("Alice", row2.getCell(0).getString());
            assertEquals("30", row2.getCell(1).getString());
        }
    }

    @Test
    void testStaticWriteMethod() throws IOException {
        Path outputFile = Paths.get("src/test/resources/runtime/test-static-write.csv");
        Files.createDirectories(outputFile.getParent());

        try (FileOutputStream fos = new FileOutputStream(outputFile.toFile())) {
            Future<Void> future = KeelCsvWriter.write(
                    fos,
                    writer -> {
                        try {
                            writer.blockWriteRow(Arrays.asList("Name", "Age", "City"));
                            writer.blockWriteRow(Arrays.asList("Alice", "30", "New York"));
                            writer.blockWriteRow(Arrays.asList("Bob", "25", "London"));
                            return Future.succeededFuture();
                        } catch (IOException e) {
                            return Future.failedFuture(e);
                        }
                    }
            );

            assertTrue(future.succeeded());
        }

        // 验证文件内容
        List<String> lines = Files.readAllLines(outputFile, StandardCharsets.UTF_8);
        assertEquals(3, lines.size());
        assertEquals("Name,Age,City", lines.get(0));
        assertEquals("Alice,30,New York", lines.get(1));
        assertEquals("Bob,25,London", lines.get(2));
    }

    @Test
    void testStaticWriteMethodWithCustomSeparatorAndCharset() throws IOException {
        Path outputFile = Paths.get("src/test/resources/runtime/test-static-write-custom.csv");
        Files.createDirectories(outputFile.getParent());

        Charset gbk = Charset.forName("GBK");
        try (FileOutputStream fos = new FileOutputStream(outputFile.toFile())) {
            Future<Void> future = KeelCsvWriter.write(
                    fos,
                    "|",
                    gbk,
                    writer -> {
                        try {
                            writer.blockWriteRow(Arrays.asList("姓名", "年龄", "城市"));
                            writer.blockWriteRow(Arrays.asList("张三", "25", "北京"));
                            return Future.succeededFuture();
                        } catch (IOException e) {
                            return Future.failedFuture(e);
                        }
                    }
            );

            assertTrue(future.succeeded());
        }

        // 使用KeelCsvReader验证
        try (FileInputStream fis = new FileInputStream(outputFile.toFile());
             KeelCsvReader reader = new KeelCsvReader(fis, gbk, "|")) {

            CsvRow row1 = reader.next();
            assertNotNull(row1);
            assertEquals("姓名", row1.getCell(0).getString());

            CsvRow row2 = reader.next();
            assertNotNull(row2);
            assertEquals("张三", row2.getCell(0).getString());
        }
    }

    @Test
    void testMixedWriteMethods() throws IOException {
        Path outputFile = Paths.get("src/test/resources/runtime/test-mixed-write-methods.csv");
        Files.createDirectories(outputFile.getParent());

        try (FileOutputStream fos = new FileOutputStream(outputFile.toFile());
             KeelCsvWriter writer = new KeelCsvWriter(fos)) {

            // 先使用blockWriteRow
            writer.blockWriteRow(Arrays.asList("Name", "Age", "City"));
            // 然后使用writeCell和writeRowEnding
            writer.writeCell("Alice");
            writer.writeCell("30");
            writer.writeCell("New York");
            writer.writeRowEnding();
            // 再使用blockWriteRow
            writer.blockWriteRow(Arrays.asList("Bob", "25", "London"));
        }

        // 验证文件内容
        List<String> lines = Files.readAllLines(outputFile, StandardCharsets.UTF_8);
        assertEquals(3, lines.size());
        assertEquals("Name,Age,City", lines.get(0));
        assertTrue(lines.get(1).startsWith("Alice,"));
        assertEquals("Bob,25,London", lines.get(2));
    }

    @Test
    void testLargeData() throws IOException {
        Path outputFile = Paths.get("src/test/resources/runtime/test-large-data.csv");
        Files.createDirectories(outputFile.getParent());

        int rowCount = 1000;
        try (FileOutputStream fos = new FileOutputStream(outputFile.toFile());
             KeelCsvWriter writer = new KeelCsvWriter(fos)) {

            writer.blockWriteRow(Arrays.asList("ID", "Name", "Value"));
            for (int i = 1; i <= rowCount; i++) {
                writer.blockWriteRow(Arrays.asList(String.valueOf(i), "Item" + i, String.valueOf(i * 10)));
            }
        }

        // 验证文件内容
        List<String> lines = Files.readAllLines(outputFile, StandardCharsets.UTF_8);
        assertEquals(rowCount + 1, lines.size());
        assertEquals("ID,Name,Value", lines.get(0));
        assertEquals("1,Item1,10", lines.get(1));
        assertEquals("1000,Item1000,10000", lines.get(rowCount));

        // 使用KeelCsvReader验证
        try (FileInputStream fis = new FileInputStream(outputFile.toFile());
             KeelCsvReader reader = new KeelCsvReader(fis, StandardCharsets.UTF_8, ",")) {

            CsvRow header = reader.next();
            assertNotNull(header);
            assertEquals("ID", header.getCell(0).getString());

            int count = 0;
            CsvRow row;
            while ((row = reader.next()) != null) {
                count++;
                assertEquals(String.valueOf(count), row.getCell(0).getString());
                assertEquals("Item" + count, row.getCell(1).getString());
                assertEquals(String.valueOf(count * 10), row.getCell(2).getString());
            }
            assertEquals(rowCount, count);
        }
    }
}