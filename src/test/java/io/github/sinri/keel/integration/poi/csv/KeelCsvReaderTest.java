package io.github.sinri.keel.integration.poi.csv;

import io.github.sinri.keel.facade.tesuto.unit.KeelJUnit5Test;
import io.vertx.core.Vertx;
import io.vertx.junit5.VertxExtension;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

@ExtendWith(VertxExtension.class)
class KeelCsvReaderTest extends KeelJUnit5Test {

    public KeelCsvReaderTest(Vertx vertx) {
        super(vertx);
    }

    @Test
    void testBasicCsvReading() throws IOException {
        String csvContent = "Name,Age,City\nJohn,25,New York\nJane,30,Los Angeles";
        ByteArrayInputStream inputStream = new ByteArrayInputStream(csvContent.getBytes(StandardCharsets.UTF_8));
        try (KeelCsvReader reader = new KeelCsvReader(inputStream, StandardCharsets.UTF_8)) {
            // Read header
            CsvRow header = reader.next();
            assertNotNull(header);
            assertEquals(3, header.size());
            assertEquals("Name", header.getCell(0).getString());
            assertEquals("Age", header.getCell(1).getString());
            assertEquals("City", header.getCell(2).getString());

            // Read first data row
            CsvRow row1 = reader.next();
            assertNotNull(row1);
            assertEquals(3, row1.size());
            assertEquals("John", row1.getCell(0).getString());
            assertEquals("25", row1.getCell(1).getString());
            assertTrue(row1.getCell(1).isNumber());
            assertEquals("New York", row1.getCell(2).getString());

            // Read second data row
            CsvRow row2 = reader.next();
            assertNotNull(row2);
            assertEquals(3, row2.size());
            assertEquals("Jane", row2.getCell(0).getString());
            assertEquals("30", row2.getCell(1).getString());
            assertTrue(row2.getCell(1).isNumber());
            assertEquals("Los Angeles", row2.getCell(2).getString());

            // End of file
            assertNull(reader.next());
        }
    }

    @Test
    void testCustomSeparator() throws IOException {
        String csvContent = "Name;Age;City\nJohn;25;New York\nJane;30;Los Angeles";
        ByteArrayInputStream inputStream = new ByteArrayInputStream(csvContent.getBytes(StandardCharsets.UTF_8));

        try (KeelCsvReader reader = new KeelCsvReader(inputStream, StandardCharsets.UTF_8, ";")) {
            CsvRow header = reader.next();
            assertNotNull(header);
            assertEquals(3, header.size());
            assertEquals("Name", header.getCell(0).getString());
            assertEquals("Age", header.getCell(1).getString());
            assertEquals("City", header.getCell(2).getString());

            CsvRow row1 = reader.next();
            assertNotNull(row1);
            assertEquals("John", row1.getCell(0).getString());
            assertEquals("25", row1.getCell(1).getString());
            assertEquals("New York", row1.getCell(2).getString());
        }
    }

    @Test
    void testQuotedFields() throws IOException {
        String csvContent = "Name,Description,Value\n\"John Doe\",\"Contains, comma\",100\n\"Jane \"\"Smith\"\"\",\"Contains \"\"quotes\"\"\",200";
        ByteArrayInputStream inputStream = new ByteArrayInputStream(csvContent.getBytes(StandardCharsets.UTF_8));

        try (KeelCsvReader reader = new KeelCsvReader(inputStream, StandardCharsets.UTF_8)) {
            // Read header
            CsvRow header = reader.next();
            assertNotNull(header);
            assertEquals(3, header.size());

            // Read first data row
            CsvRow row1 = reader.next();
            assertNotNull(row1);
            assertEquals(3, row1.size());
            assertEquals("John Doe", row1.getCell(0).getString());
            assertEquals("Contains, comma", row1.getCell(1).getString());
            assertEquals("100", row1.getCell(2).getString());
            assertTrue(row1.getCell(2).isNumber());

            // Read second data row
            CsvRow row2 = reader.next();
            assertNotNull(row2);
            assertEquals(3, row2.size());
            assertEquals("Jane \"Smith\"", row2.getCell(0).getString());
            assertEquals("Contains \"quotes\"", row2.getCell(1).getString());
            assertEquals("200", row2.getCell(2).getString());
            assertTrue(row2.getCell(2).isNumber());
        }
    }

    @Test
    void testMultiLineFields() throws IOException {
        String csvContent = "Name,Description,Value\n\"John Doe\",\"Multi\nline\ndescription\",100\nJane,Simple description,200";
        ByteArrayInputStream inputStream = new ByteArrayInputStream(csvContent.getBytes(StandardCharsets.UTF_8));

        try (KeelCsvReader reader = new KeelCsvReader(inputStream, StandardCharsets.UTF_8)) {
            // Read header
            CsvRow header = reader.next();
            assertNotNull(header);
            assertEquals(3, header.size());

            // Read first data row (multi-line)
            CsvRow row1 = reader.next();
            assertNotNull(row1);
            assertEquals(3, row1.size());
            assertEquals("John Doe", row1.getCell(0).getString());
            assertEquals("Multi\nline\ndescription", row1.getCell(1).getString());
            assertEquals("100", row1.getCell(2).getString());

            // Read second data row
            CsvRow row2 = reader.next();
            assertNotNull(row2);
            assertEquals(3, row2.size());
            assertEquals("Jane", row2.getCell(0).getString());
            assertEquals("Simple description", row2.getCell(1).getString());
            assertEquals("200", row2.getCell(2).getString());
        }
    }

    @Test
    void testEmptyFields() throws IOException {
        String csvContent = "Name,Age,City\nJohn,,New York\n,30,\n\"\",\"\",\"\"";
        ByteArrayInputStream inputStream = new ByteArrayInputStream(csvContent.getBytes(StandardCharsets.UTF_8));

        try (KeelCsvReader reader = new KeelCsvReader(inputStream, StandardCharsets.UTF_8)) {
            // Read header
            CsvRow header = reader.next();
            assertNotNull(header);
            assertEquals(3, header.size());

            // Read first data row (empty middle field)
            CsvRow row1 = reader.next();
            assertNotNull(row1);
            assertEquals(3, row1.size());
            assertEquals("John", row1.getCell(0).getString());
            assertEquals("", row1.getCell(1).getString());
            assertTrue(row1.getCell(1).isEmpty());
            assertEquals("New York", row1.getCell(2).getString());

            // Read second data row (empty first and last fields)
            CsvRow row2 = reader.next();
            assertNotNull(row2);
            assertEquals(3, row2.size());
            assertEquals("", row2.getCell(0).getString());
            assertTrue(row2.getCell(0).isEmpty());
            assertEquals("30", row2.getCell(1).getString());
            assertEquals("", row2.getCell(2).getString());
            assertTrue(row2.getCell(2).isEmpty());

            // Read third data row (all empty quoted fields)
            CsvRow row3 = reader.next();
            assertNotNull(row3);
            assertEquals(3, row3.size());
            assertEquals("", row3.getCell(0).getString());
            assertEquals("", row3.getCell(1).getString());
            assertEquals("", row3.getCell(2).getString());
        }
    }

    @Test
    void testNumberParsing() throws IOException {
        String csvContent = "Integer,Decimal,Negative,Scientific,Invalid\n123,45.67,-123.45,1.23E+10,abc\n0,0.0,-0.0,1E-5,123abc";
        ByteArrayInputStream inputStream = new ByteArrayInputStream(csvContent.getBytes(StandardCharsets.UTF_8));

        try (KeelCsvReader reader = new KeelCsvReader(inputStream, StandardCharsets.UTF_8)) {
            // Read header
            CsvRow header = reader.next();
            assertNotNull(header);
            assertEquals(5, header.size());

            // Read first data row
            CsvRow row1 = reader.next();
            assertNotNull(row1);
            assertEquals(5, row1.size());

            // Test valid numbers
            assertTrue(row1.getCell(0).isNumber());
            assertEquals("123", row1.getCell(0).getString());
            assertEquals(123, row1.getCell(0).getNumber().intValue());

            assertTrue(row1.getCell(1).isNumber());
            assertEquals("45.67", row1.getCell(1).getString());
            assertEquals(45.67, row1.getCell(1).getNumber().doubleValue(), 0.001);

            assertTrue(row1.getCell(2).isNumber());
            assertEquals("-123.45", row1.getCell(2).getString());
            assertEquals(-123.45, row1.getCell(2).getNumber().doubleValue(), 0.001);

            assertTrue(row1.getCell(3).isNumber());
            assertEquals("1.23E+10", row1.getCell(3).getString());

            // Test invalid number
            assertFalse(row1.getCell(4).isNumber());
            assertEquals("abc", row1.getCell(4).getString());
            assertThrows(NumberFormatException.class, () -> row1.getCell(4).getNumber());

            // Read second data row
            CsvRow row2 = reader.next();
            assertNotNull(row2);
            assertEquals(5, row2.size());

            // Test zero values
            assertTrue(row2.getCell(0).isNumber());
            assertEquals("0", row2.getCell(0).getString());
            assertEquals(0, row2.getCell(0).getNumber().intValue());

            assertTrue(row2.getCell(1).isNumber());
            assertEquals("0.0", row2.getCell(1).getString());
            assertEquals(0.0, row2.getCell(1).getNumber().doubleValue(), 0.001);

            assertTrue(row2.getCell(2).isNumber());
            assertEquals("-0.0", row2.getCell(2).getString());
            assertEquals(0.0, row2.getCell(2).getNumber().doubleValue(), 0.001);

            assertTrue(row2.getCell(3).isNumber());
            assertEquals("1E-5", row2.getCell(3).getString());

            // Test invalid number with digits
            assertFalse(row2.getCell(4).isNumber());
            assertEquals("123abc", row2.getCell(4).getString());
            assertThrows(NumberFormatException.class, () -> row2.getCell(4).getNumber());
        }
    }

    @Test
    void testReadFromFile(@TempDir Path tempDir) throws IOException {
        // Create a test CSV file
        Path csvFile = tempDir.resolve("test.csv");
        String csvContent = "Name,Age,City\nJohn,25,New York\nJane,30,Los Angeles";
        Files.write(csvFile, csvContent.getBytes(StandardCharsets.UTF_8));

        try (KeelCsvReader reader = new KeelCsvReader(new FileInputStream(csvFile.toFile()), StandardCharsets.UTF_8)) {
            // Read header
            CsvRow header = reader.next();
            assertNotNull(header);
            assertEquals(3, header.size());
            assertEquals("Name", header.getCell(0).getString());
            assertEquals("Age", header.getCell(1).getString());
            assertEquals("City", header.getCell(2).getString());

            // Read data rows
            CsvRow row1 = reader.next();
            assertNotNull(row1);
            assertEquals("John", row1.getCell(0).getString());
            assertEquals("25", row1.getCell(1).getString());
            assertEquals("New York", row1.getCell(2).getString());

            CsvRow row2 = reader.next();
            assertNotNull(row2);
            assertEquals("Jane", row2.getCell(0).getString());
            assertEquals("30", row2.getCell(1).getString());
            assertEquals("Los Angeles", row2.getCell(2).getString());

            assertNull(reader.next());
        }
    }

    @Test
    void testReadFromExistingTestFile() throws IOException {
        String testFile = "src/test/resources/runtime/csv/write_test.csv";

        FileInputStream fis = new FileInputStream(testFile);
        try (KeelCsvReader reader = new KeelCsvReader(fis, StandardCharsets.UTF_8)) {
            // Read header
            CsvRow header = reader.next();
            assertNotNull(header);
            assertEquals(3, header.size());
            assertEquals("Name", header.getCell(0).getString());
            assertEquals("Age", header.getCell(1).getString());
            assertEquals("Number", header.getCell(2).getString());

            // Read first data row
            CsvRow row1 = reader.next();
            assertNotNull(row1);
            assertEquals(3, row1.size());
            assertEquals("Asana", row1.getCell(0).getString());
            assertEquals("18", row1.getCell(1).getString());
            assertTrue(row1.getCell(1).isNumber());
            assertEquals("1.0", row1.getCell(2).getString());
            assertTrue(row1.getCell(2).isNumber());

            // Read second data row (contains quotes)
            CsvRow row2 = reader.next();
            assertNotNull(row2);
            assertEquals(3, row2.size());
            assertEquals("Bi'zu'mi", row2.getCell(0).getString());
            assertEquals("-73", row2.getCell(1).getString());
            assertTrue(row2.getCell(1).isNumber());
            assertEquals("10000000000000.0", row2.getCell(2).getString());
            assertTrue(row2.getCell(2).isNumber());

            // Read third data row (contains quotes and large numbers)
            CsvRow row3 = reader.next();
            assertNotNull(row3);
            assertEquals(3, row3.size());
            assertEquals("Calapa", row3.getCell(0).getString());
            assertEquals("10000000000003", row3.getCell(1).getString());
            assertTrue(row3.getCell(1).isNumber());
            assertEquals("-2.212341242345235", row3.getCell(2).getString());
            assertTrue(row3.getCell(2).isNumber());

            // Read fourth data row (first part of multi-line field)
            CsvRow row4 = reader.next();
            assertNotNull(row4);
            assertEquals(1, row4.size());
            assertEquals("Da", row4.getCell(0).getString());

            // Read fifth data row (second part of multi-line field)
            CsvRow row5 = reader.next();
            assertNotNull(row5);
            assertEquals(3, row5.size());
            assertEquals("nue", row5.getCell(0).getString());
            assertEquals("-1788888883427777", row5.getCell(1).getString());
            assertTrue(row5.getCell(1).isNumber());
            assertEquals("0.0", row5.getCell(2).getString());
            assertTrue(row5.getCell(2).isNumber());

            // Check for end of file
            CsvRow row6 = reader.next();
            assertNull(row6);
        }
    }


    @Test
    void testEmptyFile() throws IOException {
        String csvContent = "";
        ByteArrayInputStream inputStream = new ByteArrayInputStream(csvContent.getBytes(StandardCharsets.UTF_8));

        try (KeelCsvReader reader = new KeelCsvReader(inputStream, StandardCharsets.UTF_8)) {
            assertNull(reader.next());
        }
    }

    @Test
    void testSingleLine() throws IOException {
        String csvContent = "Name,Age,City";
        ByteArrayInputStream inputStream = new ByteArrayInputStream(csvContent.getBytes(StandardCharsets.UTF_8));

        try (KeelCsvReader reader = new KeelCsvReader(inputStream, StandardCharsets.UTF_8)) {
            CsvRow row = reader.next();
            assertNotNull(row);
            assertEquals(3, row.size());
            assertEquals("Name", row.getCell(0).getString());
            assertEquals("Age", row.getCell(1).getString());
            assertEquals("City", row.getCell(2).getString());

            assertNull(reader.next());
        }
    }

    @Test
    void testTrailingSeparator() throws IOException {
        String csvContent = "Name,Age,City,\nJohn,25,New York,\nJane,30,Los Angeles,";
        ByteArrayInputStream inputStream = new ByteArrayInputStream(csvContent.getBytes(StandardCharsets.UTF_8));

        try (KeelCsvReader reader = new KeelCsvReader(inputStream, StandardCharsets.UTF_8)) {
            // Read header
            CsvRow header = reader.next();
            assertNotNull(header);
            assertEquals(4, header.size());
            assertEquals("Name", header.getCell(0).getString());
            assertEquals("Age", header.getCell(1).getString());
            assertEquals("City", header.getCell(2).getString());
            assertEquals("", header.getCell(3).getString());

            // Read first data row
            CsvRow row1 = reader.next();
            assertNotNull(row1);
            assertEquals(4, row1.size());
            assertEquals("John", row1.getCell(0).getString());
            assertEquals("25", row1.getCell(1).getString());
            assertEquals("New York", row1.getCell(2).getString());
            assertEquals("", row1.getCell(3).getString());

            // Read second data row
            CsvRow row2 = reader.next();
            assertNotNull(row2);
            assertEquals(4, row2.size());
            assertEquals("Jane", row2.getCell(0).getString());
            assertEquals("30", row2.getCell(1).getString());
            assertEquals("Los Angeles", row2.getCell(2).getString());
            assertEquals("", row2.getCell(3).getString());
        }
    }

    @Test
    void testCsvCellMethods() {
        // Test null cell
        CsvCell nullCell = new CsvCell(null);
        assertTrue(nullCell.isNull());
        assertFalse(nullCell.isEmpty());
        assertFalse(nullCell.isNumber());
        assertNull(nullCell.getString());
        assertNull(nullCell.getNumber());

        // Test empty cell
        CsvCell emptyCell = new CsvCell("");
        assertFalse(emptyCell.isNull());
        assertTrue(emptyCell.isEmpty());
        assertFalse(emptyCell.isNumber());
        assertEquals("", emptyCell.getString());
        assertThrows(NumberFormatException.class, () -> emptyCell.getNumber());

        // Test number cell
        CsvCell numberCell = new CsvCell("123.45");
        assertFalse(numberCell.isNull());
        assertFalse(numberCell.isEmpty());
        assertTrue(numberCell.isNumber());
        assertEquals("123.45", numberCell.getString());
        assertEquals(123.45, numberCell.getNumber().doubleValue(), 0.001);

        // Test text cell
        CsvCell textCell = new CsvCell("Hello World");
        assertFalse(textCell.isNull());
        assertFalse(textCell.isEmpty());
        assertFalse(textCell.isNumber());
        assertEquals("Hello World", textCell.getString());
        assertThrows(NumberFormatException.class, () -> textCell.getNumber());
    }

    @Test
    void testCsvRowMethods() {
        CsvRow row = new CsvRow();
        assertEquals(0, row.size());

        CsvCell cell1 = new CsvCell("Name");
        CsvCell cell2 = new CsvCell("25");
        CsvCell cell3 = new CsvCell("City");

        row.addCell(cell1).addCell(cell2).addCell(cell3);
        assertEquals(3, row.size());

        assertEquals("Name", row.getCell(0).getString());
        assertEquals("25", row.getCell(1).getString());
        assertEquals("City", row.getCell(2).getString());

        // Test index out of bounds
        assertThrows(IndexOutOfBoundsException.class, () -> row.getCell(3));
        assertThrows(IndexOutOfBoundsException.class, () -> row.getCell(-1));
    }

    @Test
    void testTabSeparator() throws IOException {
        String csvContent = "Name\tAge\tCity\nJohn\t25\tNew York\nJane\t30\tLos Angeles";
        ByteArrayInputStream inputStream = new ByteArrayInputStream(csvContent.getBytes(StandardCharsets.UTF_8));

        try (KeelCsvReader reader = new KeelCsvReader(inputStream, StandardCharsets.UTF_8, "\t")) {
            CsvRow header = reader.next();
            assertNotNull(header);
            assertEquals(3, header.size());
            assertEquals("Name", header.getCell(0).getString());
            assertEquals("Age", header.getCell(1).getString());
            assertEquals("City", header.getCell(2).getString());

            CsvRow row1 = reader.next();
            assertNotNull(row1);
            assertEquals("John", row1.getCell(0).getString());
            assertEquals("25", row1.getCell(1).getString());
            assertEquals("New York", row1.getCell(2).getString());
        }
    }

    @Test
    void testEmptyLines() throws IOException {
        String csvContent = "Name,Age,City\n\nJohn,25,New York\n\nJane,30,Los Angeles\n";
        ByteArrayInputStream inputStream = new ByteArrayInputStream(csvContent.getBytes(StandardCharsets.UTF_8));

        try (KeelCsvReader reader = new KeelCsvReader(inputStream, StandardCharsets.UTF_8)) {
            // Read header
            CsvRow header = reader.next();
            assertNotNull(header);
            assertEquals(3, header.size());

            // Read empty line (should be a row with one empty cell)
            CsvRow emptyRow1 = reader.next();
            assertNotNull(emptyRow1);
            assertEquals(1, emptyRow1.size());
            assertEquals("", emptyRow1.getCell(0).getString());

            // Read first data row
            CsvRow row1 = reader.next();
            assertNotNull(row1);
            assertEquals(3, row1.size());
            assertEquals("John", row1.getCell(0).getString());

            // Read another empty line
            CsvRow emptyRow2 = reader.next();
            assertNotNull(emptyRow2);
            assertEquals(1, emptyRow2.size());
            assertEquals("", emptyRow2.getCell(0).getString());

            // Read second data row
            CsvRow row2 = reader.next();
            assertNotNull(row2);
            assertEquals(3, row2.size());
            assertEquals("Jane", row2.getCell(0).getString());

            // End of file
            assertNull(reader.next());
        }
    }


    @Test
    void testCsvCellNumberMethods() {
        // Test getNumberOrElse methods
        CsvCell validNumber = new CsvCell("123.45");
        CsvCell invalidNumber = new CsvCell("abc");
        CsvCell nullCell = new CsvCell(null);
        CsvCell emptyCell = new CsvCell("");

        // Test valid number
        assertEquals(123.45, validNumber.getNumberOrElse(0.0).doubleValue(), 0.001);
        assertEquals(123, validNumber.getNumberOrElse(0).intValue());
        assertEquals(123L, validNumber.getNumberOrElse(0L).longValue());
        assertEquals(123.45f, validNumber.getNumberOrElse(0.0f).floatValue(), 0.001);

        // Test invalid number
        assertEquals(999.0, invalidNumber.getNumberOrElse(999.0).doubleValue(), 0.001);
        assertEquals(999, invalidNumber.getNumberOrElse(999).intValue());
        assertEquals(999L, invalidNumber.getNumberOrElse(999L).longValue());
        assertEquals(999.0f, invalidNumber.getNumberOrElse(999.0f).floatValue(), 0.001);

        // Test null cell
        assertEquals(0.0, nullCell.getNumberOrElse(0.0).doubleValue(), 0.001);
        assertEquals(0, nullCell.getNumberOrElse(0).intValue());
        assertEquals(0L, nullCell.getNumberOrElse(0L).longValue());
        assertEquals(0.0f, nullCell.getNumberOrElse(0.0f).floatValue(), 0.001);

        // Test empty cell
        assertEquals(100.0, emptyCell.getNumberOrElse(100.0).doubleValue(), 0.001);
        assertEquals(100, emptyCell.getNumberOrElse(100).intValue());
        assertEquals(100L, emptyCell.getNumberOrElse(100L).longValue());
        assertEquals(100.0f, emptyCell.getNumberOrElse(100.0f).floatValue(), 0.001);
    }
}