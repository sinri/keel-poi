package io.github.sinri.keel.integration.poi.csv;

import org.junit.jupiter.api.Test;

import java.io.BufferedReader;
import java.io.ByteArrayOutputStream;
import java.io.StringReader;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class KeelCsvReaderNewlineReproductionTest {
    @Test
    void readsMixedRecordBoundariesAndFinalRowWithoutNewline() throws Exception {
        try (KeelCsvReader reader = new KeelCsvReader(new BufferedReader(
                new StringReader("a,b\rc,d\ne,f\r\n\r\nlast,")))) {
            for (String first : List.of("a", "c", "e", "", "last")) {
                CsvRow row = reader.next();
                assertNotNull(row);
                assertEquals(first, row.getCell(0).getString());
                assertEquals(first.isEmpty() ? 1 : 2, row.size());
                if (first.equals("last")) assertEquals("", row.getCell(1).getString());
            }
            assertNull(reader.next());
            assertNull(reader.next());
        }
    }

    @Test
    void roundTripsMixedNewlinesQuotesAndMultiCharacterSeparator() throws Exception {
        String value = "a\r\nb\rc\n\"quoted\"||tail\r\n";
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (KeelCsvWriter writer = new KeelCsvWriter(output, "||", StandardCharsets.UTF_8)) {
            writer.blockWriteRow(List.of(value, "x|y", ""));
            writer.blockWriteRow(List.of("next", "row"));
        }
        try (KeelCsvReader reader = new KeelCsvReader(new BufferedReader(
                new StringReader(output.toString(StandardCharsets.UTF_8)), 1), "||")) {
            CsvRow row = reader.next();
            assertNotNull(row);
            assertEquals(3, row.size());
            assertEquals(value, row.getCell(0).getString());
            assertEquals("x|y", row.getCell(1).getString());
            assertEquals("", row.getCell(2).getString());
            CsvRow next = reader.next();
            assertNotNull(next);
            assertEquals("next", next.getCell(0).getString());
            assertEquals("row", next.getCell(1).getString());
            assertNull(reader.next());
        }
    }

    @Test
    void preservesPartialSeparatorAtRecordBoundaryAndEof() throws Exception {
        try (KeelCsvReader reader = new KeelCsvReader(new BufferedReader(
                new StringReader("x|\ry|")), "||")) {
            assertEquals("x|", reader.next().getCell(0).getString());
            assertEquals("y|", reader.next().getCell(0).getString());
            assertNull(reader.next());
        }
    }

    @Test
    void readsCrLfUnchanged() throws Exception {
        verifyRead("a\r\nb", "\"a\r\nb\",c\r\n");
    }

    @Test
    void readsCrUnchanged() throws Exception {
        verifyRead("a\rb", "\"a\rb\",c\r\n");
    }

    @Test
    void readsLfUnchanged() throws Exception {
        verifyRead("a\nb", "\"a\nb\",c\r\n");
    }

    @Test
    void roundTripsCrLfUnchanged() throws Exception {
        verifyRoundTrip("a\r\nb");
    }

    @Test
    void roundTripsCrUnchanged() throws Exception {
        verifyRoundTrip("a\rb");
    }

    @Test
    void roundTripsLfUnchanged() throws Exception {
        verifyRoundTrip("a\nb");
    }

    private void verifyRoundTrip(String value) throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (KeelCsvWriter writer = new KeelCsvWriter(output)) {
            writer.blockWriteRow(List.of(value, "c"));
        }
        String csv = output.toString(StandardCharsets.UTF_8);
        assertEquals("\"" + value + "\",c\r\n", csv, "Writer must preserve field newlines");
        verifyRead(value, csv);
    }

    private void verifyRead(String expected, String csv) throws Exception {
        try (KeelCsvReader reader = new KeelCsvReader(new BufferedReader(new StringReader(csv)))) {
            CsvRow row = reader.next();
            assertNotNull(row);
            assertEquals(2, row.size());
            assertEquals("c", row.getCell(1).getString());
            assertNull(reader.next());
            String actual = row.getCell(0).getString();
            assertEquals(expected, actual,
                    () -> "expected=" + visible(expected) + ", actual=" + visible(actual));
        }
    }

    private String visible(String value) {
        return value.replace("\r", "<CR>").replace("\n", "<LF>");
    }
}
