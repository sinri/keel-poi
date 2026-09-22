package io.github.sinri.keel.integration.poi.csv;

import io.vertx.core.Future;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class KeelCsvWriterCharsetTest {
    @Test
    void utf8WritesPreserveContent() throws IOException {
        verifyWrites(StandardCharsets.UTF_8);
    }

    @Test
    void utf16WritesPreserveContentWithSingleLeadingBom() throws IOException {
        verifyWrites(StandardCharsets.UTF_16);
    }

    private void verifyWrites(Charset charset) throws IOException {
        for (String mode : List.of("cells", "rows", "mixed", "incompleteRow")) {
            var out = new ByteArrayOutputStream();
            try (var writer = new KeelCsvWriter(out, ",", charset)) {
                if (mode.equals("rows")) {
                    writer.blockWriteRow(List.of("a", "b"));
                } else {
                    writer.writeCell("a");
                    writer.writeCell("b");
                    if (!mode.equals("incompleteRow")) {
                        writer.writeRowEnding();
                    }
                }
                if (mode.equals("cells")) {
                    writer.writeCell("中文");
                    writer.writeCell("d,\"e\"");
                    writer.writeRowEnding();
                } else {
                    writer.blockWriteRow(List.of("中文", "d,\"e\""));
                }
            }
            String expected = "a,b\r\n中文,\"d,\"\"e\"\"\"\r\n";
            assertEquals(expected, out.toString(charset), mode);
            // Whole-string encoding contains exactly one leading BOM for UTF-16.
            assertArrayEquals(expected.getBytes(charset), out.toByteArray(), mode);
        }
    }

    @Test
    void closeWritesBufferedContentAndClosesUnderlyingStream() throws IOException {
        var out = new TrackingOutputStream();
        var writer = new KeelCsvWriter(out, ",", StandardCharsets.UTF_16);
        writer.writeCell("pending");
        writer.close();
        assertEquals("pending", out.bytes.toString(StandardCharsets.UTF_16));
        assertTrue(out.closed);
    }

    @Test
    void staticWriteClosesAndWritesContentAfterCallbackFailure() {
        for (boolean synchronous : List.of(false, true)) {
            var out = new TrackingOutputStream();
            var failure = new IllegalStateException("callback failed");
            var result = KeelCsvWriter.write(out, ",", StandardCharsets.UTF_16, writer -> {
                try {
                    writer.writeCell("pending");
                } catch (IOException e) {
                    return Future.failedFuture(e);
                }
                if (synchronous) {
                    throw failure;
                }
                return Future.failedFuture(failure);
            });
            assertSame(failure, result.cause());
            assertEquals("pending", out.bytes.toString(StandardCharsets.UTF_16));
            assertTrue(out.closed);
        }
    }

    @Test
    void staticWritePropagatesBufferedWriteFailureAndReleasesStream() {
        var out = new TrackingOutputStream();
        out.writeFailure = new IOException("write failed");
        var result = KeelCsvWriter.write(out, writer -> {
            try {
                writer.writeCell("pending");
                return Future.succeededFuture();
            } catch (IOException e) {
                return Future.failedFuture(e);
            }
        });
        assertSame(out.writeFailure, result.cause());
        assertTrue(out.closed);
    }

    @Test
    void staticWritePropagatesCloseFailureOrSuppressesItAfterCallbackFailure() {
        for (boolean callbackFails : List.of(false, true)) {
            var out = new TrackingOutputStream();
            out.closeFailure = new IOException("close failed");
            var failure = new IllegalStateException("callback failed");
            var result = KeelCsvWriter.write(out, writer -> callbackFails
                    ? Future.failedFuture(failure) : Future.succeededFuture());
            assertTrue(out.closed);
            if (callbackFails) {
                assertSame(failure, result.cause());
                assertArrayEquals(new Throwable[]{out.closeFailure}, failure.getSuppressed());
            } else {
                assertSame(out.closeFailure, result.cause());
            }
        }
    }

    private static class TrackingOutputStream extends OutputStream {
        private final ByteArrayOutputStream bytes = new ByteArrayOutputStream();
        private boolean closed;
        private IOException writeFailure;
        private IOException closeFailure;

        @Override
        public void write(int b) throws IOException {
            if (writeFailure != null) {
                throw writeFailure;
            }
            bytes.write(b);
        }

        @Override
        public void close() throws IOException {
            closed = true;
            if (closeFailure != null) {
                throw closeFailure;
            }
        }
    }
}
