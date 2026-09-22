package io.github.sinri.keel.integration.poi.excel;

import io.vertx.core.Future;
import io.vertx.core.Promise;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FilterInputStream;
import java.io.IOException;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeUnit;

import static org.junit.jupiter.api.Assertions.*;

class KeelInputStreamOwnershipTest {
    @Test
    void callerOwnsStreamUntilAndAfterAsyncCompletion() throws Exception {
        for (int mode = 0; mode < 5; mode++) {
            for (boolean fail : new boolean[]{false, true}) {
                try (TrackedInput input = new TrackedInput(workbookBytes(mode))) {
                    Promise<Void> pending = Promise.promise();
                    RuntimeException failure = new RuntimeException("usage failed");
                    Future<Void> result = KeelSheets.useSheets(options(mode, input), sheets -> {
                        assertFalse(input.closed);
                        assertEquals("header", sheets.getWorkbook().getSheetAt(0)
                                .iterator().next().getCell(0).getStringCellValue());
                        return pending.future();
                    });
                    assertFalse(result.isComplete());
                    assertFalse(input.closed);
                    if (fail) {
                        pending.fail(failure);
                        ExecutionException thrown = assertThrows(ExecutionException.class,
                                () -> result.toCompletionStage().toCompletableFuture().get(30, TimeUnit.SECONDS));
                        assertSame(failure, thrown.getCause());
                    } else {
                        pending.complete();
                        result.toCompletionStage().toCompletableFuture().get(30, TimeUnit.SECONDS);
                    }
                    assertFalse(input.closed, "mode=" + mode);
                    input.close();
                    assertTrue(input.closed);
                }
            }
        }
    }

    @Test
    void synchronousCallbackFailureLeavesOriginalStreamOpen() throws Exception {
        for (int mode = 0; mode < 5; mode++) {
            try (TrackedInput input = new TrackedInput(workbookBytes(mode))) {
                RuntimeException failure = new RuntimeException("synchronous failure");
                Future<Void> result = KeelSheets.useSheets(options(mode, input), sheets -> {
                    throw failure;
                });
                ExecutionException thrown = assertThrows(ExecutionException.class,
                        () -> result.toCompletionStage().toCompletableFuture().get(30, TimeUnit.SECONDS));
                assertSame(failure, thrown.getCause());
                assertFalse(input.closed, "mode=" + mode);
            }
        }
    }

    @Test
    void openingFailureLeavesOriginalStreamOpen() throws Exception {
        for (int mode = 0; mode < 5; mode++) {
            for (byte[] bytes : new byte[][]{new byte[0], new byte[]{1, 2, 3}}) {
                try (TrackedInput input = new TrackedInput(bytes)) {
                    Future<Void> result = KeelSheets.useSheets(options(mode, input), sheets -> {
                        fail("Invalid input must not invoke usage");
                        return Future.succeededFuture();
                    });
                    assertThrows(ExecutionException.class,
                            () -> result.toCompletionStage().toCompletableFuture().get(30, TimeUnit.SECONDS));
                    assertFalse(input.closed, "mode=" + mode);
                }
            }
        }
    }

    // Auto XLS, auto XLSX, explicit XLS, explicit XLSX, streaming XLSX.
    private static SheetsOpenOptions options(int mode, TrackedInput input) {
        SheetsOpenOptions options = new SheetsOpenOptions().setInputStream(input);
        if (mode == 2 || mode == 3) options.setUseXlsx(mode == 3);
        if (mode == 4) options.setHugeXlsxStreamingReaderBuilder(builder -> builder.rowCacheSize(2));
        return options;
    }

    private static byte[] workbookBytes(int mode) throws IOException {
        try (Workbook workbook = mode == 0 || mode == 2 ? new HSSFWorkbook() : new XSSFWorkbook();
             ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            workbook.createSheet("S").createRow(0).createCell(0).setCellValue("header");
            workbook.write(output);
            return output.toByteArray();
        }
    }

    private static final class TrackedInput extends FilterInputStream {
        private boolean closed;

        private TrackedInput(byte[] bytes) {
            super(new ByteArrayInputStream(bytes));
        }

        @Override
        public boolean markSupported() {
            return false;
        }

        @Override
        public void close() throws IOException {
            closed = true;
            super.close();
        }
    }
}
