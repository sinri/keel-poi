# Keel POI Integration

`io.github.sinri:keel-poi` is the Keel integration library for reading and writing Excel and CSV files on top of the Keel / Vert.x `Future` model.

It wraps Apache POI, excel-streaming-reader, and a lightweight CSV reader/writer behind a smaller API for common workbook, sheet, matrix, templated-row, and CSV workflows.

Latest version: **5.0.2**

- User guide: [docs/5.0.2/index.md](./docs/5.0.2/index.md)
- Review report: [docs/5.0.2/risks.md](./docs/5.0.2/risks.md)
- Version index: [docs/index.md](./docs/index.md)

## Requirements

- JDK 17
- Keel / Vert.x runtime APIs, provided transitively through `keel-core`
- License: GPL-3.0

## Installation

Maven:

```xml
<dependency>
    <groupId>io.github.sinri</groupId>
    <artifactId>keel-poi</artifactId>
    <version>5.0.2</version>
</dependency>
```

Gradle Kotlin DSL:

```kotlin
implementation("io.github.sinri:keel-poi:5.0.2")
```

## Main APIs

| API | Purpose |
|-----|---------|
| `KeelSheets` | Open or create workbooks and close them automatically after `Future` workflows complete. |
| `SheetsOpenOptions` / `SheetsCreateOptions` | Configure workbook input, output, XLS/XLSX mode, formula evaluation, large XLSX streaming reads, and stream writing. |
| `KeelSheet` | Read and write a single sheet, including matrices, templated matrices, raw rows, formulas, and embedded pictures. |
| `KeelSheetMatrix` / `KeelSheetTemplatedMatrix` | Represent table data as header + rows, with optional column-name access. |
| `KeelCsvReader` / `KeelCsvWriter` | Read and write CSV with configurable separator and charset. |

## Excel Example

```java
import io.github.sinri.keel.integration.poi.excel.KeelSheet;
import io.github.sinri.keel.integration.poi.excel.KeelSheets;
import io.github.sinri.keel.integration.poi.excel.SheetsOpenOptions;
import io.github.sinri.keel.integration.poi.excel.entity.KeelSheetMatrix;
import io.vertx.core.Future;

Future<KeelSheetMatrix> future = KeelSheets.useSheets(
        new SheetsOpenOptions()
                .setFile("/path/to/workbook.xlsx")
                .setWithFormulaEvaluator(true),
        sheets -> {
            KeelSheet sheet = sheets.generateReaderForSheet("Sheet1");
            KeelSheetMatrix matrix = sheet.readAllRowsToMatrix();
            return Future.succeededFuture(matrix);
        }
);
```

## CSV Example

```java
import io.github.sinri.keel.integration.poi.csv.CsvRow;
import io.github.sinri.keel.integration.poi.csv.KeelCsvReader;
import io.vertx.core.Future;

import java.io.IOException;
import java.nio.charset.StandardCharsets;

Future<Void> future = KeelCsvReader.read(
        inputStream,
        StandardCharsets.UTF_8,
        ",",
        reader -> {
            try {
                CsvRow row;
                while ((row = reader.next()) != null) {
                    row.stream().forEach(cell -> {
                        String value = cell.getString();
                        // handle value
                    });
                }
                return Future.succeededFuture();
            } catch (IOException e) {
                return Future.failedFuture(e);
            }
        }
);
```

## Version 5.0.2 Notes

5.0.2 focuses on stability and dependency maintenance:

- Auto-close helpers now close resources even when user callbacks throw synchronously.
- Sheet row iterators skip rows excluded by `SheetRowFilter` instead of returning `null`.
- Templated row access by column name now matches indexed access for trailing missing cells.
- Matrix and template objects defensively copy row/header data and expose readonly views.
- Regression tests cover auto-close failure paths, row filtering, templated rows, defensive copies, XLSX streaming reads, and picture extraction.

## Build And Test

```bash
./gradlew test
```

## Documentation

Full usage details are maintained under `docs/`:

- [5.0.2 user guide](./docs/5.0.2/index.md)
- [5.0.2 review report](./docs/5.0.2/risks.md)
- [5.0.1 user guide](./docs/5.0.1/index.md)
