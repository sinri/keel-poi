module io.github.sinri.keel.integration.poi {
    requires io.github.sinri.keel.core;
    requires io.vertx.core;
    requires org.jetbrains.annotations;
    requires java.desktop;

    requires org.apache.poi.poi; // Core POI functionalities (kept as code may directly use ss/usermodel)
    requires org.apache.poi.ooxml; // OOXML functionality
    requires org.apache.commons.collections4; // Apache Commons Collections
    requires org.apache.commons.compress; // Apache Commons Compress for zip handling
    requires org.apache.commons.codec; // Apache Commons Codec for encoding/decoding
    requires org.apache.commons.io; // Apache Commons IO (Required by POI)
    // requires commons.math3; // Apache Commons Math 3 (Required by POI, filename-based automatic module) - Commented out as it's a transitive dependency

    requires com.github.pjfanning.excelstreamingreader;

    exports io.github.sinri.keel.integration.poi.csv;
    exports io.github.sinri.keel.integration.poi.excel;
    exports io.github.sinri.keel.integration.poi.excel.entity;
}