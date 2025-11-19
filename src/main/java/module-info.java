module io.github.sinri.keel.integration.poi {
    requires com.github.pjfanning.excelstreamingreader;
    requires io.github.sinri.keel.core;
    requires io.vertx.core;
    requires org.apache.poi.ooxml;
    requires org.apache.poi.poi;
    requires org.jetbrains.annotations;

    exports io.github.sinri.keel.integration.poi.csv;
    exports io.github.sinri.keel.integration.poi.excel;
    exports io.github.sinri.keel.integration.poi.excel.entity;
}