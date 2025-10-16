package com.davita.botcommand.excel;

import com.davita.botcommand.excel.internal.SheetUtility; // where deleteColumns(...) lives
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

import static org.junit.Assert.assertNotNull;

public class DeleteColumnsResourceTest {

    @Test
    public void deleteColumns_fromResource_writeOutputAlongside_overwriteIfExists() throws Exception {
        URL url = getClass().getResource("/excel/Excel_TestFile.xlsx");
        assertNotNull(url);

        // Copy resource stream to a temp working file
        Path tmpDir = Files.createTempDirectory("poi_test_");
        Path input = tmpDir.resolve("Excel_TestFile.xlsx");
        try (var in = getClass().getResourceAsStream("/excel/Excel_TestFile.xlsx")) {
            Files.copy(in, input, java.nio.file.StandardCopyOption.REPLACE_EXISTING);
        }

        // Work with input and write output alongside it
        Path out = tmpDir.resolve("Excel_TestFile_out.xlsx");
        Files.deleteIfExists(out);

        // open/read-write then save
        try (var pkg = org.apache.poi.openxml4j.opc.OPCPackage.open(input.toFile(),
                org.apache.poi.openxml4j.opc.PackageAccess.READ_WRITE);
             var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(pkg)) {

            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                var sh = wb.getSheetAt(i);
                com.davita.botcommand.excel.internal.SheetUtility.deleteColumns(sh, 0, 0);
            }
//            var sh = wb.getSheet("Sheet3");
//            com.davita.botcommand.excel.internal.SheetUtility.deleteColumns(sh, 1, 1);

            try (var outStream = new java.io.FileOutputStream(out.toFile())) {
                wb.write(outStream);
            }
        }

//        if (java.awt.Desktop.isDesktopSupported()) {
//            java.awt.Desktop.getDesktop().open(out.toFile());
//        }
//
//        // optional assertion open
//        try (var pkg2 = org.apache.poi.openxml4j.opc.OPCPackage.open(out.toFile(),
//                org.apache.poi.openxml4j.opc.PackageAccess.READ);
//             var wb2 = new org.apache.poi.xssf.usermodel.XSSFWorkbook(pkg2)) {
//            assertNotNull(wb2.getSheet("Sheet3"));
//        }

    }
}
