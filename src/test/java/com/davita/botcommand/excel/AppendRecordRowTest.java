package com.davita.botcommand.excel;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.BooleanValue;
import com.automationanywhere.botcommand.data.impl.NumberValue;
import com.davita.botcommand.excel.commands.dataOperations.AppendRecordToSheet;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.data.model.Schema;
import com.automationanywhere.botcommand.data.model.record.Record;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;

import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedList;
import java.util.List;

public class AppendRecordRowTest {

    // Choose an existing sheet present in your resource workbook (e.g., "Sheet3")
    private static final String SHEET_NAME = "Sheet1";

    /**
     * Prints a clickable file path link in IntelliJ console
     */
    private void logClickableFile(String message, Path filePath) {
        // Format compatible with IntelliJ IDEA on Windows
        String clickablePath = filePath.toAbsolutePath().toString().replace("\\", "/");
        System.out.println(message.trim() + " " + clickablePath);
    }

    @Test
    public void appendRecord_usesExistingSheet_logsAndVerifies() throws Exception {
        // 1) Copy resource workbook to a temp path (no pre-open or sheet creation)
        URL url = getClass().getResource("/excel/Excel_TestFile.xlsx");
        Assert.assertNotNull("Test resource not found", url);
        Path tmpDir = Files.createTempDirectory("append_row_");
        logClickableFile("[TEST] Created test directory at:",tmpDir);

        Path work = tmpDir.resolve("Excel_TestFile_out.xlsx");
        Files.copy(url.openStream(), work, java.nio.file.StandardCopyOption.REPLACE_EXISTING);

        logClickableFile("[TEST] Copied test workbook to:",work);

        String text = "test";
        double number = 123;
        boolean flag = true;

        List<Schema> schemas = new LinkedList<>();
        List<Value> values = new LinkedList<>();
        schemas.add(new Schema("Text"));
        values.add(new StringValue(text));
        schemas.add(new Schema("Number"));
        values.add(new NumberValue(number));
        schemas.add(new Schema("Bool"));
        values.add(new BooleanValue(flag));
        Record record = new Record(schemas,values);
        System.out.println("[TEST] Created record variable");
        // 3) Execute the command (uses existing sheet, does not create new)
        AppendRecordToSheet cmd = new AppendRecordToSheet();
        logClickableFile("[TEST] Appending record to sheet '" + SHEET_NAME + "' at:",work);
        cmd.action(work.toString(), SHEET_NAME,true, record);
        System.out.println("[TEST] Append command completed, reopening for verification.");

//        if (java.awt.Desktop.isDesktopSupported()) {
//            java.awt.Desktop.getDesktop().open(work.toFile());
//        }

        // 4) Reopen and verify last data row values (no pre-open before command)
        try (OPCPackage pkg = OPCPackage.open(work.toFile(), PackageAccess.READ);
             Workbook wb = new XSSFWorkbook(pkg)) {
            Sheet sh = wb.getSheet(SHEET_NAME);
            Assert.assertNotNull("Sheet not found: " + SHEET_NAME, sh);
            int lastRow = findLastDataRow(sh);
            int firstDataCol = findFirstDataColumn(sh);
            System.out.println("[TEST] lastDataRow=" + lastRow + ", firstDataCol=" + firstDataCol);

            Row appended = sh.getRow(lastRow);
            Assert.assertNotNull("Appended row not found", appended);
            String actualText = getAsString(appended.getCell(firstDataCol));
            double actualNumber = getAsNumeric(appended.getCell(firstDataCol + 1));
            boolean actualFlag = getAsBoolean(appended.getCell(firstDataCol + 2));
            System.out.println("[TEST] Appended values: text=" + actualText + ", number=" + actualNumber + ", flag=" + actualFlag);

            Assert.assertEquals(text, actualText);
            Assert.assertEquals(number, actualNumber, 0.0);
            Assert.assertEquals(flag, actualFlag);
        }
    }

    // -------- Helpers (read-only; mirror command’s logic) --------

    private static int findLastDataRow(Sheet sheet) {
        int last = sheet.getLastRowNum();
        for (int r = last; r >= 0; r--) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            short lastCell = row.getLastCellNum();
            if (lastCell < 0) continue;
            int firstCell = Math.max(0, row.getFirstCellNum());
            for (int c = firstCell; c < lastCell; c++) {
                Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell != null && !isBlankCell(cell)) return r;
            }
        }
        return -1;
    }

    private static int findFirstDataColumn(Sheet sheet) {
        int min = Integer.MAX_VALUE;
        int lastRow = sheet.getLastRowNum();
        for (int r = 0; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            short lastCell = row.getLastCellNum();
            if (lastCell < 0) continue;
            int firstCell = Math.max(0, row.getFirstCellNum());
            for (int c = firstCell; c < lastCell; c++) {
                Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell != null && !isBlankCell(cell)) {
                    if (c < min) min = c;
                    break;
                }
            }
        }
        return (min == Integer.MAX_VALUE) ? 0 : min;
    }

    private static boolean isBlankCell(Cell cell) {
        if (cell == null) return true;
        if (cell.getCellType() == CellType.BLANK) return true;
        if (cell.getCellType() == CellType.STRING) {
            String s = cell.getStringCellValue();
            return s == null || s.trim().isEmpty();
        }
        return false;
    }

    private static String getAsString(Cell c) {
        if (c == null) return "";
        if (c.getCellType() == CellType.STRING) return c.getStringCellValue();
        DataFormatter fmt = new DataFormatter();
        return fmt.formatCellValue(c);
    }

    private static double getAsNumeric(Cell c) {
        if (c == null) return 0d;
        if (c.getCellType() == CellType.NUMERIC) return c.getNumericCellValue();
        DataFormatter fmt = new DataFormatter();
        try {
            return Double.parseDouble(fmt.formatCellValue(c).replace(",", ""));
        } catch (Exception e) {
            return 0d;
        }
    }

    private static boolean getAsBoolean(Cell c) {
        if (c == null) return false;
        if (c.getCellType() == CellType.BOOLEAN) return c.getBooleanCellValue();
        DataFormatter fmt = new DataFormatter();
        String s = fmt.formatCellValue(c);
        return "TRUE".equalsIgnoreCase(s) || "Yes".equalsIgnoreCase(s) || "1".equals(s);
    }
}
