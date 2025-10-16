package com.davita.botcommand.excel;

import com.automationanywhere.botcommand.data.impl.SessionValue;
import com.davita.botcommand.excel.commands.sessionManagement.OpenWorkbook;
import com.davita.botcommand.excel.commands.workbookOperations.SaveWorkbook;
import com.davita.botcommand.excel.commands.workbookOperations.SaveWorkbookAs;
import com.davita.botcommand.excel.commands.workbookOperations.CreateWorkbook;
import com.davita.botcommand.excel.sessions.WorkbookSession;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;

public class OpenWorkbookTest {

    private static final String BOOK_NAME = "ExcelTest";
    private static final String SHEET_NAME = "Sheet2";

    @Test
    public void open_workbook_verify() throws Exception {
        // Copy resource workbooks to temp paths
        URL urlXLS = getClass().getResource("/excel/Excel_TestFile.xls");
        Assert.assertNotNull("Test XLS resource not found", urlXLS);

        URL urlXLSX = getClass().getResource("/excel/Excel_TestFile.xlsx");
        Assert.assertNotNull("Test XLSX resource not found", urlXLSX);

        URL urlXLSM = getClass().getResource("/excel/Excel_TestFile.xlsm");
        Assert.assertNotNull("Test XLSM resource not found", urlXLSM);

        Path tmpDir = Files.createTempDirectory("open_book_");

        System.out.println("[TEST] Testing XLS workbook opening");
        open_workbook_verify(BOOK_NAME + ".xls", SHEET_NAME, tmpDir, urlXLS);

        System.out.println("[TEST] Testing XLSX workbook opening");
        open_workbook_verify(BOOK_NAME + ".xlsx", SHEET_NAME, tmpDir, urlXLSX);

        System.out.println("[TEST] Testing XLSM workbook opening");
        open_workbook_verify(BOOK_NAME + ".xlsm", SHEET_NAME, tmpDir, urlXLSM);
    }

    @Test
    public void save_workbook_verify() throws Exception {
        Path tmpDir = Files.createTempDirectory("save_book_");

        System.out.println("[TEST] Testing SaveWorkbook for XLS format");
        save_workbook_verify(tmpDir, ".xls");

        System.out.println("[TEST] Testing SaveWorkbook for XLSX format");
        save_workbook_verify(tmpDir, ".xlsx");

        System.out.println("[TEST] Testing SaveWorkbook for XLSM format");
        save_workbook_verify(tmpDir, ".xlsm");
    }

    @Test
    public void save_workbook_as_verify() throws Exception {
        Path tmpDir = Files.createTempDirectory("save_as_book_");

        System.out.println("[TEST] Testing SaveWorkbookAs XLS -> XLS");
        save_as_verify(tmpDir, ".xls", ".xls");

        System.out.println("[TEST] Testing SaveWorkbookAs XLSX -> XLSX");
        save_as_verify(tmpDir, ".xlsx", ".xlsx");

        System.out.println("[TEST] Testing SaveWorkbookAs XLSX -> different path");
        save_as_different_path_verify(tmpDir, ".xlsx");

        System.out.println("[TEST] Testing SaveWorkbookAs with overwrite=false");
        save_as_overwrite_verify(tmpDir, ".xlsx");
    }

    // ========== HELPER METHODS ==========

    private void open_workbook_verify(String bookName, String sheetName, Path tmpDir, URL url) throws IOException {
        Path work = tmpDir.resolve(bookName);
        Files.copy(url.openStream(), work, java.nio.file.StandardCopyOption.REPLACE_EXISTING);

        System.out.println("[TEST] Copied test workbook to: " + work.toAbsolutePath());

        OpenWorkbook cmd = new OpenWorkbook();
        SessionValue sessionValue = null;

        // Test READ-WRITE mode
        try {
            System.out.println("[TEST] Trying to open the test workbook in READ-WRITE mode at: " + work);
            sessionValue = cmd.action(work.toString(), null, false);
            System.out.println("[TEST] Opened test successfully");

            Assert.assertNotNull("Session should not be null", sessionValue);
            Assert.assertNotNull("Workbook should not be null", sessionValue.getSession());

            // Verify workbook is actually usable
            WorkbookSession wbSession = (WorkbookSession) sessionValue.getSession();
            Workbook wb = wbSession.getWorkbook();
            Assert.assertTrue("Workbook should have at least one sheet", wb.getNumberOfSheets() > 0);

            System.out.println("[TEST] Workbook has " + wb.getNumberOfSheets() + " sheets");

        } finally {
            if (sessionValue != null) {
                try {
                    sessionValue.getSession().close();
                    System.out.println("[TEST] Closed READ-WRITE session successfully");
                } catch (Exception e) {
                    System.err.println("[TEST] Error closing session: " + e.getMessage());
                }
            }
        }

        // Test READ-ONLY mode
        sessionValue = null;
        try {
            System.out.println("[TEST] Trying to open the test workbook in READ mode at: " + work);
            sessionValue = cmd.action(work.toString(), null, true);
            System.out.println("[TEST] Opened test successfully in READ mode");

            Assert.assertNotNull("Session should not be null", sessionValue);
            Assert.assertNotNull("Workbook should not be null", sessionValue.getSession());

            WorkbookSession wbSession = (WorkbookSession) sessionValue.getSession();
            Assert.assertTrue("Session should be marked as read-only", wbSession.isReadOnly());

            // Verify workbook is actually usable
            Workbook wb = wbSession.getWorkbook();
            Assert.assertTrue("Workbook should have at least one sheet", wb.getNumberOfSheets() > 0);

            System.out.println("[TEST] Workbook has " + wb.getNumberOfSheets() + " sheets");

        } finally {
            if (sessionValue != null) {
                try {
                    sessionValue.getSession().close();
                    System.out.println("[TEST] Closed READ session successfully");
                } catch (Exception e) {
                    System.err.println("[TEST] Error closing session: " + e.getMessage());
                }
            }
        }
    }

    private void save_workbook_verify(Path tmpDir, String extension) throws Exception {
        Path workPath = tmpDir.resolve("SaveTest" + extension);

        System.out.println("[TEST] Creating new workbook at: " + workPath);

        // Create workbook
        CreateWorkbook createCmd = new CreateWorkbook();
        SessionValue sessionValue = createCmd.action(workPath.toString(), "TestSheet");

        Assert.assertNotNull("Session should not be null", sessionValue);
        WorkbookSession wbSession = (WorkbookSession) sessionValue.getSession();

        try {
            // Add some data to the workbook
            Workbook wb = wbSession.getWorkbook();
            Sheet sheet = wb.getSheetAt(0);
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue("Test Data for Save");

            System.out.println("[TEST] Added test data to workbook");

            // Save the workbook
            SaveWorkbook saveCmd = new SaveWorkbook();
            saveCmd.action(wbSession);

            System.out.println("[TEST] SaveWorkbook command executed successfully");

            // Verify file exists and has content
            File savedFile = new File(workPath.toString());
            Assert.assertTrue("Saved file should exist", savedFile.exists());
            Assert.assertTrue("Saved file should have content", savedFile.length() > 0);

            System.out.println("[TEST] Verified saved file exists with size: " + savedFile.length() + " bytes");

        } finally {
            wbSession.close();
            System.out.println("[TEST] Closed session after save test");
        }

        // Reopen and verify the saved data
        System.out.println("[TEST] Reopening saved file to verify data persistence");

        OpenWorkbook openCmd = new OpenWorkbook();
        sessionValue = openCmd.action(workPath.toString(), null, false);

        try {
            WorkbookSession reopenedSession = (WorkbookSession) sessionValue.getSession();
            Workbook wb = reopenedSession.getWorkbook();
            Sheet sheet = wb.getSheetAt(0);
            Row row = sheet.getRow(0);
            Cell cell = row.getCell(0);

            Assert.assertEquals("Saved data should match", "Test Data for Save", cell.getStringCellValue());
            System.out.println("[TEST] Verified saved data: " + cell.getStringCellValue());

        } finally {
            sessionValue.getSession().close();
        }
    }

    private void save_as_verify(Path tmpDir, String originalExt, String newExt) throws Exception {
        Path originalPath = tmpDir.resolve("SaveAsOriginal" + originalExt);
        Path newPath = tmpDir.resolve("SaveAsNew" + newExt);

        System.out.println("[TEST] Creating workbook at: " + originalPath);

        // Create workbook
        CreateWorkbook createCmd = new CreateWorkbook();
        SessionValue sessionValue = createCmd.action(originalPath.toString(), "OriginalSheet");

        WorkbookSession wbSession = (WorkbookSession) sessionValue.getSession();

        try {
            // Add data
            Workbook wb = wbSession.getWorkbook();
            Sheet sheet = wb.getSheetAt(0);
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue("Test Data for SaveAs");

            System.out.println("[TEST] Added test data");

            // Save to original location first
            SaveWorkbook saveCmd = new SaveWorkbook();
            saveCmd.action(wbSession);

            System.out.println("[TEST] Initial save completed");

            // SaveAs to new location
            SaveWorkbookAs saveAsCmd = new SaveWorkbookAs();
            saveAsCmd.action(newPath.toString(), true, wbSession);

            System.out.println("[TEST] SaveWorkbookAs executed to: " + newPath);

            // Verify both files exist
            Assert.assertTrue("Original file should exist", new File(originalPath.toString()).exists());
            Assert.assertTrue("New file should exist", new File(newPath.toString()).exists());

            // Verify session switched to new path
            Assert.assertEquals("Session path should be updated", newPath.toAbsolutePath().toString(),
                    new File(wbSession.getFilePath()).getAbsolutePath());

            System.out.println("[TEST] Verified session switched to new path");

        } finally {
            wbSession.close();
        }

        // Reopen new file and verify data
        OpenWorkbook openCmd = new OpenWorkbook();
        sessionValue = openCmd.action(newPath.toString(), null, false);

        try {
            WorkbookSession reopenedSession = (WorkbookSession) sessionValue.getSession();
            Workbook wb = reopenedSession.getWorkbook();
            Sheet sheet = wb.getSheetAt(0);
            Row row = sheet.getRow(0);
            Cell cell = row.getCell(0);

            Assert.assertEquals("SavedAs data should match", "Test Data for SaveAs", cell.getStringCellValue());
            System.out.println("[TEST] Verified SaveAs data persisted correctly");

        } finally {
            sessionValue.getSession().close();
        }
    }

    private void save_as_different_path_verify(Path tmpDir, String extension) throws Exception {
        Path originalPath = tmpDir.resolve("DifferentPathOriginal" + extension);
        Path subDir = tmpDir.resolve("subdir");
        Files.createDirectories(subDir);
        Path newPath = subDir.resolve("DifferentPathNew" + extension);

        System.out.println("[TEST] Testing SaveAs to different directory");

        // Create workbook
        CreateWorkbook createCmd = new CreateWorkbook();
        SessionValue sessionValue = createCmd.action(originalPath.toString(), "TestSheet");

        WorkbookSession wbSession = (WorkbookSession) sessionValue.getSession();

        try {
            // Add data
            Workbook wb = wbSession.getWorkbook();
            Sheet sheet = wb.getSheetAt(0);
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue("Different Path Test");

            // Save to original
            SaveWorkbook saveCmd = new SaveWorkbook();
            saveCmd.action(wbSession);

            // SaveAs to different directory
            SaveWorkbookAs saveAsCmd = new SaveWorkbookAs();
            saveAsCmd.action(newPath.toString(), true,wbSession);

            Assert.assertTrue("New path file should exist", new File(newPath.toString()).exists());
            System.out.println("[TEST] SaveAs to different directory successful");

        } finally {
            wbSession.close();
        }
    }

    private void save_as_overwrite_verify(Path tmpDir, String extension) throws Exception {
        Path originalPath = tmpDir.resolve("OverwriteOriginal" + extension);
        Path targetPath = tmpDir.resolve("OverwriteTarget" + extension);

        System.out.println("[TEST] Testing SaveAs overwrite behavior");

        // Create target file first
        CreateWorkbook createCmd = new CreateWorkbook();
        SessionValue targetSession = createCmd.action(targetPath.toString(), "TargetSheet");
        WorkbookSession targetWb = (WorkbookSession) targetSession.getSession();
        SaveWorkbook saveCmd = new SaveWorkbook();
        saveCmd.action(targetWb);
        targetWb.close();

        System.out.println("[TEST] Created target file");

        // Create original file
        SessionValue originalSession = createCmd.action(originalPath.toString(), "OriginalSheet");
        WorkbookSession originalWb = (WorkbookSession) originalSession.getSession();

        try {
            // Try SaveAs with overwrite=false (should fail)
            SaveWorkbookAs saveAsCmd = new SaveWorkbookAs();

            boolean exceptionThrown = false;
            try {
                saveAsCmd.action(targetPath.toString(), false,originalWb);
            } catch (Exception e) {
                exceptionThrown = true;
                System.out.println("[TEST] Expected exception caught: " + e.getMessage());
                Assert.assertTrue("Exception should mention file exists",
                        e.getMessage().contains("E-EXISTS") || e.getMessage().contains("already exists"));
            }

            Assert.assertTrue("Should throw exception when overwrite=false and file exists", exceptionThrown);

            // Try with overwrite=true (should succeed)
            saveAsCmd.action(targetPath.toString(), true,originalWb);
            System.out.println("[TEST] SaveAs with overwrite=true succeeded");

        } finally {
            originalWb.close();
        }
    }
}
