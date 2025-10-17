package com.davita.botcommand.excel;

import com.automationanywhere.botcommand.data.impl.SessionValue;
import com.automationanywhere.core.security.SecureString;
import com.davita.botcommand.excel.commands.sessionManagement.OpenWorkbook;
import com.davita.botcommand.excel.commands.worksheetOperations.CreateWorksheet;
import com.davita.botcommand.excel.commands.worksheetOperations.DeleteWorksheet;
import com.davita.botcommand.excel.sessions.WorkbookSession;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Assert;
import org.junit.Test;

import java.io.IOException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;

public class EditProtectedWorkbookTest {

    private static final String BOOK_NAME = "Excel_TestFile_PasswordProtected";
    private static final String NEW_SHEET_NAME = "NewSheet";
    private static final String PASSWORD = "123";

    /**
     * Prints a clickable file path link in IntelliJ console
     */
    private void logClickableFile(String message, Path filePath) {
        // Format compatible with IntelliJ IDEA on Windows
        String clickablePath = filePath.toAbsolutePath().toString().replace("\\", "/");
        System.out.println(message.trim() + " " + clickablePath);
    }

    @Test
    public void edit_protected_workbook_verify() throws Exception {
        System.out.println("[TEST START] Starting protected workbook editing test");
        // Copy resource workbooks to temp paths
        URL urlXLS = getClass().getResource("/excel/Excel_TestFile_PasswordProtected.xls");
        Assert.assertNotNull("Test XLS resource not found", urlXLS);

        URL urlXLSX = getClass().getResource("/excel/Excel_TestFile_PasswordProtected.xlsx");
        Assert.assertNotNull("Test XLSX resource not found", urlXLSX);

        URL urlXLSM = getClass().getResource("/excel/Excel_TestFile_PasswordProtected.xlsm");
        Assert.assertNotNull("Test XLSM resource not found", urlXLSM);

        Path tmpDir = Files.createTempDirectory("edit_protected_book_");

        logClickableFile("[TEST] Created test directory at:",tmpDir);

        System.out.println("[TEST] Testing XLS protected workbook editing");
        edit_workbook_verify(BOOK_NAME + ".xls", tmpDir, urlXLS);

        System.out.println("[TEST] Testing XLSX protected workbook editing");
        edit_workbook_verify(BOOK_NAME + ".xlsx", tmpDir, urlXLSX);

        System.out.println("[TEST] Testing XLSM protected workbook editing");
        edit_workbook_verify(BOOK_NAME + ".xlsm", tmpDir, urlXLSM);

        System.out.println("[TEST END] Finished protected workbook editing test");
    }

    // ========== HELPER METHODS ==========

    private void edit_workbook_verify(String bookName, Path tmpDir, URL url) throws IOException {
        Path work = tmpDir.resolve(bookName);
        Files.copy(url.openStream(), work, java.nio.file.StandardCopyOption.REPLACE_EXISTING);

        logClickableFile("[TEST] Copied test workbook to:",work);

        OpenWorkbook openCmd = new OpenWorkbook();
        CreateWorksheet createSheetCmd = new CreateWorksheet();
        DeleteWorksheet deleteSheetCmd = new DeleteWorksheet();
        SessionValue sessionValue = null;

        // Test READ-WRITE mode
        try {
            logClickableFile("[TEST] Trying to open the test workbook in READ-WRITE mode at:",work);
            SecureString securePassword = new SecureString(PASSWORD.getBytes());
            sessionValue = openCmd.action(work.toString(), securePassword, false);
            System.out.println("[TEST] Opened test successfully");

            Assert.assertNotNull("Session should not be null", sessionValue);
            Assert.assertNotNull("Workbook should not be null", sessionValue.getSession());

            // Verify workbook is actually usable
            WorkbookSession wbSession = (WorkbookSession) sessionValue.getSession();
            Workbook wb = wbSession.getWorkbook();
            Assert.assertTrue("Workbook should have at least one sheet", wb.getNumberOfSheets() > 0);

            System.out.println("[TEST] Workbook has " + wb.getNumberOfSheets() + " sheets");

            // Create new sheet
            System.out.println("[TEST] Trying to create sheet '"+ NEW_SHEET_NAME +"'");
            createSheetCmd.action(NEW_SHEET_NAME,wbSession);
            System.out.println("[TEST] Successfully created sheet '"+ NEW_SHEET_NAME +"'");
            wbSession.save();

            // Delete one sheet
            System.out.println("[TEST] Trying to delete sheet '"+ NEW_SHEET_NAME +"'");
            deleteSheetCmd.action("name",null, NEW_SHEET_NAME,wbSession);
            System.out.println("[TEST] Successfully deleted sheet '"+ NEW_SHEET_NAME +"'");
            wbSession.save();

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
    }
}
