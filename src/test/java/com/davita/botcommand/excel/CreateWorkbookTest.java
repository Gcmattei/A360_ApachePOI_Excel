package com.davita.botcommand.excel;

import com.automationanywhere.botcommand.data.impl.SessionValue;
import com.davita.botcommand.excel.commands.workbookOperations.CreateWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.junit.Assert;
import org.junit.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

public class CreateWorkbookTest {

    private static final String BOOK_NAME = "ExcelTest";
    private static final String SHEET_NAME = "Sheet2";

    /**
     * Prints a clickable file path link in IntelliJ console
     */
    private void logClickableFile(String message, Path filePath) {
        // Format compatible with IntelliJ IDEA on Windows
        String clickablePath = filePath.toAbsolutePath().toString().replace("\\", "/");
        System.out.println(message.trim() + " " + clickablePath);
    }

    @Test
    public void create_workbook_verify() throws Exception {
        Path tmpDir = Files.createTempDirectory("create_book_");

        logClickableFile("[TEST] Created test directory at:",tmpDir);

        System.out.println("[TEST] Testing XLS workbook creation");
        create_workbook_verify(BOOK_NAME+".xls",SHEET_NAME,tmpDir);

        System.out.println("[TEST] Testing XLSX workbook creation");
        create_workbook_verify(BOOK_NAME+".xlsx",SHEET_NAME,tmpDir);
    }
    private void create_workbook_verify (String bookName,String sheetName, Path tmpDir) throws IOException {

        Path work = tmpDir.resolve(bookName);
        logClickableFile("[TEST] Trying to create a test workbook at:",work);
        CreateWorkbook cmd = new CreateWorkbook();
        SessionValue session = cmd.action(work.toString(),sheetName);
        session.getSession().close();
        System.out.println("[TEST] Created test workbook");

        System.out.println("[TEST] Opening test workbook for verification");

        try (Workbook wb = WorkbookFactory.create(work.toFile())) {
            Sheet sh = wb.getSheet(sheetName);
            Assert.assertNotNull("Sheet not found: " + sheetName, sh);
            System.out.println("[TEST] Test workbook verified");
        }
    }
}
