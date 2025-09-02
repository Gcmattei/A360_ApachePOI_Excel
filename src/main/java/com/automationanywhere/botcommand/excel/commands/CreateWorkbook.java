package com.automationanywhere.botcommand.excel.commands;

import com.automationanywhere.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.data.impl.SessionValue;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.FileExtension;
import com.automationanywhere.commandsdk.annotations.rules.LocalFile;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.automationanywhere.commandsdk.model.ReturnSettingsType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;

@BotCommand
@CommandPkg(
        name = "createWorkbook",
        label = "[[CreateWorkbook.label]]",
        node_label = "[[CreateWorkbook.node_label]]",
        description = "[[CreateWorkbook.description]]",
        icon = "excel-icon.svg",
        return_label = "[[createSession.label]]",
        return_settings = {ReturnSettingsType.SESSION_TARGET},
        return_type = DataType.SESSION,
        default_session_value="Default",
        return_required = true
)
public class CreateWorkbook {

    @Execute
    public SessionValue action(
            @Idx(index = "1", type = AttributeType.FILE)
            @Pkg(label = "[[CreateWorkbook.filePath.label]]", description = "[[CreateWorkbook.filePath.description]]")
            @NotEmpty @FileExtension(value = "xlsx,xls") @LocalFile String filePath,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "[[CreateWorkbook.sheetName.label]]", description = "[[CreateWorkbook.sheetName.description]]")
            String sheetName
    ) {
        if (sheetName == null || sheetName.isEmpty()) {
            sheetName = "Sheet1";
        }
        if (filePath == null || filePath.trim().isEmpty()) {
            throw new BotCommandException("File path cannot be empty.");
        }

        String fileExtension = getFileExtension(filePath).toLowerCase();

        Workbook workbook;
        switch (fileExtension) {
            case "xlsx":
                workbook = new XSSFWorkbook();
                break;
            case "xls":
                workbook = new HSSFWorkbook();
                break;
            default:
                throw new BotCommandException("Unsupported file extension: " + fileExtension);
        }

        workbook.createSheet(sheetName);

        File file = new File(filePath);
        File parentDir = file.getParentFile();
        if (parentDir != null && !parentDir.exists() && !parentDir.mkdirs()) {
            throw new BotCommandException(String.format("Failed to create directories for path: %s", parentDir.getAbsolutePath()));
        }

        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
        } catch (Exception e) {
            throw new BotCommandException("Failed to write workbook file: " + e.getMessage(), e);
        }

        WorkbookSession session = new WorkbookSession();
        session.setWorkbook(workbook);
        session.setFilePath(filePath);
        session.setReadOnly(false); // new workbooks default to RW
        try {
            session.acquireLock();
        } catch (Exception ioe) {
            throw new BotCommandException("Failed to lock workbook file for session: " + ioe.getMessage(), ioe);
        }

        return SessionValue.builder()
                .withSessionObject(session)
                .build();
    }

    private String getFileExtension(String fileName) {
        int dotIdx = fileName.lastIndexOf('.');
        if (dotIdx == -1 || dotIdx == fileName.length() - 1) {
            return "";
        }
        return fileName.substring(dotIdx + 1);
    }
}
