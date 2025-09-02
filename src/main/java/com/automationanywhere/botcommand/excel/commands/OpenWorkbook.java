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
import java.io.FileInputStream;
import java.io.IOException;

@BotCommand
@CommandPkg(
        name = "openWorkbook",
        label = "[[OpenWorkbook.label]]",
        node_label = "[[OpenWorkbook.node_label]]",
        description = "[[OpenWorkbook.description]]",
        icon = "excel-icon.svg",
        return_label = "[[createSession.label]]",
        return_settings = {ReturnSettingsType.SESSION_TARGET},
        return_type = DataType.SESSION,
        default_session_value="Default",
        return_required = true
)
public class OpenWorkbook {

    @Execute
    public SessionValue action(
            @Idx(index = "1", type = AttributeType.FILE)
            @Pkg(label = "[[OpenWorkbook.filePath.label]]", description = "[[OpenWorkbook.filePath.description]]")
            @NotEmpty @FileExtension(value = "xlsx,xls") @LocalFile String filePath,

            @Idx(index = "2", type = AttributeType.CHECKBOX)
            @Pkg(label = "[[OpenWorkbook.readOnly.label]]", description = "[[OpenWorkbook.readOnly.description]]", default_value_type = DataType.BOOLEAN, default_value = "False")
            @NotEmpty Boolean readOnly
    ) {
        if (filePath == null || filePath.trim().isEmpty()) {
            throw new BotCommandException("File path cannot be empty.");
        }

        if (readOnly == null) {
            readOnly = false;
        }

        File file = new File(filePath);

        if (!file.exists() || !file.isFile()) {
            throw new BotCommandException(String.format("File not found: %s", filePath));
        }

        String fileExtension = getFileExtension(filePath).toLowerCase();

        Workbook workbook;

        try (FileInputStream fis = new FileInputStream(file)) {
            switch (fileExtension) {
                case "xlsx":
                    workbook = new XSSFWorkbook(fis);
                    break;
                case "xls":
                    workbook = new HSSFWorkbook(fis);
                    break;
                default:
                    throw new BotCommandException(String.format("Unsupported file extension: %s", fileExtension));
            }
        } catch (IOException e) {
            throw new BotCommandException(String.format("Failed to open workbook: %s", e.getMessage()), e);
        }

        WorkbookSession session = new WorkbookSession();
        session.setWorkbook(workbook);
        session.setFilePath(filePath);
        session.setReadOnly(readOnly);
        try {
            session.acquireLock();
        } catch (IOException ioe) {
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
