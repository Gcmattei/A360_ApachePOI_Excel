package com.automationanywhere.botcommand.excel.commands;

import com.automationanywhere.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.FileExtension;
import com.automationanywhere.commandsdk.annotations.rules.LocalFile;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;

@BotCommand
@CommandPkg(
        name = "saveWorkbookAs",
        label = "[[SaveWorkbookAs.label]]",
        node_label = "[[SaveWorkbookAs.node_label]]",
        description = "[[SaveWorkbookAs.description]]",
        icon = "excel-icon.svg"
)
public class SaveWorkbookAs {

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.FILE)
            @Pkg(label = "[[SaveWorkbookAs.destPath.label]]",
                    description = "[[SaveWorkbookAs.destPath.description]]")
            @NotEmpty @FileExtension(value = "xlsx,xls") @LocalFile String destPath,

            @Idx(index = "2", type = AttributeType.CHECKBOX)
            @Pkg(label = "[[SaveWorkbookAs.overwrite.label]]",
                    description = "[[SaveWorkbookAs.overwrite.description]]",
                    default_value = "false", default_value_type = DataType.BOOLEAN)
            Boolean overwrite,

            @Idx(index = "3", type = AttributeType.SESSION)
            @Pkg(label = "[[existingSession.label]]",
                    description = "[[existingSession.description]]",
                    default_value = "Default",
                    default_value_type = DataType.SESSION)
            @SessionObject
            @NotEmpty WorkbookSession session
    ) {
        try {
            if (session == null || session.getWorkbook() == null) {
                throw new BotCommandException("Workbook session not initialized. Please open or create a workbook first.");
            }
            if (destPath == null || destPath.trim().isEmpty()) {
                throw new BotCommandException("Destination file path cannot be empty.");
            }

            String ext = getExtension(destPath);
            if (!"xlsx".equalsIgnoreCase(ext) && !"xls".equalsIgnoreCase(ext)) {
                throw new BotCommandException("Unsupported file extension. Use .xlsx or .xls");
            }

            session.saveAs(destPath,overwrite);

        } catch (BotCommandException e) {
            throw e;
        } catch (Exception e) {
            throw new BotCommandException("Failed to save workbook as: " + e.getMessage(), e);
        }
    }

    private String getExtension(String path) {
        int dot = path.lastIndexOf('.');
        if (dot < 0 || dot == path.length() - 1) return "";
        return path.substring(dot + 1);
    }
}
