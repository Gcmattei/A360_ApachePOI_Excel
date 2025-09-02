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
        name = "saveWorkbook",
        label = "[[SaveWorkbook.label]]",
        node_label = "[[SaveWorkbook.node_label]]",
        description = "[[SaveWorkbook.description]]",
        icon = "excel-icon.svg"
)
public class SaveWorkbook {

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "[[existingSession.label]]",
                    description = "[[existingSession.description]]",
                    default_value = "Default",
                    default_value_type = DataType.SESSION)
            @SessionObject
            @NotEmpty WorkbookSession session
    ) {
        try {
            session.saveChanges();

        } catch (BotCommandException e) {
            throw e;
        } catch (Exception e) {
            throw new BotCommandException("Failed to save workbook: " + e.getMessage(), e);
        }
    }

    private String getExtension(String path) {
        int dot = path.lastIndexOf('.');
        if (dot < 0 || dot == path.length() - 1) return "";
        return path.substring(dot + 1);
    }
}
