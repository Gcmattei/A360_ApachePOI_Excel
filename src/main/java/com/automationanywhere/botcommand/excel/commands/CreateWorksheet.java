package com.automationanywhere.botcommand.excel.commands;

import com.automationanywhere.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;

@BotCommand
@CommandPkg(
        name = "createWorksheet",
        label = "[[CreateWorksheet.label]]",
        node_label = "[[CreateWorksheet.node_label]]",
        description = "[[CreateWorksheet.description]]",
        icon = "excel-icon.svg"
)
public class CreateWorksheet {

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.TEXT)
            @Pkg(label = "[[CreateWorksheet.sheetName.label]]", description = "[[CreateWorksheet.sheetName.description]]")
            @NotEmpty String sheetName,

            @Idx(index = "2", type = AttributeType.SESSION)
            @Pkg(label = "[[existingSession.label]]", description = "[[existingSession.description]]",
                    default_value = "Default", default_value_type = DataType.SESSION)
            @SessionObject
            @NotEmpty WorkbookSession session
    ) {
        if (session == null || session.getWorkbook() == null) {
            throw new BotCommandException("Workbook session not initialized. Please create a workbook first.");
        }

        try {
            if (session.getWorkbook().getSheet(sheetName) != null) {
                throw new BotCommandException(String.format("Worksheet already exists: %s", sheetName));
            }
            session.getWorkbook().createSheet(sheetName);

        } catch (Exception e) {
            throw new BotCommandException("Failed to create worksheet: " + e.getMessage(), e);
        }
    }
}
