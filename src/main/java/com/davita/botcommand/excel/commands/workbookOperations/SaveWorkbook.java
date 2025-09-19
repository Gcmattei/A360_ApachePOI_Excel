package com.davita.botcommand.excel.commands.workbookOperations;

import com.davita.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;

@BotCommand
@CommandPkg(
        documentation_url = "https://github.com/Gcmattei/A360_ApachePOI_Excel/blob/main/docs/A360-Excel-Comprehensive-Docs.md#save-workbook",
        group_label = "[[Group.workbookOperations.label]]",
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
}
