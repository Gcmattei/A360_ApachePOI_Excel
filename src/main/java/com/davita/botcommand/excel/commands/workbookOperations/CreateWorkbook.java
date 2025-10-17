package com.davita.botcommand.excel.commands.workbookOperations;

import com.davita.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.data.impl.SessionValue;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.FileExtension;
import com.automationanywhere.commandsdk.annotations.rules.LocalFile;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.automationanywhere.commandsdk.model.ReturnSettingsType;

import java.io.IOException;

@BotCommand
@CommandPkg(
        documentation_url = "https://confluence.davita.com/spaces/PA/pages/993465649/Excel+Package+Documentation#ExcelPackageDocumentation-Createworkbook",
        group_label = "[[Group.workbookOperations.label]]",
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
    ) throws IOException {
        if (sheetName == null || sheetName.isEmpty()) {
            sheetName = "Sheet1";
        }
        if (filePath == null || filePath.trim().isEmpty()) {
            throw new BotCommandException("File path cannot be empty.");
        }

        WorkbookSession session = WorkbookSession.createWorkbook(filePath,sheetName);

        return SessionValue.builder()
                .withSessionObject(session)
                .build();
    }
}
