package com.davita.botcommand.excel.commands.sessionManagement;

import com.automationanywhere.commandsdk.annotations.rules.CredentialAllowPassword;
import com.automationanywhere.core.security.SecureString;
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
import java.util.function.Consumer;

@BotCommand
@CommandPkg(
        documentation_url = "https://confluence.davita.com/spaces/PA/pages/993465649/Excel+Package+Documentation#ExcelPackageDocumentation-Open",
        group_label = "[[Group.sessionManagement.label]]",
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

            @Idx(index = "2", type = AttributeType.CREDENTIAL)
            @Pkg(label = "[[OpenWorkbook.credential.label]]", description = "[[OpenWorkbook.credential.description]]", default_value_type = DataType.BOOLEAN, default_value = "False")
            @CredentialAllowPassword SecureString credential,

            @Idx(index = "3", type = AttributeType.CHECKBOX)
            @Pkg(label = "[[OpenWorkbook.readOnly.label]]", description = "[[OpenWorkbook.readOnly.description]]", default_value_type = DataType.BOOLEAN, default_value = "False")
            @NotEmpty Boolean readOnly
    ) throws IOException {
        if (filePath == null || filePath.trim().isEmpty()) {
            throw new BotCommandException("File path cannot be empty.");
        }

        if (readOnly == null) {
            readOnly = false;
        }

        WorkbookSession session;
        if (credential == null) {
            session = WorkbookSession.openWorkbook(filePath,readOnly);
        } else {
            String password = credential.getInsecureString();
            session = WorkbookSession.openWorkbook(filePath,password,readOnly);
        }

        return SessionValue.builder()
                .withSessionObject(session)
                .build();
    }
}
