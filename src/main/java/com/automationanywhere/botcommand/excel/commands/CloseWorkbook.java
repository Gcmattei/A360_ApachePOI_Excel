package com.automationanywhere.botcommand.excel.commands;

import com.automationanywhere.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;

import java.io.FileOutputStream;

@BotCommand
@CommandPkg(
        name = "closeWorkbook",
        label = "[[CloseWorkbook.label]]",
        node_label = "[[CloseWorkbook.node_label]]",
        description = "[[CloseWorkbook.description]]",
        icon = "excel-icon.svg"
)
public class CloseWorkbook {

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.CHECKBOX)
            @Pkg(label = "[[CloseWorkbook.saveChanges.label]]",
                    description = "[[CloseWorkbook.saveChanges.description]]",
                    default_value = "false",
                    default_value_type = DataType.BOOLEAN)
            @NotEmpty Boolean saveChanges,

            @Idx(index = "2", type = AttributeType.SESSION)
            @Pkg(label = "[[existingSession.label]]",
                    description = "[[existingSession.description]]",
                    default_value = "Default",
                    default_value_type = DataType.SESSION)
            @SessionObject
            @NotEmpty WorkbookSession session
    ) {
        // Basic null safety
        if (saveChanges == null) {
            saveChanges = false;
        }

        if (session == null || session.getWorkbook() == null) {
            throw new BotCommandException("Workbook session not initialized. Please open or create a workbook first.");
        }

        // Enforce read-only policy when attempting to save
        if (saveChanges && session.isReadOnly()) {
            throw new BotCommandException("Cannot save changes: the workbook is open in read-only mode.");
        }

        try {
            if (saveChanges) {
                // Persist changes to the same path stored in session
                String path = session.getFilePath();
                if (path == null || path.trim().isEmpty()) {
                    throw new BotCommandException("Cannot save changes because the session has no file path.");
                }
                session.saveChanges();
            }

            // Close and mark session as closed
            try {
                session.close();
            } catch (Exception e) {
                throw new BotCommandException("Failed to close the workbook: " + e.getMessage(), e);
            }

        } catch (BotCommandException e) {
            throw e;
        } catch (Exception e) {
            // Catch-all to ensure localized error is surfaced
            throw new BotCommandException("An unexpected error occurred while closing the workbook: " + e.getMessage(), e);
        }
    }
}
