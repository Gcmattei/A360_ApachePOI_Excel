package com.automationanywhere.botcommand.excel.commands;

import com.automationanywhere.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import org.apache.poi.ss.usermodel.Workbook;

@BotCommand
@CommandPkg(
        name = "renameWorksheet",
        label = "[[RenameWorksheet.label]]",
        node_label = "[[RenameWorksheet.node_label]]",
        description = "[[RenameWorksheet.description]]",
        icon = "excel-icon.svg"
)
public class RenameWorksheet {

    @Execute
    public void action(
            // Parent RADIO: select how to identify the original sheet
            @Idx(index = "1", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "1.1", pkg = @Pkg(label = "[[RenameWorksheet.select.byIndex.label]]", value = "ByIndex")),
                    @Idx.Option(index = "1.2", pkg = @Pkg(label = "[[RenameWorksheet.select.byName.label]]", value = "ByName"))
            })
            @Pkg(label = "[[RenameWorksheet.select.label]]",
                    description = "[[RenameWorksheet.select.description]]",
                    default_value = "ByIndex", default_value_type = DataType.STRING)
            @NotEmpty String selectMode,

            // Child of ByIndex: original sheet index (1-based)
            @Idx(index = "1.1.1", type = AttributeType.NUMBER)
            @Pkg(label = "[[RenameWorksheet.originalIndex.label]]",
                    description = "[[RenameWorksheet.originalIndex.description]]")
            @NotEmpty Double originalIndexOneBased,

            // Child of ByName: original sheet name
            @Idx(index = "1.2.1", type = AttributeType.TEXT)
            @Pkg(label = "[[RenameWorksheet.originalName.label]]",
                    description = "[[RenameWorksheet.originalName.description]]")
            @NotEmpty String originalName,

            // New name
            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "[[RenameWorksheet.newName.label]]",
                    description = "[[RenameWorksheet.newName.description]]")
            @NotEmpty String newName,

            // Session
            @Idx(index = "3", type = AttributeType.SESSION)
            @Pkg(label = "[[existingSession.label]]",
                    description = "[[existingSession.description]]",
                    default_value = "Default",
                    default_value_type = DataType.SESSION)
            @SessionObject
            @NotEmpty WorkbookSession session
    ) {
        if (session == null || session.getWorkbook() == null) {
            throw new BotCommandException("Workbook session not initialized. Please open or create a workbook first.");
        }

        Workbook wb = session.getWorkbook();

        // Validate new name against Excel rules
        String trimmed = newName == null ? null : newName.trim();
        validateSheetName(trimmed);
        if (sheetExists(wb, trimmed)) {
            throw new BotCommandException("A worksheet with the new name already exists: " + trimmed);
        }

        // Resolve source sheet index
        int srcIndex;
        if ("ByName".equals(selectMode)) {
            if (originalName == null || originalName.trim().isEmpty()) {
                throw new BotCommandException("Original worksheet name cannot be empty.");
            }
            srcIndex = wb.getSheetIndex(originalName.trim());
            if (srcIndex < 0) {
                throw new BotCommandException("Original worksheet not found: " + originalName);
            }
        } else if ("ByIndex".equals(selectMode)) {
            if (originalIndexOneBased == null) {
                throw new BotCommandException("Original worksheet index is required.");
            }
            if (originalIndexOneBased < 1) {
                throw new BotCommandException("Worksheet index must be 1 or greater.");
            }
            srcIndex = originalIndexOneBased.intValue() - 1;
            if (srcIndex < 0 || srcIndex >= wb.getNumberOfSheets()) {
                throw new BotCommandException("Worksheet index out of range. Total sheets: " + wb.getNumberOfSheets());
            }
        } else {
            throw new BotCommandException("Invalid selection mode. Choose By name or By index.");
        }

        // Perform rename
        try {
            wb.setSheetName(srcIndex, trimmed);
        } catch (IllegalArgumentException iae) {
            // POI throws IllegalArgumentException for invalid names/duplicates as well
            throw new BotCommandException("Failed to rename worksheet: " + iae.getMessage(), iae);
        }
        // Do not save here; persistence is managed by your save/close command.
    }

    // Excel sheet name rules similar to UI:
    // - Not blank; length <= 31
    // - No: : \ / ? * [ ]
    // - Cannot start/end with apostrophe
    // - Cannot be "History"
    private void validateSheetName(String name) {
        if (name == null || name.isEmpty()) {
            throw new BotCommandException("New worksheet name cannot be empty.");
        }
        if (name.length() > 31) {
            throw new BotCommandException("New worksheet name cannot exceed 31 characters.");
        }
        if (name.startsWith("'") || name.endsWith("'")) {
            throw new BotCommandException("New worksheet name cannot begin or end with an apostrophe (').");
        }
        String illegal = "\\/:?*[]";
        for (int i = 0; i < illegal.length(); i++) {
            char ch = illegal.charAt(i);
            if (name.indexOf(ch) >= 0) {
                throw new BotCommandException("New worksheet name cannot contain: / \\ ? * : [ ]");
            }
        }
        if ("History".equalsIgnoreCase(name)) {
            throw new BotCommandException("New worksheet name cannot be 'History'.");
        }
    }

    private boolean sheetExists(Workbook wb, String target) {
        return wb.getSheetIndex(target) >= 0;
    }
}
