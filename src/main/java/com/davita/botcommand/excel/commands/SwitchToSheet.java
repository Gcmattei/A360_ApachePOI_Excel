package com.davita.botcommand.excel.commands;

import com.davita.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

@BotCommand
@CommandPkg(
        name = "switchToSheet",
        label = "[[SwitchToSheet.label]]",
        node_label = "[[SwitchToSheet.node_label]]",
        description = "[[SwitchToSheet.description]]",
        icon = "excel-icon.svg"
)
public class SwitchToSheet {

    private static final String BY_NAME = "BY_NAME";
    private static final String BY_INDEX = "BY_INDEX";

    @Execute
    public void action(
            // Selection mode
            @Idx(index = "1", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "1.1", pkg = @Pkg(value = BY_INDEX, label = "[[SwitchToSheet.mode.byIndex]]")),
                    @Idx.Option(index = "1.2", pkg = @Pkg(value = BY_NAME,  label = "[[SwitchToSheet.mode.byName]]"))
            })
            @Pkg(label = "[[SwitchToSheet.mode.label]]",
                    description = "[[SwitchToSheet.mode.description]]",
                    default_value = BY_NAME, default_value_type = DataType.STRING)
            @NotEmpty String mode,

            // Sheet index (1-based)
            @Idx(index = "1.1.1", type = AttributeType.NUMBER)
            @Pkg(label = "[[SwitchToSheet.sheetIndex.label]]",
                    description = "[[SwitchToSheet.sheetIndex.description]]",
                    default_value = "1", default_value_type = DataType.NUMBER)
            @NotEmpty Double sheetIndexOneBased,

            // Sheet name
            @Idx(index = "1.2.1", type = AttributeType.TEXT)
            @Pkg(label = "[[SwitchToSheet.sheetName.label]]",
                    description = "[[SwitchToSheet.sheetName.description]]")
            @NotEmpty String sheetName,

            // Session
            @Idx(index = "2", type = AttributeType.SESSION)
            @Pkg(label = "[[existingSession.label]]",
                    description = "[[existingSession.description]]",
                    default_value = "Default", default_value_type = DataType.SESSION)
            @SessionObject @NotEmpty WorkbookSession session
    ) {
        if (session == null || session.getWorkbook() == null) {
            throw new BotCommandException("Workbook session not initialized. Please open or create a workbook first.");
        }

        Workbook wb = session.getWorkbook();
        int idx;

        if (BY_NAME.equals(mode)) {
            if (sheetName == null || sheetName.trim().isEmpty()) {
                throw new BotCommandException("Sheet name cannot be empty when selecting by name.");
            }
            Sheet sheet = wb.getSheet(sheetName.trim());
            if (sheet == null) {
                throw new BotCommandException("Worksheet not found: " + sheetName);
            }
            idx = wb.getSheetIndex(sheet);
            if (idx < 0) {
                throw new BotCommandException("Unable to determine index for sheet: " + sheetName);
            }
        } else if (BY_INDEX.equals(mode)) {
            if (sheetIndexOneBased == null) {
                throw new BotCommandException("Sheet index is required when selecting by index.");
            }
            if (sheetIndexOneBased < 1) {
                throw new BotCommandException("Sheet index must be 1 or greater.");
            }
            idx = sheetIndexOneBased.intValue() - 1;
            if (idx < 0 || idx >= wb.getNumberOfSheets()) {
                throw new BotCommandException("Sheet index out of range. Total sheets: " + wb.getNumberOfSheets());
            }
        } else {
            throw new BotCommandException("Invalid selection mode. Choose by name or by index.");
        }

        // Ensure only one sheet is selected and active:
        // 1) Deselect all sheets (clears grouped selection)
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            wb.getSheetAt(i).setSelected(false);
        }
        // 2) Set active tab and selected tab at workbook level
        wb.setActiveSheet(idx);
        wb.setSelectedTab(idx);
        // 3) Mark only the target sheet as selected
        wb.getSheetAt(idx).setSelected(true);

        // No save here; persist via the Close Workbook command
    }
}
