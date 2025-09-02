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
        name = "deleteWorksheet",
        label = "[[DeleteWorksheet.label]]",
        node_label = "[[DeleteWorksheet.node_label]]",
        description = "[[DeleteWorksheet.description]]",
        icon = "excel-icon.svg"
)
public class DeleteWorksheet {

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "1.1", pkg = @Pkg(label = "[[DeleteWorksheet.sheetOption.index.label]]", value = "index")),
                    @Idx.Option(index = "1.2", pkg = @Pkg(label = "[[DeleteWorksheet.sheetOption.name.label]]", value = "name"))})
            @Pkg(label = "[[DeleteWorksheet.sheetOption.label]]", description = "[[DeleteWorksheet.sheetOption.description]]", default_value = "index", default_value_type = DataType.STRING)
            @NotEmpty String sheetOption,

            @Idx(index = "1.1.1", type = AttributeType.NUMBER)
            @Pkg(label = "[[DeleteWorksheet.sheetIndex.label]]",
                    description = "[[DeleteWorksheet.sheetIndex.description]]",
                    default_value = "1",default_value_type = DataType.NUMBER)
            @NotEmpty Double sheetIndex,

            @Idx(index = "1.2.1", type = AttributeType.TEXT)
            @Pkg(label = "[[DeleteWorksheet.sheetName.label]]",
                    description = "[[DeleteWorksheet.sheetName.description]]")
            @NotEmpty String sheetName,

            @Idx(index = "2", type = AttributeType.SESSION)
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

        int index = -1;

        if (sheetOption.equals("name")){
            if (sheetName == null || sheetName.trim().isEmpty()) {
                throw new BotCommandException("Sheet name cannot be empty.");
            }

            Sheet target = wb.getSheet(sheetName);
            if (target == null) {
                throw new BotCommandException("Worksheet not found: " + sheetName);
            }

            index = wb.getSheetIndex(target);
        } else {
            if ((sheetIndex % 1) != 0) {
                throw new BotCommandException("Sheet index must be an integer.");
            }

            index = sheetIndex.intValue();
        }

        if (index < 0) {
            throw new BotCommandException("Failed to resolve worksheet index for: " + sheetName);
        }

        // Prevent deleting the last remaining sheet (Excel requires at least one sheet)
        if (wb.getNumberOfSheets() <= 1) {
            throw new BotCommandException("Cannot delete the only remaining worksheet in the workbook.");
        }

        // Delete the sheet in-memory only. Do NOT save to disk here.
        wb.removeSheetAt(index);

        // No file saving here by design. CloseWorkbook will decide whether to persist changes.
    }
}
