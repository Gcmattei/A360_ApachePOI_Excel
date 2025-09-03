package com.davita.botcommand.excel.commands;

import com.automationanywhere.botcommand.data.impl.StringValue;
import com.davita.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.CommandPkg;
import com.automationanywhere.commandsdk.annotations.Execute;
import com.automationanywhere.commandsdk.annotations.Idx;
import com.automationanywhere.commandsdk.annotations.Pkg;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

@BotCommand
@CommandPkg(
        name = "getCurrentWorksheetName",
        label = "[[GetCurrentWorksheetName.label]]",
        node_label = "[[GetCurrentWorksheetName.node_label]]",
        description = "[[GetCurrentWorksheetName.description]]",
        icon = "excel-icon.svg",
        return_type = DataType.STRING,
        return_label = "[[GetCurrentWorksheetName.return.label]]",
        return_required = true
)
public class GetCurrentWorksheetName {

    @Execute
    public StringValue action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "[[existingSession.label]]", description = "[[existingSession.description]]", default_value = "Default", default_value_type = DataType.SESSION)
            @SessionObject
            @NotEmpty  WorkbookSession session
    ) {
        if (session == null || session.getWorkbook() == null) {
            throw new BotCommandException("Workbook session not initialized. Please open or create a workbook first.");
        }
        Workbook wb = session.getWorkbook();
        int activeSheetIdx = wb.getActiveSheetIndex();
        if (activeSheetIdx < 0 || activeSheetIdx >= wb.getNumberOfSheets()) {
            throw new BotCommandException("No active worksheet found in the workbook.");
        }

        // Either approach is fine; both are supported by the SS usermodel:
        // Option A: via Sheet
        Sheet sheet = wb.getSheetAt(activeSheetIdx);
        return new StringValue(sheet.getSheetName());
        // Option B: directly from Workbook:
        // return wb.getSheetName(activeSheetIdx);
    }
}
