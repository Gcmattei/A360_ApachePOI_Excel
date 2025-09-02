package com.davita.botcommand.excel.commands;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.NumberValue;
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
        name = "getRowCount",
        label = "[[GetRowCount.label]]",
        node_label = "[[GetRowCount.node_label]]",
        description = "[[GetRowCount.description]]",
        icon = "excel-icon.svg",
        return_label = "[[GetRowCount.return_label]]",
        return_type = DataType.NUMBER,
        return_required = true
)
public class GetRowCount {

    @Execute
    public NumberValue action(
            // RADIO: Worksheet selection by Name or Index
            @Idx(index = "1", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "1.1", pkg = @Pkg(label = "[[GetRowCount.sheetSelect.byIndex.label]]", value = "ByIndex")),
                    @Idx.Option(index = "1.2", pkg = @Pkg(label = "[[GetRowCount.sheetSelect.byName.label]]", value = "ByName"))
            })
            @Pkg(label = "[[GetRowCount.sheetSelect.label]]",
                    description = "[[GetRowCount.sheetSelect.description]]",
                    default_value = "ByIndex", default_value_type = DataType.STRING)
            @NotEmpty String sheetSelect,

            // Dependent input when selecting by Index (1-based)
            @Idx(index = "1.1.1", type = AttributeType.NUMBER)
            @Pkg(label = "[[GetRowCount.sheetIndex.label]]", description = "[[GetRowCount.sheetIndex.description]]",default_value = "1",default_value_type = DataType.NUMBER)
            @NotEmpty Double sheetIndexOneBased,

            // Dependent input when selecting by Name
            @Idx(index = "1.2.1", type = AttributeType.TEXT)
            @Pkg(label = "[[GetRowCount.sheetName.label]]", description = "[[GetRowCount.sheetName.description]]")
            @NotEmpty String sheetName,

            // RADIO: Count mode
            @Idx(index = "2", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "[[GetRowCount.countMode.nonEmpty.label]]", value = "NonEmpty")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "[[GetRowCount.countMode.total.label]]", value = "Total"))
            })
            @Pkg(label = "[[GetRowCount.countMode.label]]",
                    description = "[[GetRowCount.countMode.description]]",
                    default_value = "NonEmpty", default_value_type = DataType.STRING)
            @NotEmpty String countMode,

            // Session input
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

        // Resolve sheet
        Sheet sheet;
        if ("ByName".equals(sheetSelect)) {
            if (sheetName == null || sheetName.trim().isEmpty()) {
                throw new BotCommandException("Worksheet name cannot be empty when selecting by name.");
            }
            sheet = wb.getSheet(sheetName.trim());
            if (sheet == null) {
                throw new BotCommandException("Worksheet not found: " + sheetName);
            }
        } else if ("ByIndex".equals(sheetSelect)) {
            if (sheetIndexOneBased == null) {
                throw new BotCommandException("Worksheet index is required when selecting by index.");
            }
            if (sheetIndexOneBased < 1) {
                throw new BotCommandException("Worksheet index must be 1 or greater.");
            }
            int zeroBased = sheetIndexOneBased.intValue() - 1;
            if (zeroBased < 0 || zeroBased >= wb.getNumberOfSheets()) {
                throw new BotCommandException("Worksheet index out of range. Total sheets: " + wb.getNumberOfSheets());
            }
            sheet = wb.getSheetAt(zeroBased);
        } else {
            throw new BotCommandException("Invalid worksheet selection mode.");
        }

        // Count logic
        int result;
        if ("NonEmpty".equals(countMode)) {
            // Physical rows: counts initialized rows (data or formatting)
            result = sheet.getPhysicalNumberOfRows();
        } else if ("Total".equals(countMode)) {
            // Total rows with data based on lastRowNum + 1 when there is at least one row
            int last = sheet.getLastRowNum(); // 0-based
            if (last >= 0) {
                result = last + 1;
            } else {
                result = 0;
            }
        } else {
            throw new BotCommandException("Invalid count mode.");
        }

        return new NumberValue(result);
    }
}
