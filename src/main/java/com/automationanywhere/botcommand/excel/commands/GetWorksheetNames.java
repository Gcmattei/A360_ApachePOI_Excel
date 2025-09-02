package com.automationanywhere.botcommand.excel.commands;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.ListValue;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;
import java.util.StringJoiner;

@BotCommand
@CommandPkg(
        name = "getWorksheetNames",
        label = "[[GetWorksheetNames.label]]",
        node_label = "[[GetWorksheetNames.node_label]]",
        description = "[[GetWorksheetNames.description]]",
        icon = "excel-icon.svg",
        return_label = "[[GetWorksheetNames.return_label]]",
        return_required = true,
        return_type = DataType.LIST,
        return_sub_type = DataType.STRING
)
public class GetWorksheetNames {

    @Execute
    public ListValue<String> action(
            @Idx(index = "1", type = AttributeType.SESSION)
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
        int n = wb.getNumberOfSheets(); // count sheets [web:182][web:183]
        List<Value> names = new ArrayList<>(Math.max(n, 0)); // A360 list of Value

        for (int i = 0; i < n; i++) {
            String name = wb.getSheetName(i); // sheet name by index [web:182][web:183]
            names.add(new StringValue(name));
        }

        ListValue<String> result = new ListValue<>();
        result.set(names); // Return as A360 ListValue [web:173]
        return result;
    }
}
