package com.automationanywhere.botcommand.excel.commands;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.DateTimeValue;
import com.automationanywhere.botcommand.data.impl.TableValue;
import com.automationanywhere.botcommand.data.model.Schema;
import com.automationanywhere.botcommand.data.model.table.Row;
import com.automationanywhere.botcommand.data.model.table.Table;
import com.automationanywhere.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import org.apache.poi.ss.usermodel.*;

import java.time.LocalDateTime;
import java.time.ZonedDateTime;
import java.util.List;

@BotCommand
@CommandPkg(
        name = "writeTableToWorksheet",
        label = "[[WriteTableToWorksheet.label]]",
        node_label = "[[WriteTableToWorksheet.node_label]]",
        description = "[[WriteTableToWorksheet.description]]",
        icon = "excel-icon.svg"
)
public class WriteTableToWorksheet {

    private static final String TARGET_ACTIVE = "ACTIVE";
    private static final String TARGET_SPECIFIC = "SPECIFIC";

    @Execute
    public void action(
            // Input: Data table to write
            @Idx(index = "1", type = AttributeType.TABLE)
            @Pkg(label = "[[WriteTableToWorksheet.dataTable.label]]",
                    description = "[[WriteTableToWorksheet.dataTable.description]]")
            @NotEmpty Table table,

            // Target sheet selection mode
            @Idx(index = "2", type = AttributeType.RADIO,
                    options = {
                        @Idx.Option(index = "2.1", pkg = @Pkg(value = TARGET_ACTIVE, label = "[[WriteTableToWorksheet.targetSheetMode.active]]")),
                        @Idx.Option(index = "2.2", pkg = @Pkg(value = TARGET_SPECIFIC, label = "[[WriteTableToWorksheet.targetSheetMode.specific]]"))
                    }
            )
            @Pkg(label = "[[WriteTableToWorksheet.targetSheetMode.label]]",
                    description = "[[WriteTableToWorksheet.targetSheetMode.description]]",
                    default_value = TARGET_ACTIVE, default_value_type = DataType.STRING)
            @NotEmpty String targetSheetMode,

            // Specific sheet name (used when mode == SPECIFIC)
            @Idx(index = "2.2.1", type = AttributeType.TEXT)
            @Pkg(label = "[[WriteTableToWorksheet.sheetName.label]]",
                    description = "[[WriteTableToWorksheet.sheetName.description]]")
            @NotEmpty String sheetName,

            // Start cell (A1) for top-left corner
            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "[[WriteTableToWorksheet.startCell.label]]",
                    description = "[[WriteTableToWorksheet.startCell.description]]",
                    default_value = "A1", default_value_type = DataType.STRING)
            @NotEmpty String startCellA1,

            // Retain data types or write as string
            @Idx(index = "4", type = AttributeType.CHECKBOX)
            @Pkg(label = "[[WriteTableToWorksheet.retainTypes.label]]",
                    description = "[[WriteTableToWorksheet.retainTypes.description]]",
                    default_value = "false", default_value_type = DataType.BOOLEAN)
            Boolean retainTypes,

            // Session
            @Idx(index = "5", type = AttributeType.SESSION)
            @Pkg(label = "[[existingSession.label]]",
                    description = "[[existingSession.description]]",
                    default_value = "Default", default_value_type = DataType.SESSION)
            @SessionObject
            @NotEmpty WorkbookSession session
    ) {
        try {
            if (session == null || session.getWorkbook() == null) {
                throw new BotCommandException("Workbook session not initialized. Please open or create a workbook first.");
            }
            if (table == null) {
                throw new BotCommandException("Input data table cannot be null.");
            }
            if (retainTypes == null) {
                retainTypes = false;
            }

            Workbook wb = session.getWorkbook();

            // Resolve target sheet
            Sheet sheet;
            if (TARGET_ACTIVE.equals(targetSheetMode)) {
                int activeIndex = wb.getActiveSheetIndex();
                if (activeIndex < 0 || activeIndex >= wb.getNumberOfSheets()) {
                    throw new BotCommandException("Active sheet is not set or out of range.");
                }
                sheet = wb.getSheetAt(activeIndex);
            } else if (TARGET_SPECIFIC.equals(targetSheetMode)) {
                if (sheetName == null || sheetName.trim().isEmpty()) {
                    throw new BotCommandException("Worksheet name cannot be empty when selecting a specific sheet.");
                }
                sheet = wb.getSheet(sheetName.trim());
                if (sheet == null) {
                    throw new BotCommandException("Worksheet not found: " + sheetName);
                }
            } else {
                throw new BotCommandException("Invalid target sheet mode. Choose active or specific.");
            }

            // Parse start cell A1
            int[] start = a1ToRowCol(startCellA1);
            int startRow = start[0];
            int startCol = start[1];
            if (startRow < 0 || startCol < 0) {
                throw new BotCommandException("Start cell coordinates cannot be negative.");
            }

            // Extract table data
            List<Schema> schema = table.getSchema();
            List<Row> rows = table.getRows();
            int numCols = (schema != null) ? schema.size() : 0;
            if (numCols <= 0) {
                throw new BotCommandException("Input data table has no columns.");
            }

            // Write rows
            for (int r = 0; r < rows.size(); r++) {
                Row sourceRow = rows.get(r);
                org.apache.poi.ss.usermodel.Row targetRow = sheet.getRow(startRow + r);
                if (targetRow == null) targetRow = sheet.createRow(startRow + r);

                List<Value> values = sourceRow.getValues();
                for (int c = 0; c < numCols; c++) {
                    Cell cell = targetRow.getCell(startCol + c);
                    if (cell == null) cell = targetRow.createCell(startCol + c);

                    Value v = (values != null && c < values.size()) ? values.get(c) : null;
                    Object raw = (v == null) ? null : v.get();

                    writeWithType(cell, raw, retainTypes);
                }
            }

            // Do not save here; CloseWorkbook will handle persistence. Respect readOnly by not saving.

            // Optionally set active sheet to the target one (no file save)
            wb.setActiveSheet(wb.getSheetIndex(sheet));

        } catch (BotCommandException e) {
            throw e;
        } catch (Exception e) {
            throw new BotCommandException("Failed to write data table: " + e.getMessage(), e);
        }
    }

    // Helpers

    private void writeWithType(Cell cell, Object raw, boolean retainType) {
        if (!retainType) {
            cell.setCellStyle(null);
        }
        if (raw == null){
            cell.setBlank();;
        } else if (raw instanceof Number) {
            cell.setCellValue(((Number) raw).doubleValue());
        } else if (raw instanceof Boolean) {
            cell.setCellValue((Boolean) raw);
        } else if (raw instanceof ZonedDateTime) {
            cell.setCellValue(((ZonedDateTime) raw).toLocalDateTime());
        } else {
            cell.setCellValue(String.valueOf(raw));
        }
    }

    private int[] a1ToRowCol(String a1) {
        if (a1 == null || a1.trim().isEmpty()) {
            throw new BotCommandException("Start cell (A1) cannot be empty.");
        }
        String s = a1.trim().toUpperCase();
        int i = 0;
        while (i < s.length() && Character.isLetter(s.charAt(i))) i++;
        if (i == 0 || i == s.length()) {
            throw new BotCommandException("Invalid A1 address: " + a1);
        }
        String colPart = s.substring(0, i);
        String rowPart = s.substring(i);
        int col = colLettersToIndex(colPart);
        int row1 = Integer.parseInt(rowPart);
        if (row1 < 1) throw new BotCommandException("Row index must be 1 or greater in A1 address: " + a1);
        return new int[]{row1 - 1, col};
    }

    private int colLettersToIndex(String letters) {
        int col = 0;
        for (int k = 0; k < letters.length(); k++) {
            char ch = letters.charAt(k);
            if (ch < 'A' || ch > 'Z') throw new BotCommandException("Invalid column letters: " + letters);
            col = col * 26 + (ch - 'A' + 1);
        }
        return col - 1; // zero-based
    }
}
