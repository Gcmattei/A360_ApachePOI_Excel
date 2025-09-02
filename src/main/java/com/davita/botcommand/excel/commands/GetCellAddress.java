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
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;

@BotCommand
@CommandPkg(
        name = "getCellAddress",
        label = "[[GetCellAddress.label]]",
        node_label = "[[GetCellAddress.node_label]]",
        description = "[[GetCellAddress.description]]",
        icon = "excel-icon.svg",
        return_type = DataType.STRING,
        return_label = "[[GetCellAddress.return.label]]",
        return_required = true
)
public class GetCellAddress {

    @Execute
    public StringValue action(
            // 1) Mode: Active vs Specific Cell
            @Idx(index = "1", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "1.1", pkg = @Pkg(label = "[[GetCellAddress.mode.active.label]]", value = "ACTIVE")),
                    @Idx.Option(index = "1.2", pkg = @Pkg(label = "[[GetCellAddress.mode.specific.label]]", value = "SPECIFIC"))
            })
            @Pkg(
                    label = "[[GetCellAddress.mode.label]]",
                    description = "[[GetCellAddress.mode.description]]",
                    default_value = "ACTIVE",
                    default_value_type = DataType.STRING
            )
            @NotEmpty String mode,

            // 1.2.1) Column title (Specific)
            @Idx(index = "1.2.1", type = AttributeType.TEXT)
            @Pkg(label = "[[GetCellAddress.columnTitle.label]]", description = "[[GetCellAddress.columnTitle.description]]")
            @NotEmpty String columnTitle,

            // 1.2.2) Position relative from the title (1-based, first data row under header is 1)
            @Idx(index = "1.2.2", type = AttributeType.NUMBER)
            @Pkg(label = "[[GetCellAddress.position.label]]", description = "[[GetCellAddress.position.description]]", default_value = "1", default_value_type = DataType.NUMBER)
            @NotEmpty Double position,

            // 2) Existing session
            @Idx(index = "2", type = AttributeType.SESSION)
            @Pkg(label = "[[existingSession.label]]", description = "[[existingSession.description]]", default_value = "Default", default_value_type = DataType.SESSION)
            @SessionObject
            @NotEmpty WorkbookSession session
    ) {
        if (session == null || session.getWorkbook() == null) {
            throw new BotCommandException("Workbook session not initialized. Please open or create a workbook first.");
        }
        Workbook wb = session.getWorkbook();
        int activeSheetIndex = wb.getActiveSheetIndex();
        if (activeSheetIndex < 0 || activeSheetIndex >= wb.getNumberOfSheets()) {
            throw new BotCommandException("No active worksheet found in the workbook.");
        }
        Sheet sheet = wb.getSheetAt(activeSheetIndex);

        if ("ACTIVE".equalsIgnoreCase(mode)) {
            CellAddress addr = sheet.getActiveCell();
            if (addr == null) {
                addr = new CellAddress(0, 0); // default A1 if not set
            }
            return new StringValue(addr.formatAsString());
        }

        if (!"SPECIFIC".equalsIgnoreCase(mode)) {
            throw new BotCommandException("Invalid mode. Choose either 'Active' or 'Specific'.");
        }

        // Validate Specific inputs
        if (columnTitle == null || columnTitle.trim().isEmpty()) {
            throw new BotCommandException("Column title must not be empty for 'Specific Cell'.");
        }
        if (position == null || (position % 1) != 0) {
            throw new BotCommandException("Position must be an integer for 'Specific Cell'.");
        }
        int offset = position.intValue();
        if (offset <= 0) {
            throw new BotCommandException("Position must be a positive integer (1 = first row under the header).");
        }

        // Find header row: first non-empty row
        int headerRowIdx = findFirstNonEmptyRow(sheet);
        if (headerRowIdx < 0) {
            throw new BotCommandException("Header row not found. The sheet appears to be empty.");
        }
        Row headerRow = sheet.getRow(headerRowIdx);

        // Find column by title (case-insensitive compare of displayed value)
        int colIdx = findColumnIndexByTitle(wb, headerRow, columnTitle.trim());
        if (colIdx < 0) {
            throw new BotCommandException("Column title not found in header row: " + columnTitle);
        }

        // Compute target row index: header + offset
        int targetRow = headerRowIdx + offset;
        int maxRow = wb.getSpreadsheetVersion().getLastRowIndex();
        if (targetRow < 0 || targetRow > maxRow) {
            throw new BotCommandException("Computed row is out of supported range for this workbook format.");
        }
        int maxCol = wb.getSpreadsheetVersion().getLastColumnIndex();
        if (colIdx < 0 || colIdx > maxCol) {
            throw new BotCommandException("Computed column is out of supported range for this workbook format.");
        }
        return new StringValue(new CellAddress(targetRow, colIdx).formatAsString());
    }

    private int findFirstNonEmptyRow(Sheet sheet) {
        int first = Math.max(0, sheet.getFirstRowNum());
        int last = Math.max(first, sheet.getLastRowNum());
        for (int r = first; r <= last; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            short firstCell = row.getFirstCellNum();
            short lastCell = row.getLastCellNum();
            if (firstCell == -1 || lastCell == -1) continue;
            for (int c = firstCell; c < lastCell; c++) {
                Cell cell = row.getCell(c);
                if (cell != null && cell.getCellType() != CellType.BLANK) {
                    return r;
                }
            }
        }
        return -1;
    }

    private int findColumnIndexByTitle(Workbook wb, Row headerRow, String wanted) {
        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();
        short firstCell = headerRow.getFirstCellNum();
        short lastCell = headerRow.getLastCellNum();
        for (int c = firstCell; c < lastCell; c++) {
            Cell cell = headerRow.getCell(c);
            if (cell == null) continue;
            String text = formatter.formatCellValue(cell, eval);
            if (text != null && text.trim().equalsIgnoreCase(wanted)) {
                return c;
            }
        }
        return -1;
    }
}
