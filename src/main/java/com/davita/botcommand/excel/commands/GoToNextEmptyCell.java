package com.davita.botcommand.excel.commands;

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
        name = "goToNextEmptyCell",
        label = "[[GoToNextEmptyCell.label]]",
        node_label = "[[GoToNextEmptyCell.node_label]]",
        description = "[[GoToNextEmptyCell.description]]",
        icon = "excel-icon.svg"
)
public class GoToNextEmptyCell {

    @Execute
    public void action(
            // 1) Start: Active Cell vs Specific Cell
            @Idx(index = "1", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "1.1", pkg = @Pkg(label = "[[GoToNextEmptyCell.mode.active.label]]", value = "ACTIVE")),
                    @Idx.Option(index = "1.2", pkg = @Pkg(label = "[[GoToNextEmptyCell.mode.specific.label]]", value = "SPECIFIC"))
            })
            @Pkg(
                    label = "[[GoToNextEmptyCell.mode.label]]",
                    description = "[[GoToNextEmptyCell.mode.description]]",
                    default_value = "ACTIVE",
                    default_value_type = DataType.STRING
            )
            @NotEmpty String mode,

            // 1.2.1) Specific starting cell (A1) shown only for SPECIFIC
            @Idx(index = "1.2.1", type = AttributeType.TEXT)
            @Pkg(
                    label = "[[GoToNextEmptyCell.startCell.label]]",
                    description = "[[GoToNextEmptyCell.startCell.description]]"
            )
            @NotEmpty String startCell,

            // 2) Direction
            @Idx(index = "2", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "[[GoToNextEmptyCell.dir.left]]", value = "LEFT")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "[[GoToNextEmptyCell.dir.right]]", value = "RIGHT")),
                    @Idx.Option(index = "2.3", pkg = @Pkg(label = "[[GoToNextEmptyCell.dir.up]]", value = "UP")),
                    @Idx.Option(index = "2.4", pkg = @Pkg(label = "[[GoToNextEmptyCell.dir.down]]", value = "DOWN"))
            })
            @Pkg(
                    label = "[[GoToNextEmptyCell.dir.label]]",
                    description = "[[GoToNextEmptyCell.dir.description]]",
                    default_value = "RIGHT",
                    default_value_type = DataType.STRING
            )
            @NotEmpty String direction,

            // 3) Existing session
            @Idx(index = "3", type = AttributeType.SESSION)
            @Pkg(
                    label = "[[existingSession.label]]",
                    description = "[[existingSession.description]]",
                    default_value = "Default",
                    default_value_type = DataType.SESSION
            )
            @SessionObject
            @NotEmpty WorkbookSession session
    ) {
        // Validate session and workbook
        if (session == null || session.getWorkbook() == null) {
            throw new BotCommandException("Workbook session not initialized. Please open or create a workbook first.");
        }
        Workbook wb = session.getWorkbook();
        int sheetIdx = wb.getActiveSheetIndex();
        if (sheetIdx < 0 || sheetIdx >= wb.getNumberOfSheets()) {
            throw new BotCommandException("No active worksheet found in the workbook.");
        }
        Sheet sheet = wb.getSheetAt(sheetIdx);

        // Determine start address
        CellAddress start;
        if ("ACTIVE".equalsIgnoreCase(mode)) {
            start = sheet.getActiveCell();
            if (start == null) {
                start = new CellAddress(0, 0); // default to A1 if none is set
            }
        } else if ("SPECIFIC".equalsIgnoreCase(mode)) {
            if (startCell == null || startCell.trim().isEmpty()) {
                throw new BotCommandException("Starting cell address must not be empty when 'Specific Cell' is selected.");
            }
            String ref = startCell.trim();
            if (ref.contains(":")) {
                throw new BotCommandException("A cell range is not supported. Provide a single A1 address.");
            }
            try {
                start = new CellAddress(ref);
            } catch (IllegalArgumentException e) {
                throw new BotCommandException("Invalid A1 cell address: " + ref);
            }
        } else {
            throw new BotCommandException("Invalid mode. Choose either 'Active Cell' or 'Specific Cell'.");
        }

        // Bounds
        int maxRow = wb.getSpreadsheetVersion().getLastRowIndex();
        int maxCol = wb.getSpreadsheetVersion().getLastColumnIndex();

        // Step from the cell next to the start in the chosen direction
        int r = start.getRow();
        int c = start.getColumn();
        int dr = 0, dc = 0;
        switch (direction == null ? "" : direction.toUpperCase()) {
            case "LEFT":  dr = 0;  dc = -1; break;
            case "RIGHT": dr = 0;  dc =  1; break;
            case "UP":    dr = -1; dc =  0; break;
            case "DOWN":  dr =  1; dc =  0; break;
            default:
                throw new BotCommandException("Unsupported direction. Choose Left, Right, Up, or Down.");
        }

        r += dr; c += dc;

        // Scan until boundary for the first empty cell
        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        while (r >= 0 && r <= maxRow && c >= 0 && c <= maxCol) {
            if (isEmptyCell(sheet, r, c, formatter, evaluator)) {
                sheet.setActiveCell(new CellAddress(r, c));
                return;
            }
            r += dr;
            c += dc;
        }

        throw new BotCommandException("No empty cell found in the specified direction within the sheet boundaries.");
    }

    private boolean isEmptyCell(Sheet sheet, int rowIdx, int colIdx, DataFormatter formatter, FormulaEvaluator evaluator) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) {
            return true;
        }
        Cell cell = row.getCell(colIdx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) {
            return true;
        }
        if (cell.getCellType() == CellType.BLANK) {
            return true;
        }
        String text = formatter.formatCellValue(cell, evaluator);
        return text == null || text.trim().isEmpty();
    }
}
