package com.davita.botcommand.excel.commands;

import com.davita.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;

@BotCommand
@CommandPkg(
        name = "getNextEmptyCell",
        label = "[[FindNextEmptyCell.label]]",
        node_label = "[[FindNextEmptyCell.node_label]]",
        description = "[[FindNextEmptyCell.description]]",
        icon = "excel-icon.svg",
        return_type = DataType.STRING,
        return_label = "[[FindNextEmptyCell.return_label]]",
        return_required = true
)
public class FindNextEmptyCell {

    // Radio choices
    private static final String BY_ROW = "BY_ROW";
    private static final String BY_COLUMN = "BY_COLUMN";
    private static final String START_FROM_ACTIVE = "ACTIVE";
    private static final String START_FROM_SPECIFIC = "SPECIFIC";

    @Execute
    public Value<String> action(
            // Traverse direction (radio)
            @Idx(index = "1", type = AttributeType.RADIO,
                    options = {
                    @Idx.Option(index = "1.1", pkg = @Pkg(value = BY_ROW, label = "[[FindNextEmptyCell.traverseMode.byRow]]")),
                    @Idx.Option(index = "1.2", pkg = @Pkg(value = BY_COLUMN, label = "[[FindNextEmptyCell.traverseMode.byColumn]]"))
                    }
            )
            @Pkg(label = "[[FindNextEmptyCell.traverseMode.label]]",
                    description = "[[FindNextEmptyCell.traverseMode.description]]",

                    default_value = BY_ROW,
                    default_value_type = DataType.STRING)
            @NotEmpty String traverseMode,

            // Start mode (radio)
            @Idx(index = "2", type = AttributeType.RADIO,
                    options = {
                            @Idx.Option(index = "2.1", pkg = @Pkg(value = START_FROM_ACTIVE, label = "[[FindNextEmptyCell.startMode.active]]")),
                            @Idx.Option(index = "2.2", pkg = @Pkg(value = START_FROM_SPECIFIC, label = "[[FindNextEmptyCell.startMode.specific]]"))
                    }
            )
            @Pkg(label = "[[FindNextEmptyCell.startMode.label]]",
                    description = "[[FindNextEmptyCell.startMode.description]]",
                    default_value = START_FROM_ACTIVE,
                    default_value_type = DataType.STRING)
            @NotEmpty String startMode,

            // Specific start cell (only used when startMode == SPECIFIC). Expect A1-style like "B3"
            @Idx(index = "2.2.1", type = AttributeType.TEXT)
            @Pkg(label = "[[FindNextEmptyCell.startCell.label]]",
                    description = "[[FindNextEmptyCell.startCell.description]]")
            @NotEmpty String startCell,

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
        int activeIndex = wb.getActiveSheetIndex();
        if (activeIndex < 0 || activeIndex >= wb.getNumberOfSheets()) {
            throw new BotCommandException("Active sheet is not set or out of range.");
        }
        Sheet sheet = wb.getSheetAt(activeIndex);

        if (sheet == null) {
            throw new BotCommandException("Active Worksheet not found.");
        }

        // Determine starting row/col (0-based)
        int startRow;
        int startCol;

        if (START_FROM_SPECIFIC.equals(startMode)) {
            if (startCell == null || startCell.trim().isEmpty()) {
                throw new BotCommandException("Specific start cell is required when 'Start from specific cell' is selected.");
            }
            int[] rc = a1ToRowCol(startCell.trim());
            startRow = rc[0];
            startCol = rc[1];
        } else {
            // Active
            startRow = sheet.getActiveCell().getRow();
            startCol = sheet.getActiveCell().getColumn();
        }

        if (startRow < 0 || startCol < 0) {
            throw new BotCommandException("Start coordinates cannot be negative.");
        }

        CellAddress nextEmpty = findNextEmptyCell(sheet, startRow, startCol, BY_COLUMN.equals(traverseMode));
        if (nextEmpty == null) {
            throw new BotCommandException("No empty cell found from the specified starting point in the selected direction.");
        }

        // Return address in A1 notation
        String a1 = toA1(nextEmpty.getRow(), nextEmpty.getColumn());
        return new StringValue(a1);
    }

    // Find next empty cell scanning from (row, col) inclusive moving either across row or down column.
    // If byColumn==false: move right within the same row until a blank is found.
    // If byColumn==true: move down within the same column until a blank is found.
    private CellAddress findNextEmptyCell(Sheet sheet, int row, int col, boolean byColumn) {
        if (!byColumn) {
            // Traverse by row → move right on same row
            Row r = sheet.getRow(row);
            // Determine an upper bound: use lastCellNum if row exists, else 0; keep probing beyond lastCellNum until first null/blank
            int c = Math.max(col, 0);
            while (c <= Short.MAX_VALUE) { // practical guard
                Cell cell = (r == null) ? null : r.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell == null || cell.getCellType() == CellType.BLANK) {
                    return new CellAddress(row, c);
                }
                c++;
                // Update row reference if it was null at start
                if (r == null) r = sheet.getRow(row);
            }
        } else {
            // Traverse by column → move down on same column
            int rIndex = Math.max(row, 0);
            while (rIndex <= sheet.getLastRowNum() + 10000) { // soft guard to allow empty rows past lastRowNum
                Row r = sheet.getRow(rIndex);
                Cell cell = (r == null) ? null : r.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell == null || cell.getCellType() == CellType.BLANK) {
                    return new CellAddress(rIndex, col);
                }
                rIndex++;
            }
        }
        return null;
    }

    // Parse an A1 string like "B3" to 0-based row/col
    private int[] a1ToRowCol(String a1) {
        // Split letters and digits
        int i = 0;
        while (i < a1.length() && Character.isLetter(a1.charAt(i))) i++;
        if (i == 0 || i == a1.length()) {
            throw new BotCommandException("Invalid A1 address: " + a1);
        }
        String colPart = a1.substring(0, i).toUpperCase();
        String rowPart = a1.substring(i);
        int col = colLettersToIndex(colPart);
        int row1 = Integer.parseInt(rowPart);
        if (row1 < 1) throw new BotCommandException("Row index in A1 must be 1 or greater: " + a1);
        return new int[]{row1 - 1, col};
    }

    private int colLettersToIndex(String letters) {
        int col = 0;
        for (int k = 0; k < letters.length(); k++) {
            char ch = letters.charAt(k);
            if (ch < 'A' || ch > 'Z') throw new BotCommandException("Invalid column letters in A1: " + letters);
            col = col * 26 + (ch - 'A' + 1);
        }
        return col - 1; // zero-based
    }

    private String toA1(int row, int col) {
        return columnIndexToLetters(col) + (row + 1);
    }

    private String columnIndexToLetters(int index) {
        StringBuilder sb = new StringBuilder();
        int n = index + 1;
        while (n > 0) {
            int rem = (n - 1) % 26;
            sb.append((char) ('A' + rem));
            n = (n - 1) / 26;
        }
        return sb.reverse().toString();
    }
}
