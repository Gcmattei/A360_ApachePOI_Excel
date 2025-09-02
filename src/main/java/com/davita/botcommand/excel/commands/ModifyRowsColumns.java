package com.davita.botcommand.excel.commands;

import com.davita.botcommand.excel.internal.SheetUtility;
import com.davita.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import org.apache.poi.ss.usermodel.*;

@BotCommand
@CommandPkg(
        name = "modifyRowsColumns",
        label = "[[ModifyRowsColumns.label]]",
        node_label = "[[ModifyRowsColumns.node_label]]",
        description = "[[ModifyRowsColumns.description]]",
        icon = "excel-icon.svg"
)
public class ModifyRowsColumns {

    @Execute
    public void action(
            // RADIO: choose operation group (Row operations vs Column operations)
            @Idx(index = "1", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "1.1", pkg = @Pkg(label = "[[ModifyRowsColumns.group.rows.label]]", value = "ROWS")),
                    @Idx.Option(index = "1.2", pkg = @Pkg(label = "[[ModifyRowsColumns.group.cols.label]]", value = "COLUMNS"))
            })
            @Pkg(label = "[[ModifyRowsColumns.group.label]]",
                    description = "[[ModifyRowsColumns.group.description]]",
                    default_value = "ROWS", default_value_type = DataType.STRING)
            @NotEmpty String group,

            // Child RADIO under Rows: Insert or Delete
            @Idx(index = "1.1.1", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "1.1.1.1", pkg = @Pkg(label = "[[ModifyRowsColumns.rows.insert.label]]", value = "INSERT")),
                    @Idx.Option(index = "1.1.1.2", pkg = @Pkg(label = "[[ModifyRowsColumns.rows.delete.label]]", value = "DELETE"))
            })
            @Pkg(label = "[[ModifyRowsColumns.rows.operation.label]]",
                    description = "[[ModifyRowsColumns.rows.operation.description]]",
                    default_value = "INSERT", default_value_type = DataType.STRING)
            String rowOperation,

            // Rows target: single row "2" or range "1:4" (1-based)
            @Idx(index = "1.1.1.1.1", type = AttributeType.TEXT)
            @Pkg(label = "[[ModifyRowsColumns.rows.target.label]]",
                    description = "[[ModifyRowsColumns.rows.target.description]]")
            @NotEmpty String rowsTargetInsert,

            // Rows target: single row "2" or range "1:4" (1-based)
            @Idx(index = "1.1.1.2.1", type = AttributeType.TEXT)
            @Pkg(label = "[[ModifyRowsColumns.rows.target.label]]",
                    description = "[[ModifyRowsColumns.rows.target.description]]")
            @NotEmpty String rowsTargetDelete,

            // Child RADIO under Columns: Insert or Delete
            @Idx(index = "1.2.1", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "1.2.1.1", pkg = @Pkg(label = "[[ModifyRowsColumns.cols.insert.label]]", value = "INSERT")),
                    @Idx.Option(index = "1.2.1.2", pkg = @Pkg(label = "[[ModifyRowsColumns.cols.delete.label]]", value = "DELETE"))
            })
            @Pkg(label = "[[ModifyRowsColumns.cols.operation.label]]",
                    description = "[[ModifyRowsColumns.cols.operation.description]]",
                    default_value = "INSERT", default_value_type = DataType.STRING)
            String colOperation,

            // Columns target: single col "B" or range "B:D"
            @Idx(index = "1.2.1.1.1", type = AttributeType.TEXT)
            @Pkg(label = "[[ModifyRowsColumns.cols.target.label]]",
                    description = "[[ModifyRowsColumns.cols.target.description]]")
            @NotEmpty String colsTargetInsert,

            // Columns target: single col "B" or range "B:D"
            @Idx(index = "1.2.1.2.1", type = AttributeType.TEXT)
            @Pkg(label = "[[ModifyRowsColumns.cols.target.label]]",
                    description = "[[ModifyRowsColumns.cols.target.description]]")
            @NotEmpty String colsTargetDelete,

            // Session
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

        int activeIndex = wb.getActiveSheetIndex();
        if (activeIndex < 0 || activeIndex >= wb.getNumberOfSheets()) {
            throw new BotCommandException("Active sheet is not set or out of range.");
        }
        Sheet sheet = wb.getSheetAt(activeIndex);

        try {
            if ("ROWS".equals(group)) {
                String rowsTarget = null;
                if (rowsTargetInsert != null) {
                    rowsTarget = rowsTargetInsert;
                } else if (rowsTargetDelete != null){
                    rowsTarget = rowsTargetDelete;
                }
                if (rowsTarget == null || rowsTarget.trim().isEmpty()) {
                    throw new BotCommandException("Rows target is required (e.g., 2 or 1:4).");
                }
                Range r = parseRowRange(rowsTarget.trim());
                if ("INSERT".equals(rowOperation)) {
                    insertRows(sheet, r.first, r.last);
                } else if ("DELETE".equals(rowOperation)) {
                    SheetUtility.deleteRows(sheet, r.first, r.last);
                } else {
                    throw new BotCommandException("Invalid row operation.");
                }
            } else if ("COLUMNS".equals(group)) {
                String colsTarget = null;
                if (colsTargetInsert != null) {
                    colsTarget = colsTargetInsert;
                } else if (colsTargetDelete != null){
                    colsTarget = colsTargetDelete;
                }
                if (colsTarget == null || colsTarget.trim().isEmpty()) {
                    throw new BotCommandException("Columns target is required (e.g., B or B:D).");
                }
                Range c = parseColumnRange(colsTarget.trim());
                if ("INSERT".equals(colOperation)) {
                    insertColumns(sheet, c.first, c.last);
                } else if ("DELETE".equals(colOperation)) {
                    SheetUtility.deleteColumns(sheet,c.first,c.last);
                } else {
                    throw new BotCommandException("Invalid column operation.");
                }
            } else {
                throw new BotCommandException("Invalid group selection.");
            }
        } catch (IllegalArgumentException iae) {
            throw new BotCommandException(iae.getMessage(), iae);
        } catch (Exception e) {
            throw new BotCommandException("Failed to modify rows/columns: " + e.getMessage(), e);
        }
    }

    // Helpers

    // 1-based input for rows: "2" or "1:4"
    private Range parseRowRange(String text) {
        if (text.contains(":")) {
            String[] parts = text.split(":");
            if (parts.length != 2) throw new IllegalArgumentException("Invalid row range: " + text);
            int first = parsePositiveInt(parts[0].trim(), "row");
            int last = parsePositiveInt(parts[1].trim(), "row");
            if (last < first) throw new IllegalArgumentException("Row range end must be >= start.");
            return new Range(first - 1, last - 1);
        } else {
            int idx = parsePositiveInt(text, "row");
            return new Range(idx - 1, idx - 1);
        }
    }

    // Column letters: "B" or "B:D"
    private Range parseColumnRange(String text) {
        if (text.contains(":")) {
            String[] parts = text.split(":");
            if (parts.length != 2) throw new IllegalArgumentException("Invalid column range: " + text);
            int first = colLettersToIndex(parts[0].trim());
            int last = colLettersToIndex(parts[1].trim());
            if (last < first) throw new IllegalArgumentException("Column range end must be >= start.");
            return new Range(first, last);
        } else {
            int idx = colLettersToIndex(text);
            return new Range(idx, idx);
        }
    }

    private int parsePositiveInt(String s, String what) {
        int v = Integer.parseInt(s);
        if (v < 1) throw new IllegalArgumentException("The " + what + " must be >= 1.");
        return v;
    }

    private int colLettersToIndex(String letters) {
        if (letters == null || letters.isEmpty()) throw new IllegalArgumentException("Column letters cannot be empty.");
        String up = letters.toUpperCase();
        int col = 0;
        for (int i = 0; i < up.length(); i++) {
            char ch = up.charAt(i);
            if (ch < 'A' || ch > 'Z') throw new IllegalArgumentException("Invalid column letters: " + letters);
            col = col * 26 + (ch - 'A' + 1);
        }
        return col - 1; // 0-based
    }

    // Insert rows: shift down from first to end by count; create blanks
    private void insertRows(Sheet sheet, int firstRow0, int lastRow0) {
        int count = (lastRow0 - firstRow0 + 1);
        int last = sheet.getLastRowNum();
        if (last >= firstRow0) {
            sheet.shiftRows(firstRow0, last, count, true, true); // shift down [web:21][web:158]
        }
        // Create empty rows for the inserted area
        for (int r = firstRow0; r < firstRow0 + count; r++) {
            if (sheet.getRow(r) == null) sheet.createRow(r);
        }
    }

    private void insertColumns(Sheet sheet, int firstCol0, int lastCol0) {
        int count = lastCol0 - firstCol0 + 1;
        if (count <= 0) return;

        int sheetLastCol = Math.max(getSheetLastUsedColumn(sheet), lastCol0);
        if (sheetLastCol >= firstCol0) {
            // Shift only the used tail to the right
            sheet.shiftColumns(firstCol0, sheetLastCol, count); // [11][2]
        }
        // No need to pre-create cells; Excel will show blanks. Create only if the caller requires initialized cells.
    }

    // Compute the true last used column across all rows to bound shiftColumns work
    private int getSheetLastUsedColumn(Sheet sheet) {
        int lastCol = -1;
        int firstRow = Math.max(sheet.getFirstRowNum(), 0);
        int lastRow = sheet.getLastRowNum();
        for (int r = firstRow; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            short lc = row.getLastCellNum(); // 1-based, -1 if none
            if (lc > 0) lastCol = Math.max(lastCol, lc - 1);
        }
        return lastCol; // -1 if sheet has no cells
    }

    // Clear a specific column interval inclusive, for all rows
    private void clearColumnRange(Sheet sheet, int firstCol0, int lastCol0) {
        if (lastCol0 < firstCol0) return;
        int firstRow = Math.max(sheet.getFirstRowNum(), 0);
        int lastRow = sheet.getLastRowNum();
        for (int r = firstRow; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = firstCol0; c <= lastCol0; c++) {
                Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell != null) row.removeCell(cell);
            }
        }
    }

    // Simple pair for 0-based ranges
    private static class Range { final int first; final int last; Range(int f, int l){ this.first=f; this.last=l; } }
}
