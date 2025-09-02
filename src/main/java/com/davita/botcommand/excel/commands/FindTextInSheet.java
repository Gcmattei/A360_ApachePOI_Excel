package com.davita.botcommand.excel.commands;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.ListValue;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.davita.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SelectModes;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

@BotCommand
@CommandPkg(
        name = "findTextInSheet",
        label = "[[FindTextInSheet.label]]",
        node_label = "[[FindTextInSheet.node_label]]",
        description = "[[FindTextInSheet.description]]",
        icon = "excel-icon.svg",
        return_label = "[[FindTextInSheet.return_label]]",
        return_required = true,
        return_type = DataType.LIST,
        return_sub_type = DataType.STRING
)
public class FindTextInSheet {

    @Execute
    public ListValue<String> action(
            // FROM anchor
            @Idx(index = "1", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "1.1", pkg = @Pkg(value = "BEGIN", label = "[[FindTextInSheet.from.begin]]")),
                    @Idx.Option(index = "1.2", pkg = @Pkg(value = "END", label = "[[FindTextInSheet.from.end]]")),
                    @Idx.Option(index = "1.3", pkg = @Pkg(value = "ACTIVE", label = "[[FindTextInSheet.from.active]]")),
                    @Idx.Option(index = "1.4", pkg = @Pkg(value = "SPECIFIC", label = "[[FindTextInSheet.from.specific]]"))
            })
            @Pkg(label = "[[FindTextInSheet.from.label]]",
                    description = "[[FindTextInSheet.from.description]]",
                    default_value = "BEGIN", default_value_type = DataType.STRING)
            @NotEmpty String fromKind,

            @Idx(index = "1.4.1", type = AttributeType.TEXT)
            @Pkg(label = "[[FindTextInSheet.from.cell.label]]",
                    description = "[[FindTextInSheet.from.cell.description]]")
            @NotEmpty String fromCellA1,

            // TO anchor
            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(value = "BEGIN", label = "[[FindTextInSheet.to.begin]]")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(value = "END", label = "[[FindTextInSheet.to.end]]")),
                    @Idx.Option(index = "2.3", pkg = @Pkg(value = "ACTIVE", label = "[[FindTextInSheet.to.active]]")),
                    @Idx.Option(index = "2.4", pkg = @Pkg(value = "SPECIFIC", label = "[[FindTextInSheet.to.specific]]"))
            })
            @Pkg(label = "[[FindTextInSheet.to.label]]",
                    description = "[[FindTextInSheet.to.description]]",
                    default_value = "END", default_value_type = DataType.STRING)
            @NotEmpty String toKind,

            @Idx(index = "2.4.1", type = AttributeType.TEXT)
            @Pkg(label = "[[FindTextInSheet.to.cell.label]]",
                    description = "[[FindTextInSheet.to.cell.description]]")
            @NotEmpty String toCellA1,

            // String to find
            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "[[FindTextInSheet.query.label]]",
                    description = "[[FindTextInSheet.query.description]]")
            String query,

            // Search direction
            @Idx(index = "4", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "4.1", pkg = @Pkg(label = "[[FindTextInSheet.dir.byRow]]", value = "BY_ROW")),
                    @Idx.Option(index = "4.2", pkg = @Pkg(label = "[[FindTextInSheet.dir.byCol]]", value = "BY_COL"))
            })
            @Pkg(label = "[[FindTextInSheet.dir.label]]",
                    description = "[[FindTextInSheet.dir.description]]",
                    default_value = "BY_ROW", default_value_type = DataType.STRING)
            @SelectModes @NotEmpty String direction,

            // New checkboxes
            @Idx(index = "5", type = AttributeType.CHECKBOX)
            @Pkg(label = "[[FindTextInSheet.matchCase.label]]",
                    description = "[[FindTextInSheet.matchCase.description]]",
                    default_value = "false", default_value_type = DataType.BOOLEAN)
            @NotEmpty Boolean matchCase,

            @Idx(index = "6", type = AttributeType.CHECKBOX)
            @Pkg(label = "[[FindTextInSheet.matchWhole.label]]",
                    description = "[[FindTextInSheet.matchWhole.description]]",
                    default_value = "false", default_value_type = DataType.BOOLEAN)
            @NotEmpty Boolean matchWhole,

            // Session
            @Idx(index = "7", type = AttributeType.SESSION)
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
        int activeSheetIdx = wb.getActiveSheetIndex();
        if (activeSheetIdx < 0 || activeSheetIdx >= wb.getNumberOfSheets()) {
            throw new BotCommandException("Active sheet is not set or out of range.");
        }
        Sheet sheet = wb.getSheetAt(activeSheetIdx);

        // Defaults for checkboxes
        boolean mc = Boolean.TRUE.equals(matchCase);
        boolean mw = Boolean.TRUE.equals(matchWhole);

        // Resolve anchors -> coordinates
        CellAddress fromAddr = resolveAnchor(sheet, fromKind, fromCellA1, true);
        CellAddress toAddr = resolveAnchor(sheet, toKind, toCellA1, false);

        int r0 = Math.min(fromAddr.getRow(), toAddr.getRow());
        int r1 = Math.max(fromAddr.getRow(), toAddr.getRow());
        int c0 = Math.min(fromAddr.getColumn(), toAddr.getColumn());
        int c1 = Math.max(fromAddr.getColumn(), toAddr.getColumn());

        // Formatter/evaluator to read displayed values (incl. formula results)
        DataFormatter formatter = new DataFormatter(Locale.getDefault());
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

        String needle = query;
        if (!mc) {
            needle = needle.toLowerCase(Locale.ROOT);
        }

        List<Value> results = new ArrayList<>();

        if ("BY_ROW".equals(direction)) {
            for (int r = r0; r <= r1; r++) {
                Row row = sheet.getRow(r);
                for (int c = c0; c <= c1; c++) {
                    String text = getCellText(row, c, formatter, evaluator);
                    if (matches(text, needle, mc, mw)) {
                        results.add(new StringValue(new CellReference(r, c).formatAsString()));
                    }
                }
            }
        } else if ("BY_COL".equals(direction)) {
            for (int c = c0; c <= c1; c++) {
                for (int r = r0; r <= r1; r++) {
                    Row row = sheet.getRow(r);
                    String text = getCellText(row, c, formatter, evaluator);
                    if (matches(text, needle, mc, mw)) {
                        results.add(new StringValue(new CellReference(r, c).formatAsString()));
                    }
                }
            }
        } else {
            throw new BotCommandException("Invalid search direction.");
        }

        ListValue<String> out = new ListValue<>();
        out.set(results);
        return out;
    }

    private boolean matches(String cellText, String needlePrepared, boolean matchCase, boolean matchWhole) {
        if (cellText == null) return false;
        String hay = matchCase ? cellText : cellText.toLowerCase(Locale.ROOT);
        if (matchWhole) {
            return hay.equals(needlePrepared);
        } else {
            return hay.contains(needlePrepared);
        }
    }

    private CellAddress resolveAnchor(Sheet sheet, String kind, String a1IfSpecific, boolean isFrom) {
        switch (kind) {
            case "ACTIVE":
                CellAddress active = sheet.getActiveCell();
                if (active == null) {
                    throw new BotCommandException("Active cell is not set for the active worksheet.");
                }
                return active;
            case "SPECIFIC":
                if (a1IfSpecific == null || a1IfSpecific.trim().isEmpty()) {
                    throw new BotCommandException((isFrom ? "From" : "To") + " cell is required when using Specific Cell.");
                }
                CellReference ref = new CellReference(a1IfSpecific.trim());
                return new CellAddress(ref.getRow(), ref.getCol());
            case "BEGIN":
                int firstRow = Math.max(sheet.getFirstRowNum(), 0);
                int firstCol = 0;
                Row fr = sheet.getRow(firstRow);
                if (fr != null && fr.getLastCellNum() > 0) {
                    firstCol = Math.max(0, fr.getFirstCellNum());
                }
                return new CellAddress(firstRow, firstCol);
            case "END":
                int lastRow = sheet.getLastRowNum();
                if (lastRow < 0) return new CellAddress(0, 0);
                int lastCol = 0;
                for (int r = lastRow; r >= sheet.getFirstRowNum(); r--) {
                    Row row = sheet.getRow(r);
                    if (row != null && row.getLastCellNum() > 0) {
                        lastCol = row.getLastCellNum() - 1;
                        lastRow = r;
                        break;
                    }
                }
                return new CellAddress(Math.max(0, lastRow), Math.max(0, lastCol));
            default:
                throw new BotCommandException("Invalid anchor selection: " + kind);
        }
    }

    private String getCellText(Row row, int col, DataFormatter formatter, FormulaEvaluator evaluator) {
        if (row == null) return "";
        Cell cell = row.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return "";
        return formatter.formatCellValue(cell, evaluator);
    }
}
