package com.davita.botcommand.excel.commands;

import com.davita.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;

// Autofill helpers (internal POI formula API)
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.formula.FormulaShifter;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;

import java.text.Format;

@BotCommand
@CommandPkg(
        name = "setCellFormula",
        label = "[[SetCell.label]]",
        node_label = "[[SetCell.node_label]]",
        description = "[[SetCell.description]]",
        icon = "excel-icon.svg"
)
public class SetCell {

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "1.1", pkg = @Pkg(label = "[[SetCell.target.active.label]]", value = "ACTIVE")),
                    @Idx.Option(index = "1.2", pkg = @Pkg(label = "[[SetCell.target.specific.label]]", value = "SPECIFIC"))
            })
            @Pkg(label = "[[SetCell.target.label]]",
                    description = "[[SetCell.target.description]]",
                    default_value = "ACTIVE", default_value_type = DataType.STRING)
            @NotEmpty String targetMode,

            @Idx(index = "1.2.1", type = AttributeType.TEXT)
            @Pkg(label = "[[SetCell.cellOrRange.label]]",
                    description = "[[SetCell.cellOrRange.description]]")
            String cellOrRangeA1,

            // Input can be "=A1*2" (formula) or "42" / "hello" (raw value)
            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "[[SetCell.formula.label]]",
                    description = "[[SetCell.formula.description]]")
            @NotEmpty String input,

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
        int activeSheetIdx = wb.getActiveSheetIndex();
        if (activeSheetIdx < 0 || activeSheetIdx >= wb.getNumberOfSheets()) {
            throw new BotCommandException("Active sheet is not set or out of range.");
        }
        Sheet sheet = wb.getSheetAt(activeSheetIdx);

        // Resolve target region (single cell or range)
        CellRangeAddress region;
        if ("ACTIVE".equalsIgnoreCase(targetMode)) {
            region = getActiveSelectionOrCell(sheet);
        } else if ("SPECIFIC".equalsIgnoreCase(targetMode)) {
            if (cellOrRangeA1 == null || cellOrRangeA1.trim().isEmpty()) {
                throw new BotCommandException("Cell or range (A1) is required when using Specific.");
            }
            region = parseCellOrRange(cellOrRangeA1.trim());
        } else {
            throw new BotCommandException("Invalid target mode. Choose Active or Specific.");
        }

        final boolean isFormula = input.trim().startsWith("="); // leading '=' means formula
        final String payload = isFormula ? input.trim().substring(1) : input; // POI setCellFormula expects no '='

        // Anchor cell for autofill
        final int baseRow = region.getFirstRow();
        final int baseCol = region.getFirstColumn();

        try {
            for (int r = region.getFirstRow(); r <= region.getLastRow(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) row = sheet.createRow(r);
                for (int c = region.getFirstColumn(); c <= region.getLastColumn(); c++) {
                    Cell cell = row.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                    cell.setBlank();

                    if (isFormula) {
                        // Autofill: shift references by (r-baseRow, c-baseCol)
                        String fForCell = translateFormulaForCell(payload, wb, sheet, baseRow, baseCol, r, c);
                        cell.setCellFormula(fForCell);
                    } else {
                        // Raw value: try boolean/number, else write as string
                        if (isBoolean(payload)) {
                            cell.setCellValue(Boolean.parseBoolean(payload));
                        } else if (isNumeric(payload)) {
                            cell.setCellValue(Double.parseDouble(payload));
                        } else {
                            cell.setCellValue(payload);
                        }
                    }
                }
            }
        } catch (Exception e) {
            throw new BotCommandException("Failed to write value/formula: " + e.getMessage(), e);
        }

        // Optionally recompute cached results later:
        // FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        // evaluator.evaluateAll();
    }

    // Accepts "A1" or "A1:C10"
    private CellRangeAddress parseCellOrRange(String a1) {
        if (a1.contains(":")) {
            try {
                return CellRangeAddress.valueOf(a1);
            } catch (IllegalArgumentException iae) {
                throw new BotCommandException("Invalid A1 range: " + a1, iae);
            }
        } else {
            CellReference ref = new CellReference(a1);
            int r = ref.getRow();
            int c = ref.getCol();
            return new CellRangeAddress(r, r, c, c);
        }
    }

    // Prefer active selection; fallback to active cell (or A1)
    private CellRangeAddress getActiveSelectionOrCell(Sheet sheet) {
        if (sheet instanceof XSSFSheet) {
            try {
                XSSFSheet xs = (XSSFSheet) sheet;
                if (xs.getCTWorksheet().isSetSheetViews()
                        && xs.getCTWorksheet().getSheetViews().sizeOfSheetViewArray() > 0
                        && xs.getCTWorksheet().getSheetViews().getSheetViewArray(0).sizeOfSelectionArray() > 0) {

                    Object sel = xs.getCTWorksheet().getSheetViews().getSheetViewArray(0).getSelectionArray(0);
                    String sqrefStr = null;
                    try {
                        Object raw = sel.getClass().getMethod("getSqref").invoke(sel);
                        if (raw instanceof java.util.List) {
                            java.util.List<?> list = (java.util.List<?>) raw;
                            if (!list.isEmpty() && list.get(0) != null) sqrefStr = list.get(0).toString();
                        } else if (raw != null) {
                            sqrefStr = raw.toString();
                        }
                    } catch (ReflectiveOperationException ignore) {}

                    if (sqrefStr != null && !sqrefStr.trim().isEmpty()) {
                        String firstArea = sqrefStr.trim().split("\\s+")[0];
                        return CellRangeAddress.valueOf(firstArea);
                    }
                }
            } catch (Throwable ignored) {}
        }

        CellAddress addr = sheet.getActiveCell();
        if (addr == null) {
            return new CellRangeAddress(0, 0, 0, 0);
        }
        return new CellRangeAddress(addr.getRow(), addr.getRow(), addr.getColumn(), addr.getColumn());
    }

    // -------- Helpers --------

    private static boolean isBoolean(String s) {
        return "true".equalsIgnoreCase(s) || "false".equalsIgnoreCase(s);
    }

    private static boolean isNumeric(String s) {
        try {
            Double.parseDouble(s);
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    // Translate a base formula string to the target cell by applying row/col copy shifts
    private static String translateFormulaForCell(
            String baseFormula, Workbook wb, Sheet sheet,
            int baseRow, int baseCol, int targetRow, int targetCol
    ) {
        final int sheetIndex = wb.getSheetIndex(sheet);
        final SpreadsheetVersion ver = (wb instanceof XSSFWorkbook)
                ? SpreadsheetVersion.EXCEL2007 : SpreadsheetVersion.EXCEL97;

        // Parse formula into tokens bound to workbook
        Ptg[] ptgs;
        if (wb instanceof XSSFWorkbook) {
            ptgs = FormulaParser.parse(baseFormula,
                    XSSFEvaluationWorkbook.create((XSSFWorkbook) wb),
                    FormulaType.CELL, sheetIndex, baseRow);
        } else {
            ptgs = FormulaParser.parse(baseFormula,
                    HSSFEvaluationWorkbook.create((HSSFWorkbook) wb),
                    FormulaType.CELL, sheetIndex, baseRow);
        }

        // Apply row and column copy shifts from anchor (baseRow/baseCol) to target (targetRow/targetCol)
        int dRow = targetRow - baseRow;
        int dCol = targetCol - baseCol;

        if (dRow != 0) {
            FormulaShifter rowCopy = FormulaShifter.createForRowCopy(
                    sheetIndex, sheet.getSheetName(), baseRow, baseRow, dRow, ver);
            rowCopy.adjustFormula(ptgs, sheetIndex);
        }

        if (dCol != 0) {
            FormulaShifter colCopy = FormulaShifter.createForColumnCopy(
                    sheetIndex, sheet.getSheetName(), baseCol, baseCol, dCol, ver);
            colCopy.adjustFormula(ptgs, sheetIndex);
        }

        // Render back to Excel text
        if (wb instanceof XSSFWorkbook) {
            return FormulaRenderer.toFormulaString(
                    XSSFEvaluationWorkbook.create((XSSFWorkbook) wb), ptgs);
        } else {
            return FormulaRenderer.toFormulaString(
                    HSSFEvaluationWorkbook.create((HSSFWorkbook) wb), ptgs);
        }
    }
}
