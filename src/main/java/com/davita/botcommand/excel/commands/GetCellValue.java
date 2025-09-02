package com.davita.botcommand.excel.commands;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.davita.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;

@BotCommand
@CommandPkg(
        name = "getCellValue",
        label = "[[GetCellValue.label]]",
        node_label = "[[GetCellValue.node_label]]",
        description = "[[GetCellValue.description]]",
        icon = "excel-icon.svg",
        return_label = "[[GetCellValue.return_label]]",
        return_type = DataType.STRING,
        return_required = true
)
public class GetCellValue {

    @Execute
    public Value<String> action(
            // RADIO: choose Active cell or Specific cell
            @Idx(index = "1", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "1.1", pkg = @Pkg(label = "[[GetCellValue.targetMode.active.label]]", value = "ACTIVE")),
                    @Idx.Option(index = "1.2", pkg = @Pkg(label = "[[GetCellValue.targetMode.specific.label]]", value = "SPECIFIC"))
            })
            @Pkg(label = "[[GetCellValue.targetMode.label]]",
                    description = "[[GetCellValue.targetMode.description]]",
                    default_value = "ACTIVE", default_value_type = DataType.STRING)
            @NotEmpty String targetMode,

            // Child for SPECIFIC: A1 address (must be a single cell; ranges not allowed)
            @Idx(index = "1.2.1", type = AttributeType.TEXT)
            @Pkg(label = "[[GetCellValue.cellA1.label]]",
                    description = "[[GetCellValue.cellA1.description]]")
            @NotEmpty String cellA1,

            // RADIO: read visible text or the underlying cell value
            @Idx(index = "2", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "[[GetCellValue.read.visible.label]]", value = "VISIBLE", description = "[[GetCellValue.read.visible.description]]")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "[[GetCellValue.read.value.label]]", value = "VALUE", description = "[[GetCellValue.read.value.description]]"))
            })
            @Pkg(label = "[[GetCellValue.read.label]]",
                    description = "[[GetCellValue.read.description]]",
                    default_value = "VISIBLE", default_value_type = DataType.STRING)
            @NotEmpty String readMode,


            @Idx(index = "2.1.1", type = AttributeType.HELP)
            @Pkg(label = "", description = "[[GetCellValue.read.visible.description]]",default_value_type = DataType.STRING) String visibleHelp,
            @Idx(index = "2.2.1", type = AttributeType.HELP)
            @Pkg(label = "", description = "[[GetCellValue.read.value.description]]",default_value_type = DataType.STRING) String valueHelp,

            // Session
            @Idx(index = "3", type = AttributeType.SESSION)
            @Pkg(label = "[[existingSession.label]]",
                    description = "[[existingSession.description]]",
                    default_value = "Default",
                    default_value_type = DataType.SESSION)
            @SessionObject
            @NotEmpty WorkbookSession session
    ) {
        // Preconditions
        if (session == null || session.getWorkbook() == null) {
            throw new BotCommandException("Workbook session not initialized. Please open or create a workbook first.");
        }

        Workbook wb = session.getWorkbook();
        int activeSheetIdx = wb.getActiveSheetIndex();
        if (activeSheetIdx < 0 || activeSheetIdx >= wb.getNumberOfSheets()) {
            throw new BotCommandException("Active sheet is not set or out of range.");
        }
        Sheet sheet = wb.getSheetAt(activeSheetIdx);

        // Resolve target cell coordinates
        int rowIdx, colIdx;
        if ("ACTIVE".equalsIgnoreCase(targetMode)) {
            CellAddress addr = sheet.getActiveCell();
            if (addr == null) {
                throw new BotCommandException("Active cell is not set on the active worksheet.");
            }
            rowIdx = addr.getRow();
            colIdx = addr.getColumn();
        } else if ("SPECIFIC".equalsIgnoreCase(targetMode)) {
            if (cellA1 == null || cellA1.trim().isEmpty()) {
                throw new BotCommandException("Cell address (A1) is required when using Specific cell.");
            }
            String a1 = cellA1.trim();
            if (a1.contains(":")) {
                throw new BotCommandException("A range was provided. Please specify a single cell address (e.g., C5).");
            }
            CellReference ref = new CellReference(a1);
            rowIdx = ref.getRow();
            colIdx = ref.getCol();
        } else {
            throw new BotCommandException("Invalid target mode. Choose Active or Specific.");
        }

        // Get cell (blank if missing)
        Row row = sheet.getRow(rowIdx);
        Cell cell = (row == null) ? null : row.getCell(colIdx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

        // Read value
        String result;
        if ("VISIBLE".equalsIgnoreCase(readMode)) {
            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            result = readVisible(cell, formatter,evaluator);
        } else if ("VALUE".equalsIgnoreCase(readMode)) {
            Object raw = readRaw(cell);
            result = raw == null ? "" : String.valueOf(raw);
        } else {
            throw new BotCommandException("Invalid read mode. Choose Visible or Value.");
        }

        return new StringValue(result == null ? "" : result);
    }

    // Visible text using DataFormatter + FormulaEvaluator (mirrors user-visible content)
    private String readVisible(Cell cell, DataFormatter  formatter, FormulaEvaluator evaluator) {
        if (cell == null) return "";
        return formatter.formatCellValue(cell, evaluator);
    }

    // Underlying value (number/boolean/string/formula-as-result or formula text)
    private Object readRaw(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return cell.getNumericCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                // Return cached result rather than the formula text to align with "value" semantics
                switch (cell.getCachedFormulaResultType()) {
                    case STRING:  return cell.getStringCellValue();
                    case NUMERIC: return cell.getNumericCellValue();
                    case BOOLEAN: return cell.getBooleanCellValue();
                    default:      return cell.getCellFormula(); // fallback to formula string
                }
            case BLANK:
                return null;
            default:
                return cell.toString();
        }
    }
}
