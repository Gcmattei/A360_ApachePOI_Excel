package com.davita.botcommand.excel.commands;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.data.impl.TableValue;
import com.automationanywhere.botcommand.data.model.Schema;
import com.automationanywhere.botcommand.data.model.table.Table;
import com.davita.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;

import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SelectModes;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

@BotCommand
@CommandPkg(
        name = "getWorksheetAsTable",
        label = "[[GetWorksheetAsTable.label]]",
        node_label = "[[GetWorksheetAsTable.node_label]]",
        description = "[[GetWorksheetAsTable.description]]",
        icon = "excel-icon.svg",
        return_label = "[[GetWorksheetAsTable.return_label]]",
        return_type = DataType.TABLE,
        return_required = true
)
public class GetWorksheetAsTable {

    @Execute
    public TableValue action(
            // Sheet selection: Active or Specific
            @Idx(index = "1", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "1.1", pkg = @Pkg(label = "[[GetWorksheetAsTable.sheetOption.active.label]]", value = "Active")),
                    @Idx.Option(index = "1.2", pkg = @Pkg(label = "[[GetWorksheetAsTable.sheetOption.specific.label]]", value = "Specific"))
            })
            @Pkg(label = "[[GetWorksheetAsTable.sheetOption.label]]",
                    description = "[[GetWorksheetAsTable.sheetOption.description]]",
                    default_value = "Active", default_value_type = DataType.STRING)
            @NotEmpty String sheetOption,

            @Idx(index = "1.2.1", type = AttributeType.TEXT)
            @Pkg(label = "[[GetWorksheetAsTable.sheetName.label]]",
                    description = "[[GetWorksheetAsTable.sheetName.description]]")
            String sheetName,

            // Area selection: entire sheet or A1 range
            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "[[GetWorksheetAsTable.areaOption.sheet.label]]", value = "Sheet")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "[[GetWorksheetAsTable.areaOption.range.label]]", value = "Range"))
            })
            @Pkg(label = "[[GetWorksheetAsTable.areaOption.label]]",
                    description = "[[GetWorksheetAsTable.areaOption.description]]",
                    default_value = "Sheet", default_value_type = DataType.STRING)
            @SelectModes @NotEmpty String areaOption,

            @Idx(index = "2.2.1", type = AttributeType.TEXT)
            @Pkg(label = "[[GetWorksheetAsTable.rangeA1.label]]",
                    description = "[[GetWorksheetAsTable.rangeA1.description]]")
            String rangeA1,

            // Rows selection: all or numeric window (1-based)
            @Idx(index = "3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "3.1", pkg = @Pkg(label = "[[GetWorksheetAsTable.rowsOption.all.label]]", value = "All")),
                    @Idx.Option(index = "3.2", pkg = @Pkg(label = "[[GetWorksheetAsTable.rowsOption.range.label]]", value = "Range"))
            })
            @Pkg(label = "[[GetWorksheetAsTable.rowsOption.label]]",
                    description = "[[GetWorksheetAsTable.rowsOption.description]]",
                    default_value = "All", default_value_type = DataType.STRING)
            @SelectModes @NotEmpty String rowsOption,

            @Idx(index = "3.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "[[GetWorksheetAsTable.firstRow.label]]",
                    description = "[[GetWorksheetAsTable.firstRow.description]]")
            @NotEmpty Double firstRowOneBased,

            @Idx(index = "3.2.2", type = AttributeType.NUMBER)
            @Pkg(label = "[[GetWorksheetAsTable.lastRow.label]]",
                    description = "[[GetWorksheetAsTable.lastRow.description]]")
            @NotEmpty Double lastRowOneBased,

            // Header and read mode
            @Idx(index = "4", type = AttributeType.CHECKBOX)
            @Pkg(label = "[[GetWorksheetAsTable.hasHeader.label]]",
                    description = "[[GetWorksheetAsTable.hasHeader.description]]",
                    default_value = "true", default_value_type = DataType.BOOLEAN)
            @NotEmpty Boolean hasHeader,

            @Idx(index = "5", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "5.1", pkg = @Pkg(label = "[[GetWorksheetAsTable.readOption.visible.label]]",
                            description = "[[GetWorksheetAsTable.readOption.visible.description]]",
                            value = "visible")),
                    @Idx.Option(index = "5.2", pkg = @Pkg(label = "[[GetWorksheetAsTable.readOption.value.label]]",
                            description = "[[GetWorksheetAsTable.readOption.value.description]]",
                            value = "value"))
            })
            @Pkg(label = "[[GetWorksheetAsTable.readOption.label]]",
                    description = "[[GetWorksheetAsTable.readOption.description]]",
                    default_value = "visible", default_value_type = DataType.STRING)
            @NotEmpty String readOption,

            @Idx(index = "5.1.1", type = AttributeType.HELP)
            @Pkg(label = "", description = "[[GetWorksheetAsTable.readOption.visible.description]]",default_value_type = DataType.STRING) String visibleHelp,
            @Idx(index = "5.2.1", type = AttributeType.HELP)
            @Pkg(label = "", description = "[[GetWorksheetAsTable.readOption.value.description]]",default_value_type = DataType.STRING) String valueHelp,

            // Session
            @Idx(index = "6", type = AttributeType.SESSION)
            @Pkg(label = "[[existingSession.label]]",
                    description = "[[existingSession.description]]",
                    default_value = "Default",
                    default_value_type = DataType.SESSION)
            @SessionObject
            @NotEmpty WorkbookSession session
    ) {
        try {
            if (session == null || session.getWorkbook() == null) {
                throw new BotCommandException("Workbook session not initialized. Please open or create a workbook first.");
            }
            if ("Specific".equalsIgnoreCase(sheetOption)) {
                if (sheetName == null || sheetName.trim().isEmpty()) {
                    throw new BotCommandException("Sheet name cannot be empty.");
                }
            }
            if (hasHeader == null) hasHeader = true;

            Workbook wb = session.getWorkbook();
            Sheet sheet;

            if ("Active".equalsIgnoreCase(sheetOption)) {
                int idx = wb.getActiveSheetIndex();
                if (idx < 0 || idx >= wb.getNumberOfSheets()) {
                    throw new BotCommandException("Active sheet is not set or out of range.");
                }
                sheet = wb.getSheetAt(idx);
            } else {
                sheet = wb.getSheet(sheetName);
                if (sheet == null) {
                    throw new BotCommandException("Worksheet not found: " + sheetName);
                }
            }

            // 1) Sheet-wide bounds
            int sheetFirstRow = sheet.getFirstRowNum();
            int sheetLastRow  = sheet.getLastRowNum();

            // 2) Area bounds (Sheet vs A1 Range)
            int firstCol, lastCol, areaFirstRow, areaLastRow;
            if ("Range".equalsIgnoreCase(areaOption)) {
                if (rangeA1 == null || rangeA1.trim().isEmpty()) {
                    throw new BotCommandException("Range (A1) cannot be empty when Area=Range.");
                }
                CellRangeAddress a = CellRangeAddress.valueOf(rangeA1.trim());
                firstCol     = a.getFirstColumn();
                lastCol      = a.getLastColumn();
                areaFirstRow = a.getFirstRow();
                areaLastRow  = a.getLastRow();
            } else {
                int maxCols = calculateMaxColumns(sheet, sheetFirstRow, sheetLastRow);
                firstCol     = 0;
                lastCol      = Math.max(0, maxCols - 1);
                areaFirstRow = sheetFirstRow;
                areaLastRow  = sheetLastRow;
            }

            // 3) Selection bounds
            int firstRowNum = areaFirstRow;
            int lastRowNum  = areaLastRow;
            int columnCount = (lastCol - firstCol + 1);

            Table table = new Table();
            if (lastRowNum < firstRowNum || columnCount <= 0) {
                table.setSchema(new ArrayList<>());
                table.setRows(new ArrayList<>());
                return new TableValue(table);
            }

            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();



            // Headers
            List<Schema> headers = buildHeadersInArea(sheet, hasHeader, firstRowNum, firstCol, columnCount, readOption,formatter,evaluator);
            table.setSchema(headers);

            // 4) Relative data rows window (header-aware)
            int dataStartBase = hasHeader ? firstRowNum + 1 : firstRowNum; // first data row inside selection
            int dataFirst = dataStartBase;
            int dataLast  = lastRowNum;

            if ("Range".equalsIgnoreCase(rowsOption)) {
                // Validate Double inputs and convert to integers (must be whole, 1-based)
                int relFirst = toRowIndex("First row", firstRowOneBased);
                int relLast  = toRowIndex("Last row",  lastRowOneBased);
                if (relLast < relFirst) {
                    throw new BotCommandException("Last row must be greater than or equal to first row.");
                }
                int absFirst = dataStartBase + (relFirst - 1);
                int absLast  = dataStartBase + (relLast  - 1);
                dataFirst = Math.max(dataStartBase, absFirst);
                dataLast  = Math.min(lastRowNum,     absLast);
            }

            if (dataLast < dataFirst) {
                table.setRows(new ArrayList<>());
                return new TableValue(table);
            }

            // 5) Read rows
            List<com.automationanywhere.botcommand.data.model.table.Row> rows = new ArrayList<>();
            for (int r = dataFirst; r <= dataLast; r++) {
                Row poiRow = sheet.getRow(r);
                com.automationanywhere.botcommand.data.model.table.Row record =
                        new com.automationanywhere.botcommand.data.model.table.Row();
                List<Value> values = new ArrayList<>(headers.size());
                for (int c = 0; c < headers.size(); c++) {
                    Cell cell = (poiRow == null) ? null : poiRow.getCell(firstCol + c);

                    String cellVal = null;
                    if ("visible".equalsIgnoreCase(readOption)) {
                        cellVal = readVisible(cell,formatter,evaluator);
                    } else {
                        cellVal = String.valueOf(readRaw(cell));
                    }
                    values.add(new StringValue(cellVal == null ? "" : cellVal));
                }
                record.setValues(values);
                rows.add(record);
            }

            table.setRows(rows);
            return new TableValue(table);

        } catch (Exception e) {
            throw new BotCommandException(e.getMessage());
        }
    }

    // ---------- Helpers ----------

    private int toRowIndex(String name, Double value) {
        if (value == null) {
            throw new BotCommandException(name + " is required when Rows=Range.");
        }
        if (value.isNaN() || value.isInfinite()) {
            throw new BotCommandException(name + " must be a finite number.");
        }
        if (value <= 0d) {
            throw new BotCommandException(name + " must be positive (1-based).");
        }
        double frac = value - Math.floor(value);
        if (frac != 0d) {
            throw new BotCommandException(name + " must be a whole number (no decimals).");
        }
        // safe to cast
        return value.intValue();
    }

    private int calculateMaxColumns(Sheet sheet, int firstRowNum, int lastRowNum) {
        int maxColumns = 0;
        for (int r = firstRowNum; r <= lastRowNum; r++) {
            Row row = sheet.getRow(r);
            if (row != null) {
                short last = row.getLastCellNum();
                maxColumns = Math.max(maxColumns, last > 0 ? last : 0);
            }
        }
        return maxColumns;
    }

    private List<Schema> buildHeadersInArea(Sheet sheet, boolean hasHeader, int headerRowIndex,
                                            int firstCol, int columnCount, String readOption, DataFormatter  formatter, FormulaEvaluator evaluator) {
        List<Schema> headers = new ArrayList<>();
        if (columnCount <= 0) return headers;

        if (hasHeader) {
            Row headerRow = sheet.getRow(headerRowIndex);
            for (int c = 0; c < columnCount; c++) {
                String headerName = null;
                if (headerRow != null) {
                    Cell cell = headerRow.getCell(firstCol + c);
                    String cellValue = null;
                    if ("visible".equalsIgnoreCase(readOption)) {
                        cellValue = readVisible(cell,formatter,evaluator);
                    } else {
                        cellValue = String.valueOf(readRaw(cell));
                    }
                    if (cellValue != null) headerName = cellValue;
                }
                if (headerName == null || headerName.trim().isEmpty()) {
                    headerName = "Column" + (c + 1);
                }
                headers.add(new Schema(headerName, com.automationanywhere.botcore.api.dto.AttributeType.STRING));
            }
        } else {
            for (int c = 0; c < columnCount; c++) {
                headers.add(new Schema("Column" + (c + 1), com.automationanywhere.botcore.api.dto.AttributeType.STRING));
            }
        }
        return headers;
    }

    private Object readCellValue(Cell cell, String readOption) {
        if (cell == null) return null;
        if ("visible".equalsIgnoreCase(readOption)) {
            DataFormatter df = new DataFormatter();
            return df.formatCellValue(cell);
        } else {
            return cell.getNumericCellValue();
        }
    }

    private Object readCachedFormulaValue(Cell cell, String readOption) {
        if ("visible".equalsIgnoreCase(readOption)) {
            DataFormatter formatter = new DataFormatter();
            Workbook workbook = cell.getSheet().getWorkbook();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            return formatter.formatCellValue(cell, evaluator);
        } else {
            return cell.getCellFormula();
        }
    }

    private String asString(Cell cell) {
        if (cell == null) return null;
        try {
            cell.setCellType(CellType.STRING);
            return cell.getStringCellValue();
        } catch (Exception e) {
            return cell.toString();
        }
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