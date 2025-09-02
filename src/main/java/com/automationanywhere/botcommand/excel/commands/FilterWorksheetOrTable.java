package com.automationanywhere.botcommand.excel.commands;

import com.automationanywhere.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SelectModes;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFTable;

import com.automationanywhere.botcommand.excel.internal.TableRange;

import java.util.Locale;

@BotCommand
@CommandPkg(
        name = "filterWorksheetOrTable",
        label = "[[FilterWorksheetOrTable.label]]",
        node_label = "[[FilterWorksheetOrTable.node_label]]",
        description = "[[FilterWorksheetOrTable.description]]",
        icon = "excel-icon.svg"
)
public class FilterWorksheetOrTable {

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "1.1", pkg = @Pkg(label = "[[FilterWorksheetOrTable.mode.table.label]]", value = "TABLE")),
                    @Idx.Option(index = "1.2", pkg = @Pkg(label = "[[FilterWorksheetOrTable.mode.worksheet.label]]", value = "WORKSHEET"))
            })
            @Pkg(label = "[[FilterWorksheetOrTable.mode.label]]",
                    description = "[[FilterWorksheetOrTable.mode.description]]",
                    default_value = "TABLE", default_value_type = DataType.STRING)
            @SelectModes @NotEmpty String mode,

            // Table branch
            @Idx(index = "1.1.1", type = AttributeType.TEXT)
            @Pkg(label = "[[FilterWorksheetOrTable.table.name.label]]",
                    description = "[[FilterWorksheetOrTable.table.name.description]]")
            @NotEmpty String tableName,

            @Idx(index = "1.1.2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "1.1.2.1", pkg = @Pkg(label = "[[FilterWorksheetOrTable.table.colByName.label]]", value = "BY_NAME")),
                    @Idx.Option(index = "1.1.2.2", pkg = @Pkg(label = "[[FilterWorksheetOrTable.table.colByIndex.label]]", value = "BY_INDEX"))
            })
            @Pkg(label = "[[FilterWorksheetOrTable.table.colSelector.label]]",
                    description = "[[FilterWorksheetOrTable.table.colSelector.description]]",
                    default_value = "BY_NAME", default_value_type = DataType.STRING)
            @SelectModes @NotEmpty String tableColSelector,

            @Idx(index = "1.1.2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "[[FilterWorksheetOrTable.table.colName.label]]",
                    description = "[[FilterWorksheetOrTable.table.colName.description]]")
            @NotEmpty String tableColumnName,

            @Idx(index = "1.1.2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "[[FilterWorksheetOrTable.table.colIndex.label]]",
                    description = "[[FilterWorksheetOrTable.table.colIndex.description]]")
            @NotEmpty Double tableColumnIndexOneBased,

            // Worksheet branch
            @Idx(index = "1.2.1", type = AttributeType.TEXT)
            @Pkg(label = "[[FilterWorksheetOrTable.ws.name.label]]",
                    description = "[[FilterWorksheetOrTable.ws.name.description]]")
            @NotEmpty String sheetName,

            @Idx(index = "1.2.2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "1.2.2.1", pkg = @Pkg(label = "[[FilterWorksheetOrTable.ws.rangeAll.label]]", value = "ALL")),
                    @Idx.Option(index = "1.2.2.2", pkg = @Pkg(label = "[[FilterWorksheetOrTable.ws.rangeSpecific.label]]", value = "SPECIFIC"))
            })
            @Pkg(label = "[[FilterWorksheetOrTable.ws.rangeMode.label]]",
                    description = "[[FilterWorksheetOrTable.ws.rangeMode.description]]",
                    default_value = "ALL", default_value_type = DataType.STRING)
            @NotEmpty String wsRangeMode,

            @Idx(index = "1.2.2.2.1", type = AttributeType.TEXT)
            @Pkg(label = "[[FilterWorksheetOrTable.ws.rangeText.label]]",
                    description = "[[FilterWorksheetOrTable.ws.rangeText.description]]")
            @NotEmpty String wsRangeA1,

            @Idx(index = "1.2.3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "1.2.3.1", pkg = @Pkg(label = "[[FilterWorksheetOrTable.ws.colByName.label]]", value = "BY_NAME")),
                    @Idx.Option(index = "1.2.3.2", pkg = @Pkg(label = "[[FilterWorksheetOrTable.ws.colByIndex.label]]", value = "BY_INDEX"))
            })
            @Pkg(label = "[[FilterWorksheetOrTable.ws.colSelector.label]]",
                    description = "[[FilterWorksheetOrTable.ws.colSelector.description]]",
                    default_value = "BY_NAME", default_value_type = DataType.STRING)
            @SelectModes @NotEmpty String wsColSelector,

            @Idx(index = "1.2.3.1.1", type = AttributeType.TEXT)
            @Pkg(label = "[[FilterWorksheetOrTable.ws.colName.label]]",
                    description = "[[FilterWorksheetOrTable.ws.colName.description]]")
            @NotEmpty String wsColumnName,

            @Idx(index = "1.2.3.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "[[FilterWorksheetOrTable.ws.colIndex.label]]",
                    description = "[[FilterWorksheetOrTable.ws.colIndex.description]]")
            @NotEmpty Double wsColumnIndexOneBased,

            // Filter type
            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "[[FilterWorksheetOrTable.filterType.number.label]]", value = "NUMBER")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "[[FilterWorksheetOrTable.filterType.text.label]]", value = "TEXT"))
            })
            @Pkg(label = "[[FilterWorksheetOrTable.filterType.label]]",
                    description = "[[FilterWorksheetOrTable.filterType.description]]",
                    default_value = "TEXT", default_value_type = DataType.STRING)
            @SelectModes @NotEmpty String filterType,

            // Number operators + values
            @Idx(index = "2.1.1", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1.1.1", pkg = @Pkg(value = "EQ",  label = "[[FilterWorksheetOrTable.num.eq]]")),
                    @Idx.Option(index = "2.1.1.2", pkg = @Pkg(value = "NEQ", label = "[[FilterWorksheetOrTable.num.neq]]")),
                    @Idx.Option(index = "2.1.1.3", pkg = @Pkg(value = "GT",  label = "[[FilterWorksheetOrTable.num.gt]]")),
                    @Idx.Option(index = "2.1.1.4", pkg = @Pkg(value = "GTE", label = "[[FilterWorksheetOrTable.num.gte]]")),
                    @Idx.Option(index = "2.1.1.5", pkg = @Pkg(value = "LT",  label = "[[FilterWorksheetOrTable.num.lt]]")),
                    @Idx.Option(index = "2.1.1.6", pkg = @Pkg(value = "LTE", label = "[[FilterWorksheetOrTable.num.lte]]")),
                    @Idx.Option(index = "2.1.1.7", pkg = @Pkg(value = "BETWEEN", label = "[[FilterWorksheetOrTable.num.between]]"))
            })
            @Pkg(label = "[[FilterWorksheetOrTable.num.op.label]]",
                    description = "[[FilterWorksheetOrTable.num.op.description]]",
                    default_value = "EQ", default_value_type = DataType.STRING)
            @NotEmpty String numOperator,

            @Idx(index = "2.1.1.1.1", type = AttributeType.NUMBER) @Pkg() @NotEmpty Double numValueEq,
            @Idx(index = "2.1.1.2.1", type = AttributeType.NUMBER) @Pkg() @NotEmpty Double numValueNeq,
            @Idx(index = "2.1.1.3.1", type = AttributeType.NUMBER) @Pkg() @NotEmpty Double numValueGt,
            @Idx(index = "2.1.1.4.1", type = AttributeType.NUMBER) @Pkg() @NotEmpty Double numValueGte,
            @Idx(index = "2.1.1.5.1", type = AttributeType.NUMBER) @Pkg() @NotEmpty Double numValueLt,
            @Idx(index = "2.1.1.6.1", type = AttributeType.NUMBER) @Pkg() @NotEmpty Double numValueLte,
            @Idx(index = "2.1.1.7.1", type = AttributeType.NUMBER) @Pkg() @NotEmpty Double numValueBetween,
            @Idx(index = "2.1.1.7.2", type = AttributeType.NUMBER) @Pkg() @NotEmpty Double numValueBetween2,

            // Text operators + values
            @Idx(index = "2.2.1", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.2.1.1", pkg = @Pkg(value = "EQ",   label = "[[FilterWorksheetOrTable.txt.eq]]")),
                    @Idx.Option(index = "2.2.1.2", pkg = @Pkg(value = "NEQ",  label = "[[FilterWorksheetOrTable.txt.neq]]")),
                    @Idx.Option(index = "2.2.1.3", pkg = @Pkg(value = "BEG",  label = "[[FilterWorksheetOrTable.txt.begins]]")),
                    @Idx.Option(index = "2.2.1.4", pkg = @Pkg(value = "END",  label = "[[FilterWorksheetOrTable.txt.ends]]")),
                    @Idx.Option(index = "2.2.1.5", pkg = @Pkg(value = "CON",  label = "[[FilterWorksheetOrTable.txt.contains]]")),
                    @Idx.Option(index = "2.2.1.6", pkg = @Pkg(value = "NCON", label = "[[FilterWorksheetOrTable.txt.notcontains]]"))
            })
            @Pkg(label = "[[FilterWorksheetOrTable.txt.op.label]]",
                    description = "[[FilterWorksheetOrTable.txt.op.description]]",
                    default_value = "EQ", default_value_type = DataType.STRING)
            @NotEmpty String textOperator,

            @Idx(index = "2.2.1.1.1", type = AttributeType.TEXT) @Pkg() @NotEmpty String textValueEq,
            @Idx(index = "2.2.1.2.1", type = AttributeType.TEXT) @Pkg() @NotEmpty String textValueNeq,
            @Idx(index = "2.2.1.3.1", type = AttributeType.TEXT) @Pkg() @NotEmpty String textValueBeg,
            @Idx(index = "2.2.1.4.1", type = AttributeType.TEXT) @Pkg() @NotEmpty String textValueEnd,
            @Idx(index = "2.2.1.5.1", type = AttributeType.TEXT) @Pkg() @NotEmpty String textValueCon,
            @Idx(index = "2.2.1.6.1", type = AttributeType.TEXT) @Pkg() @NotEmpty String textValueNcon,

            // Session
            @Idx(index = "3", type = AttributeType.SESSION)
            @Pkg(label = "[[existingSession.label]]",
                    description = "[[existingSession.description]]",
                    default_value = "Default",
                    default_value_type = DataType.SESSION)
            @SessionObject @NotEmpty WorkbookSession session
    ) {
        if (session == null || session.getWorkbook() == null) {
            throw new BotCommandException("Workbook session not initialized. Open or create a workbook first.");
        }

        try {
            Workbook wb = session.getWorkbook();
            SpreadsheetVersion ver = wb.getSpreadsheetVersion();

            Sheet sheet;
            int headerRow0;
            int dataStartRow0;
            int firstCol0;
            int lastCol0;
            int filterCol0;
            int lastRow0;

            if ("TABLE".equals(mode)) {
                if (ver != SpreadsheetVersion.EXCEL2007 || !(wb instanceof XSSFWorkbook)) {
                    throw new BotCommandException("Table filtering requires an .xlsx workbook.");
                }
                if (isBlank(tableName)) {
                    throw new BotCommandException("Table name is required.");
                }

                // Resolve table area
                AreaReference ar = TableRange.getTableArea((XSSFWorkbook) wb, tableName.trim());
                if (ar == null) {
                    throw new BotCommandException("Table not found: " + tableName);
                }
                CellReference tl = ar.getFirstCell();
                CellReference br = ar.getLastCell();
                if (tl == null || br == null) {
                    throw new BotCommandException("Invalid table area for: " + tableName);
                }

                // Robust sheet resolution: sheet name may be absent in tl; scan tables by name
                XSSFSheet xsheet = resolveSheetForTable((XSSFWorkbook) wb, tableName, tl);
                sheet = xsheet;

                // Clear any sheet-level AutoFilter to avoid conflicts with table parts (no UI filters used)
                if (xsheet.getCTWorksheet().getAutoFilter() != null) {
                    xsheet.getCTWorksheet().unsetAutoFilter();
                }

                headerRow0 = tl.getRow();
                dataStartRow0 = headerRow0 + 1;
                firstCol0 = tl.getCol();
                lastCol0 = br.getCol();
                lastRow0 = br.getRow();

                // Validate header presence
                if (sheet.getRow(headerRow0) == null) {
                    throw new BotCommandException("Table header row not found at: " + headerRow0);
                }

                filterCol0 = resolveColumnIndexFromSelector(sheet, headerRow0, firstCol0, lastCol0, tableColSelector, tableColumnName, tableColumnIndexOneBased);
                if (filterCol0 < firstCol0 || filterCol0 > lastCol0) {
                    throw new BotCommandException("Selected filter column is outside the table range.");
                }
            } else if ("WORKSHEET".equals(mode)) {
                if (isBlank(sheetName)) {
                    throw new BotCommandException("Worksheet name is required.");
                }
                sheet = wb.getSheet(sheetName.trim());
                if (sheet == null) {
                    throw new BotCommandException("Worksheet not found: " + sheetName);
                }

                if ("ALL".equals(wsRangeMode)) {
                    headerRow0 = sheet.getFirstRowNum();
                    Row headerRow = sheet.getRow(headerRow0);
                    if (headerRow == null) {
                        throw new BotCommandException("Worksheet appears to be empty.");
                    }
                    firstCol0 = 0;
                    short lastCell = headerRow.getLastCellNum(); // -1 if none
                    lastCol0 = lastCell > 0 ? lastCell - 1 : 0;
                    dataStartRow0 = headerRow0 + 1;
                    lastRow0 = sheet.getLastRowNum();
                } else if ("SPECIFIC".equals(wsRangeMode)) {
                    if (isBlank(wsRangeA1)) {
                        throw new BotCommandException("Cell range is required when selecting a specific range.");
                    }
                    int[] b = parseA1Range(wsRangeA1.trim()); // [r0,c0,r1,c1]
                    headerRow0 = b[0];
                    dataStartRow0 = headerRow0 + 1;
                    firstCol0 = b[1];
                    lastCol0 = b[3];
                    lastRow0 = b[2];
                } else {
                    throw new BotCommandException("Invalid worksheet range mode.");
                }

                // Validate header presence
                if (sheet.getRow(headerRow0) == null) {
                    throw new BotCommandException("Header row not found at: " + headerRow0);
                }

                filterCol0 = resolveColumnIndexFromSelector(sheet, headerRow0, firstCol0, lastCol0, wsColSelector, wsColumnName, wsColumnIndexOneBased);
                if (filterCol0 < firstCol0 || filterCol0 > lastCol0) {
                    throw new BotCommandException("Selected filter column is outside the specified range.");
                }
                // No AutoFilter UI: do not set or modify any autoFilter on the sheet
            } else {
                throw new BotCommandException("Invalid mode. Choose TABLE or WORKSHEET.");
            }

            if (lastRow0 < dataStartRow0) {
                throw new BotCommandException("No data rows found to apply filter.");
            }

            // Parse only the active operator set
            Double numValue1 = null, numValue2 = null;
            String textValue = null;

            if ("NUMBER".equals(filterType)) {
                if (numOperator == null) {
                    throw new BotCommandException("Number operator is required for Number filter type.");
                }
                switch (numOperator) {
                    case "EQ":  requireNotNull(numValueEq,  "Number value (equals)");           numValue1 = numValueEq;  break;
                    case "NEQ": requireNotNull(numValueNeq, "Number value (not equals)");       numValue1 = numValueNeq; break;
                    case "GT":  requireNotNull(numValueGt,  "Number value (greater than)");     numValue1 = numValueGt;  break;
                    case "GTE": requireNotNull(numValueGte, "Number value (greater or equal)"); numValue1 = numValueGte; break;
                    case "LT":  requireNotNull(numValueLt,  "Number value (less than)");        numValue1 = numValueLt;  break;
                    case "LTE": requireNotNull(numValueLte, "Number value (less or equal)");    numValue1 = numValueLte; break;
                    case "BETWEEN":
                        requireNotNull(numValueBetween,  "Number value 1 (between)");
                        requireNotNull(numValueBetween2, "Number value 2 (between)");
                        numValue1 = numValueBetween;
                        numValue2 = numValueBetween2;
                        break;
                    default:
                        throw new BotCommandException("Invalid number operator: " + numOperator);
                }
            } else if ("TEXT".equals(filterType)) {
                if (textOperator == null) {
                    throw new BotCommandException("Text operator is required for Text filter type.");
                }
                switch (textOperator) {
                    case "EQ":   textValue = requireNotBlank(textValueEq,  "Text value (equals)");        break;
                    case "NEQ":  textValue = requireNotBlank(textValueNeq, "Text value (not equals)");    break;
                    case "BEG":  textValue = requireNotBlank(textValueBeg, "Text value (begins with)");   break;
                    case "END":  textValue = requireNotBlank(textValueEnd, "Text value (ends with)");     break;
                    case "CON":  textValue = requireNotBlank(textValueCon, "Text value (contains)");      break;
                    case "NCON": textValue = requireNotBlank(textValueNcon,"Text value (not contains)");  break;
                    default:
                        throw new BotCommandException("Invalid text operator: " + textOperator);
                }
            } else {
                throw new BotCommandException("Invalid filter type. Choose NUMBER or TEXT.");
            }

            applyCriteriaByHidingRows(sheet, dataStartRow0, lastRow0, filterCol0, filterType,
                    numOperator, numValue1, numValue2, textOperator, textValue);

        } catch (BotCommandException e) {
            throw e;
        } catch (Throwable t) {
            String context = "Filter failed"
                    + " [mode=" + safe(mode)
                    + ", table=" + safe(tableName)
                    + ", sheet=" + safe(sheetName)
                    + ", filterType=" + safe(filterType)
                    + ", numOp=" + safe(numOperator)
                    + ", textOp=" + safe(textOperator)
                    + "]";
            throw new BotCommandException(context + ": " + t.getMessage(), t);
        }
    }

    // Resolve the sheet that contains a table by name, tolerating a null CellReference sheet name
    private XSSFSheet resolveSheetForTable(XSSFWorkbook wb, String tableName, CellReference tl) {
        String tlSheet = (tl != null) ? tl.getSheetName() : null;
        if (tlSheet != null) {
            XSSFSheet s = wb.getSheet(tlSheet);
            if (s != null) return s;
        }
        // Scan all sheets/tables for matching table name
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            XSSFSheet s = wb.getSheetAt(i);
            for (XSSFTable t : s.getTables()) {
                String n1 = t.getName();
                String n2 = t.getCTTable().getName();
                if ((n1 != null && n1.equalsIgnoreCase(tableName)) ||
                        (n2 != null && n2.equalsIgnoreCase(tableName))) {
                    return s;
                }
            }
        }
        // Fallback to active sheet
        return wb.getSheetAt(wb.getActiveSheetIndex());
    }

    private int resolveColumnIndexFromSelector(Sheet sheet, int headerRow0, int firstCol0, int lastCol0, String selectorMode,
                                               String colName, Double colIndexOneBased) {
        if ("BY_NAME".equals(selectorMode)) {
            if (isBlank(colName)) {
                throw new BotCommandException("Column name cannot be empty.");
            }
            Row header = sheet.getRow(headerRow0);
            if (header == null) throw new BotCommandException("Header row not found.");
            String target = colName.trim();
            int last = header.getLastCellNum() > 0 ? header.getLastCellNum() - 1 : 0;
            for (int c = firstNonNegative(0); c <= last; c++) {
                Cell hc = header.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (hc != null) {
                    String h = getCellString(hc);
                    if (target.equalsIgnoreCase(h)) return c;
                }
            }
            throw new BotCommandException("Column not found in header: " + target);
        } else if ("BY_INDEX".equals(selectorMode)) {
            int colRange = lastCol0 - firstCol0 + 1;
            if (colIndexOneBased == null || colIndexOneBased < 1) {
                throw new BotCommandException("Column index must be 1 or greater.");
            } else if (colIndexOneBased > colRange) {
                throw new BotCommandException("Column index must be smaller than the range width.");
            }
            return colIndexOneBased.intValue() - 1 + firstCol0;
        } else {
            throw new BotCommandException("Invalid column selector mode.");
        }
    }

    private int firstNonNegative(int v) { return Math.max(0, v); }

    private void applyCriteriaByHidingRows(Sheet sheet,
                                           int dataStartRow0,
                                           int lastRow0,
                                           int filterCol0,
                                           String filterType,
                                           String numOp,
                                           Double v1,
                                           Double v2,
                                           String txtOp,
                                           String tval) {
        DataFormatter fmt = new DataFormatter(Locale.getDefault());
        for (int r = dataStartRow0; r <= lastRow0; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            Cell cell = row.getCell(filterCol0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

            boolean match;
            if ("NUMBER".equals(filterType)) {
                Double val = readNumeric(cell);
                match = evaluateNumeric(val, numOp, v1, v2);
            } else if ("TEXT".equals(filterType)) {
                String val = (cell == null) ? "" : fmt.formatCellValue(cell);
                match = evaluateText(val, txtOp, tval == null ? "" : tval);
            } else {
                throw new BotCommandException("Invalid filter type. Choose NUMBER or TEXT.");
            }
            row.setZeroHeight(!match);
        }
    }

    private Double readNumeric(Cell cell) {
        if (cell == null) return null;
        if (cell.getCellType() == CellType.NUMERIC && !DateUtil.isCellDateFormatted(cell)) {
            return cell.getNumericCellValue();
        }
        if (cell.getCellType() == CellType.FORMULA) {
            switch (cell.getCachedFormulaResultType()) {
                case NUMERIC:
                    if (!DateUtil.isCellDateFormatted(cell)) {
                        return cell.getNumericCellValue();
                    }
                    break;
                default:
                    return null;
            }
        }
        return null;
    }

    private boolean evaluateNumeric(Double actual, String op, Double a, Double b) {
        if (actual == null) return false;
        if ("EQ".equals(op))  return a != null && Double.compare(actual, a) == 0;
        if ("NEQ".equals(op)) return a != null && Double.compare(actual, a) != 0;
        if ("GT".equals(op))  return a != null && actual > a;
        if ("GTE".equals(op)) return a != null && actual >= a;
        if ("LT".equals(op))  return a != null && actual < a;
        if ("LTE".equals(op)) return a != null && actual <= a;
        if ("BETWEEN".equals(op)) return a != null && b != null && actual >= Math.min(a, b) && actual <= Math.max(a, b);
        throw new BotCommandException("Invalid number operator.");
    }

    private boolean evaluateText(String actual, String op, String val) {
        String A = actual == null ? "" : actual;
        String V = val == null ? "" : val;
        String aL = A.toLowerCase();
        String vL = V.toLowerCase();
        if ("EQ".equals(op))   return aL.equals(vL);
        if ("NEQ".equals(op))  return !aL.equals(vL);
        if ("BEG".equals(op))  return aL.startsWith(vL);
        if ("END".equals(op))  return aL.endsWith(vL);
        if ("CON".equals(op))  return aL.contains(vL);
        if ("NCON".equals(op)) return !aL.contains(vL);
        throw new BotCommandException("Invalid text operator.");
    }

    // Parse A1 range like A1:D200 → [row0, col0, row1, col1]
    private int[] parseA1Range(String a1) {
        String[] parts = a1.split(":");
        if (parts.length != 2) throw new BotCommandException("Invalid A1 range: " + a1);
        CellReference a = new CellReference(parts[0]);
        CellReference b = new CellReference(parts[1]);
        int r0 = Math.min(a.getRow(), b.getRow());
        int c0 = Math.min(a.getCol(), b.getCol());
        int r1 = Math.max(a.getRow(), b.getRow());
        int c1 = Math.max(a.getCol(), b.getCol());
        return new int[]{r0, c0, r1, c1};
    }

    private String getCellString(Cell c) {
        if (c == null) return "";
        switch (c.getCellType()) {
            case STRING:  return c.getStringCellValue();
            case NUMERIC: return DateUtil.isCellDateFormatted(c) ? "" : String.valueOf(c.getNumericCellValue());
            case BOOLEAN: return String.valueOf(c.getBooleanCellValue());
            case FORMULA:
                switch (c.getCachedFormulaResultType()) {
                    case STRING:  return c.getStringCellValue();
                    case NUMERIC: return DateUtil.isCellDateFormatted(c) ? "" : String.valueOf(c.getNumericCellValue());
                    case BOOLEAN: return String.valueOf(c.getBooleanCellValue());
                    default: return "";
                }
            default: return "";
        }
    }

    private static boolean isBlank(String s) { return s == null || s.trim().isEmpty(); }
    private static String safe(Object o) { return o == null ? "null" : String.valueOf(o); }

    private static void requireNotNull(Object val, String label) {
        if (val == null) throw new BotCommandException(label + " is required.");
    }

    private static String requireNotBlank(String val, String label) {
        if (val == null || val.trim().isEmpty()) {
            throw new BotCommandException(label + " is required.");
        }
        return val;
    }
}
