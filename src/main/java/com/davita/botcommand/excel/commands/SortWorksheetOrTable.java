package com.davita.botcommand.excel.commands;

import com.davita.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.CommandPkg;
import com.automationanywhere.commandsdk.annotations.Execute;
import com.automationanywhere.commandsdk.annotations.Idx;
import com.automationanywhere.commandsdk.annotations.Pkg;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SelectModes;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.davita.botcommand.excel.internal.TableRange;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;

@BotCommand
@CommandPkg(
        name = "sortWorksheetOrTable",
        label = "[[SortWorksheetOrTable.label]]",
        node_label = "[[SortWorksheetOrTable.node_label]]",
        description = "[[SortWorksheetOrTable.description]]",
        icon = "excel-icon.svg"
)
public class SortWorksheetOrTable {

    @Execute
    public void action(
            // 1) Mode: Table vs Worksheet
            @Idx(index = "1", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "1.1", pkg = @Pkg(label = "[[SortWorksheetOrTable.mode.table.label]]", value = "TABLE")),
                    @Idx.Option(index = "1.2", pkg = @Pkg(label = "[[SortWorksheetOrTable.mode.worksheet.label]]", value = "WORKSHEET"))
            })
            @Pkg(label = "[[SortWorksheetOrTable.mode.label]]",
                    description = "[[SortWorksheetOrTable.mode.description]]",
                    default_value = "TABLE", default_value_type = DataType.STRING)
            @SelectModes @NotEmpty String mode,

            // 1.1) TABLE branch
            @Idx(index = "1.1.1", type = AttributeType.TEXT)
            @Pkg(label = "[[SortWorksheetOrTable.table.name.label]]",
                    description = "[[SortWorksheetOrTable.table.name.description]]")
            @NotEmpty String tableName,

            @Idx(index = "1.1.2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "1.1.2.1", pkg = @Pkg(label = "[[SortWorksheetOrTable.table.colByName.label]]", value = "BY_NAME")),
                    @Idx.Option(index = "1.1.2.2", pkg = @Pkg(label = "[[SortWorksheetOrTable.table.colByIndex.label]]", value = "BY_INDEX"))
            })
            @Pkg(label = "[[SortWorksheetOrTable.table.colSelector.label]]",
                    description = "[[SortWorksheetOrTable.table.colSelector.description]]",
                    default_value = "BY_NAME", default_value_type = DataType.STRING)
            @SelectModes @NotEmpty String tableColSelector,

            @Idx(index = "1.1.2.1.1", type = AttributeType.TEXT)
            @Pkg(label = "[[SortWorksheetOrTable.table.colName.label]]",
                    description = "[[SortWorksheetOrTable.table.colName.description]]")
            @NotEmpty String tableColumnName,

            @Idx(index = "1.1.2.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "[[SortWorksheetOrTable.table.colIndex.label]]",
                    description = "[[SortWorksheetOrTable.table.colIndex.description]]")
            @NotEmpty Double tableColumnIndexOneBased,

            // 1.2) WORKSHEET branch
            @Idx(index = "1.2.1", type = AttributeType.TEXT)
            @Pkg(label = "[[SortWorksheetOrTable.ws.name.label]]",
                    description = "[[SortWorksheetOrTable.ws.name.description]]")
            @NotEmpty String sheetName,

            @Idx(index = "1.2.2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "1.2.2.1", pkg = @Pkg(label = "[[SortWorksheetOrTable.ws.rangeAll.label]]", value = "ALL")),
                    @Idx.Option(index = "1.2.2.2", pkg = @Pkg(label = "[[SortWorksheetOrTable.ws.rangeSpecific.label]]", value = "SPECIFIC"))
            })
            @Pkg(label = "[[SortWorksheetOrTable.ws.rangeMode.label]]",
                    description = "[[SortWorksheetOrTable.ws.rangeMode.description]]",
                    default_value = "ALL", default_value_type = DataType.STRING)
            @NotEmpty String wsRangeMode,

            @Idx(index = "1.2.2.2.1", type = AttributeType.TEXT)
            @Pkg(label = "[[SortWorksheetOrTable.ws.rangeText.label]]",
                    description = "[[SortWorksheetOrTable.ws.rangeText.description]]")
            @NotEmpty String wsRangeA1,

            @Idx(index = "1.2.3", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "1.2.3.1", pkg = @Pkg(label = "[[SortWorksheetOrTable.ws.colByName.label]]", value = "BY_NAME")),
                    @Idx.Option(index = "1.2.3.2", pkg = @Pkg(label = "[[SortWorksheetOrTable.ws.colByIndex.label]]", value = "BY_INDEX"))
            })
            @Pkg(label = "[[SortWorksheetOrTable.ws.colSelector.label]]",
                    description = "[[SortWorksheetOrTable.ws.colSelector.description]]",
                    default_value = "BY_NAME", default_value_type = DataType.STRING)
            @SelectModes @NotEmpty String wsColSelector,

            @Idx(index = "1.2.3.1.1", type = AttributeType.TEXT)
            @Pkg(label = "[[SortWorksheetOrTable.ws.colName.label]]",
                    description = "[[SortWorksheetOrTable.ws.colName.description]]")
            @NotEmpty String wsColumnName,

            @Idx(index = "1.2.3.2.1", type = AttributeType.NUMBER)
            @Pkg(label = "[[SortWorksheetOrTable.ws.colIndex.label]]",
                    description = "[[SortWorksheetOrTable.ws.colIndex.description]]")
            @NotEmpty Double wsColumnIndexOneBased,

            @Idx(index = "1.2.4", type = AttributeType.CHECKBOX)
            @Pkg(label = "[[SortWorksheetOrTable.ws.hasHeader.label]]",
                    description = "[[SortWorksheetOrTable.ws.hasHeader.description]]",
                    default_value = "True", default_value_type = DataType.BOOLEAN)
            @NotEmpty Boolean wsHasHeader,

            // 2) Sort type
            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "[[SortWorksheetOrTable.sortType.number.label]]", value = "NUMBER")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "[[SortWorksheetOrTable.sortType.text.label]]", value = "TEXT"))
            })
            @Pkg(label = "[[SortWorksheetOrTable.sortType.label]]",
                    description = "[[SortWorksheetOrTable.sortType.description]]",
                    default_value = "TEXT", default_value_type = DataType.STRING)
            @SelectModes @NotEmpty String sortType,

            // 2.1) Number order
            @Idx(index = "2.1.1", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.1.1.1", pkg = @Pkg(label = "[[SortWorksheetOrTable.order.num.asc]]", value = "ASC")),
                    @Idx.Option(index = "2.1.1.2", pkg = @Pkg(label = "[[SortWorksheetOrTable.order.num.desc]]", value = "DESC"))
            })
            @Pkg(label = "[[SortWorksheetOrTable.order.num.label]]",
                    description = "[[SortWorksheetOrTable.order.num.description]]",
                    default_value = "ASC", default_value_type = DataType.STRING)
            @NotEmpty String numberOrder,

            // 2.2) Text order
            @Idx(index = "2.2.1", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "2.2.1.1", pkg = @Pkg(label = "[[SortWorksheetOrTable.order.txt.asc]]", value = "ASC")),
                    @Idx.Option(index = "2.2.1.2", pkg = @Pkg(label = "[[SortWorksheetOrTable.order.txt.desc]]", value = "DESC"))
            })
            @Pkg(label = "[[SortWorksheetOrTable.order.txt.label]]",
                    description = "[[SortWorksheetOrTable.order.txt.description]]",
                    default_value = "ASC", default_value_type = DataType.STRING)
            @NotEmpty String textOrder,

            // 3) Session
            @Idx(index = "3", type = AttributeType.SESSION)
            @Pkg(label = "[[existingSession.label]]",
                    description = "[[existingSession.description]]",
                    default_value = "Default", default_value_type = DataType.SESSION)
            @SessionObject @NotEmpty WorkbookSession session
    ) {
        if (session == null || session.getWorkbook() == null) {
            throw new BotCommandException("Workbook session not initialized. Open or create a workbook first.");
        }
        Workbook wb = session.getWorkbook();
        SpreadsheetVersion ver = wb.getSpreadsheetVersion();

        Sheet sheet;
        int headerRow0;
        int dataStartRow0;
        int firstCol0;
        int lastCol0;
        int sortCol0;
        int lastRow0;

        try {
            if ("TABLE".equals(mode)) {
                if (ver != SpreadsheetVersion.EXCEL2007 || !(wb instanceof XSSFWorkbook)) {
                    throw new BotCommandException("Table sorting requires an .xlsx workbook.");
                }
                if (isBlank(tableName)) {
                    throw new BotCommandException("Table name is required.");
                }
                AreaReference ar = TableRange.getTableArea((XSSFWorkbook) wb, tableName.trim());
                if (ar == null) throw new BotCommandException("Table not found: " + tableName);
                CellReference tl = ar.getFirstCell();
                CellReference br = ar.getLastCell();
                if (tl == null || br == null) throw new BotCommandException("Invalid table area for: " + tableName);

                XSSFSheet xsheet = resolveSheetForTable((XSSFWorkbook) wb, tableName, tl);
                sheet = xsheet;

                headerRow0 = tl.getRow();
                dataStartRow0 = headerRow0 + 1;
                firstCol0 = tl.getCol();
                lastCol0 = br.getCol();
                lastRow0 = br.getRow();

                if (sheet.getRow(headerRow0) == null) {
                    throw new BotCommandException("Table header row not found at: " + headerRow0);
                }

                sortCol0 = resolveColumnIndexFromSelector(sheet, headerRow0, firstCol0, lastCol0,
                        tableColSelector, tableColumnName, tableColumnIndexOneBased);
                if (sortCol0 < firstCol0 || sortCol0 > lastCol0) {
                    throw new BotCommandException("Selected sort column is outside the table range.");
                }
            } else if ("WORKSHEET".equals(mode)) {
                if (isBlank(sheetName)) throw new BotCommandException("Worksheet name is required.");
                sheet = wb.getSheet(sheetName.trim());
                if (sheet == null) throw new BotCommandException("Worksheet not found: " + sheetName);

                if ("ALL".equals(wsRangeMode)) {
                    headerRow0 = sheet.getFirstRowNum();
                    Row headerRow = sheet.getRow(headerRow0);
                    if (headerRow == null) throw new BotCommandException("Worksheet appears to be empty.");
                    firstCol0 = 0;
                    short lastCell = headerRow.getLastCellNum();
                    lastCol0 = lastCell > 0 ? lastCell - 1 : 0;
                    lastRow0 = sheet.getLastRowNum();
                    dataStartRow0 = wsHasHeader != null && wsHasHeader ? headerRow0 + 1 : headerRow0;
                } else if ("SPECIFIC".equals(wsRangeMode)) {
                    if (isBlank(wsRangeA1)) {
                        throw new BotCommandException("Cell range is required when selecting a specific range.");
                    }
                    int[] b = parseA1Range(wsRangeA1.trim()); // [r0,c0,r1,c1]
                    headerRow0 = b[0];
                    firstCol0 = b[1];
                    lastRow0 = b[2];
                    lastCol0 = b[3];
                    dataStartRow0 = wsHasHeader != null && wsHasHeader ? headerRow0 + 1 : headerRow0;
                } else {
                    throw new BotCommandException("Invalid worksheet range mode.");
                }

                if (wsHasHeader != null && wsHasHeader && sheet.getRow(headerRow0) == null) {
                    throw new BotCommandException("Header row not found at: " + headerRow0);
                }

                sortCol0 = resolveColumnIndexFromSelector(sheet, headerRow0, firstCol0, lastCol0,
                        wsColSelector, wsColumnName, wsColumnIndexOneBased);
                if (sortCol0 < firstCol0 || sortCol0 > lastCol0) {
                    throw new BotCommandException("Selected sort column is outside the specified range.");
                }
            } else {
                throw new BotCommandException("Invalid mode. Choose TABLE or WORKSHEET.");
            }

            if (lastRow0 < dataStartRow0) {
                return; // nothing to sort
            }

            // Snapshot rows within the sort range
            List<RowSnapshot> rows = new ArrayList<>();
            for (int r = dataStartRow0; r <= lastRow0; r++) {
                rows.add(snapshotRow(sheet, r, firstCol0, lastCol0));
            }

            // Build comparator with explicit types to aid inference
            Comparator<RowSnapshot> cmp;
            if ("NUMBER".equalsIgnoreCase(sortType)) {
                final boolean asc = !"DESC".equalsIgnoreCase(numberOrder);
                cmp = Comparator.comparing(
                        (RowSnapshot rs) -> rs.numericAt(sheet, sortCol0),
                        (Double a, Double b) -> compareNumbersNullsLast(a, b, asc)
                ).thenComparingInt((RowSnapshot rs) -> rs.rowIndex);
            } else if ("TEXT".equalsIgnoreCase(sortType)) {
                final boolean asc = !"DESC".equalsIgnoreCase(textOrder);
                final DataFormatter fmt = new DataFormatter(Locale.getDefault());
                final FormulaEvaluator eval = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
                cmp = Comparator.comparing(
                        (RowSnapshot rs) -> rs.textAt(sheet, sortCol0, fmt, eval),
                        (String a, String b) -> compareStringsNullsLast(a, b, asc)
                ).thenComparingInt((RowSnapshot rs) -> rs.rowIndex);
            } else {
                throw new BotCommandException("Invalid sort type. Choose NUMBER or TEXT.");
            }

            rows.sort(cmp);

            // Write back in order
            int target = dataStartRow0;
            for (RowSnapshot rs : rows) {
                rs.writeBackTo(sheet, target++, firstCol0, lastCol0);
            }

            // Set active cell to top-left of sorted region
            sheet.setActiveCell(new CellAddress(dataStartRow0, firstCol0));

        } catch (BotCommandException e) {
            throw e;
        } catch (Throwable t) {
            String context = "Sort failed"
                    + " [mode=" + safe(mode)
                    + ", table=" + safe(tableName)
                    + ", sheet=" + safe(sheetName)
                    + ", sortType=" + safe(sortType)
                    + "]";
            throw new BotCommandException(context + ": " + t.getMessage(), t);
        }
    }

    // Snapshot of a row subset [c0..c1]
    private static class RowSnapshot {
        final int rowIndex;
        final CellSnapshot[] cells;

        RowSnapshot(int rowIndex, int width) {
            this.rowIndex = rowIndex;
            this.cells = new CellSnapshot[width];
        }

        // Read numeric value from the current sheet for the given absolute column
        Double numericAt(Sheet sheet, int absCol) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) return null;
            Cell cell = row.getCell(absCol, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
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

        // Read display text from the current sheet using DataFormatter
        String textAt(Sheet sheet, int absCol, DataFormatter fmt, FormulaEvaluator eval) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) return "";
            Cell cell = row.getCell(absCol, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (cell == null) return "";
            return fmt.formatCellValue(cell, eval);
        }

        void writeBackTo(Sheet sheet, int targetRowIdx, int c0, int c1) {
            Row row = sheet.getRow(targetRowIdx);
            if (row == null) row = sheet.createRow(targetRowIdx);
            for (int c = c0; c <= c1; c++) {
                int ix = c - c0;
                CellSnapshot snap = cells[ix];
                Cell cell = row.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                if (snap == null) {
                    cell.setBlank();
                } else {
                    snap.applyTo(cell);
                }
            }
        }
    }

    // Snapshot of a single cell: type, value/formula, and style
    private static class CellSnapshot {
        final CellType type;
        final String str;
        final Double num;
        final Boolean bool;
        final String formula;
        final CellStyle style;

        private CellSnapshot(CellType type, String str, Double num, Boolean bool, String formula, CellStyle style) {
            this.type = type;
            this.str = str;
            this.num = num;
            this.bool = bool;
            this.formula = formula;
            this.style = style;
        }

        static CellSnapshot fromCell(Cell cell) {
            if (cell == null) return null;
            CellStyle style = cell.getCellStyle();
            switch (cell.getCellType()) {
                case STRING:
                    return new CellSnapshot(CellType.STRING, cell.getStringCellValue(), null, null, null, style);
                case NUMERIC:
                    return new CellSnapshot(CellType.NUMERIC, null, cell.getNumericCellValue(), null, null, style);
                case BOOLEAN:
                    return new CellSnapshot(CellType.BOOLEAN, null, null, cell.getBooleanCellValue(), null, style);
                case FORMULA:
                    return new CellSnapshot(CellType.FORMULA, null, null, null, cell.getCellFormula(), style);
                case BLANK:
                default:
                    return new CellSnapshot(CellType.BLANK, null, null, null, null, style);
            }
        }

        void applyTo(Cell cell) {
            if (style != null) cell.setCellStyle(style);
            switch (type) {
                case STRING:
                    cell.setCellValue(str == null ? "" : str);
                    break;
                case NUMERIC:
                    cell.setCellValue(num == null ? 0d : num);
                    break;
                case BOOLEAN:
                    cell.setCellValue(bool != null && bool);
                    break;
                case FORMULA:
                    cell.setCellFormula(formula == null ? "" : formula);
                    break;
                case BLANK:
                default:
                    cell.setBlank();
                    break;
            }
        }
    }

    // Create a snapshot of row r for c in [c0..c1]
    private RowSnapshot snapshotRow(Sheet sheet, int r, int c0, int c1) {
        Row row = sheet.getRow(r);
        RowSnapshot rs = new RowSnapshot(r, c1 - c0 + 1);
        for (int c = c0; c <= c1; c++) {
            Cell cell = (row == null) ? null : row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            rs.cells[c - c0] = CellSnapshot.fromCell(cell);
        }
        return rs;
    }

    // Resolve table's sheet (tolerates null sheet name on CellReference)
    private XSSFSheet resolveSheetForTable(XSSFWorkbook wb, String tableName, CellReference tl) {
        String tlSheet = (tl != null) ? tl.getSheetName() : null;
        if (tlSheet != null) {
            XSSFSheet s = wb.getSheet(tlSheet);
            if (s != null) return s;
        }
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            XSSFSheet s = wb.getSheetAt(i);
            if (s == null) continue;
            for (org.apache.poi.xssf.usermodel.XSSFTable t : s.getTables()) {
                String n1 = t.getName();
                String n2 = t.getCTTable().getName();
                if ((n1 != null && n1.equalsIgnoreCase(tableName)) ||
                        (n2 != null && n2.equalsIgnoreCase(tableName))) {
                    return s;
                }
            }
        }
        return wb.getSheetAt(wb.getActiveSheetIndex());
    }

    // Column selector helper (BY_NAME searches headerRow0; BY_INDEX is 1-based within [firstCol0..lastCol0])
    private int resolveColumnIndexFromSelector(Sheet sheet, int headerRow0, int firstCol0, int lastCol0,
                                               String selectorMode, String colName, Double colIndexOneBased) {
        if ("BY_NAME".equals(selectorMode)) {
            if (isBlank(colName)) throw new BotCommandException("Column name cannot be empty.");
            Row header = sheet.getRow(headerRow0);
            if (header == null) throw new BotCommandException("Header row not found.");
            String target = colName.trim();
            int last = header.getLastCellNum() > 0 ? header.getLastCellNum() - 1 : 0;
            for (int c = Math.max(0, firstCol0); c <= Math.min(lastCol0, last); c++) {
                Cell hc = header.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (hc != null) {
                    String h = getCellString(hc);
                    if (target.equalsIgnoreCase(h)) return c;
                }
            }
            throw new BotCommandException("Column not found in header: " + target);
        } else if ("BY_INDEX".equals(selectorMode)) {
            if (colIndexOneBased == null || colIndexOneBased < 1) {
                throw new BotCommandException("Column index must be 1 or greater.");
            }
            int colRange = lastCol0 - firstCol0 + 1;
            if (colIndexOneBased > colRange) {
                throw new BotCommandException("Column index must be smaller than the range width.");
            }
            return colIndexOneBased.intValue() - 1 + firstCol0;
        } else {
            throw new BotCommandException("Invalid column selector mode.");
        }
    }

    // Parse A1 range "A1:D200" -> [row0, col0, row1, col1]
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
            case STRING: return c.getStringCellValue();
            case NUMERIC: return DateUtil.isCellDateFormatted(c) ? "" : String.valueOf(c.getNumericCellValue());
            case BOOLEAN: return String.valueOf(c.getBooleanCellValue());
            case FORMULA:
                switch (c.getCachedFormulaResultType()) {
                    case STRING: return c.getStringCellValue();
                    case NUMERIC: return DateUtil.isCellDateFormatted(c) ? "" : String.valueOf(c.getNumericCellValue());
                    case BOOLEAN: return String.valueOf(c.getBooleanCellValue());
                    default: return "";
                }
            default: return "";
        }
    }

    private static boolean isBlank(String s) { return s == null || s.trim().isEmpty(); }
    private static String safe(Object o) { return o == null ? "null" : String.valueOf(o); }

    private static int compareNumbersNullsLast(Double a, Double b, boolean asc) {
        if (a == null && b == null) return 0;
        if (a == null) return 1;   // nulls last
        if (b == null) return -1;  // nulls last
        int cmp = Double.compare(a, b);
        return asc ? cmp : -cmp;
    }

    private static int compareStringsNullsLast(String a, String b, boolean asc) {
        String aa = a == null ? "" : a;
        String bb = b == null ? "" : b;
        boolean aEmpty = aa.trim().isEmpty();
        boolean bEmpty = bb.trim().isEmpty();
        if (aEmpty && bEmpty) return 0;
        if (aEmpty) return 1;
        if (bEmpty) return -1;
        int cmp = aa.compareToIgnoreCase(bb);
        return asc ? cmp : -cmp;
    }
}
