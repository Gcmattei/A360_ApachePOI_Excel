package com.davita.botcommand.excel.iterators;

import com.davita.botcommand.excel.sessions.WorkbookSession;
import com.davita.botcommand.excel.internal.TableRange;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.BooleanValue;
import com.automationanywhere.botcommand.data.impl.NumberValue;
import com.automationanywhere.botcommand.data.impl.RecordValue;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.data.model.Schema;
import com.automationanywhere.botcommand.data.model.record.Record;

import com.automationanywhere.botcommand.exception.BotCommandException;

import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.BotCommand.CommandType;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SelectModes;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

// IMPORTANT: make the following helpers from GetWorksheetAsTable public static (or move to a shared utils class)
// so they can be reused here via static import:
//   toRowIndex(String, Double)
//   calculateMaxColumns(Sheet, int, int)
//   buildHeadersInArea(Sheet, boolean, int, int, int, String, DataFormatter, FormulaEvaluator)
//   readVisible(Cell, DataFormatter, FormulaEvaluator)
//   readRaw(Cell)
import static com.davita.botcommand.excel.commands.GetWorksheetAsTable.toRowIndex;
import static com.davita.botcommand.excel.commands.GetWorksheetAsTable.calculateMaxColumns;
import static com.davita.botcommand.excel.commands.GetWorksheetAsTable.buildHeadersInArea;
import static com.davita.botcommand.excel.commands.GetWorksheetAsTable.readVisible;
import static com.davita.botcommand.excel.commands.GetWorksheetAsTable.readRaw;
import static com.davita.botcommand.excel.commands.SortWorksheetOrTable.resolveSheetForTable;

@BotCommand(commandType = CommandType.Iterator)
@CommandPkg(
        name = "iterateWorksheetOrTable",
        label = "[[IterateWorksheetOrTable.label]]",
        node_label = "[[IterateWorksheetOrTable.node_label]]",
        description = "[[IterateWorksheetOrTable.description]]",
        return_type = DataType.RECORD,
        return_label = "[[IterateWorksheetOrTable.return.label]]",
        icon = "excel-icon.svg"
)
public class IterateWorksheetOrTable {

    // Mode: TABLE or WORKSHEET
    @Idx(index = "1", type = AttributeType.SELECT, options = {
            @Idx.Option(index = "1.1", pkg = @Pkg(label = "[[IterateWorksheetOrTable.mode.table.label]]", value = "TABLE")),
            @Idx.Option(index = "1.2", pkg = @Pkg(label = "[[IterateWorksheetOrTable.mode.worksheet.label]]", value = "WORKSHEET"))
    })
    @Pkg(label = "[[IterateWorksheetOrTable.mode.label]]",
            description = "[[IterateWorksheetOrTable.mode.description]]",
            default_value = "TABLE", default_value_type = DataType.STRING)
    @SelectModes @Inject @NotEmpty
    private String mode;

    // TABLE inputs
    @Idx(index = "1.1.1", type = AttributeType.TEXT)
    @Pkg(label = "[[IterateWorksheetOrTable.table.name.label]]",
            description = "[[IterateWorksheetOrTable.table.name.description]]")
    @Inject
    private String tableName;

    @Idx(index = "1.1.2", type = AttributeType.SELECT, options = {
            @Idx.Option(index = "1.1.2.1", pkg = @Pkg(label = "[[IterateWorksheetOrTable.rows.all.label]]", value = "ALL_ROWS")),
            @Idx.Option(index = "1.1.2.2", pkg = @Pkg(label = "[[IterateWorksheetOrTable.rows.specific.label]]", value = "SPECIFIC_ROWS"))
    })
    @Pkg(label = "[[IterateWorksheetOrTable.table.rowsMode.label]]",
            description = "[[IterateWorksheetOrTable.table.rowsMode.description]]",
            default_value = "ALL_ROWS", default_value_type = DataType.STRING)
    @SelectModes @Inject @NotEmpty
    private String tableRowsMode;

    @Idx(index = "1.1.2.2.1", type = AttributeType.NUMBER)
    @Pkg(label = "[[IterateWorksheetOrTable.rows.start.label]]",
            description = "[[IterateWorksheetOrTable.rows.start.description]]")
    @Inject
    private Double tableStartRowOneBased; // optional (1-based within data region)

    @Idx(index = "1.1.2.2.2", type = AttributeType.NUMBER)
    @Pkg(label = "[[IterateWorksheetOrTable.rows.end.label]]",
            description = "[[IterateWorksheetOrTable.rows.end.description]]")
    @Inject
    private Double tableEndRowOneBased;   // optional

    // WORKSHEET inputs
    @Idx(index = "1.2.1", type = AttributeType.TEXT)
    @Pkg(label = "[[IterateWorksheetOrTable.ws.name.label]]",
            description = "[[IterateWorksheetOrTable.ws.name.description]]")
    @Inject
    private String sheetName;

    @Idx(index = "1.2.2", type = AttributeType.SELECT, options = {
            @Idx.Option(index = "1.2.2.1", pkg = @Pkg(label = "[[IterateWorksheetOrTable.ws.rangeAll.label]]", value = "ALL")),
            @Idx.Option(index = "1.2.2.2", pkg = @Pkg(label = "[[IterateWorksheetOrTable.ws.rangeSpecific.label]]", value = "SPECIFIC"))
    })
    @Pkg(label = "[[IterateWorksheetOrTable.ws.rangeMode.label]]",
            description = "[[IterateWorksheetOrTable.ws.rangeMode.description]]",
            default_value = "ALL", default_value_type = DataType.STRING)
    @Inject @NotEmpty
    private String wsRangeMode;

    @Idx(index = "1.2.2.2.1", type = AttributeType.TEXT)
    @Pkg(label = "[[IterateWorksheetOrTable.ws.rangeText.label]]",
            description = "[[IterateWorksheetOrTable.ws.rangeText.description]]")
    @Inject @NotEmpty
    private String wsRangeA1;

    @Idx(index = "1.2.3", type = AttributeType.SELECT, options = {
            @Idx.Option(index = "1.2.3.1", pkg = @Pkg(label = "[[IterateWorksheetOrTable.rows.all.label]]", value = "ALL_ROWS")),
            @Idx.Option(index = "1.2.3.2", pkg = @Pkg(label = "[[IterateWorksheetOrTable.rows.specific.label]]", value = "SPECIFIC_ROWS"))
    })
    @Pkg(label = "[[IterateWorksheetOrTable.ws.rowsMode.label]]",
            description = "[[IterateWorksheetOrTable.ws.rowsMode.description]]",
            default_value = "ALL_ROWS", default_value_type = DataType.STRING)
    @SelectModes @Inject @NotEmpty
    private String wsRowsMode;

    @Idx(index = "1.2.3.2.1", type = AttributeType.NUMBER)
    @Pkg(label = "[[IterateWorksheetOrTable.rows.start.label]]",
            description = "[[IterateWorksheetOrTable.rows.start.description]]")
    @Inject
    private Double wsStartRowOneBased; // optional (1-based within data region)

    @Idx(index = "1.2.3.2.2", type = AttributeType.NUMBER)
    @Pkg(label = "[[IterateWorksheetOrTable.rows.end.label]]",
            description = "[[IterateWorksheetOrTable.rows.end.description]]")
    @Inject
    private Double wsEndRowOneBased;   // optional

    @Idx(index = "1.2.4", type = AttributeType.CHECKBOX)
    @Pkg(label = "[[IterateWorksheetOrTable.ws.hasHeader.label]]",
            description = "[[IterateWorksheetOrTable.ws.hasHeader.description]]",
            default_value = "True", default_value_type = DataType.BOOLEAN)
    @Inject @NotEmpty
    private Boolean wsHasHeader;

    // Read mode: visible vs raw
    @Idx(index = "2", type = AttributeType.RADIO, options = {
            @Idx.Option(index = "2.1", pkg = @Pkg(label = "[[IterateWorksheetOrTable.read.visible.label]]", value = "visible")),
            @Idx.Option(index = "2.2", pkg = @Pkg(label = "[[IterateWorksheetOrTable.read.value.label]]", value = "value"))
    })
    @Pkg(label = "[[IterateWorksheetOrTable.read.mode.label]]",
            description = "[[IterateWorksheetOrTable.read.mode.description]]",
            default_value = "visible", default_value_type = DataType.STRING)
    @Inject @NotEmpty
    private String readMode;

    @Idx(index = "2.1.1", type = AttributeType.HELP)
    @Pkg(label = "", description = "[[IterateWorksheetOrTable.read.visible.description]]",default_value_type = DataType.STRING) String visibleHelp;
    @Idx(index = "2.2.1", type = AttributeType.HELP)
    @Pkg(label = "", description = "[[IterateWorksheetOrTable.read.value.description]]",default_value_type = DataType.STRING) String valueHelp;

    // Session
    @Idx(index = "3", type = AttributeType.SESSION)
    @Pkg(label = "[[existingSession.label]]",
            description = "[[existingSession.description]]",
            default_value = "Default", default_value_type = DataType.SESSION)
    @SessionObject @Inject @NotEmpty
    private WorkbookSession session;

    // State
    private transient Workbook wb;
    private transient Sheet sheet;
    private transient DataFormatter formatter;
    private transient FormulaEvaluator evaluator;

    private transient int headerRow0;
    private transient int firstCol0;
    private transient int lastCol0;
    private transient int dataStartRow0;
    private transient int dataEndRow0;
    private transient int currentRow0;

//    private transient List<String> headerNames; // from buildHeadersInArea
    private transient List<Schema> recordSchema;      // A360 Record schema

    @HasNext
    public boolean hasNext() {
        ensureInitialized();
        return currentRow0 <= dataEndRow0;
    }

    @Next
    public RecordValue next() throws Exception {
        ensureInitialized();
        if (currentRow0 > dataEndRow0) {
            throw new Exception("No more rows available to iterate.");
        }

        // Build a Record for currentRow0
        List<Value> fields = new ArrayList<>(recordSchema.size());
//        List<Value> fields = new LinkedHashMap<>();
        for (int c = firstCol0; c <= lastCol0; c++) {
            String colName = recordSchema.get(c - firstCol0).getName();
            Value v = readCellAsValue(currentRow0, c);
            fields.add(v);
//            fields.put(colName, v);
        }

        Record rec = new Record();
        rec.setSchema(recordSchema);

        rec.setValues(fields);

        currentRow0++;

        RecordValue recordValue = new RecordValue();
        recordValue.set(rec);
        return recordValue;
    }

    // ---------------- Initialization ----------------

    private void ensureInitialized() {
        if (wb != null) return;

        if (session == null || session.getWorkbook() == null) {
            throw new BotCommandException("Workbook session not initialized. Open or create a workbook first.");
        }
        wb = session.getWorkbook();
        formatter = new DataFormatter(Locale.getDefault());
        evaluator = wb.getCreationHelper().createFormulaEvaluator();

        if ("TABLE".equalsIgnoreCase(mode)) {
            initFromTable();
        } else if ("WORKSHEET".equalsIgnoreCase(mode)) {
            initFromWorksheet();
        } else {
            throw new BotCommandException("Invalid mode. Choose TABLE or WORKSHEET.");
        }

        // Build header names via existing helper, then map to Record schema used by A360 RecordValue
        int columnCount = (lastCol0 - firstCol0 + 1);
        boolean hasHeader = ("TABLE".equalsIgnoreCase(mode)) || (wsHasHeader != null && wsHasHeader);

        // Reuse GetWorksheetAsTable.buildHeadersInArea to create header names consistently
        // Note: This returns a list of Schema objects defined in that command; we only need the names.
        recordSchema = buildHeadersInArea(sheet, hasHeader, headerRow0, firstCol0, columnCount, readMode, formatter, evaluator);
//        headerNames = new ArrayList<>(columnCount);
//        for (Object s : schemas) {
//            // Expect a simple getName() on the Schema type used by GetWorksheetAsTable
//            try {
//                String name = (String) s.getClass().getMethod("getName").invoke(s);
//                headerNames.add((name == null || name.trim().isEmpty()) ? "Column" + (headerNames.size() + 1) : name.trim());
//            } catch (Exception reflectionError) {
//                // Fallback: if helper Schema lacks getName, synthesize names
//                headerNames.add("Column" + (headerNames.size() + 1));
//            }
//        }

        // Build Record schema for A360 Record value (all STRING for broad compatibility)
//        recordSchema = schemas;

        // Start iterating at first data row
        currentRow0 = dataStartRow0;
    }

    private void initFromTable() {
        if (!(wb instanceof XSSFWorkbook)) {
            throw new BotCommandException("Table iteration requires an .xlsx workbook.");
        }
        if (tableName == null || tableName.trim().isEmpty()) {
            throw new BotCommandException("Table name is required.");
        }

        AreaReference ar = TableRange.getTableArea((XSSFWorkbook) wb, tableName.trim());
        if (ar == null) throw new BotCommandException("Table not found: " + tableName);
        CellReference tl = ar.getFirstCell();
        CellReference br = ar.getLastCell();
        if (tl == null || br == null) throw new BotCommandException("Invalid table area for: " + tableName);

        sheet = resolveSheetForTable((XSSFWorkbook) wb, tableName, tl);

//        sheet = (tlSheet != null) ? wb.getSheet(tlSheet) : wb.getSheetAt(wb.getActiveSheetIndex());
        if (sheet == null) {
            throw new BotCommandException("Failed to resolve worksheet for the specified table.");
        }

        headerRow0 = tl.getRow();
        firstCol0 = tl.getCol();
        lastCol0 = br.getCol();
        int tableLastRow = br.getRow();

        // Tables always have a header row at headerRow0
        dataStartRow0 = headerRow0 + 1;
        int baseEnd = Math.max(dataStartRow0 - 1, tableLastRow);

        int start = dataStartRow0;
        int end = baseEnd;

        if ("SPECIFIC_ROWS".equalsIgnoreCase(tableRowsMode)) {
            if (tableStartRowOneBased != null) {
                int s = toRowIndex("Start row", tableStartRowOneBased);
                start = dataStartRow0 + (s - 1);
            }
            if (tableEndRowOneBased != null) {
                int e = toRowIndex("End row", tableEndRowOneBased);
                end = dataStartRow0 + (e - 1);
            }
        }

        dataStartRow0 = Math.max(dataStartRow0, start);
        dataEndRow0 = Math.min(baseEnd, end);
        if (dataEndRow0 < dataStartRow0) dataEndRow0 = dataStartRow0 - 1;
    }

    private void initFromWorksheet() {
        if (sheetName == null || sheetName.trim().isEmpty()) {
            sheet = wb.getSheetAt(wb.getActiveSheetIndex());
//            throw new BotCommandException("Worksheet name is required.");
        } else {
            sheet = wb.getSheet(sheetName.trim());
        }
        if (sheet == null) {
            throw new BotCommandException("Worksheet not found: " + sheetName);
        }

        if ("ALL".equalsIgnoreCase(wsRangeMode)) {
            headerRow0 = sheet.getFirstRowNum();
            int lastRow = sheet.getLastRowNum();

            // Use existing helper to compute max columns across the used rows
            int maxCols = calculateMaxColumns(sheet, headerRow0, lastRow);
            firstCol0 = 0;
            lastCol0 = Math.max(0, maxCols - 1);

            boolean hasHeader = wsHasHeader != null && wsHasHeader;
            dataStartRow0 = hasHeader ? headerRow0 + 1 : headerRow0;

            int baseEnd = Math.max(dataStartRow0 - 1, lastRow);
            int start = dataStartRow0;
            int end = baseEnd;

            if ("SPECIFIC_ROWS".equalsIgnoreCase(wsRowsMode)) {
                if (wsStartRowOneBased != null) {
                    int s = toRowIndex("Start row", wsStartRowOneBased);
                    start = dataStartRow0 + (s - 1);
                }
                if (wsEndRowOneBased != null) {
                    int e = toRowIndex("End row", wsEndRowOneBased);
                    end = dataStartRow0 + (e - 1);
                }
            }

            dataStartRow0 = Math.max(dataStartRow0, start);
            dataEndRow0 = Math.min(baseEnd, end);

        } else if ("SPECIFIC".equalsIgnoreCase(wsRangeMode)) {
            if (wsRangeA1 == null || wsRangeA1.trim().isEmpty()) {
                throw new BotCommandException("Cell range is required for specific range.");
            }
            // Parse A1 range with CellReference on both ends (no extra helper required)
            String a1 = wsRangeA1.trim();
            String[] parts = a1.split(":");
            if (parts.length != 2) {
                throw new BotCommandException("Invalid A1 range: " + a1);
            }
            CellReference a = new CellReference(parts[0]);
            CellReference b = new CellReference(parts[1]);
            headerRow0 = Math.min(a.getRow(), b.getRow());
            firstCol0 = Math.min(a.getCol(), b.getCol());
            int r1 = Math.max(a.getRow(), b.getRow());
            lastCol0 = Math.max(a.getCol(), b.getCol());

            boolean hasHeader = wsHasHeader != null && wsHasHeader;
            dataStartRow0 = hasHeader ? headerRow0 + 1 : headerRow0;

            int baseEnd = Math.max(dataStartRow0 - 1, r1);
            int start = dataStartRow0;
            int end = baseEnd;

            if ("SPECIFIC_ROWS".equalsIgnoreCase(wsRowsMode)) {
                if (wsStartRowOneBased != null) {
                    int s = toRowIndex("Start row", wsStartRowOneBased);
                    start = dataStartRow0 + (s - 1);
                }
                if (wsEndRowOneBased != null) {
                    int e = toRowIndex("End row", wsEndRowOneBased);
                    end = dataStartRow0 + (e - 1);
                }
            }

            dataStartRow0 = Math.max(dataStartRow0, start);
            dataEndRow0 = Math.min(baseEnd, end);

        } else {
            throw new BotCommandException("Invalid worksheet range mode.");
        }

        if (dataEndRow0 < dataStartRow0) dataEndRow0 = dataStartRow0 - 1;
    }

    // ---------------- Cell reading (reuses existing helpers) ----------------

    private Value readCellAsValue(int r, int c) {
        Row row = sheet.getRow(r);
        Cell cell = (row == null) ? null : row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

        if ("visible".equalsIgnoreCase(readMode)) {
            String txt = readVisible(cell, formatter, evaluator);
            return new StringValue(txt == null ? "" : txt);
        } else {
            Object raw = readRaw(cell);
            if (raw == null) return new StringValue("");
            if (raw instanceof String)  return new StringValue((String) raw);
            if (raw instanceof Boolean) return new BooleanValue((Boolean) raw);
            if (raw instanceof Number)  return new NumberValue(((Number) raw).doubleValue());
            return new StringValue(String.valueOf(raw));
        }
    }

    // ---------------- Setters for injection ----------------

    public void setMode(String mode) { this.mode = mode; }
    public void setTableName(String tableName) { this.tableName = tableName; }
    public void setTableRowsMode(String tableRowsMode) { this.tableRowsMode = tableRowsMode; }
    public void setTableStartRowOneBased(Double tableStartRowOneBased) { this.tableStartRowOneBased = tableStartRowOneBased; }
    public void setTableEndRowOneBased(Double tableEndRowOneBased) { this.tableEndRowOneBased = tableEndRowOneBased; }
    public void setSheetName(String sheetName) { this.sheetName = sheetName; }
    public void setWsRangeMode(String wsRangeMode) { this.wsRangeMode = wsRangeMode; }
    public void setWsRangeA1(String wsRangeA1) { this.wsRangeA1 = wsRangeA1; }
    public void setWsRowsMode(String wsRowsMode) { this.wsRowsMode = wsRowsMode; }
    public void setWsStartRowOneBased(Double wsStartRowOneBased) { this.wsStartRowOneBased = wsStartRowOneBased; }
    public void setWsEndRowOneBased(Double wsEndRowOneBased) { this.wsEndRowOneBased = wsEndRowOneBased; }
    public void setWsHasHeader(Boolean wsHasHeader) { this.wsHasHeader = wsHasHeader; }
    public void setReadMode(String readMode) { this.readMode = readMode; }
    public void setSession(WorkbookSession session) { this.session = session; }
}
