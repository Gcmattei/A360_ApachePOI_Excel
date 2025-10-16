package com.davita.botcommand.excel.commands.dataOperations;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.BooleanValue;
import com.automationanywhere.botcommand.data.impl.NumberValue;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.data.model.Schema;
import com.automationanywhere.botcommand.data.model.record.Record;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.CommandPkg;
import com.automationanywhere.commandsdk.annotations.Execute;
import com.automationanywhere.commandsdk.annotations.Idx;
import com.automationanywhere.commandsdk.annotations.Pkg;
import com.automationanywhere.commandsdk.annotations.rules.FileExtension;
import com.automationanywhere.commandsdk.annotations.rules.LocalFile;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.AttributeType;

import com.davita.botcommand.excel.sessions.WorkbookSession;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.util.List;

@BotCommand
@CommandPkg(
        name = "appendRecordToSheet",
        label = "[[AppendRecordToSheet.label]]",
        node_label = "[[AppendRecordToSheet.node_label]]",
        description = "[[AppendRecordToSheet.description]]",
        icon = "excel-icon.svg"
)
public class AppendRecordToSheet {

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.FILE)
            @Pkg(label = "[[AppendRecordToSheet.file.label]]", description = "[[AppendRecordToSheet.file.description]]")
            @NotEmpty @LocalFile @FileExtension(value = "xlsx,xls")
            String workbookPath,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "[[AppendRecordToSheet.sheet.label]]", description = "[[AppendRecordToSheet.sheet.description]]")
            @NotEmpty
            String sheetName,

            @Idx(index = "3", type = AttributeType.CHECKBOX)
            @Pkg(label = "[[AppendRecordToSheet.writeHeader.label]]", description = "[[AppendRecordToSheet.writeHeader.description]]")
            @NotEmpty
            Boolean writeHeader,

            @Idx(index = "4", type = AttributeType.RECORD)
            @Pkg(label = "[[AppendRecordToSheet.record.label]]", description = "[[AppendRecordToSheet.record.description]]")
            @NotEmpty
            Record record
    ) {
        if (workbookPath == null || workbookPath.trim().isEmpty()) {
            throw new BotCommandException("Workbook path cannot be empty.");
        }
        if (sheetName == null || sheetName.trim().isEmpty()) {
            throw new BotCommandException("Sheet name cannot be empty.");
        }
        if (record == null) {
            throw new BotCommandException("Record cannot be null.");
        }
        File f = new File(workbookPath);
        if (!f.exists() || !f.isFile()) {
            throw new BotCommandException("File not found: " + workbookPath);
        }

        try {
            // Open via your session
            WorkbookSession session = WorkbookSession.openWorkbook(workbookPath,false);

            Workbook wb = session.getWorkbook();

            Sheet sheet = wb.getSheet(sheetName.trim());
            if (sheet == null) throw new BotCommandException("Worksheet not found: " + sheetName);

            // Compute append target
            int firstDataCol = findFirstDataColumn(sheet);
            int lastDataRow  = findLastDataRow(sheet);
            if (lastDataRow < 0) { lastDataRow = -1; firstDataCol = 0; }
            int targetRowIdx = lastDataRow + 1;
            Row row = sheet.getRow(targetRowIdx);
            if (row == null) row = sheet.createRow(targetRowIdx);

            // Write header if requested
            if (writeHeader) {
                List<Schema> schemas= record.getSchema();
                // Write left‑to‑right starting at firstDataCol
                int c = firstDataCol;
                for (Schema schema : schemas) {
                    Cell cell = row.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    writeValue(cell, new StringValue(schema.getName()));
                    c++;
                }
                targetRowIdx++;
                row = sheet.getRow(targetRowIdx);
                if (row == null) row = sheet.createRow(targetRowIdx);
            }

            List<Value> values= record.getValues();
            // Write left‑to‑right starting at firstDataCol
            int c = firstDataCol;
            for (Value value : values) {
                Cell cell = row.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                writeValue(cell, value);
                c++;
            }

            session.save();
            session.close();

        } catch (Exception ex) {
            throw (ex instanceof BotCommandException)
                    ? (BotCommandException) ex
                    : new BotCommandException("Failed to append record: " + ex.getMessage(), ex);
        }
    }

    // ---------- Helpers ----------

    private static void writeValue(Cell cell, Value v) {
        if (v == null) { cell.setBlank(); return; }
        if (v instanceof StringValue) {
            cell.setCellValue(((StringValue) v).get());
        } else if (v instanceof NumberValue) {
            Double d = ((NumberValue) v).get();
            cell.setCellValue(d == null ? 0d : d);
        } else if (v instanceof BooleanValue) {
            Boolean b = ((BooleanValue) v).get();
            cell.setCellValue(b != null && b);
        } else {
            Object o = v.get();
            if (o == null) cell.setBlank();
            else if (o instanceof Number)  cell.setCellValue(((Number) o).doubleValue());
            else if (o instanceof Boolean) cell.setCellValue((Boolean) o);
            else                            cell.setCellValue(String.valueOf(o));
        }
    }

    private static int findLastDataRow(Sheet sheet) {
        int last = sheet.getLastRowNum();
        for (int r = last; r >= 0; r--) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            short lastCell = row.getLastCellNum();
            if (lastCell < 0) continue;
            int firstCell = Math.max(0, row.getFirstCellNum());
            for (int c = firstCell; c < lastCell; c++) {
                Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell != null && !isBlankCell(cell)) return r;
            }
        }
        return -1;
    }

    private static int findFirstDataColumn(Sheet sheet) {
        int min = Integer.MAX_VALUE;
        int lastRow = sheet.getLastRowNum();
        for (int r = 0; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            short lastCell = row.getLastCellNum();
            if (lastCell < 0) continue;
            int firstCell = Math.max(0, row.getFirstCellNum());
            for (int c = firstCell; c < lastCell; c++) {
                Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell != null && !isBlankCell(cell)) {
                    if (c < min) min = c;
                    break;
                }
            }
        }
        return (min == Integer.MAX_VALUE) ? 0 : min;
    }

    private static boolean isBlankCell(Cell cell) {
        if (cell == null) return true;
        if (cell.getCellType() == CellType.BLANK) return true;
        if (cell.getCellType() == CellType.STRING) {
            String s = cell.getStringCellValue();
            return s == null || s.trim().isEmpty();
        }
        return false;
    }
}
