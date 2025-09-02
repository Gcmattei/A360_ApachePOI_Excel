package com.automationanywhere.botcommand.excel.internal;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * JDK 11–compatible helpers for locating XLSX table ranges.
 * Requires an .xlsx workbook (XSSFWorkbook). Not applicable to .xls (HSSF).
 */
public final class TableRange {

    private TableRange() {}

    /**
     * Finds a table by name (case-insensitive) and returns its AreaReference.
     * Matches either XSSFTable.getName() or getDisplayName().
     *
     * @param wb        the XSSFWorkbook (xlsx only)
     * @param tableName the target table name
     * @return AreaReference covering the entire table (header to last data cell)
     * @throws IllegalArgumentException if workbook/table not found or no range could be determined
     */
    public static AreaReference getTableArea(XSSFWorkbook wb, String tableName) {
        if (wb == null) {
            throw new IllegalArgumentException("Workbook cannot be null.");
        }
        if (tableName == null || tableName.trim().isEmpty()) {
            throw new IllegalArgumentException("Table name cannot be empty.");
        }
        final String target = tableName.trim();

        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            XSSFSheet sheet = wb.getSheetAt(i);
            for (XSSFTable t : sheet.getTables()) {
                final String nm = t.getName();
                final String disp = t.getDisplayName();
                if ((nm != null && nm.equalsIgnoreCase(target)) ||
                        (disp != null && disp.equalsIgnoreCase(target))) {

                    // Try direct start/end references first
                    CellReference tl = t.getStartCellReference();
                    CellReference br = t.getEndCellReference();
                    if (tl != null && br != null) {
                        return new AreaReference(tl, br, SpreadsheetVersion.EXCEL2007);
                    }

                    // Try table-provided AreaReference
                    AreaReference ar = t.getCellReferences();
                    if (ar != null) {
                        return ar;
                    }

                    // As a fallback, refresh internal references, then retry
                    t.updateReferences();
                    tl = t.getStartCellReference();
                    br = t.getEndCellReference();
                    if (tl != null && br != null) {
                        return new AreaReference(tl, br, SpreadsheetVersion.EXCEL2007);
                    }

                    throw new IllegalArgumentException("Table found but range references are unavailable: " + target);
                }
            }
        }
        throw new IllegalArgumentException("Table not found: " + target);
    }

    /**
     * Returns the table range in A1 notation (e.g., A1:D200).
     *
     * @param wb        XSSFWorkbook
     * @param tableName table name
     * @return A1 range string
     */
    public static String getTableAreaA1(XSSFWorkbook wb, String tableName) {
        AreaReference ar = getTableArea(wb, tableName);
        CellReference first = ar.getFirstCell();
        CellReference last = ar.getLastCell();
        return first.formatAsString() + ":" + last.formatAsString();
    }

    /**
     * Returns the sheet that contains the given table, or null if not found.
     *
     * @param wb        XSSFWorkbook
     * @param tableName target table name
     * @return XSSFSheet or null
     */
    public static XSSFSheet getTableSheet(XSSFWorkbook wb, String tableName) {
        if (wb == null || tableName == null || tableName.trim().isEmpty()) return null;
        final String target = tableName.trim();

        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            XSSFSheet sheet = wb.getSheetAt(i);
            for (XSSFTable t : sheet.getTables()) {
                final String nm = t.getName();
                final String disp = t.getDisplayName();
                if ((nm != null && nm.equalsIgnoreCase(target)) ||
                        (disp != null && disp.equalsIgnoreCase(target))) {
                    return sheet;
                }
            }
        }
        return null;
    }
}
