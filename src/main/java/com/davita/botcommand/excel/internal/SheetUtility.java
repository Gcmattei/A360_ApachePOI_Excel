package com.davita.botcommand.excel.internal;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumns;

import java.util.ArrayList;
import java.util.List;

public final class SheetUtility {

    public enum ShiftDirection { LEFT, UP }

    private SheetUtility() {}

    // ===== Rows: delete a contiguous block [firstRow..lastRow] =====
    public static void deleteRows(Sheet sheet, int firstRow, int lastRow) {
        if (sheet == null) return; // [2]
        if (lastRow < firstRow) return; // [2]

        int lastRowNum = sheet.getLastRowNum();
        if (firstRow < 0) firstRow = 0;
        if (firstRow > lastRowNum) return;
        if (lastRow > lastRowNum) lastRow = lastRowNum;
        int n = lastRow - firstRow + 1;

        // Remove merged regions intersecting the deleted rows to avoid invalid merges
        removeMergedRegionsIntersecting(sheet, firstRow, lastRow, 0, Integer.MAX_VALUE); // [9][5]

        if (lastRow < lastRowNum) {
            // Shift rows below up by n; POI handles some structures during shift

            sheet.shiftRows(lastRow + 1, lastRowNum, -n, true, false); // [3]
        }

        // Clear trailing rows at the bottom
        int newLastRowNum = sheet.getLastRowNum();
        for (int r = Math.min(newLastRowNum + 1, lastRowNum - n + 1); r <= lastRowNum; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            clearRow(row);
            sheet.removeRow(row);
        }

        Workbook wb = sheet.getWorkbook();
        if (wb != null) wb.setForceFormulaRecalculation(true); // [7]
    }

    // ===== Columns: delete a contiguous block [firstCol..lastCol] =====
//    public static void deleteColumns(Sheet sheet, int firstCol, int lastCol) {
//        if (sheet == null) return; // [2]
//        if (lastCol < firstCol) return; // [2]
//
//        int lastRowNum = sheet.getLastRowNum();
//        if (firstCol < 0) firstCol = 0;
//
//        // Remove merged regions intersecting the deleted columns
//        removeMergedRegionsIntersecting(sheet, 0, Integer.MAX_VALUE, firstCol, lastCol); // [9][5]
//
//        int width = lastCol - firstCol + 1;
//
//        // Manual per-row left shift for portability across HSSF/XSSF
//        for (int r = 0; r <= lastRowNum; r++) {
//            Row row = sheet.getRow(r);
//            if (row == null) continue;
//            int lastUsedCol = lastUsedColumnIndex(row);
//            if (lastUsedCol < 0 || firstCol > lastUsedCol) continue;
//
//            // Move cells to the left by 'width'
//            for (int c = lastCol + 1; c <= lastUsedCol; c++) {
//                Cell src = row.getCell(c);
//                Cell dst = ensureCell(row, c - width, src);
//                moveCellContent(src, dst);
//                if (src != null) row.removeCell(src);
//            }
//
//            // Clear trailing columns
//            int startTrailing = Math.max(lastUsedCol - width + 1, firstCol);
//            for (int c = startTrailing; c <= lastUsedCol; c++) {
//                clearCell(row, c);
//            }
//        }
//
//        Workbook wb = sheet.getWorkbook();
//        if (wb != null) wb.setForceFormulaRecalculation(true); // [7]
//    }

    public static void deleteColumns(Sheet sheet, int firstCol, int lastColInclusive) {
        if (sheet == null) throw new IllegalArgumentException("Sheet cannot be null.");
        if (firstCol < 0 || lastColInclusive < firstCol) {
            throw new IllegalArgumentException("Invalid column range.");
        }
        final int removeCount = lastColInclusive - firstCol + 1;

        // Snapshot table areas BEFORE edits to compute new bounds later.
        List<TableSnapshot> snapshots = new ArrayList<TableSnapshot>();
        if (sheet instanceof XSSFSheet) {
            XSSFSheet xs = (XSSFSheet) sheet;
            for (XSSFTable tbl : xs.getTables()) {
                CTTable ct = tbl.getCTTable();
                AreaReference ar = tbl.getArea();
                if (ar == null && ct.getRef() != null) {
                    ar = new AreaReference(ct.getRef(), SpreadsheetVersion.EXCEL2007);
                }
                if (ar == null) continue;
                snapshots.add(new TableSnapshot(tbl, ar.getFirstCell(), ar.getLastCell()));
            }
        }

        // Manually shift every row’s cells left within the used range.
        int lastRow = sheet.getLastRowNum();
        for (int r = 0; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            short lastCellNum = row.getLastCellNum();
            if (lastCellNum < 0) continue;

            int lastUsed = lastCellNum - 1;
            if (firstCol > lastUsed) continue;

            // Copy cells from right to left: src -> dst
            for (int c = firstCol; c <= lastUsed; c++) {
                int srcCol = c + removeCount;
                Cell dst = row.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                if (srcCol <= lastUsed) {
                    Cell src = row.getCell(srcCol, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    copyCellValueAndStyle(src, dst);
                } else {
                    dst.setBlank();
                }
            }

            // Remove trailing cells on the right that are now outside the logical range
            int tailStart = Math.max(firstCol, lastUsed - removeCount + 1);
            for (int c = tailStart; c <= lastUsed; c++) {
                Cell tail = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (tail != null) row.removeCell(tail);
            }
        }

        // Update tables that intersect or lie to the right of the deleted band.
        if (sheet instanceof XSSFSheet) {
            XSSFSheet xs = (XSSFSheet) sheet;
            for (TableSnapshot snap : snapshots) {
                XSSFTable table = snap.table;
                CTTable ct = table.getCTTable();

                // Original table bounds
                int tFirstRow = snap.tl.getRow();
                int tLastRow  = snap.br.getRow();
                int tFirstCol = snap.tl.getCol();
                int tLastCol  = snap.br.getCol();

                // 1) How many deleted columns are STRICTLY before the table start?
                int countBefore = 0;
                if (firstCol < tFirstCol) {
                    int beforeEnd = Math.min(tFirstCol - 1, lastColInclusive);
                    if (beforeEnd >= firstCol) {
                        countBefore = (beforeEnd - firstCol + 1);
                    }
                }

                // 2) How many deleted columns overlap the table span?
                int ovStart = Math.max(firstCol, tFirstCol);
                int ovEnd   = Math.min(lastColInclusive, tLastCol);
                int overlap = (ovStart <= ovEnd) ? (ovEnd - ovStart + 1) : 0;

                // 3) New bounds after deletion:
                // - Start shifts left by countBefore
                // - Width shrinks by overlap
                int newFirstCol = tFirstCol - countBefore;
                int newLastCol  = tLastCol  - countBefore - overlap;

                // Clamp to 0 to avoid negative first column
                if (newFirstCol < 0) {
                    int deficit = -newFirstCol;
                    newFirstCol = 0;
                    newLastCol  = Math.max(newFirstCol, newLastCol - deficit);
                }

                // If zero-width, remove the table
                if (newLastCol < newFirstCol) {
                    // 1) Remove the table relation from the sheet
                    xs.removeTable(table);                                    // drops the part/relation [web:341]

                    // 2) Clean worksheet artifacts that may still reference the table
                    cleanupWorksheetTableArtifacts(xs);                       // tableParts + x14:tableParts + sheet AutoFilter [web:364][web:365][web:195]

                    // 3) Clean workbook defined name used by sheet-level AutoFilter (if present)
                    removeFilterDatabaseDefinedName(xs.getWorkbook(), xs);    // _xlnm._FilterDatabase for this sheet [web:369][web:373]

                    // Skip to next table
                    continue;
                }

                // 4) Apply the new area and rebuild columns from header
                AreaReference newArea = new AreaReference(
                        new CellReference(tFirstRow, newFirstCol),
                        new CellReference(tLastRow,  newLastCol),
                        SpreadsheetVersion.EXCEL2007
                );
                table.setArea(newArea);                                // update CTTable.ref via user model [web:341]
                rebuildTableColumnsFromHeader(xs, table, newFirstCol, newLastCol, tFirstRow); // ids 1..N, names from header [web:354]
                table.updateReferences();                              // sync caches [web:341]
                table.updateHeaders();                                 // sync header captions [web:341]

                rebuildSheetTableParts((org.apache.poi.xssf.usermodel.XSSFSheet) sheet);
            }
        }
    }

    // Remove CTWorksheet.tableParts and the x14 extension list's tableParts,
    // and unset a sheet-level AutoFilter if present.
    private static void cleanupWorksheetTableArtifacts(org.apache.poi.xssf.usermodel.XSSFSheet xs) {
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet cws = xs.getCTWorksheet();

        // Unset the standard tableParts list
        if (cws.isSetTableParts()) {
            cws.unsetTableParts();                                // remove r:id table pointers [web:364][web:365]
        }

        // If no tables remain on this sheet, it is safe to drop the entire extLst;
        // this also removes x14:tableParts which Excel checks in newer files.
        if (xs.getTables() == null || xs.getTables().isEmpty()) {
            if (cws.isSetExtLst()) {
                cws.unsetExtLst();                                // remove <extLst> (clears x14:tableParts) [web:365]
            }
            if (cws.isSetAutoFilter()) {
                cws.unsetAutoFilter();                            // clear stray sheet-level AutoFilter [web:195]
            }
        }
    }

    // Remove _xlnm._FilterDatabase defined name for the given sheet (if it exists).
    // Excel uses this name for sheet-level AutoFilter; when the table/range is gone, it must be removed.
    private static void removeFilterDatabaseDefinedName(Workbook wb, org.apache.poi.xssf.usermodel.XSSFSheet xs) {
        if (!(wb instanceof org.apache.poi.xssf.usermodel.XSSFWorkbook)) return;
        org.apache.poi.xssf.usermodel.XSSFWorkbook xwb = (org.apache.poi.xssf.usermodel.XSSFWorkbook) wb;
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook ctwb = xwb.getCTWorkbook();
        if (!ctwb.isSetDefinedNames()) return;

        int sheetIndex = xwb.getSheetIndex(xs);
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedNames dns = ctwb.getDefinedNames();

        java.util.List<org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedName> keep = new java.util.ArrayList<>();
        for (int i = 0; i < dns.sizeOfDefinedNameArray(); i++) {
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedName dn = dns.getDefinedNameArray(i);
            String nm = dn.getName();
            long localId = dn.isSetLocalSheetId() ? dn.getLocalSheetId() : -1L;

            // Drop only this sheet's _xlnm._FilterDatabase
            if ("_xlnm._FilterDatabase".equals(nm) && localId == sheetIndex) {
                continue;                                         // remove it [web:369][web:373]
            }
            keep.add(dn);
        }
        // Re-write list if anything was removed
        if (keep.size() != dns.sizeOfDefinedNameArray()) {
            dns.setDefinedNameArray(new org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedName[0]);
            for (org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedName dn : keep) {
                dns.addNewDefinedName().set(dn);
            }
            if (keep.isEmpty()) {
                ctwb.unsetDefinedNames();
            }
        }
    }

    private static void rebuildSheetTableParts(org.apache.poi.xssf.usermodel.XSSFSheet xs) {
        // Rebuild CTWorksheet.tableParts to match xs.getTables()
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet cws = xs.getCTWorksheet();

        // Remove existing tableParts entirely; we’ll recreate from the live relations
        if (cws.isSetTableParts()) {
            cws.unsetTableParts();
        }

        java.util.List<org.apache.poi.xssf.usermodel.XSSFTable> tables = xs.getTables();
        if (tables == null || tables.isEmpty()) {
            return; // no table parts to write
        }

        // Create a fresh parts list with exact count
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableParts parts =
                cws.addNewTableParts();
        parts.setCount(tables.size());

        for (org.apache.poi.xssf.usermodel.XSSFTable t : tables) {
            // Get the sheet->table relationship id so Excel can resolve the part
            String rid = xs.getRelationId(t);
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTablePart tp =
                    parts.addNewTablePart();
            tp.setId(rid);
        }
    }


    private static void copyCellValueAndStyle(Cell src, Cell dst) {
        if (src == null) {
            dst.setBlank();
            return;
        }
        try {
            dst.setCellStyle(src.getCellStyle());
        } catch (Exception ignore) {}
        switch (src.getCellType()) {
            case STRING:
                dst.setCellValue(src.getStringCellValue());
                break;
            case NUMERIC:
                dst.setCellValue(src.getNumericCellValue());
                break;
            case BOOLEAN:
                dst.setCellValue(src.getBooleanCellValue());
                break;
            case FORMULA:
                dst.setCellFormula(src.getCellFormula());
                break;
            case BLANK:
            default:
                dst.setBlank();
                break;
        }
    }

//    private static void rebuildTableColumnsFromHeader(XSSFSheet xs, XSSFTable table,
//                                                      int firstCol0, int lastCol0, int headerRow0) {
//        CTTable ct = table.getCTTable();
//        int width = lastCol0 - firstCol0 + 1;
//        if (width <= 0) return;
//
//        CTTableColumns newCols = CTTableColumns.Factory.newInstance();
//        newCols.setCount(width);
//
//        Row header = xs.getRow(headerRow0);
//        DataFormatter fmt = new DataFormatter();
//        FormulaEvaluator eval = xs.getWorkbook().getCreationHelper().createFormulaEvaluator();
//
//        for (int i = 0; i < width; i++) {
//            String name = "Column" + (i + 1);
//            if (header != null) {
//                Cell hc = header.getCell(firstCol0 + i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
//                if (hc != null) {
//                    String txt = fmt.formatCellValue(hc, eval);
//                    if (txt != null && !txt.trim().isEmpty()) name = txt.trim();
//                }
//            }
//            CTTableColumn col = newCols.addNewTableColumn();
//            col.setId(i + 1);           // 1..N
//            col.setName(name);          // header text or ColumnN
//        }
//        ct.setTableColumns(newCols);
//
//        // Keep CT refs in sync with table area (includes header/total rows).
//        AreaReference ar = table.getArea();
//        String a1 = ar.formatAsString();
//        ct.setRef(a1);
//        if (ct.isSetAutoFilter()) {
//            CTAutoFilter af = ct.getAutoFilter();
//            af.setRef(a1);
//        }
//    }

    private static void rebuildTableColumnsFromHeader(XSSFSheet xs, XSSFTable table,
                                                      int firstCol0, int lastCol0, int headerRow0) {
        CTTable ct = table.getCTTable();
        if (firstCol0 < 0) firstCol0 = 0;                 // guard
        int width = lastCol0 - firstCol0 + 1;
        if (width <= 0) return;

        CTTableColumns newCols = CTTableColumns.Factory.newInstance();
        newCols.setCount(width);

        Row header = xs.getRow(Math.max(0, headerRow0));  // guard
        DataFormatter fmt = new DataFormatter();
        FormulaEvaluator eval = xs.getWorkbook().getCreationHelper().createFormulaEvaluator();

        for (int i = 0; i < width; i++) {
            String name = "Column" + (i + 1);
            if (header != null) {
                int colIdx = firstCol0 + i;
                if (colIdx >= 0) {
                    Cell hc = header.getCell(colIdx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if (hc != null) {
                        String txt = fmt.formatCellValue(hc, eval);
                        if (txt != null && !txt.trim().isEmpty()) name = txt.trim();
                    }
                }
            }
            CTTableColumn col = newCols.addNewTableColumn();
            col.setId(i + 1);
            col.setName(name);
        }
        ct.setTableColumns(newCols);

        AreaReference ar = table.getArea();
        String a1 = ar.formatAsString();
        ct.setRef(a1);
        if (ct.isSetAutoFilter()) ct.getAutoFilter().setRef(a1);

        table.updateReferences();
        table.updateHeaders();
    }


    private static final class TableSnapshot {
        final XSSFTable table;
        final CellReference tl;
        final CellReference br;
        TableSnapshot(XSSFTable table, CellReference tl, CellReference br) {
            this.table = table; this.tl = tl; this.br = br;
        }
    }

    private static int maxUsedColumn(Sheet sheet) {
        int max = 0;
        int lastRow = sheet.getLastRowNum();
        for (int r = 0; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            short last = row.getLastCellNum();
            if (last > 0) max = Math.max(max, last - 1);
        }
        return max;
    }

    // ===== Delete rectangular block and shift LEFT (rows) or UP (columns) =====
    public static void deleteCellsInRangeAndShift(
            Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, ShiftDirection direction) {
        if (sheet == null) return; // [2]
        if (lastRow < firstRow || lastCol < firstCol) return; // [2]

        int lastRowNum = sheet.getLastRowNum();
        if (firstRow < 0) firstRow = 0;
        if (firstRow > lastRowNum) return;
        if (lastRow > lastRowNum) lastRow = lastRowNum;
        if (firstCol < 0) firstCol = 0;

        // Preflight: disallow touching array formula groups to avoid IllegalStateException mid-move
//        ensureNoArrayFormulaInRange(sheet, firstRow, lastRow, firstCol, lastCol); // [6]

        // Remove merged regions intersecting the block
        removeMergedRegionsIntersecting(sheet, firstRow, lastRow, firstCol, lastCol); // [9][5]

        switch (direction) {
            case LEFT:
                deleteBlockShiftLeft(sheet, firstRow, lastRow, firstCol, lastCol);
                break;
            case UP:
                deleteBlockShiftUp(sheet, firstRow, lastRow, firstCol, lastCol);
                break;
        }

        Workbook wb = sheet.getWorkbook();
        if (wb != null) wb.setForceFormulaRecalculation(true); // [7]
    }

    // --- LEFT: for each affected row, shift slice right of the block left by block width, then clear trailing cells
    private static void deleteBlockShiftLeft(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        int width = lastCol - firstCol + 1;

        for (int r = firstRow; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            int lastUsedCol = lastUsedColumnIndex(row);
            if (lastUsedCol < 0 || firstCol > lastUsedCol) continue;

            for (int c = lastCol + 1; c <= lastUsedCol; c++) {
                Cell src = row.getCell(c);
                Cell dst = ensureCell(row, c - width, src);
                moveCellContent(src, dst);
                if (src != null) row.removeCell(src);
            }

            int startTrailing = Math.max(lastUsedCol - width + 1, firstCol);
            for (int c = startTrailing; c <= lastUsedCol; c++) {
                clearCell(row, c);
            }
        }
    }

    // --- UP: for each affected column, pull cells below the block up by block height, then clear the bottom cells
    private static void deleteBlockShiftUp(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        int height = lastRow - firstRow + 1;
        int lastRowNum = sheet.getLastRowNum();

        for (int c = firstCol; c <= lastCol; c++) {
            for (int r = lastRow + 1; r <= lastRowNum; r++) {
                Row srcRow = sheet.getRow(r);
                Row dstRow = sheet.getRow(r - height);
                if (dstRow == null) dstRow = sheet.createRow(r - height);

                Cell src = (srcRow != null) ? srcRow.getCell(c) : null;
                Cell dst = ensureCell(dstRow, c, src);
                moveCellContent(src, dst);
                if (srcRow != null && src != null) srcRow.removeCell(src);
            }

            for (int r = Math.max(lastRowNum - height + 1, firstRow); r <= lastRowNum; r++) {
                Row bottom = sheet.getRow(r);
                if (bottom != null) clearCell(bottom, c);
            }
        }
    }

    // ---- Helpers ----

    private static void clearRow(Row row) {
        if (row == null) return; // [2]
        // Collect cells first to avoid concurrent modification
        List<Integer> cols = new ArrayList<>();
        short last = row.getLastCellNum();
        for (int c = 0; c < (last > 0 ? last : 0); c++) {
            cols.add(c);
        }
        for (int c : cols) {
            clearCell(row, c);
        }
    }

    private static int lastUsedColumnIndex(Row row) {
        short last = row.getLastCellNum();
        return last > 0 ? (last - 1) : -1; // [2]
    }

    private static Cell ensureCell(Row row, int col, Cell srcForStyle) {
        if (col < 0) col = 0;
        Cell dst = row.getCell(col);
        if (dst == null) dst = row.createCell(col);
        if (srcForStyle != null && srcForStyle.getCellStyle() != null) {
            dst.setCellStyle(srcForStyle.getCellStyle());
        }
        return dst;
    }

    private static void moveCellContent(Cell src, Cell dst) {
        if (dst == null) return; // [2]
        if (src == null) {
            if (dst.getCellComment() != null) dst.removeCellComment();
            if (dst.getHyperlink() != null) dst.removeHyperlink();
            dst.setBlank();
            return;
        }

        // Comments
        if (src.getCellComment() != null) {
            Comment cm = src.getCellComment();
            try {
                cm.setAddress(new CellAddress(dst.getRowIndex(), dst.getColumnIndex()));
            } catch (Throwable ignore) {
            }
            dst.setCellComment(cm);
            src.removeCellComment();
        } else if (dst.getCellComment() != null) {
            dst.removeCellComment();
        }

        // Hyperlinks
        if (src.getHyperlink() != null) {
            Hyperlink old = src.getHyperlink();
            CreationHelper ch = dst.getSheet().getWorkbook().getCreationHelper();
            Hyperlink repl = ch.createHyperlink(old.getType());
            repl.setAddress(old.getAddress());
            dst.setHyperlink(repl);
            src.removeHyperlink();
        } else if (dst.getHyperlink() != null) {
            dst.removeHyperlink();
        }

        // Copy value/formula
        switch (src.getCellType()) {
            case FORMULA:
                // Do not attempt to split/move array formula groups
                if (src.isPartOfArrayFormulaGroup()) {
                    throw new IllegalStateException("Cell " + new CellAddress(src) + " is part of an array formula group."); // [6]
                }
                dst.setCellFormula(src.getCellFormula());
                break;
            case STRING:
                dst.setCellValue(src.getStringCellValue());
                break;
            case NUMERIC:
                dst.setCellValue(src.getNumericCellValue());
                break;
            case BOOLEAN:
                dst.setCellValue(src.getBooleanCellValue());
                break;
            case ERROR:
                dst.setCellErrorValue(src.getErrorCellValue());
                break;
            case BLANK:
            default:
                dst.setBlank();
                break;
        }
    }

    private static void clearCell(Row row, int col) {
        Cell c = row.getCell(col);
        if (c == null) return; // [2]

        if (c.getCellComment() != null) c.removeCellComment();
        if (c.getHyperlink() != null) c.removeHyperlink();

        if (c.getCellType() == CellType.FORMULA) {
            if (c.isPartOfArrayFormulaGroup()) {
                throw new IllegalStateException("Cannot clear cell " + new CellAddress(c) + " inside an array formula group."); // [6]
            }
            // Remove the formula to avoid keeping a stale cached result
            try {
                c.removeFormula();
                c.setBlank();
            } catch (Throwable t) {
                // Fallback for older POI without removeFormula
                c.setCellFormula(null);
            }
        } else {
            c.setBlank();
        }

        row.removeCell(c);
    }

    private static void removeMergedRegionsIntersecting(Sheet sheet, int r0, int r1, int c0, int c1) {
        int count = sheet.getNumMergedRegions();
        for (int i = count - 1; i >= 0; i--) {
            CellRangeAddress cra = sheet.getMergedRegion(i);
            if (rangesIntersect(r0, r1, c0, c1, cra)) {
                sheet.removeMergedRegion(i);
            }
        }
    }

    private static boolean rangesIntersect(int r0, int r1, int c0, int c1, CellRangeAddress cra) {
        int rr0 = cra.getFirstRow();
        int rr1 = cra.getLastRow();
        int cc0 = cra.getFirstColumn();
        int cc1 = cra.getLastColumn();
        boolean rowsOverlap = rr0 <= r1 && rr1 >= r0;
        boolean colsOverlap = cc0 <= c1 && cc1 >= c0;
        return rowsOverlap && colsOverlap;
    }

}
