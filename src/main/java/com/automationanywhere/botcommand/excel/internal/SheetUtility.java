package com.automationanywhere.botcommand.excel.internal;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

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
    public static void deleteColumns(Sheet sheet, int firstCol, int lastCol) {
        if (sheet == null) return; // [2]
        if (lastCol < firstCol) return; // [2]

        int lastRowNum = sheet.getLastRowNum();
        if (firstCol < 0) firstCol = 0;

        // Remove merged regions intersecting the deleted columns
        removeMergedRegionsIntersecting(sheet, 0, Integer.MAX_VALUE, firstCol, lastCol); // [9][5]

        int width = lastCol - firstCol + 1;

        // Manual per-row left shift for portability across HSSF/XSSF
        for (int r = 0; r <= lastRowNum; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            int lastUsedCol = lastUsedColumnIndex(row);
            if (lastUsedCol < 0 || firstCol > lastUsedCol) continue;

            // Move cells to the left by 'width'
            for (int c = lastCol + 1; c <= lastUsedCol; c++) {
                Cell src = row.getCell(c);
                Cell dst = ensureCell(row, c - width, src);
                moveCellContent(src, dst);
                if (src != null) row.removeCell(src);
            }

            // Clear trailing columns
            int startTrailing = Math.max(lastUsedCol - width + 1, firstCol);
            for (int c = startTrailing; c <= lastUsedCol; c++) {
                clearCell(row, c);
            }
        }

        Workbook wb = sheet.getWorkbook();
        if (wb != null) wb.setForceFormulaRecalculation(true); // [7]
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
