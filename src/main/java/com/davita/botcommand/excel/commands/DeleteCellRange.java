package com.davita.botcommand.excel.commands;

import com.davita.botcommand.excel.internal.SheetUtility;
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

@BotCommand
@CommandPkg(
        name = "deleteCellRange",
        label = "[[DeleteCellRange.label]]",
        node_label = "[[DeleteCellRange.node_label]]",
        description = "[[DeleteCellRange.description]]",
        icon = "excel-icon.svg"
)
public class DeleteCellRange {

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "1.1", pkg = @Pkg(label = "[[DeleteCellRange.target.active.label]]", value = "ACTIVE")),
                    @Idx.Option(index = "1.2", pkg = @Pkg(label = "[[DeleteCellRange.target.specific.label]]", value = "SPECIFIC"))
            })
            @Pkg(label = "[[DeleteCellRange.target.label]]",
                    description = "[[DeleteCellRange.target.description]]",
                    default_value = "ACTIVE", default_value_type = DataType.STRING)
            @NotEmpty String targetMode,

            @Idx(index = "1.2.1", type = AttributeType.TEXT)
            @Pkg(label = "[[DeleteCellRange.cellOrRange.label]]",
                    description = "[[DeleteCellRange.cellOrRange.description]]")
            String cellOrRangeA1,

            @Idx(index = "2", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "[[DeleteCellRange.option.shiftLeft.label]]", value = "SHIFT_LEFT")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "[[DeleteCellRange.option.shiftUp.label]]", value = "SHIFT_UP")),
                    @Idx.Option(index = "2.3", pkg = @Pkg(label = "[[DeleteCellRange.option.entireRow.label]]", value = "ENTIRE_ROW")),
                    @Idx.Option(index = "2.4", pkg = @Pkg(label = "[[DeleteCellRange.option.entireCol.label]]", value = "ENTIRE_COLUMN"))
            })
            @Pkg(label = "[[DeleteCellRange.option.label]]",
                    description = "[[DeleteCellRange.option.description]]",
                    default_value = "SHIFT_LEFT", default_value_type = DataType.STRING)
            @NotEmpty String deleteOption,

            @Idx(index = "3", type = AttributeType.SESSION)
            @Pkg(label = "[[existingSession.label]]",
                    description = "[[existingSession.description]]",
                    default_value = "Default", default_value_type = DataType.SESSION)
            @SessionObject
            @NotEmpty WorkbookSession session
    ) {
        if (session == null || session.getWorkbook() == null) {
            throw new BotCommandException("Workbook session not initialized. Open or create a workbook before deleting."); // [1]
        }

        Workbook wb = session.getWorkbook();
        int activeSheetIdx = wb.getActiveSheetIndex();
        if (activeSheetIdx < 0 || activeSheetIdx >= wb.getNumberOfSheets()) {
            throw new BotCommandException("Active sheet is not set or out of range."); // [1]
        }
        Sheet sheet = wb.getSheetAt(activeSheetIdx);

        // Resolve region
        CellRangeAddress region;
        if ("ACTIVE".equalsIgnoreCase(targetMode)) {
            region = getActiveCellRegion(sheet);
        } else if ("SPECIFIC".equalsIgnoreCase(targetMode)) {
            if (cellOrRangeA1 == null || cellOrRangeA1.trim().isEmpty()) {
                throw new BotCommandException("Cell or range (A1) is required for Specific mode."); // [1]
            }
            region = parseA1Range(cellOrRangeA1.trim());
        } else {
            throw new BotCommandException("Invalid target mode. Choose ACTIVE or SPECIFIC."); // [1]
        }

        final int r0 = Math.max(region.getFirstRow(), 0);
        final int r1 = Math.max(region.getLastRow(), r0);
        final int c0 = Math.max(region.getFirstColumn(), 0);
        final int c1 = Math.max(region.getLastColumn(), c0);

        try {
            switch (deleteOption) {
                case "SHIFT_LEFT":
                    SheetUtility.deleteCellsInRangeAndShift(sheet, r0, r1, c0, c1, SheetUtility.ShiftDirection.LEFT);
                    break;
                case "SHIFT_UP":
                    SheetUtility.deleteCellsInRangeAndShift(sheet, r0, r1, c0, c1, SheetUtility.ShiftDirection.UP);
                    break;
                case "ENTIRE_ROW":
                    SheetUtility.deleteRows(sheet, r0, r1);
                    break;
                case "ENTIRE_COLUMN":
                    SheetUtility.deleteColumns(sheet, c0, c1);
                    break;
                default:
                    throw new BotCommandException("Invalid delete option: " + deleteOption);
            }
            wb.setForceFormulaRecalculation(true);
        } catch (IllegalStateException ise) {
            // Typical for array formula group issues
            throw new BotCommandException("Delete failed due to array formulas in the affected range: " + ise.getMessage(), ise); // [6]
        } catch (UnsupportedOperationException uoe) {
            throw new BotCommandException("Delete failed due to unsupported operation in this workbook/format: " + uoe.getMessage(), uoe); // [4][5]
        } catch (Exception e) {
            throw new BotCommandException("Failed to delete cells: " + e.getMessage(), e); // [2]
        }
    }

    // Accepts "A1" or "A1:C10"
    private CellRangeAddress parseA1Range(String a1) {
        if (a1.contains(":")) {
            try {
                return CellRangeAddress.valueOf(a1);
            } catch (IllegalArgumentException iae) {
                throw new BotCommandException("Invalid A1 range: " + a1, iae); // [1]
            }
        } else {
            CellReference ref = new CellReference(a1);
            int r = ref.getRow();
            int c = ref.getCol();
            return new CellRangeAddress(r, r, c, c);
        }
    }

    // Use the sheet's active cell as a stable selection fallback
    private CellRangeAddress getActiveCellRegion(Sheet sheet) {
        CellAddress addr = sheet.getActiveCell();
        if (addr == null) {
            return new CellRangeAddress(0, 0, 0, 0);
        }
        return new CellRangeAddress(addr.getRow(), addr.getRow(), addr.getColumn(), addr.getColumn());
    }
}
