package com.davita.botcommand.excel.commands;

import com.davita.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.CommandPkg;
import com.automationanywhere.commandsdk.annotations.Execute;
import com.automationanywhere.commandsdk.annotations.Idx;
import com.automationanywhere.commandsdk.annotations.Pkg;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;

@BotCommand
@CommandPkg(
        name = "setActiveCell",
        label = "[[SetActiveCell.label]]",
        node_label = "[[SetActiveCell.node_label]]",
        description = "[[SetActiveCell.description]]",
        icon = "excel-icon.svg"
)
public class SetActiveCell {

    @Execute
    public void action(
            // 1) Mode: Specific Cell vs Active Cell (relative)
            @Idx(index = "1", type = AttributeType.RADIO, options = {
                    @Idx.Option(index = "1.1", pkg = @Pkg(label = "[[SetActiveCell.mode.specific.label]]", value = "specific")),
                    @Idx.Option(index = "1.2", pkg = @Pkg(label = "[[SetActiveCell.mode.active.label]]", value = "active"))
            })
            @Pkg(
                    label = "[[SetActiveCell.mode.label]]",
                    description = "[[SetActiveCell.mode.description]]",
                    default_value = "specific",
                    default_value_type = DataType.STRING
            )
            @NotEmpty String mode,

            // 1.1.1) Specific cell or range (e.g., A1 or A1:B3 -> will use A1)
            @Idx(index = "1.1.1", type = AttributeType.TEXT)
            @Pkg(
                    label = "[[SetActiveCell.cellOrRange.label]]",
                    description = "[[SetActiveCell.cellOrRange.description]]"
            )
            @NotEmpty String cellOrRangeA1,

            // 1.2.1) Relative move from current active cell (dropdown)
            @Idx(index = "1.2.1", type = AttributeType.SELECT, options = {
                    @Idx.Option(index = "1.2.1.1", pkg = @Pkg(label = "[[SetActiveCell.relative.left]]", value = "LEFT")),
                    @Idx.Option(index = "1.2.1.2", pkg = @Pkg(label = "[[SetActiveCell.relative.right]]", value = "RIGHT")),
                    @Idx.Option(index = "1.2.1.3", pkg = @Pkg(label = "[[SetActiveCell.relative.up]]", value = "UP")),
                    @Idx.Option(index = "1.2.1.4", pkg = @Pkg(label = "[[SetActiveCell.relative.down]]", value = "DOWN")),
                    @Idx.Option(index = "1.2.1.5", pkg = @Pkg(label = "[[SetActiveCell.relative.beginRow]]", value = "BEGIN_ROW")),
                    @Idx.Option(index = "1.2.1.6", pkg = @Pkg(label = "[[SetActiveCell.relative.endRow]]", value = "END_ROW")),
                    @Idx.Option(index = "1.2.1.7", pkg = @Pkg(label = "[[SetActiveCell.relative.beginCol]]", value = "BEGIN_COL")),
                    @Idx.Option(index = "1.2.1.8", pkg = @Pkg(label = "[[SetActiveCell.relative.endCol]]", value = "END_COL"))
            })
            @Pkg(
                    label = "[[SetActiveCell.relative.label]]",
                    description = "[[SetActiveCell.relative.description]]",
                    default_value = "RIGHT",
                    default_value_type = DataType.STRING
            )
            @NotEmpty String relativeMove,

            // 2) Existing session
            @Idx(index = "2", type = AttributeType.SESSION)
            @Pkg(
                    label = "[[existingSession.label]]",
                    description = "[[existingSession.description]]",
                    default_value = "Default",
                    default_value_type = DataType.SESSION
            )
            @SessionObject
            @NotEmpty WorkbookSession session
    ) {
        // Validate session
        if (session == null || session.getWorkbook() == null) {
            throw new BotCommandException("Workbook session not initialized. Please open or create a workbook first.");
        }
        Workbook wb = session.getWorkbook();
        int activeSheetIdx = wb.getActiveSheetIndex();
        if (activeSheetIdx < 0 || activeSheetIdx >= wb.getNumberOfSheets()) {
            throw new BotCommandException("No active worksheet found in the workbook.");
        }
        Sheet sheet = wb.getSheetAt(activeSheetIdx);

        // Compute target CellAddress
        CellAddress target;
        if ("specific".equalsIgnoreCase(mode)) {
            if (cellOrRangeA1 == null || cellOrRangeA1.trim().isEmpty()) {
                throw new BotCommandException("Cell or range must not be empty when 'Specific Cell' is selected.");
            }
            target = resolveTopLeftAddress(cellOrRangeA1.trim(), wb);
        } else if ("active".equalsIgnoreCase(mode)) {
            CellAddress current = sheet.getActiveCell();
            if (current == null) {
                current = new CellAddress(0, 0); // default to A1 if none set
            }
            target = moveRelative(sheet, wb, current, relativeMove);
        } else {
            throw new BotCommandException("Invalid mode. Choose either 'Specific Cell' or 'Active Cell'.");
        }

        // Apply
        sheet.setActiveCell(target);
    }

    // Resolve A1 or A1:B5 to the top-left cell safely
    private CellAddress resolveTopLeftAddress(String ref, Workbook wb) {
        try {
            if (ref.contains(":")) {
                // Range like Sheet1!A1:B5 -> use top-left
                AreaReference area = new AreaReference(ref, wb.getSpreadsheetVersion());
                CellReference first = area.getFirstCell();
                return new CellAddress(first); // <-- fix
            } else {
                // Single reference like A1, $A$1, or Sheet1!A1
                CellReference cr = new CellReference(ref);
                return new CellAddress(cr); // <-- fix
            }
        } catch (IllegalArgumentException e) {
            throw new BotCommandException("Invalid cell or range reference: " + ref);
        }
    }


    // Compute relative moves bounded by the workbook’s spreadsheet version limits and existing content
    private CellAddress moveRelative(Sheet sheet, Workbook wb, CellAddress current, String move) {
        SpreadsheetVersion ver = wb.getSpreadsheetVersion();
        int maxRow = ver.getLastRowIndex();
        int maxCol = ver.getLastColumnIndex();

        int row = current.getRow();
        int col = current.getColumn();

        if (move == null || move.trim().isEmpty()) {
            throw new BotCommandException("Relative move option must be selected for 'Active Cell' mode.");
        }

        Row r;

        switch (move.toUpperCase()) {
            case "LEFT":
                col = Math.max(0, col - 1);
                break;
            case "RIGHT":
                col = Math.min(maxCol, col + 1);
                break;
            case "UP":
                row = Math.max(0, row - 1);
                break;
            case "DOWN":
                row = Math.min(maxRow, row + 1);
                break;
            case "BEGIN_ROW":
                r = sheet.getRow(row);
                if (r != null && r.getFirstCellNum() > 0) {
                    col = Math.max(0, r.getFirstCellNum());
                } else {
                    col = 0;
                }
                break;
            case "END_ROW": {
                r = sheet.getRow(row);
                if (r != null && r.getLastCellNum() > 0) {
                    col = Math.min(maxCol, r.getLastCellNum() - 1);
                } else {
                    col = 0;
                }
                break;
            }
            case "BEGIN_COL": {
                int firstRow = Math.max(0, sheet.getFirstRowNum());
                int lastRow = Math.max(firstRow, sheet.getLastRowNum());
                int candidate = -1;
                for (int i = firstRow; i <= lastRow; i++) {
                    Row rr = sheet.getRow(i);
                    if (rr != null && rr.getCell(col) != null) {
                        candidate = i;
                        break;
                    }
                }
                row = Math.max(candidate, 0);
                break;
            }
            case "END_COL": {
                int firstRow = Math.max(0, sheet.getFirstRowNum());
                int lastRow = Math.max(firstRow, sheet.getLastRowNum());
                int candidate = -1;
                for (int i = lastRow; i >= firstRow; i--) {
                    Row rr = sheet.getRow(i);
                    if (rr != null && rr.getCell(col) != null) {
                        candidate = i;
                        break;
                    }
                }
                row = Math.max(candidate, 0);
                break;
            }
            default:
                throw new BotCommandException("Unsupported relative move option: " + move);
        }

        // Bound checks
        row = Math.max(0, Math.min(maxRow, row));
        col = Math.max(0, Math.min(maxCol, col));
        return new CellAddress(row, col);
    }
}
