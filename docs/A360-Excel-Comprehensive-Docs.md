# Excel Commands Documentation

This comprehensive documentation covers all available commands and iterators in the A360 Apache POI Excel package. This is designed as an in-depth reference for troubleshooting and advanced usage scenarios in Automation Anywhere.

## Table of Contents

- [Session Management](#session-management)
- [Workbook Operations](#workbook-operations)
- [Worksheet Operations](#worksheet-operations)
- [Cell Operations](#cell-operations)
- [Data Operations](#data-operations)
- [Utility Operations](#utility-operations)
- [Iterator Operations](#iterator-operations)

---

## Session Management

### Open

**Purpose**: Creates a connection to an existing Excel workbook file and establishes a session for all subsequent Excel operations. This command is the entry point for working with existing spreadsheets and must be used before any other Excel operations can be performed. The session acts as a handle that maintains the workbook state, file locks, and provides context for all other commands in the package.

When you open a workbook, the system creates a file lock to prevent other processes from modifying the file while your automation is running. This ensures data integrity and prevents conflicts. The read-only option is particularly useful when you need to extract data from files that should not be modified, or when multiple automations need to access the same file simultaneously.

**Inputs**:
- `Workbook file path` (FILE, required): Full path to the Excel file. Must be an existing .xlsx (Excel 2007+) or .xls (Excel 97-2003) file on the local file system or accessible network drive. UNC paths are supported.
- `Open as Read-Only` (CHECKBOX, optional): When enabled, opens the workbook in read-only mode, preventing any modifications from being saved back to the original file. This is useful for data extraction scenarios or when working with protected/shared files. Default: false.

**Outputs**:
- Returns a `SESSION` object that serves as the handle for all subsequent Excel operations

**Limitations**:
- Only supports Microsoft Excel formats (.xlsx and .xls). Other spreadsheet formats like OpenDocument (.ods) or CSV files are not supported
- File must be physically accessible on the local machine or mapped network drive
- Creates an exclusive file lock that may prevent other applications from modifying the file
- Read-only mode completely prevents any save operations, even through Save As commands
- Large files (>100MB) may require significant memory and longer opening times
- Password-protected files are not currently supported

---

### Close

**Purpose**: Properly terminates the Excel workbook session and releases all associated resources including file locks and memory. This command ensures that any pending changes can be optionally saved before the session is closed. Proper session cleanup is critical to prevent file locks from remaining active and to free up system resources.

The save option allows you to persist any modifications made during the automation without requiring a separate Save command. This is particularly useful at the end of data processing workflows where you want to ensure changes are committed. If the workbook was opened in read-only mode, the save operation will be skipped even if requested.

**Inputs**:
- `Save changes on close` (CHECKBOX, optional): When enabled, automatically saves any pending modifications to the original file path before closing the session. This provides a convenient way to commit changes without requiring a separate Save command. The operation will fail if the session was opened in read-only mode. Default: false.
- `Session name` (SESSION, required): The session identifier that was returned from the Open or Create workbook command

**Limitations**:
- Cannot save changes if the original session was opened in read-only mode - the save request will be silently ignored
- Once closed, the session becomes completely invalid and cannot be reused
- Any unsaved changes are permanently lost if the save option is disabled
- Attempting to use other Excel commands with a closed session will result in errors

---

## Workbook Operations

### Create workbook

**Purpose**: Generates a new Excel workbook in memory and establishes a session for subsequent operations. This command is used when you need to build spreadsheets from scratch rather than modifying existing files. The workbook is created in the modern .xlsx format with support for all Excel features including formulas, formatting, and multiple worksheets.

Unlike opening an existing file, creating a workbook gives you complete control over the structure and content from the beginning. The workbook exists only in memory until explicitly saved, which makes this approach suitable for temporary data processing or when generating reports that may not always need to be persisted.

**Inputs**:
- `File path` (FILE, required): Target location where the workbook will be saved. The file extension (.xlsx) will determine the format. Parent directories will be automatically created if they don't exist. This path is stored with the session for use with Save operations.
- `Sheet name` (TEXT, optional): Name for the initial worksheet that is automatically created. If not specified, a default name will be used. This sheet becomes the active sheet for immediate operations.
- `Session name` (SESSION, required): Session identifier that will be used to reference this workbook in subsequent commands

**Outputs**:
- Returns a `SESSION` object that serves as the handle for all subsequent Excel operations on this new workbook

**Limitations**:
- Always creates workbooks in .xlsx format regardless of file extension specified
- The workbook exists only in memory until explicitly saved, making it vulnerable to data loss if the automation fails
- Automatically creates one default worksheet - empty workbooks without sheets are not supported
- File path must be writable and the parent directory must exist or be creatable

---

### Save workbook

**Purpose**: Commits all modifications made to the workbook back to its original file path. This command is essential for persisting changes made during automation and ensuring data is not lost when the session is closed. The save operation overwrites the original file with all current changes.

This command is most commonly used in workflows where you open an existing file, make modifications (adding data, formulas, formatting), and then need to preserve those changes. It's more efficient than Save As when you don't need to change the file location.

**Inputs**:
- `Session name` (SESSION, required): Active workbook session that contains the changes to be saved

**Limitations**:
- Cannot be used if the session was opened in read-only mode - will throw an error
- Requires that the original file path is available (not applicable for newly created workbooks without a prior save)
- Overwrites the existing file completely, which cannot be undone
- May fail if the original file is locked by another application or if disk space is insufficient

---

### Save as

**Purpose**: Creates a copy of the current workbook at a specified location while preserving the original file. This operation is particularly useful for creating backups, generating reports with different names, or moving workbooks to different directories. After saving, the session can optionally be updated to point to the new location, effectively switching your working context to the new file.

The save as operation automatically converts the session from read-only to read-write mode, allowing subsequent modifications to be saved. This makes it useful for scenarios where you open a template or protected file in read-only mode, make changes, and then save to a new location for editing.

**Inputs**:
- `File path` (FILE, required): Destination path for the workbook copy. Parent directories will be created automatically if they don't exist. File extension determines the output format (.xlsx or .xls).
- `Replace existing file` (CHECKBOX, optional): When enabled, overwrites any existing file at the destination path without warning. When disabled, the operation will fail if a file already exists at the target location.
- `Update session to new path` (CHECKBOX, optional): When enabled, changes the session's active file path to the new location, making subsequent Save operations target the new file instead of the original. This effectively "moves" your working context to the new file.
- `Session name` (SESSION, required): Active workbook session to be saved

**Limitations**:
- Output format is determined by file extension - .xlsx for modern format, .xls for legacy format
- Cannot save to read-only locations or when insufficient disk space is available
- The replace file option provides no recovery mechanism if you accidentally overwrite important files
- Large workbooks may take significant time to save and require substantial temporary disk space

---

## Worksheet Operations

### Create worksheet

**Purpose**: Adds a new worksheet (tab) to an existing workbook, expanding the available workspace for data organization. This command is essential for creating multi-sheet workbooks where different types of data or calculations need to be separated into logical groupings. The new worksheet is created with default Excel settings and becomes available immediately for data operations.

Worksheets provide a way to organize related data within a single workbook file. Common use cases include separating raw data from calculated results, creating summary sheets, or organizing data by time periods, departments, or categories.

**Inputs**:
- `Name for new worksheet` (TEXT, required): Display name for the new worksheet tab. Must be unique within the workbook and follow Excel naming conventions (no invalid characters like : \ / ? * [ ])
- `Session name` (SESSION, required): Active workbook session where the worksheet will be added

**Limitations**:
- Worksheet names must be unique within the workbook - attempting to create duplicate names will fail
- Names cannot exceed Excel's 31-character limit
- Cannot contain the characters: : \ / ? * [ ] which are reserved by Excel
- Excel has a maximum limit of 255 worksheets per workbook, though practical limits may be lower based on content and memory

---

### Delete worksheet

**Purpose**: Permanently removes a worksheet from the workbook along with all its data, formulas, and formatting. This operation cannot be undone and is useful for cleaning up unnecessary sheets or removing temporary worksheets created during processing. Care should be taken as this operation will break any formulas in other sheets that reference the deleted worksheet.

The command provides flexibility by allowing you to specify the target worksheet either by its position number (useful when dealing with predictable sheet structures) or by its exact name (useful when working with dynamically named sheets).

**Inputs**:
- `Delete sheet by` (RADIO, required): Selection method for the target worksheet - either "Index" for position-based selection or "Name" for name-based selection
- `Sheet index` (NUMBER, conditional): Position of the worksheet to delete, using 1-based numbering where 1 represents the first (leftmost) sheet. Only required when "Index" method is selected.
- `Sheet name` (TEXT, conditional): Exact name of the worksheet to delete, case-sensitive. Only required when "Name" method is selected.
- `Session name` (SESSION, required): Active workbook session containing the worksheet to be deleted

**Index Convention**: Sheet indexes are **1-based** (1 = first sheet, 2 = second sheet, etc.)

**Limitations**:
- Cannot delete the last remaining worksheet in a workbook - Excel requires at least one sheet
- Operation cannot be undone - all data and formulas on the deleted sheet are permanently lost
- Formulas in other sheets that reference the deleted worksheet will result in #REF! errors
- Sheet names are case-sensitive and must match exactly when using name-based deletion

---

### Rename worksheet

**Purpose**: Changes the display name of an existing worksheet tab without affecting its content or position. This operation is useful for organizing workbooks with descriptive names, updating sheet names based on processing results, or standardizing naming conventions across multiple workbooks. The rename operation preserves all data, formulas, and formatting while only changing the tab label.

Formula references to the renamed sheet are automatically updated by Excel to use the new name, maintaining the integrity of cross-sheet calculations. This makes renaming safe from a data perspective, though it may affect any external references to the workbook.

**Inputs**:
- `Select worksheet by` (RADIO, required): Method for identifying the worksheet to rename - either "Index" for position-based selection or "Name" for current name-based selection
- `Original sheet index` (NUMBER, conditional): Current position of the worksheet using 1-based numbering. Only required when "Index" method is selected.
- `Original sheet name` (TEXT, conditional): Current exact name of the worksheet, case-sensitive. Only required when "Name" method is selected.
- `Enter new name for worksheet` (TEXT, required): New display name for the worksheet, must be unique within the workbook and follow Excel naming conventions
- `Session name` (SESSION, required): Active workbook session containing the worksheet to be renamed

**Index Convention**: Sheet indexes are **1-based** for user input

**Limitations**:
- New names must be unique within the workbook
- Names cannot exceed 31 characters or contain invalid characters (: \ / ? * [ ])
- Original sheet identification must be exact - names are case-sensitive
- External workbook references may be broken if they rely on the old sheet name

---

### Switch to sheet

**Purpose**: Changes the active worksheet within the workbook, setting the context for subsequent cell and data operations. This command is crucial for multi-sheet workflows where you need to perform operations on different worksheets in sequence. The active sheet determines where commands like "Get single cell" or "Set cell/range" will operate when using "Active" target modes.

Switching between sheets is essential for complex data processing workflows that involve reading from one sheet, processing the data, and writing results to another sheet. This command provides the navigation mechanism needed to orchestrate such multi-step processes.

**Inputs**:
- `Activate sheet by` (RADIO, required): Method for identifying the target worksheet - either "Name" for name-based selection or "Index" for position-based selection
- `Sheet name` (TEXT, conditional): Exact name of the worksheet to activate, case-sensitive. Only required when "Name" method is selected.
- `Sheet index` (NUMBER, conditional): Position of the worksheet using 1-based numbering where 1 is the leftmost sheet. Only required when "Index" method is selected.
- `Session name` (SESSION, required): Active workbook session containing the worksheet to be activated

**Index Convention**: Sheet indexes are **1-based** (1 = first sheet, 2 = second sheet, etc.)

**Limitations**:
- Sheet index must be within the valid range (1 to total number of sheets)
- Sheet names must match exactly and are case-sensitive
- Switching to non-existent sheets will cause the command to fail
- The active sheet setting affects all subsequent operations that use "Active" target modes

---

### Get current worksheet name

**Purpose**: Retrieves the display name of the currently active worksheet and stores it in a string variable. This command is valuable for dynamic workflows where the active sheet may change during processing, and you need to track or log which sheet is currently being processed. It's also useful for conditional logic that needs to behave differently based on the current worksheet context.

This information can be used for logging, creating dynamic file names, or implementing conditional processing logic that adapts based on which sheet is active.

**Inputs**:
- `Session name` (SESSION, required): Active workbook session from which to retrieve the current worksheet name

**Outputs**:
- Returns a `STRING` containing the display name of the currently active worksheet

**Limitations**:
- Returns null or empty string if no worksheet is currently active (rare but possible in corrupted workbooks)
- The returned name reflects the current state and may change if other operations switch the active sheet

---

### Get worksheet names

**Purpose**: Extracts the display names of all worksheets in the workbook and returns them as a list collection. This command is essential for dynamic processing scenarios where you need to iterate through all sheets, perform operations based on sheet names, or create summary reports that include information from multiple worksheets. The names are returned in their tab order from left to right.

This functionality enables data processing workflows that can adapt to workbooks with varying numbers of sheets or different sheet configurations, making automations more robust and flexible.

**Inputs**:
- `Session name` (SESSION, required): Active workbook session from which to retrieve all worksheet names

**Outputs**:
- Returns a `LIST` of strings containing all worksheet names in their tab order

**Index Convention**: Names are returned in sheet order (left to right as they appear in Excel tabs)

**Limitations**:
- Returns an empty list for workbooks with no sheets (which shouldn't occur in valid Excel files)
- The order reflects the current tab arrangement and may change if sheets are moved
- Hidden sheets are included in the results - there's no filtering for visibility

---

### Get number of rows

**Purpose**: Determines the extent of data in a specified worksheet by counting rows that contain any data. This command is crucial for dynamic data processing where the amount of data may vary between runs, and you need to determine loop boundaries or validate data completeness. The count reflects actual usage rather than theoretical maximums.

Understanding data boundaries is essential for efficient processing and preventing infinite loops when iterating through data. This command provides the intelligence needed to adapt processing to actual data volumes.

**Inputs**:
- `Select worksheet by` (RADIO, required): Method for identifying the target worksheet - either "Index" for position-based selection or "Name" for name-based selection
- `Sheet index` (NUMBER, conditional): Position of the target worksheet using 1-based numbering. Only required when "Index" method is selected.
- `Sheet name` (TEXT, conditional): Exact name of the target worksheet, case-sensitive. Only required when "Name" method is selected.
- `Count mode` (RADIO, required): Counting method - "Non-empty rows" counts only rows with visible data, "Total rows with data" includes rows that may appear empty but contain formulas or formatting
- `Session name` (SESSION, required): Active workbook session containing the target worksheet

**Outputs**:
- Returns a `NUMBER` representing the count of rows containing data

**Index Convention**: Returns actual count (not zero-based index); sheet indexes are **1-based**

**Limitations**:
- Count is based on Apache POI's assessment of data presence, which may not match Excel's visual assessment
- Empty rows in the middle of data ranges may not be counted consistently
- Very large datasets may impact performance of the counting operation
- The count may change if data is added or removed after the measurement

---

## Cell Operations

### Set cell/range

**Purpose**: Writes values, text, or formulas to individual cells or ranges of cells with intelligent auto-fill capabilities. This command is the primary mechanism for populating spreadsheets with data and calculations. When applied to ranges, formulas automatically adjust their references to maintain relative positioning, similar to Excel's native fill-down or fill-right functionality.

The auto-fill feature is particularly powerful for creating calculated columns, applying formulas across data ranges, or setting up template structures that adapt to different data sizes. The command intelligently distinguishes between literal values and formulas based on the leading equals sign.

**Inputs**:
- `Set cell/range value or formula for` (RADIO, required): Target specification method - "Active cell/range" uses the current selection state, "Specific cell/range" allows precise addressing
- `Cell or range address` (TEXT, conditional): A1-style notation specifying the target location. Single cells like "A5" or ranges like "B10:D20" are supported. Only required when "Specific" method is selected.
- `Cell/range value or formula` (TEXT, required): Content to write to the target location. Text and numbers are written as literal values. Formulas must start with "=" and will auto-fill with relative reference adjustments across ranges. Use $ to create absolute references (e.g., $A$1).
- `Session name` (SESSION, required): Active workbook session where the operation will be performed

**Index Convention**: Uses A1 notation (A1, B2, A1:C5, etc.)

**Limitations**:
- Formula syntax must be valid Excel syntax or the operation will fail
- Auto-fill adjusts only relative references - absolute references (with $) remain fixed
- Very large ranges may impact performance and memory usage
- The active cell/range method depends on proper selection state being maintained by previous operations
- Complex array formulas or functions not supported by Apache POI may not work correctly

---

### Get single cell

**Purpose**: Extracts the content of a single cell and returns it as text, providing flexibility in how the data is interpreted. This command supports two distinct reading modes: visible text (which respects cell formatting and displays data as the user would see it) and raw value (which returns the underlying stored value). This distinction is crucial for data processing scenarios where formatted display may differ from actual values.

The reading mode selection affects how numbers, dates, percentages, and formulas are returned, making this command adaptable to different data processing needs. For example, a percentage cell might display as "50%" but have an underlying value of 0.5.

**Inputs**:
- `Cell option` (RADIO, required): Target specification method - "Active cell" uses the current cursor position, "Specific cell" allows precise addressing
- `Cell address` (TEXT, conditional): A1-style notation for the target cell location (e.g., "A5", "B10"). Only required when "Specific cell" method is selected.
- `Read option` (RADIO, required): Data interpretation method - "Read visible text in cell" returns formatted display text, "Read cell value" returns underlying stored values
- `Session name` (SESSION, required): Active workbook session containing the target cell

**Outputs**:
- Returns a `STRING` containing the cell content according to the selected read mode

**Index Convention**: Uses A1 notation for single cells only (A1, B2, etc.)

**Limitations**:
- Only works with individual cells - range addresses will cause errors
- Visible text mode may return formatted strings that cannot be used in numerical calculations
- Raw value mode may return scientific notation or full precision numbers that differ from display
- Formula cells return their calculated results, not the formula text itself
- Error cells (#DIV/0!, #REF!, etc.) are returned as error strings

---

### Go to cell

**Purpose**: Changes the active cell position within the worksheet, establishing the context for subsequent operations that use "Active" target modes. This command is essential for navigation-based workflows where you need to move through data systematically or position the cursor at specific locations for data entry or extraction operations.

The relative movement options provide sophisticated navigation capabilities that can adapt to data layouts, such as moving to row/column boundaries or navigating by single cells. This makes it possible to create dynamic navigation patterns that work with varying data structures.

**Inputs**:
- `Cell option` (RADIO, required): Navigation method - "Specific cell" for direct addressing, "Active cell" for relative movement from current position
- `Cell or range address` (TEXT, conditional): A1-style target location. If a range is specified, the top-leftmost cell becomes active. Only required when "Specific cell" method is selected.
- `Relative movement` (RADIO, conditional): Direction and distance for movement when "Active cell" method is selected. Options include single-cell movements (left, right, up, down) and boundary movements (beginning/end of row/column).
- `Session name` (SESSION, required): Active workbook session where navigation will occur

**Index Convention**: Uses A1 notation (A1, B2, etc.)

**Limitations**:
- When targeting ranges, only the top-leftmost cell becomes active - range selection is not maintained
- Relative movements that would go beyond worksheet boundaries are constrained to valid cell addresses
- The command does not create new cells - target addresses must be within Excel's valid range
- Boundary movements depend on data presence and may not behave predictably in sparse data regions

---

### Get cell address

**Purpose**: Retrieves the A1-style address of the currently active cell or locates cells based on column header matching. This command is valuable for tracking position during navigation-based processing, creating dynamic references for formulas, or implementing search functionality that locates data based on column headers rather than fixed positions.

The header-based search functionality is particularly useful for processing structured data where column positions may vary but header names remain consistent. This makes automations more robust when dealing with varying report formats.

**Inputs**:
- `Cell option` (RADIO, required): Address source method - "Active cell" returns the current cursor position, "Specific cell" searches for cells based on header matching
- `Column title` (TEXT, conditional): Text to search for in the header row (first non-empty row of the worksheet). The search is case-sensitive and looks for exact matches. Only required when "Specific cell" method is selected.
- `Position under header (1-based)` (NUMBER, conditional): Row offset below the matched header to return the address for. Value of 1 means the first data row immediately below the header. Only required when "Specific cell" method is selected.
- `Session name` (SESSION, required): Active workbook session for address retrieval

**Outputs**:
- Returns a `STRING` containing the A1-style cell address

**Index Convention**: Returns standard A1 notation; position numbers are **1-based**

**Limitations**:
- Active cell method requires a valid active cell to be set - returns null if none is set
- Header search is case-sensitive and requires exact text matching
- Header search only examines the first non-empty row of the worksheet
- Position offset must be within valid worksheet boundaries
- Multiple columns with identical headers may produce unpredictable results

---

### Delete Cell/Range

**Purpose**: Removes cells from the worksheet and automatically adjusts the layout by shifting remaining cells to fill the gap. This command is more sophisticated than simple content clearing because it actually removes the cell structure itself, which affects the positioning of surrounding data. The shift direction determines how the remaining cells reorganize to fill the space.

This operation is commonly used for data cleaning workflows where invalid or unnecessary rows/columns need to be completely removed from the dataset. The shift behavior mimics Excel's native delete cell functionality.

**Inputs**:
- `Cell option` (RADIO, required): Target specification method - "Active cell/range" uses current selection, "Specific cell/range" allows precise addressing
- `Cell or range address` (TEXT, conditional): A1-style notation for the target location (e.g., "A5" for single cell, "B10:D20" for range). Only required when "Specific" method is selected.
- `Delete option` (RADIO, required): Reorganization method after deletion - "Shift cells left" moves cells horizontally, "Shift cells up" moves cells vertically, "Entire row" removes complete rows, "Entire column" removes complete columns
- `Session name` (SESSION, required): Active workbook session where the deletion will occur

**Index Convention**: Uses A1 notation for ranges (A1:C5, etc.)

**Limitations**:
- Operation cannot be undone - deleted cells and their content are permanently lost
- Shift operations may disrupt carefully structured data layouts
- Formulas that reference deleted cells will result in #REF! errors
- Entire row/column deletions affect the entire worksheet structure
- Very large deletion operations may impact performance

---

## Data Operations

### Get worksheet as data table

**Purpose**: Converts structured worksheet data into A360's DataTable format for use with other automation commands and data manipulation operations. This command is essential for bridging the gap between Excel-based data storage and A360's data processing capabilities. It intelligently handles headers, data types, and empty cells to create clean, structured data sets that can be used for filtering, sorting, and analysis.

The extraction process can be customized to work with specific data regions, handle headers appropriately, and control how cell values are interpreted. This flexibility makes it suitable for processing various data layouts and formats commonly found in business spreadsheets.

**Inputs**:
- `Enter worksheet name` (RADIO, required): Source specification method - "Active worksheet" uses the currently selected sheet, "Specific worksheet" allows targeting any sheet by name
- `Worksheet name` (TEXT, conditional): Exact name of the target worksheet, case-sensitive. Only required when "Specific worksheet" method is selected.
- `Cell range` (RADIO, conditional): Data scope selection - "Entire sheet" processes all data in the worksheet, "Specific range" allows targeting precise cell ranges
- `Range address` (TEXT, conditional): A1-style range notation defining the data boundaries (e.g., "A1:D100", "A:C" for entire columns). Only required when "Specific range" is selected.
- `Row selection` (RADIO, conditional): Data subset control - "All rows" processes the complete range, "Specific rows" allows limiting to particular row ranges
- `From row` (NUMBER, conditional): Starting row number using 1-based indexing. Only required when "Specific rows" is selected.
- `To row` (NUMBER, conditional): Ending row number using 1-based indexing. Only required when "Specific rows" is selected.
- `Sheet contains a header` (CHECKBOX, required): Header handling flag - when enabled, the first row is treated as column headers and excluded from data processing
- `Read option` (RADIO, required): Value interpretation method - "Read visible text in cell" preserves formatting, "Read cell value" extracts raw underlying values
- `Session name` (SESSION, required): Active workbook session containing the source data

**Outputs**:
- Returns a `TABLE` (DataTable) object containing the structured data

**Index Convention**: Uses A1 notation for ranges; row numbers are **1-based**

**Limitations**:
- Very large datasets may consume significant memory and processing time
- Mixed data types within columns may cause conversion issues
- Empty cells are converted to null values in the DataTable
- Formula cells are converted to their calculated results, losing the original formulas
- Date and time formatting may not transfer perfectly to the DataTable format

---

### Write from data table

**Purpose**: Transfers data from A360's DataTable format back into worksheet cells, enabling the reverse operation of data table extraction. This command is crucial for report generation workflows where processed data needs to be written back to Excel for presentation or further analysis. The writing process handles data type conversion and can optionally preserve or reset cell formatting.

The positioning flexibility allows for precise control over where data is placed, making it possible to write results to specific report templates or append data to existing worksheets without disrupting the layout.

**Inputs**:
- `Enter data table variable` (TABLE, required): Source DataTable containing the data to be written to the worksheet
- `Enter worksheet name` (RADIO, required): Target specification method - "Active worksheet" writes to the currently selected sheet, "Specific worksheet" allows targeting any sheet by name
- `Worksheet name` (TEXT, conditional): Exact name of the target worksheet, case-sensitive. Only required when "Specific worksheet" method is selected.
- `Specify the first cell` (TEXT, required): A1-style address where the data table's first cell will be positioned (e.g., "A5", "B10"). Data expands down and right from this position.
- `Retain cell data type` (CHECKBOX, optional): Formatting preservation flag - when enabled, maintains original cell formatting; when disabled, clears formatting before writing new data
- `Session name` (SESSION, required): Active workbook session where data will be written

**Index Convention**: Uses A1 notation for start position

**Limitations**:
- Overwrites existing data without warning - no undo capability
- DataTable structure must be compatible with Excel's row/column limitations
- Large DataTables may extend beyond worksheet boundaries and cause errors
- Data type conversions may not be perfect for complex formats
- Formatting retention may not work correctly with all data types

---

### Filter

**Purpose**: Applies filtering criteria to Excel data to show only rows that meet specific conditions, similar to Excel's native AutoFilter functionality. This command works with both structured Excel tables (in .xlsx files) and regular worksheet ranges, providing flexible data filtering capabilities for analysis and reporting. The filtering system supports both numeric and text-based criteria with various comparison operators.

Filtering is particularly useful for data analysis workflows where you need to focus on specific subsets of data, create reports for particular criteria, or clean datasets by identifying records that meet certain conditions. The filtered results remain in place within the Excel worksheet.

**Inputs**:
- `Filter mode` (SELECT, required): Data structure type - "Table" for working with Excel table objects, "Worksheet" for filtering regular cell ranges
- `Table name` (TEXT, conditional): Name of the target Excel table object. Only required when "Table" mode is selected.
- `Filter for` (RADIO, conditional): Column identification method when working with tables - "Column name" uses header text, "Column position" uses numeric positioning
- `Column name` (TEXT, conditional): Exact text of the column header to filter by. Case-sensitive. Required when table mode uses column name identification.
- `Column position` (NUMBER, conditional): 1-based column number within the table structure. Required when table mode uses position identification.
- `Worksheet name` (TEXT, conditional): Exact name of the target worksheet when using worksheet mode. Case-sensitive.
- `Cell range` (RADIO, conditional): Data scope when using worksheet mode - "All cells" applies to the entire worksheet, "Specific" targets a defined range
- `Range` (TEXT, conditional): A1-style range notation for the data area to filter. Required when worksheet mode uses specific range.
- `Filter` (RADIO, required): Data type classification for the filtering criteria - "Number" for numeric comparisons, "Text" for string-based matching
- Number filter options: "Equals", "Does not equal", "Greater than", "Greater than or equal", "Less than", "Less than or equal", "Between"
- Text filter options: "Equals", "Does not equal", "Begins with", "Ends with", "Contains", "Does not contain"
- Filter values (TEXT/NUMBER, required): Specific criteria values based on the selected filter type and operator
- `Session name` (SESSION, required): Active workbook session containing the data to be filtered

**Index Convention**: Column positions are **1-based**; uses A1 notation for ranges

**Limitations**:
- Table mode requires .xlsx format files and existing Excel table structures
- Only supports single-column filtering per command execution
- Text filtering is case-sensitive and uses exact matching
- Complex filter expressions or multiple simultaneous criteria are not supported
- Filtered state persists in the worksheet and may affect other operations
- Worksheet mode filtering may not persist as reliably as table-based filtering

---

### Sort

**Purpose**: Arranges worksheet data in ascending or descending order based on values in a specified column, providing essential data organization capabilities. This command works with both Excel table structures and regular worksheet ranges, automatically handling headers and maintaining row integrity during the sort process. The sorting operation preserves relationships between data in different columns within the same rows.

Sorting is fundamental for data analysis, report generation, and preparing data for further processing. The command provides both numeric and alphabetical sorting with proper handling of different data types within the sort column.

**Inputs**:
- `Sort mode` (SELECT, required): Data structure type - "Table" for Excel table objects, "Worksheet" for regular cell ranges
- `Table name` (TEXT, conditional): Name of the target Excel table. Required when "Table" mode is selected.
- `Sort for` (RADIO, conditional): Column identification method for table mode - "Column name" uses header text, "Column position" uses numeric positioning
- `Column name` (TEXT, conditional): Exact text of the column header to sort by. Case-sensitive. Required when table mode uses name identification.
- `Column position` (NUMBER, conditional): 1-based column number within the table. Required when table mode uses position identification.
- `Sheet name` (TEXT, conditional): Exact name of the worksheet when using worksheet mode. Case-sensitive.
- `Cell range` (RADIO, conditional): Data scope for worksheet mode - "All cells" sorts the entire worksheet, "Specific" targets a defined range
- `Range` (TEXT, conditional): A1-style range notation defining the sort area. Required when worksheet mode uses specific range.
- `Data has headers` (CHECKBOX, conditional): Header handling flag for worksheet mode - when enabled, excludes the first row from sorting
- `Sort order` (RADIO, required): Sort criteria type - "Number" for numeric sorting, "Text" for alphabetical sorting
- Number order options: "Smallest to largest" (ascending), "Largest to smallest" (descending)
- Text order options: "A to Z" (ascending), "Z to A" (descending)
- `Session name` (SESSION, required): Active workbook session containing the data to be sorted

**Index Convention**: Column positions are **1-based**; uses A1 notation for ranges

**Limitations**:
- Table mode requires .xlsx format and existing Excel table structures
- Only supports single-column sorting - multi-level sorts require multiple command executions
- Mixed data types within the sort column may produce unexpected results
- Sort operation cannot be undone through the command interface
- Very large datasets may impact performance significantly
- Custom sort orders or special sorting rules are not supported

---

### Insert/Delete rows/columns

**Purpose**: Modifies the worksheet structure by adding new rows/columns or removing existing ones, automatically adjusting references and maintaining data integrity. This command is essential for dynamic data processing where the worksheet structure needs to adapt to changing data requirements. All formulas and references are automatically updated to account for the structural changes.

This operation is commonly used for data preparation, report formatting, and creating space for additional data entry. The command handles both single and multiple row/column operations efficiently.

**Inputs**:
- `Operation group` (RADIO, required): Target structure type - "Row operations" for horizontal changes, "Column operations" for vertical changes
- Row operations:
  - Operation type (RADIO): "Insert Row(s) at" for adding rows, "Delete Row(s) at" for removing rows
  - `Target rows` (TEXT, required): Row specification - single row number (e.g., "10") or range (e.g., "1:10")
- Column operations:
  - Operation type (RADIO): "Insert Column(s) at" for adding columns, "Delete Column(s) at" for removing columns
  - `Target columns` (TEXT, required): Column specification - single column letter (e.g., "B") or range (e.g., "B:D")
- `Session name` (SESSION, required): Active workbook session where the structural changes will be made

**Index Convention**: Row positions are **1-based** (1 = first row); column letters follow Excel standard (A, B, C, etc.)

**Limitations**:
- Target positions must be within valid worksheet boundaries
- Delete operations cannot remove all rows or columns from a worksheet
- Insert operations may cause formulas to adjust their references automatically
- Large insert/delete operations may significantly impact performance
- Operations cannot be undone through the command interface
- Structural changes may affect other sheets that reference the modified areas

---

## Utility Operations

### Find

**Purpose**: Searches through worksheet content to locate cells containing specific text strings, returning the addresses of all matching cells. This command provides comprehensive search capabilities with options for case sensitivity, whole cell matching, and directional searching patterns. It's invaluable for data validation, content location, and dynamic processing workflows that need to adapt to varying data layouts.

The search functionality can be constrained to specific regions and configured to search in different patterns (by rows or by columns), making it suitable for structured data analysis and quality control processes.

**Inputs**:
- `Find text` (TEXT, required): Text string to search for within the worksheet cells. The search examines visible cell content, not formulas.
- `From` (RADIO, required): Search starting point - "Active Cell" begins from cursor position, "Specific Cell" starts from defined address, "Beginning" starts from A1, "End" starts from the last used cell
- `From cell` (TEXT, conditional): A1-style address for the search starting point. Required when "Specific Cell" is selected.
- `To` (RADIO, required): Search ending boundary - "Active Cell" ends at cursor position, "Specific Cell" ends at defined address, "Beginning" ends at A1, "End" ends at the last used cell
- `To cell` (TEXT, conditional): A1-style address for the search ending point. Required when "Specific Cell" is selected.
- `Search direction` (RADIO, required): Search pattern - "By rows" searches left-to-right then down, "By columns" searches top-to-bottom then right
- `Match case` (CHECKBOX, optional): Case sensitivity flag - when enabled, "Text" and "text" are treated as different strings
- `Match entire cell contents` (CHECKBOX, optional): Whole cell matching - when enabled, only cells containing exactly the search string (and nothing else) are matched
- `Session name` (SESSION, required): Active workbook session to search within

**Outputs**:
- Returns a `LIST` of strings containing A1-style addresses of all matching cells

**Index Convention**: Returns A1 notation addresses

**Limitations**:
- Only searches visible cell values, not underlying formulas or hidden content
- Search is limited to the currently active worksheet
- Very large worksheets may impact search performance
- Complex search patterns or regular expressions are not supported
- Search results are returned as static addresses that may become invalid if worksheet structure changes

---

### Find next empty cell

**Purpose**: Locates the next cell containing no data when traversing from a specified starting point in a particular direction. This command is essential for dynamic data processing where you need to find insertion points, determine data boundaries, or locate available space for new content. The search pattern can be configured to move by rows or columns depending on data layout requirements.

This functionality is particularly useful for append operations, data validation (checking for gaps), and creating adaptive processing workflows that can handle varying data sizes and structures.

**Inputs**:
- `Traverse by` (RADIO, required): Search direction pattern - "row" searches horizontally (left/right), "column" searches vertically (up/down)
- `Start from` (RADIO, required): Starting point specification - "active cell" begins from current cursor position, "specific cell" begins from defined address
- `Cell address` (TEXT, conditional): A1-style address for the search starting point. Required when "specific cell" is selected.
- `Session name` (SESSION, required): Active workbook session to search within

**Outputs**:
- Returns a `STRING` containing the A1-style address of the next empty cell

**Index Convention**: Returns A1 notation addresses

**Limitations**:
- Search stops at worksheet boundaries if no empty cell is found
- May not find empty cells in very dense data regions
- Returns the starting cell address if it's already empty
- Search direction is determined by the traverse setting and cannot be dynamically adjusted
- Performance may degrade in worksheets with large amounts of data

---

### Go to next empty cell

**Purpose**: Navigates the active cell cursor to the next empty cell in a specified direction, combining search and navigation functionality in a single operation. This command is particularly useful for data entry workflows where you need to systematically move through available positions, or for processing patterns that require positioning at empty locations for subsequent operations.

Unlike the find command, this operation actually changes the active cell position, making it ready for immediate data entry or extraction operations without requiring additional navigation commands.

**Inputs**:
- `Start from` (RADIO, required): Starting point specification - "Active Cell" begins from current cursor position, "Specific Cell" begins from defined address
- `Cell address` (TEXT, conditional): A1-style address for the search starting point. Required when "Specific Cell" is selected.
- `Towards` (RADIO, required): Search direction - "Left", "Right", "Up", or "Down" from the starting position
- `Session name` (SESSION, required): Active workbook session where navigation will occur

**Index Convention**: Uses standard cell navigation conventions

**Limitations**:
- Navigation stops at worksheet boundaries if no empty cell is found in the specified direction
- May not move if the starting position is already empty
- Direction is fixed per command execution - dynamic direction changes require multiple commands
- The active cell position change affects all subsequent operations using "Active" target modes

---

## Iterator Operations

### For each row in worksheet or table

**Purpose**: Provides systematic row-by-row processing of Excel data by iterating through each row and returning the cell values as structured record objects. This iterator is fundamental for data processing workflows where each row represents a distinct record that needs individual processing. The iterator handles headers automatically, supports both Excel tables and worksheet ranges, and provides flexible data reading options.

The iterator maintains state between iterations, allowing for consistent processing of large datasets without loading all data into memory simultaneously. This makes it suitable for processing very large Excel files efficiently while maintaining low memory usage.

**Inputs**:
- `Iterator mode` (SELECT, required): Data source type - "Table" for Excel table objects, "Worksheet" for regular cell ranges

**TABLE Mode Inputs**:
- `Table name` (TEXT, required): Name of the Excel table object to iterate through
- `Row selection` (SELECT, required): Data subset control - "All rows" processes every data row, "Specific rows" allows limiting to particular ranges
- `From row` (NUMBER, conditional): Starting row number within the data region using 1-based indexing (excludes headers). Required when "Specific rows" is selected.
- `To row` (NUMBER, conditional): Ending row number within the data region using 1-based indexing. Required when "Specific rows" is selected.

**WORKSHEET Mode Inputs**:
- `Sheet name` (TEXT, optional): Name of the worksheet to process. If empty, uses the currently active sheet.
- `Range mode` (SELECT, required): Data scope - "Entire sheet" processes all data in the worksheet, "Specific range" targets defined cell ranges
- `Range address` (TEXT, conditional): A1-style range notation defining the processing area. Required when "Specific range" is selected.
- `Row selection` (SELECT, required): Data subset control - "All rows" processes the complete range, "Specific rows" allows limiting to particular row ranges
- `From row` (NUMBER, conditional): Starting row number within the data region using 1-based indexing. Required when "Specific rows" is selected.
- `To row` (NUMBER, conditional): Ending row number within the data region using 1-based indexing. Required when "Specific rows" is selected.
- `Sheet contains a header` (CHECKBOX, required): Header handling flag - when enabled, treats the first row as column headers and excludes it from iteration

**Common Inputs**:
- `Read option` (RADIO, required): Value interpretation method - "Read visible text in cell" preserves formatting and returns display values, "Read cell value" extracts underlying stored values
- `Session name` (SESSION, required): Active workbook session containing the data to iterate

**Outputs**:
- Returns `RECORD` objects containing row data with column names as field identifiers

**Index Convention**: 
- Row numbers are **1-based** for user input
- A1 notation for range specifications
- Column headers become field names in the returned records

**Limitations**:
- Table mode requires .xlsx format workbooks with properly defined Excel table structures
- Iterator state is not persistent across automation runs - cannot resume from previous positions
- Very large datasets may impact performance, though memory usage is optimized through streaming
- Row number specifications are relative to the data region (headers are excluded from counting)
- Field names in returned records are derived from column headers or auto-generated if headers are not present
- Complex data types may not convert perfectly to record field values

---

## General Notes

### Session Management Best Practices
- Always establish sessions before performing Excel operations - sessions provide context and resource management
- Close sessions properly to release file locks and free system resources
- Use meaningful session names for complex workflows involving multiple workbooks
- Consider read-only mode for data extraction workflows to prevent accidental modifications

### Index and Addressing Conventions
- **Sheet positions**: Consistently use 1-based indexing (first sheet = 1, second sheet = 2)
- **Row and column positions**: Follow 1-based numbering (first row = 1, first column = 1) 
- **Cell addresses**: Standard Excel A1 notation (A1, B2, C3) for consistency with Excel interface
- **Range specifications**: Use colon notation (A1:B5) for defining rectangular areas
- **Column references**: Use Excel letter notation (A, B, C) for column identification

### Data Reading and Writing Options
- **Visible text mode**: Preserves Excel formatting and returns values as they appear to users (e.g., "50%" for percentage cells)
- **Cell value mode**: Returns underlying stored values for calculations (e.g., 0.5 for a 50% cell)
- **Formula handling**: Formulas are evaluated and results are returned, not the formula text itself
- **Data type conversion**: Mixed data types within operations may produce unexpected results

### Target Selection Methods
- **Active selections**: Depend on cursor position and current selection state maintained by previous operations
- **Specific selections**: Provide precise control but require exact addressing and may be less flexible
- **Range operations**: Support both individual cells and rectangular areas using A1 notation

### Performance Considerations
- **Large datasets**: May require significant memory and processing time, especially for iterative operations
- **File I/O operations**: Opening, saving, and closing large workbooks may introduce delays
- **Complex formulas**: Advanced Excel functions may not be fully supported by Apache POI
- **Memory management**: Iterator operations are optimized for memory efficiency with large datasets
- **Concurrent access**: File locking prevents multiple operations on the same file simultaneously

### Error Handling and Limitations
- **Formula errors**: Cells containing #DIV/0!, #REF!, or other Excel errors are returned as error strings
- **Data validation**: Invalid addresses, non-existent sheets, or out-of-range positions cause command failures
- **File format restrictions**: Advanced Excel features may not be available in legacy .xls format
- **Undo limitations**: Most operations cannot be reversed - always work with backup copies when possible

### Excel Format Compatibility
- **.xlsx files (Excel 2007+)**: Full feature support including tables, advanced filtering, and modern Excel functions
- **.xls files (Excel 97-2003)**: Basic functionality only - table operations and some advanced features not available
- **Password protection**: Protected workbooks are not currently supported
- **External links**: References to other workbooks may not be maintained correctly

### Advanced Features
- **Auto-fill capabilities**: Formula operations automatically adjust references when applied to ranges
- **Header recognition**: Commands intelligently detect and handle header rows for structured data processing
- **Filter and sort persistence**: Applied filters and sorts remain active in the worksheet until manually cleared
- **Cross-sheet operations**: Commands can work across multiple worksheets within the same session