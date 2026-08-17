# Power-Query-File-Comparison

Yes. I would restructure it as a small Power Query reconciliation framework rather than four independent queries that each reread the folders.

The optimized architecture is:

```text
tblParameters
      │
      ▼
Parameters
      │
      ▼
File Inventory
      │
      ├──────────────► File Comparison
      │
      ▼
Sheet Inventory
      │
      ├──────────────► Sheet Comparison
      │
      ▼
Cell Differences
      │
      ▼
Summary
```

The key performance improvements are:

* Folder paths are controlled by an Excel table rather than hard-coded.
* Each workbook is opened with `Excel.Workbook` only once in the shared inventory query.
* Matching uses normalized, case-insensitive keys.
* Cell comparison is performed **row-by-row**, rather than creating a massive Cartesian product of every possible cell.
* Only actual differences are returned.
* Missing files and sheets are handled separately.
* `Summary` aggregates the exception data rather than reprocessing the workbooks.
* The comparison reads sheets with `useHeaders=false`, so the comparison includes the actual worksheet cells starting at A1. `Excel.Workbook` supports this header behavior and the delayed type evaluation used below. ([Berenguer Formation Conseil][1])

## 1. Create the parameter table

On a worksheet, create an Excel table named:

**`tblParameters`**

with exactly these columns:

| Parameter | Value               |
| --------- | ------------------- |
| Folder1   | `C:\Data\Version_A` |
| Folder2   | `C:\Data\Version_B` |

The paths can be changed at any time without editing the Power Query code.

---

# 2. `Parameters`

Create a blank query named **`Parameters`**.

```powerquery
let
    Source =
        Excel.CurrentWorkbook(){[Name="tblParameters"]}[Content],

    Typed =
        Table.TransformColumnTypes(
            Source,
            {
                {"Parameter", type text},
                {"Value", type text}
            }
        ),

    Clean =
        Table.TransformColumns(
            Typed,
            {
                {"Parameter", each Text.Trim(_), type text},
                {"Value", each Text.Trim(_), type text}
            }
        ),

    GetParameter =
        (ParameterName as text) as text =>
            let
                Matches =
                    Table.SelectRows(
                        Clean,
                        each Text.Upper([Parameter])
                            = Text.Upper(ParameterName)
                    ),

                Value =
                    if Table.RowCount(Matches) = 0 then
                        error "Missing parameter: " & ParameterName
                    else
                        Matches{0}[Value]
            in
                Value,

    Folder1 =
        GetParameter("Folder1"),

    Folder2 =
        GetParameter("Folder2"),

    Output =
        [
            Folder1 = Folder1,
            Folder2 = Folder2
        ]
in
    Output
```

This query returns a record containing the two paths.

---

# 3. `File Inventory`

This is the primary staging query.

Create a query named **`File Inventory`**.

```powerquery
let
    //===========================================================
    // Configuration
    //===========================================================

    Folder1 =
        Parameters[Folder1],

    Folder2 =
        Parameters[Folder2],

    //===========================================================
    // Function: Get files from one folder
    //===========================================================

    GetFiles =
        (FolderPath as text, SourceID as text) as table =>
        let
            Source =
                Folder.Files(FolderPath),

            ExcelFiles =
                Table.SelectRows(
                    Source,
                    each
                        List.Contains(
                            {".xlsx", ".xlsm", ".xlsb"},
                            Text.Lower([Extension])
                        )
                        and not Text.StartsWith(
                            [Name],
                            "~$"
                        )
                ),

            Selected =
                Table.SelectColumns(
                    ExcelFiles,
                    {
                        "Name",
                        "Extension",
                        "Content",
                        "Date modified",
                        "Folder Path"
                    }
                ),

            AddSource =
                Table.AddColumn(
                    Selected,
                    "Source",
                    each SourceID,
                    type text
                ),

            AddFileKey =
                Table.AddColumn(
                    AddSource,
                    "FileKey",
                    each
                        Text.Upper(
                            Text.Trim([Name])
                        ),
                    type text
                ),

            AddRelativePath =
                Table.AddColumn(
                    AddFileKey,
                    "RelativePath",
                    each
                        Text.AfterDelimiter(
                            [Folder Path],
                            FolderPath,
                            {0, RelativePosition.FromEnd}
                        ),
                    type text
                )
        in
            AddRelativePath,

    //===========================================================
    // Read both folders
    //===========================================================

    Source1 =
        GetFiles(
            Folder1,
            "Folder 1"
        ),

    Source2 =
        GetFiles(
            Folder2,
            "Folder 2"
        ),

    //===========================================================
    // Combine
    //===========================================================

    Combined =
        Table.Combine(
            {
                Source1,
                Source2
            }
        ),

    //===========================================================
    // Validate duplicate workbook names within a folder
    //===========================================================

    DuplicateCheck =
        Table.Group(
            Combined,
            {
                "Source",
                "FileKey"
            },
            {
                {
                    "FileCount",
                    each Table.RowCount(_),
                    Int64.Type
                }
            }
        ),

    Duplicates =
        Table.SelectRows(
            DuplicateCheck,
            each [FileCount] > 1
        ),

    AddDuplicateFlag =
        Table.AddColumn(
            Combined,
            "DuplicateFile",
            each
                Table.RowCount(
                    Table.SelectRows(
                        Duplicates,
                        (r) =>
                            r[Source] = [Source]
                            and r[FileKey] = [FileKey]
                    )
                ) > 0,
            type logical
        ),

    //===========================================================
    // Keep content because downstream queries use it
    //===========================================================

    Result =
        Table.Buffer(
            AddDuplicateFlag
        )
in
    Result
```

### Why this query is important

It becomes the common source for everything else.

It also detects this situation:

```text
Folder 1
    Rates.xlsx
    Rates.xlsx       <-- duplicate filename in subfolders

Folder 2
    Rates.xlsx
```

The `DuplicateFile` flag prevents an ambiguous workbook comparison from silently producing incorrect results.

---

# 4. `File Comparison`

Create a query named **`File Comparison`**.

This compares the workbook inventory without opening workbook contents again.

```powerquery
let
    Source =
        #"File Inventory",

    Folder1 =
        Table.SelectRows(
            Source,
            each [Source] = "Folder 1"
        ),

    Folder2 =
        Table.SelectRows(
            Source,
            each [Source] = "Folder 2"
        ),

    //===========================================================
    // Folder 1 -> Folder 2
    //===========================================================

    JoinFolder2 =
        Table.NestedJoin(
            Folder1,
            {"FileKey"},
            Folder2,
            {"FileKey"},
            "Folder2File",
            JoinKind.LeftOuter
        ),

    ExpandFolder2 =
        Table.ExpandTableColumn(
            JoinFolder2,
            "Folder2File",
            {
                "Name",
                "Date modified",
                "Folder Path",
                "Content",
                "DuplicateFile"
            },
            {
                "Folder2FileName",
                "Folder2Modified",
                "Folder2Path",
                "Folder2Content",
                "Folder2Duplicate"
            }
        ),

    //===========================================================
    // Status
    //===========================================================

    AddStatus =
        Table.AddColumn(
            ExpandFolder2,
            "Status",
            each
                if [DuplicateFile] then
                    "DUPLICATE IN FOLDER 1"

                else if [Folder2Content] = null then
                    "MISSING FROM FOLDER 2"

                else if [Folder2Duplicate] then
                    "DUPLICATE IN FOLDER 2"

                else
                    "MATCHED",
            type text
        ),

    //===========================================================
    // Find files that exist only in Folder 2
    //===========================================================

    Folder2Only =
        Table.SelectRows(
            Folder2,
            each
                not List.Contains(
                    Folder1[FileKey],
                    [FileKey]
                )
        ),

    Folder2OnlyFormatted =
        Table.AddColumn(
            Folder2Only,
            "Status",
            each "MISSING FROM FOLDER 1",
            type text
        ),

    Folder2OnlySelected =
        Table.SelectColumns(
            Folder2OnlyFormatted,
            {
                "FileKey",
                "Name",
                "Date modified",
                "Folder Path",
                "Status"
            }
        ),

    Folder1Selected =
        Table.SelectColumns(
            AddStatus,
            {
                "FileKey",
                "Name",
                "Date modified",
                "Folder Path",
                "Folder2FileName",
                "Folder2Modified",
                "Folder2Path",
                "Status"
            }
        ),

    RenameColumns =
        Table.RenameColumns(
            Folder1Selected,
            {
                {"Name", "Folder1FileName"},
                {"Date modified", "Folder1Modified"},
                {"Folder Path", "Folder1Path"}
            }
        ),

    AddMissingColumns =
        Table.AddColumn(
            RenameColumns,
            "Folder2FileName2",
            each null,
            type text
        ),

    //===========================================================
    // Final structure
    //===========================================================

    Folder1Result =
        Table.SelectColumns(
            AddMissingColumns,
            {
                "FileKey",
                "Folder1FileName",
                "Folder1Modified",
                "Folder1Path",
                "Folder2FileName",
                "Folder2Modified",
                "Folder2Path",
                "Status"
            }
        ),

    Folder2Result =
        Table.RenameColumns(
            Folder2OnlySelected,
            {
                {"Name", "Folder2FileName"},
                {"Date modified", "Folder2Modified"},
                {"Folder Path", "Folder2Path"}
            }
        ),

    Folder2ResultWithColumns =
        Table.AddColumn(
            Folder2Result,
            "Folder1FileName",
            each null,
            type text
        ),

    Folder2ResultWithModified =
        Table.AddColumn(
            Folder2ResultWithColumns,
            "Folder1Modified",
            each null
        ),

    Folder2ResultWithPath =
        Table.AddColumn(
            Folder2ResultWithModified,
            "Folder1Path",
            each null,
            type text
        ),

    Folder2ResultReordered =
        Table.SelectColumns(
            Folder2ResultWithPath,
            {
                "FileKey",
                "Folder1FileName",
                "Folder1Modified",
                "Folder1Path",
                "Folder2FileName",
                "Folder2Modified",
                "Folder2Path",
                "Status"
            }
        ),

    Result =
        Table.Sort(
            Table.Combine(
                {
                    Folder1Result,
                    Folder2ResultReordered
                }
            ),
            {
                {"Status", Order.Ascending},
                {"FileKey", Order.Ascending}
            }
        )
in
    Result
```

The result gives you a workbook-level control report.

---

# 5. `Sheet Inventory`

This is the key optimization layer.

Create a query named **`Sheet Inventory`**.

```powerquery
let
    Source =
        #"File Inventory",

    //===========================================================
    // Only process unique, valid workbooks
    //===========================================================

    ValidFiles =
        Table.SelectRows(
            Source,
            each [DuplicateFile] = false
        ),

    //===========================================================
    // Open each workbook ONCE
    //===========================================================

    AddWorkbook =
        Table.AddColumn(
            ValidFiles,
            "Workbook",
            each
                try
                    Excel.Workbook(
                        [Content],
                        false,
                        true
                    )
                otherwise
                    null
        ),

    //===========================================================
    // Keep worksheets only
    //===========================================================

    AddSheets =
        Table.AddColumn(
            AddWorkbook,
            "Sheets",
            each
                if [Workbook] = null then
                    null
                else
                    Table.SelectRows(
                        [Workbook],
                        each [Kind] = "Sheet"
                    )
        ),

    //===========================================================
    // Expand worksheet metadata
    //===========================================================

    ExpandSheets =
        Table.ExpandTableColumn(
            AddSheets,
            "Sheets",
            {
                "Name",
                "Data",
                "Item",
                "Kind",
                "Hidden"
            },
            {
                "SheetName",
                "Data",
                "Item",
                "Kind",
                "Hidden"
            }
        ),

    //===========================================================
    // Normalize sheet key
    //===========================================================

    AddSheetKey =
        Table.AddColumn(
            ExpandSheets,
            "SheetKey",
            each
                Text.Upper(
                    Text.Trim([SheetName])
                ),
            type text
        ),

    //===========================================================
    // Add sheet position
    //===========================================================

    Grouped =
        Table.Group(
            AddSheetKey,
            {
                "Source",
                "FileKey"
            },
            {
                {
                    "Rows",
                    each
                        Table.AddIndexColumn(
                            _,
                            "SheetIndex",
                            1,
                            1,
                            Int64.Type
                        ),
                    type table
                }
            }
        ),

    Expanded =
        Table.ExpandTableColumn(
            Grouped,
            "Rows",
            {
                "SheetName",
                "SheetKey",
                "Data",
                "Hidden",
                "SheetIndex"
            },
            {
                "SheetName",
                "SheetKey",
                "Data",
                "Hidden",
                "SheetIndex"
            }
        ),

    Result =
        Table.SelectColumns(
            Expanded,
            {
                "Source",
                "FileKey",
                "SheetName",
                "SheetKey",
                "Data",
                "Hidden",
                "SheetIndex"
            }
        )
in
    Result
```

### Important optimization

This line:

```powerquery
Excel.Workbook([Content], false, true)
```

is deliberately used with `false` for headers. That means the comparison operates on the physical worksheet cells rather than interpreting row 1 as column headers.

That is critical if your requirement is genuinely:

> "Compare all sheets."

Power Query represents an entire worksheet with blank cells as `null`, so the comparison can identify changes in the physical sheet contents. ([GitHub][2])

---

# 6. `Sheet Comparison`

Create a query named **`Sheet Comparison`**.

```powerquery
let
    Source =
        #"Sheet Inventory",

    Folder1 =
        Table.SelectRows(
            Source,
            each [Source] = "Folder 1"
        ),

    Folder2 =
        Table.SelectRows(
            Source,
            each [Source] = "Folder 2"
        ),

    //===========================================================
    // Folder 1 -> Folder 2
    //===========================================================

    JoinFolder2 =
        Table.NestedJoin(
            Folder1,
            {
                "FileKey",
                "SheetKey"
            },
            Folder2,
            {
                "FileKey",
                "SheetKey"
            },
            "F2",
            JoinKind.LeftOuter
        ),

    ExpandF2 =
        Table.ExpandTableColumn(
            JoinFolder2,
            "F2",
            {
                "SheetName",
                "Data",
                "Hidden",
                "SheetIndex"
            },
            {
                "Folder2SheetName",
                "Data2",
                "Folder2Hidden",
                "Folder2SheetIndex"
            }
        ),

    AddStatus =
        Table.AddColumn(
            ExpandF2,
            "Status",
            each
                if [Data2] = null then
                    "MISSING FROM FOLDER 2"
                else
                    "MATCHED",
            type text
        ),

    Folder1Result =
        Table.SelectColumns(
            AddStatus,
            {
                "FileKey",
                "SheetKey",
                "SheetName",
                "Folder2SheetName",
                "SheetIndex",
                "Folder2SheetIndex",
                "Hidden",
                "Folder2Hidden",
                "Data",
                "Data2",
                "Status"
            }
        ),

    //===========================================================
    // Sheets only in Folder 2
    //===========================================================

    Folder2Only =
        Table.SelectRows(
            Folder2,
            each
                not List.Contains(
                    Table.SelectRows(
                        Folder1,
                        (x) =>
                            x[FileKey] = [FileKey]
                            and x[SheetKey] = [SheetKey]
                    )[SheetKey],
                    [SheetKey]
                )
        ),

    Folder2OnlyResult =
        Table.AddColumn(
            Folder2Only,
            "Status",
            each "MISSING FROM FOLDER 1",
            type text
        ),

    Folder2OnlyFormatted =
        Table.SelectColumns(
            Folder2OnlyResult,
            {
                "FileKey",
                "SheetKey",
                "SheetName",
                "Hidden",
                "SheetIndex",
                "Data",
                "Status"
            }
        ),

    AddFolder2SheetName =
        Table.AddColumn(
            Folder2OnlyFormatted,
            "Folder2SheetName",
            each [SheetName],
            type text
        ),

    AddFolder2Index =
        Table.AddColumn(
            AddFolder2SheetName,
            "Folder2SheetIndex",
            each [SheetIndex],
            Int64.Type
        ),

    AddFolder2Hidden =
        Table.AddColumn(
            AddFolder2Index,
            "Folder2Hidden",
            each [Hidden],
            type logical
        ),

    AddData2 =
        Table.AddColumn(
            AddFolder2Hidden,
            "Data2",
            each null
        ),

    Reordered =
        Table.SelectColumns(
            AddData2,
            {
                "FileKey",
                "SheetKey",
                "SheetName",
                "Folder2SheetName",
                "SheetIndex",
                "Folder2SheetIndex",
                "Hidden",
                "Folder2Hidden",
                "Data",
                "Data2",
                "Status"
            }
        ),

    Combined =
        Table.Combine(
            {
                Folder1Result,
                Reordered
            }
        ),

    //===========================================================
    // Identify sheet-level duplicate names
    //===========================================================

    DuplicateSheets =
        Table.Group(
            Source,
            {
                "Source",
                "FileKey",
                "SheetKey"
            },
            {
                {
                    "SheetCount",
                    each Table.RowCount(_),
                    Int64.Type
                }
            }
        ),

    AddDuplicateFlag =
        Table.AddColumn(
            Combined,
            "DuplicateSheet",
            each
                Table.RowCount(
                    Table.SelectRows(
                        DuplicateSheets,
                        (x) =>
                            x[FileKey] = [FileKey]
                            and x[SheetKey] = [SheetKey]
                            and x[SheetCount] > 1
                    )
                ) > 0,
            type logical
        ),

    Result =
        Table.Sort(
            AddDuplicateFlag,
            {
                {"FileKey", Order.Ascending},
                {"SheetIndex", Order.Ascending}
            }
        )
in
    Result
```

---

# 7. `fxCompareSheets`

I recommend using a dedicated function for the cell comparison rather than embedding the logic inside `Cell Differences`.

Create a blank query named:

**`fxCompareSheets`**

```powerquery
(
    Table1 as nullable table,
    Table2 as nullable table
)
as table =>
let

    OutputColumns =
        {
            "Row",
            "Column",
            "Folder1Value",
            "Folder2Value",
            "Status"
        },

    Empty =
        #table(
            OutputColumns,
            {}
        ),

    Result =
        if Table1 = null or Table2 = null then
            Empty

        else
            let

                //=================================================
                // Convert tables to row lists.
                // This is substantially more efficient than
                // generating Row x Column Cartesian products.
                //=================================================

                Rows1 =
                    Table.ToRows(Table1),

                Rows2 =
                    Table.ToRows(Table2),

                RowCount =
                    List.Max(
                        {
                            List.Count(Rows1),
                            List.Count(Rows2)
                        }
                    ),

                ColumnCount =
                    List.Max(
                        {
                            Table.ColumnCount(Table1),
                            Table.ColumnCount(Table2)
                        }
                    ),

                //=================================================
                // Compare one row
                //=================================================

                CompareRow =
                    (RowNumber as number) as list =>
                    let

                        Index =
                            RowNumber - 1,

                        R1 =
                            if Index < List.Count(Rows1)
                            then Rows1{Index}
                            else {},

                        R2 =
                            if Index < List.Count(Rows2)
                            then Rows2{Index}
                            else {},

                        Differences =
                            List.Combine(
                                List.Transform(
                                    {0..ColumnCount - 1},
                                    (ColumnIndex) =>
                                    let

                                        V1 =
                                            if ColumnIndex
                                                < List.Count(R1)
                                            then
                                                R1{ColumnIndex}
                                            else
                                                null,

                                        V2 =
                                            if ColumnIndex
                                                < List.Count(R2)
                                            then
                                                R2{ColumnIndex}
                                            else
                                                null,

                                        Equal =
                                            if V1 = null and V2 = null then
                                                true
                                            else if Value.Is(V1, type number)
                                                and Value.Is(V2, type number) then
                                                V1 = V2
                                            else
                                                try
                                                    Value.Equals(V1, V2)
                                                otherwise
                                                    Text.From(V1)
                                                    = Text.From(V2)

                                    in
                                        if Equal then
                                            {}
                                        else
                                            {
                                                {
                                                    RowNumber,
                                                    ColumnIndex + 1,
                                                    V1,
                                                    V2,
                                                    "DIFFERENT"
                                                }
                                            }
                                )
                            )

                    in
                        Differences,

                //=================================================
                // Process rows
                //=================================================

                DifferenceRows =
                    if RowCount = 0
                    then {}
                    else
                        List.Combine(
                            List.Transform(
                                {1..RowCount},
                                each CompareRow(_)
                            )
                        ),

                Output =
                    if List.IsEmpty(DifferenceRows)
                    then
                        Empty
                    else
                        #table(
                            OutputColumns,
                            DifferenceRows
                        )

            in
                Output
in
    Result
```

### Why this version is materially better

A naive implementation might create:

```text
Rows × Columns
```

records for every worksheet and then filter them.

This implementation keeps the data as row lists and only emits a record when a cell actually differs.

`Table.ToRows` converts a table into nested row lists, which is specifically useful for this type of positional comparison. ([Microsoft Learn][3])

---

# 8. `Cell Differences`

Now create the main exception query.

Name it:

**`Cell Differences`**

```powerquery
let
    Source =
        #"Sheet Comparison",

    //===========================================================
    // Only compare sheets existing in both workbooks
    //===========================================================

    Comparable =
        Table.SelectRows(
            Source,
            each
                [Status] = "MATCHED"
                and [DuplicateSheet] = false
        ),

    //===========================================================
    // Run cell comparison
    //===========================================================

    AddDifferences =
        Table.AddColumn(
            Comparable,
            "Differences",
            each
                fxCompareSheets(
                    [Data],
                    [Data2]
                )
        ),

    //===========================================================
    // Expand only actual differences
    //===========================================================

    ExpandDifferences =
        Table.ExpandTableColumn(
            AddDifferences,
            "Differences",
            {
                "Row",
                "Column",
                "Folder1Value",
                "Folder2Value",
                "Status"
            },
            {
                "Row",
                "Column",
                "Folder1Value",
                "Folder2Value",
                "CellStatus"
            }
        ),

    //===========================================================
    // Add Excel-style address
    //===========================================================

    ColumnLetter =
        (ColumnNumber as number) as text =>
        let
            Letters =
                List.Generate(
                    () =>
                        [
                            N = ColumnNumber,
                            Result = ""
                        ],

                    each [N] > 0,

                    each
                        [
                            N =
                                Number.IntegerDivide(
                                    [N] - 1,
                                    26
                                ),

                            Result =
                                Character.FromNumber(
                                    65
                                    +
                                    Number.Mod(
                                        [N] - 1,
                                        26
                                    )
                                )
                                & [Result]
                        ],

                    each [Result]
                ),

            Result =
                List.Last(Letters)
        in
            Result,

    AddCellAddress =
        Table.AddColumn(
            ExpandDifferences,
            "Cell",
            each
                ColumnLetter([Column])
                & Text.From([Row]),
            type text
        ),

    //===========================================================
    // Remove table payloads
    //===========================================================

    Result =
        Table.SelectColumns(
            AddCellAddress,
            {
                "FileKey",
                "SheetKey",
                "SheetName",
                "Cell",
                "Row",
                "Column",
                "Folder1Value",
                "Folder2Value",
                "CellStatus"
            }
        ),

    Sorted =
        Table.Sort(
            Result,
            {
                {"FileKey", Order.Ascending},
                {"SheetName", Order.Ascending},
                {"Row", Order.Ascending},
                {"Column", Order.Ascending}
            }
        )
in
    Sorted
```

The result becomes your detailed exception report:

| FileKey     | SheetName | Cell | Folder1Value | Folder2Value | CellStatus |
| ----------- | --------- | ---- | ------------ | ------------ | ---------- |
| RATES.XLSX  | Rates     | D15  | 0.035        | 0.0375       | DIFFERENT  |
| RATES.XLSX  | Rates     | H27  | 100          | 125          | DIFFERENT  |
| POLICY.XLSX | Benefits  | B42  | Basic        | Enhanced     | DIFFERENT  |

---

# 9. `Summary`

Finally create **`Summary`**.

```powerquery
let
    //===========================================================
    // FILE SUMMARY
    //===========================================================

    FileComparison =
        #"File Comparison",

    FileSummary =
        Table.Group(
            FileComparison,
            {"Status"},
            {
                {
                    "Count",
                    each Table.RowCount(_),
                    Int64.Type
                }
            }
        ),

    AddFileCategory =
        Table.AddColumn(
            FileSummary,
            "Category",
            each "FILE",
            type text
        ),

    RenameFileCount =
        Table.RenameColumns(
            AddFileCategory,
            {
                {"Count", "ExceptionCount"}
            }
        ),

    //===========================================================
    // SHEET SUMMARY
    //===========================================================

    SheetComparison =
        #"Sheet Comparison",

    SheetSummary =
        Table.Group(
            SheetComparison,
            {"Status"},
            {
                {
                    "Count",
                    each Table.RowCount(_),
                    Int64.Type
                }
            }
        ),

    AddSheetCategory =
        Table.AddColumn(
            SheetSummary,
            "Category",
            each "SHEET",
            type text
        ),

    RenameSheetCount =
        Table.RenameColumns(
            AddSheetCategory,
            {
                {"Count", "ExceptionCount"}
            }
        ),

    //===========================================================
    // CELL SUMMARY
    //===========================================================

    CellDifferences =
        #"Cell Differences",

    CellSummary =
        Table.Group(
            CellDifferences,
            {},
            {
                {
                    "Count",
                    each Table.RowCount(_),
                    Int64.Type
                }
            }
        ),

    AddCellStatus =
        Table.AddColumn(
            CellSummary,
            "Status",
            each "CELL DIFFERENCES",
            type text
        ),

    AddCellCategory =
        Table.AddColumn(
            AddCellStatus,
            "Category",
            each "CELL",
            type text
        ),

    RenameCellCount =
        Table.RenameColumns(
            AddCellCategory,
            {
                {"Count", "ExceptionCount"}
            }
        ),

    //===========================================================
    // COMBINE
    //===========================================================

    Combined =
        Table.Combine(
            {
                RenameFileCount,
                RenameSheetCount,
                RenameCellCount
            }
        ),

    Reorder =
        Table.SelectColumns(
            Combined,
            {
                "Category",
                "Status",
                "ExceptionCount"
            }
        ),

    Sort =
        Table.Sort(
            Reorder,
            {
                {"Category", Order.Ascending},
                {"Status", Order.Ascending}
            }
        ),

    //===========================================================
    // OVERALL STATUS
    //===========================================================

    FileExceptions =
        List.Sum(
            Table.SelectRows(
                RenameFileCount,
                each
                    [Status] <> "MATCHED"
            )[ExceptionCount]
        ),

    SheetExceptions =
        List.Sum(
            Table.SelectRows(
                RenameSheetCount,
                each
                    [Status] <> "MATCHED"
            )[ExceptionCount]
        ),

    CellExceptionCount =
        if Table.RowCount(CellDifferences) = 0
        then 0
        else Table.RowCount(CellDifferences),

    OverallStatus =
        if FileExceptions = 0
            and SheetExceptions = 0
            and CellExceptionCount = 0
        then
            "PASS"
        else
            "FAIL",

    AddOverall =
        Table.AddColumn(
            Sort,
            "OverallStatus",
            each OverallStatus,
            type text
        )
in
    AddOverall
```

---

## 10. Recommended workbook layout

I would set the workbook up like this:

```text
Comparison.xlsx

├── Parameters
│   └── tblParameters
│
├── Summary
│
├── File Comparison
│
├── Sheet Comparison
│
└── Cell Differences
```

The **inventory queries should be connection-only**:

```text
Parameters
File Inventory
Sheet Inventory
fxCompareSheets
```

Load only:

```text
Summary
File Comparison
Sheet Comparison
Cell Differences
```

---

## 11. What the finished system gives you

### File-level control

```text
PASS / FAIL
    │
    ├── Missing workbook
    ├── Duplicate workbook
    └── Matched workbook
```

### Sheet-level control

```text
Workbook
    │
    ├── Sheet exists in both
    ├── Sheet missing from Folder 1
    ├── Sheet missing from Folder 2
    └── Duplicate sheet
```

### Cell-level control

```text
Workbook
    └── Sheet
         └── Cell
              ├── Folder 1 value
              └── Folder 2 value
```

This gives you an auditable reconciliation hierarchy rather than simply a list of changed cells.

---

## 12. One improvement I strongly recommend

For large actuarial/modeling workbooks, I would make the comparison **configurable** rather than always comparing every physical cell.

For example, add parameters:

| Parameter            | Value          |
| -------------------- | -------------- |
| Folder1              | `C:\Model\Old` |
| Folder2              | `C:\Model\New` |
| CompareHiddenSheets  | `FALSE`        |
| IgnoreBlankCells     | `TRUE`         |
| IgnoreFormulaChanges | `FALSE`        |
| CompareFormulas      | `TRUE`         |
| CompareValues        | `TRUE`         |
| CaseSensitive        | `FALSE`        |

That would allow the engine to distinguish between:

```text
VALUE CHANGE
100 → 105
```

and:

```text
FORMULA CHANGE
=A1+B1 → =A1+B1+C1
```

It could also report:

```text
CELL CHANGED
FORMULA CHANGED
FORMAT CHANGED
SHEET ADDED
SHEET REMOVED
WORKBOOK ADDED
WORKBOOK REMOVED
```

That would be a more robust **Excel workbook reconciliation engine**, particularly if these files are actuarial models, rate tables, assumptions, or production model outputs.

The architecture above already establishes the right foundation for that extension.

[1]: https://berenguer-formation-conseil.fr/wp-content/uploads/2020/05/Langage-M-POwer-Query.pdf?utm_source=chatgpt.com "Power Query M function reference | Microsoft Docs"
[2]: https://github.com/MicrosoftDocs/powerquery-docs/blob/main/powerquery-docs/connectors/excel.md?utm_source=chatgpt.com "powerquery-docs/powerquery-docs/connectors/excel.md at main · MicrosoftDocs/powerquery-docs · GitHub"
[3]: https://learn.microsoft.com/es-es/powerquery-m/table-functions?utm_source=chatgpt.com "Funciones de tabla - PowerQuery M | Microsoft Learn"
