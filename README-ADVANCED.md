Make the parameterization more comprehensive and also correct one limitation in the prior design: **standard `Excel.Workbook()` exposes worksheet data/values, not the underlying cell formulas or formatting**, so `CompareFormulas` and `CompareFormatting` cannot honestly be implemented as native `Excel.Workbook` comparisons. Microsoft documents `Excel.Workbook` as returning workbook contents/data, with options such as `DelayTypes` and `InferSheetDimensions`; it does not expose a formula/formatting object model. ([Microsoft Learn][1])

The revised version below therefore implements the configurable **value comparison engine** fully, while adding configuration hooks for formula/format comparison and explicitly reporting those capabilities as unsupported rather than producing misleading results.

## Configuration

Use this Excel table:

**`tblParameters`**

| Parameter                | Value                |
| ------------------------ | -------------------- |
| Folder1                  | `C:\Model\Version_A` |
| Folder2                  | `C:\Model\Version_B` |
| CompareHiddenSheets      | `FALSE`              |
| IgnoreBlankCells         | `TRUE`               |
| CompareValues            | `TRUE`               |
| CompareFormulas          | `FALSE`              |
| CompareFormatting        | `FALSE`              |
| CaseSensitive            | `FALSE`              |
| TrimText                 | `TRUE`               |
| TreatNullAndBlankAsEqual | `TRUE`               |
| InferSheetDimensions     | `FALSE`              |
| FileExtensions           | `.xlsx,.xlsm,.xlsb`  |

This gives you a single control surface for the reconciliation.

---

# 1. `Parameters`

Replace the previous `Parameters` query with:

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
                {
                    "Parameter",
                    each Text.Trim(_),
                    type text
                },
                {
                    "Value",
                    each Text.Trim(_),
                    type text
                }
            }
        ),

    GetText =
        (ParameterName as text, optional Default as nullable text) as text =>
        let
            Matches =
                Table.SelectRows(
                    Clean,
                    each
                        Text.Upper([Parameter])
                        =
                        Text.Upper(ParameterName)
                )
        in
            if Table.RowCount(Matches) = 0 then
                if Default <> null then
                    Default
                else
                    error
                        "Missing parameter: "
                        & ParameterName
            else
                Matches{0}[Value],

    GetLogical =
        (
            ParameterName as text,
            optional Default as nullable logical
        ) as logical =>
        let
            Raw =
                GetText(
                    ParameterName,
                    if Default = null
                    then null
                    else if Default
                    then "TRUE"
                    else "FALSE"
                ),

            Result =
                if List.Contains(
                    {"TRUE", "YES", "Y", "1"},
                    Text.Upper(Raw)
                )
                then
                    true

                else if List.Contains(
                    {"FALSE", "NO", "N", "0"},
                    Text.Upper(Raw)
                )
                then
                    false

                else
                    error
                        "Parameter '"
                        & ParameterName
                        & "' must be TRUE/FALSE."
        in
            Result,

    Folder1 =
        GetText("Folder1"),

    Folder2 =
        GetText("Folder2"),

    CompareHiddenSheets =
        GetLogical(
            "CompareHiddenSheets",
            false
        ),

    IgnoreBlankCells =
        GetLogical(
            "IgnoreBlankCells",
            true
        ),

    CompareValues =
        GetLogical(
            "CompareValues",
            true
        ),

    CompareFormulas =
        GetLogical(
            "CompareFormulas",
            false
        ),

    CompareFormatting =
        GetLogical(
            "CompareFormatting",
            false
        ),

    CaseSensitive =
        GetLogical(
            "CaseSensitive",
            false
        ),

    TrimText =
        GetLogical(
            "TrimText",
            true
        ),

    TreatNullAndBlankAsEqual =
        GetLogical(
            "TreatNullAndBlankAsEqual",
            true
        ),

    InferSheetDimensions =
        GetLogical(
            "InferSheetDimensions",
            false
        ),

    ExtensionText =
        GetText(
            "FileExtensions",
            ".xlsx,.xlsm,.xlsb"
        ),

    FileExtensions =
        List.Transform(
            Text.Split(
                ExtensionText,
                ","
            ),
            each
                Text.Lower(
                    Text.Trim(_)
                )
        ),

    Output =
        [
            Folder1 = Folder1,
            Folder2 = Folder2,
            CompareHiddenSheets = CompareHiddenSheets,
            IgnoreBlankCells = IgnoreBlankCells,
            CompareValues = CompareValues,
            CompareFormulas = CompareFormulas,
            CompareFormatting = CompareFormatting,
            CaseSensitive = CaseSensitive,
            TrimText = TrimText,
            TreatNullAndBlankAsEqual =
                TreatNullAndBlankAsEqual,
            InferSheetDimensions =
                InferSheetDimensions,
            FileExtensions = FileExtensions
        ]
in
    Output
```

---

# 2. `File Inventory`

This version avoids unnecessary workbook processing and buffers the file binary.

```powerquery
let
    Folder1 =
        Parameters[Folder1],

    Folder2 =
        Parameters[Folder2],

    Extensions =
        Parameters[FileExtensions],

    GetFiles =
        (
            FolderPath as text,
            SourceID as text
        ) as table =>
        let
            Source =
                Folder.Files(FolderPath),

            Filtered =
                Table.SelectRows(
                    Source,
                    each
                        List.Contains(
                            Extensions,
                            Text.Lower([Extension])
                        )
                        and
                        not Text.StartsWith(
                            [Name],
                            "~$"
                        )
                ),

            Selected =
                Table.SelectColumns(
                    Filtered,
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

            BufferContent =
                Table.TransformColumns(
                    AddFileKey,
                    {
                        {
                            "Content",
                            Binary.Buffer,
                            type binary
                        }
                    }
                )
        in
            BufferContent,

    Folder1Files =
        GetFiles(
            Folder1,
            "Folder 1"
        ),

    Folder2Files =
        GetFiles(
            Folder2,
            "Folder 2"
        ),

    Combined =
        Table.Combine(
            {
                Folder1Files,
                Folder2Files
            }
        ),

    DuplicateGroups =
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

    AddDuplicateFlag =
        Table.AddColumn(
            Combined,
            "DuplicateFile",
            each
                let
                    SourceID = [Source],
                    Key = [FileKey],
                    Match =
                        Table.SelectRows(
                            DuplicateGroups,
                            each
                                [Source] = SourceID
                                and
                                [FileKey] = Key
                        )
                in
                    if Table.RowCount(Match) = 0
                    then false
                    else Match{0}[FileCount] > 1,
            type logical
        )
in
    AddDuplicateFlag
```

### Optimization

`Binary.Buffer` is intentional here. It prevents the binary content from being repeatedly retrieved from the filesystem while downstream queries evaluate the same workbook.

---

# 3. `File Comparison`

```powerquery
let
    Source =
        #"File Inventory",

    F1 =
        Table.SelectRows(
            Source,
            each [Source] = "Folder 1"
        ),

    F2 =
        Table.SelectRows(
            Source,
            each [Source] = "Folder 2"
        ),

    F1Keys =
        List.Buffer(
            F1[FileKey]
        ),

    F2Keys =
        List.Buffer(
            F2[FileKey]
        ),

    AllKeys =
        List.Sort(
            List.Distinct(
                List.Combine(
                    {
                        F1Keys,
                        F2Keys
                    }
                )
            )
        ),

    KeyTable =
        Table.FromList(
            AllKeys,
            Splitter.SplitByNothing(),
            {"FileKey"}
        ),

    JoinF1 =
        Table.NestedJoin(
            KeyTable,
            {"FileKey"},
            F1,
            {"FileKey"},
            "F1",
            JoinKind.LeftOuter
        ),

    JoinF2 =
        Table.NestedJoin(
            JoinF1,
            {"FileKey"},
            F2,
            {"FileKey"},
            "F2",
            JoinKind.LeftOuter
        ),

    ExpandF1 =
        Table.ExpandTableColumn(
            JoinF2,
            "F1",
            {
                "Name",
                "Date modified",
                "Folder Path",
                "DuplicateFile"
            },
            {
                "Folder1FileName",
                "Folder1Modified",
                "Folder1Path",
                "Folder1Duplicate"
            }
        ),

    ExpandF2 =
        Table.ExpandTableColumn(
            ExpandF1,
            "F2",
            {
                "Name",
                "Date modified",
                "Folder Path",
                "DuplicateFile"
            },
            {
                "Folder2FileName",
                "Folder2Modified",
                "Folder2Path",
                "Folder2Duplicate"
            }
        ),

    AddStatus =
        Table.AddColumn(
            ExpandF2,
            "Status",
            each
                if [Folder1FileName] = null then
                    "MISSING FROM FOLDER 1"

                else if [Folder2FileName] = null then
                    "MISSING FROM FOLDER 2"

                else if [Folder1Duplicate] then
                    "DUPLICATE IN FOLDER 1"

                else if [Folder2Duplicate] then
                    "DUPLICATE IN FOLDER 2"

                else
                    "MATCHED",
            type text
        ),

    AddFileDifference =
        Table.AddColumn(
            AddStatus,
            "FileModifiedDifference",
            each
                if [Folder1Modified] = null
                    or [Folder2Modified] = null
                then
                    null
                else
                    [Folder1Modified]
                    <> [Folder2Modified],
            type logical
        )
in
    AddFileDifference
```

---

# 4. `Sheet Inventory`

This is where the hidden-sheet and dimension controls are applied.

```powerquery
let
    Source =
        #"File Inventory",

    CompareHiddenSheets =
        Parameters[CompareHiddenSheets],

    InferSheetDimensions =
        Parameters[InferSheetDimensions],

    ValidFiles =
        Table.SelectRows(
            Source,
            each
                [DuplicateFile] = false
        ),

    AddWorkbook =
        Table.AddColumn(
            ValidFiles,
            "Workbook",
            each
                try
                    Excel.Workbook(
                        [Content],
                        [
                            UseHeaders = false,
                            DelayTypes = true,
                            InferSheetDimensions =
                                InferSheetDimensions
                        ]
                    )
                otherwise
                    null
        ),

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
                        each
                            [Kind] = "Sheet"
                            and
                            (
                                CompareHiddenSheets
                                or
                                [Hidden] <> true
                            )
                    )
        ),

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

    SelectColumns =
        Table.SelectColumns(
            AddSheetKey,
            {
                "Source",
                "FileKey",
                "SheetName",
                "SheetKey",
                "Data",
                "Hidden"
            }
        )
in
    SelectColumns
```

`InferSheetDimensions` is useful for generated workbooks whose stored worksheet dimensions are inaccurate; Microsoft specifically documents this option for Open XML workbooks. ([Microsoft Learn][2])

---

# 5. `Sheet Comparison`

I would simplify the sheet comparison and make it more efficient than the previous version.

```powerquery
let
    Source =
        #"Sheet Inventory",

    F1 =
        Table.SelectRows(
            Source,
            each [Source] = "Folder 1"
        ),

    F2 =
        Table.SelectRows(
            Source,
            each [Source] = "Folder 2"
        ),

    F1Keys =
        List.Buffer(
            Table.AddColumn(
                F1,
                "JoinKey",
                each
                    [FileKey]
                    & "¦"
                    & [SheetKey]
            )[JoinKey]
        ),

    F2Keys =
        List.Buffer(
            Table.AddColumn(
                F2,
                "JoinKey",
                each
                    [FileKey]
                    & "¦"
                    & [SheetKey]
            )[JoinKey]
        ),

    AddJoinKey1 =
        Table.AddColumn(
            F1,
            "JoinKey",
            each
                [FileKey]
                & "¦"
                & [SheetKey],
            type text
        ),

    AddJoinKey2 =
        Table.AddColumn(
            F2,
            "JoinKey",
            each
                [FileKey]
                & "¦"
                & [SheetKey],
            type text
        ),

    AllKeys =
        List.Sort(
            List.Distinct(
                List.Combine(
                    {
                        F1Keys,
                        F2Keys
                    }
                )
            )
        ),

    KeyTable =
        Table.FromList(
            AllKeys,
            Splitter.SplitByNothing(),
            {"JoinKey"}
        ),

    JoinF1 =
        Table.NestedJoin(
            KeyTable,
            {"JoinKey"},
            AddJoinKey1,
            {"JoinKey"},
            "F1",
            JoinKind.LeftOuter
        ),

    JoinF2 =
        Table.NestedJoin(
            JoinF1,
            {"JoinKey"},
            AddJoinKey2,
            {"JoinKey"},
            "F2",
            JoinKind.LeftOuter
        ),

    ExpandF1 =
        Table.ExpandTableColumn(
            JoinF2,
            "F1",
            {
                "FileKey",
                "SheetKey",
                "SheetName",
                "Data",
                "Hidden"
            },
            {
                "FileKey",
                "SheetKey",
                "Folder1SheetName",
                "Data1",
                "Folder1Hidden"
            }
        ),

    ExpandF2 =
        Table.ExpandTableColumn(
            ExpandF1,
            "F2",
            {
                "SheetName",
                "Data",
                "Hidden"
            },
            {
                "Folder2SheetName",
                "Data2",
                "Folder2Hidden"
            }
        ),

    AddStatus =
        Table.AddColumn(
            ExpandF2,
            "Status",
            each
                if [Folder1SheetName] = null then
                    "MISSING FROM FOLDER 1"

                else if [Folder2SheetName] = null then
                    "MISSING FROM FOLDER 2"

                else
                    "MATCHED",
            type text
        ),

    AddSheetNameDifference =
        Table.AddColumn(
            AddStatus,
            "SheetNameDifference",
            each
                if [Folder1SheetName] = null
                    or [Folder2SheetName] = null
                then
                    null
                else
                    [Folder1SheetName]
                    <> [Folder2SheetName],
            type logical
        )
in
    AddSheetNameDifference
```

---

# 6. `fxNormalizeValue`

This function centralizes the comparison rules.

Create a query named:

**`fxNormalizeValue`**

```powerquery
(
    Value as any
)
as any =>
let
    TrimText =
        Parameters[TrimText],

    CaseSensitive =
        Parameters[CaseSensitive],

    TreatNullAndBlankAsEqual =
        Parameters[TreatNullAndBlankAsEqual],

    Result =
        if Value = null then
            null

        else if Value is text then
            let
                Trimmed =
                    if TrimText
                    then Text.Trim(Value)
                    else Value,

                Normalized =
                    if not CaseSensitive
                    then Text.Upper(Trimmed)
                    else Trimmed
            in
                Normalized

        else
            Value
in
    Result
```

This means:

```text
"ABC"
"abc"
" ABC "
```

are treated as equal when:

```text
CaseSensitive = FALSE
TrimText = TRUE
```

---

# 7. `fxCompareSheets`

Replace the earlier function with this parameter-aware version.

```powerquery
(
    Table1 as nullable table,
    Table2 as nullable table
)
as table =>
let

    IgnoreBlankCells =
        Parameters[IgnoreBlankCells],

    CompareValues =
        Parameters[CompareValues],

    TreatNullAndBlankAsEqual =
        Parameters[TreatNullAndBlankAsEqual],

    Empty =
        #table(
            {
                "Row",
                "Column",
                "Folder1Value",
                "Folder2Value",
                "DifferenceType"
            },
            {}
        ),

    Normalize =
        (Value as any) as any =>
            fxNormalizeValue(Value),

    IsBlank =
        (Value as any) as logical =>
            Value = null
            or
            (
                Value is text
                and
                Text.Trim(Value) = ""
            ),

    ValuesEqual =
        (V1 as any, V2 as any) as logical =>
        let
            N1 =
                Normalize(V1),

            N2 =
                Normalize(V2),

            Equal =
                if TreatNullAndBlankAsEqual then
                    (
                        IsBlank(N1)
                        and
                        IsBlank(N2)
                    )
                    or
                    (
                        N1 <> null
                        and
                        N2 <> null
                        and
                        Value.Equals(N1, N2)
                    )

                else
                    Value.Equals(
                        N1,
                        N2
                    )
        in
            Equal,

    Result =
        if Table1 = null
            or Table2 = null
            or not CompareValues
        then
            Empty

        else
            let

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

                        Results =
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

                                    Blank =
                                        IsBlank(V1)
                                        and
                                        IsBlank(V2),

                                    Equal =
                                        ValuesEqual(
                                            V1,
                                            V2
                                        )
                                in

                                    if
                                        IgnoreBlankCells
                                        and
                                        Blank
                                    then
                                        null

                                    else if Equal then
                                        null

                                    else
                                        {
                                            RowNumber,
                                            ColumnIndex + 1,
                                            V1,
                                            V2,
                                            "VALUE DIFFERENCE"
                                        }
                            ),

                        NonNull =
                            List.RemoveNulls(
                                Results
                            )
                    in
                        NonNull,

                AllDifferences =
                    if RowCount = 0
                    then
                        {}
                    else
                        List.Combine(
                            List.Transform(
                                {1..RowCount},
                                each CompareRow(_)
                            )
                        ),

                Output =
                    if List.IsEmpty(
                        AllDifferences
                    )
                    then
                        Empty
                    else
                        #table(
                            {
                                "Row",
                                "Column",
                                "Folder1Value",
                                "Folder2Value",
                                "DifferenceType"
                            },
                            AllDifferences
                        )
            in
                Output
in
    Result
```

---

# 8. `Cell Differences`

```powerquery
let
    Source =
        #"Sheet Comparison",

    Comparable =
        Table.SelectRows(
            Source,
            each [Status] = "MATCHED"
        ),

    AddComparison =
        Table.AddColumn(
            Comparable,
            "Differences",
            each
                fxCompareSheets(
                    [Data1],
                    [Data2]
                )
        ),

    Expand =
        Table.ExpandTableColumn(
            AddComparison,
            "Differences",
            {
                "Row",
                "Column",
                "Folder1Value",
                "Folder2Value",
                "DifferenceType"
            },
            {
                "Row",
                "Column",
                "Folder1Value",
                "Folder2Value",
                "DifferenceType"
            }
        ),

    ColumnLetter =
        (N as number) as text =>
        let
            Convert =
                (Number as number) as text =>
                    if Number <= 26 then
                        Character.FromNumber(
                            64 + Number
                        )
                    else
                        @Convert(
                            Number.IntegerDivide(
                                Number - 1,
                                26
                            )
                        )
                        &
                        Character.FromNumber(
                            65
                            +
                            Number.Mod(
                                Number - 1,
                                26
                            )
                        )
        in
            Convert(N),

    AddCell =
        Table.AddColumn(
            Expand,
            "Cell",
            each
                ColumnLetter([Column])
                &
                Text.From([Row]),
            type text
        ),

    Result =
        Table.SelectColumns(
            AddCell,
            {
                "FileKey",
                "SheetKey",
                "Folder1SheetName",
                "Folder2SheetName",
                "Cell",
                "Row",
                "Column",
                "Folder1Value",
                "Folder2Value",
                "DifferenceType"
            }
        )
in
    Result
```

---

# 9. Add a `Configuration` query

I recommend making the current configuration visible in Excel.

Create:

**`Configuration`**

```powerquery
let
    P =
        Parameters,

    Output =
        #table(
            {
                "Parameter",
                "Value"
            },
            {
                {"Folder 1", P[Folder1]},
                {"Folder 2", P[Folder2]},
                {
                    "Compare Hidden Sheets",
                    Text.From(
                        P[CompareHiddenSheets]
                    )
                },
                {
                    "Ignore Blank Cells",
                    Text.From(
                        P[IgnoreBlankCells]
                    )
                },
                {
                    "Compare Values",
                    Text.From(
                        P[CompareValues]
                    )
                },
                {
                    "Compare Formulas",
                    Text.From(
                        P[CompareFormulas]
                    )
                },
                {
                    "Compare Formatting",
                    Text.From(
                        P[CompareFormatting]
                    )
                },
                {
                    "Case Sensitive",
                    Text.From(
                        P[CaseSensitive]
                    )
                },
                {
                    "Trim Text",
                    Text.From(
                        P[TrimText]
                    )
                },
                {
                    "Null = Blank",
                    Text.From(
                        P[TreatNullAndBlankAsEqual]
                    )
                },
                {
                    "Infer Sheet Dimensions",
                    Text.From(
                        P[InferSheetDimensions]
                    )
                }
            }
        )
in
    Output
```

Load this to a worksheet alongside `Summary`.

---

# 10. Revised `Summary`

This version provides a real reconciliation status.

```powerquery
let
    Files =
        #"File Comparison",

    Sheets =
        #"Sheet Comparison",

    Cells =
        #"Cell Differences",

    //===========================================================
    // File metrics
    //===========================================================

    TotalFiles =
        Table.RowCount(Files),

    MissingFiles =
        Table.RowCount(
            Table.SelectRows(
                Files,
                each
                    [Status]
                    = "MISSING FROM FOLDER 1"
                    or
                    [Status]
                    = "MISSING FROM FOLDER 2"
            )
        ),

    DuplicateFiles =
        Table.RowCount(
            Table.SelectRows(
                Files,
                each
                    Text.StartsWith(
                        [Status],
                        "DUPLICATE"
                    )
            )
        ),

    MatchedFiles =
        Table.RowCount(
            Table.SelectRows(
                Files,
                each
                    [Status] = "MATCHED"
            )
        ),

    //===========================================================
    // Sheet metrics
    //===========================================================

    TotalSheets =
        Table.RowCount(Sheets),

    MissingSheets =
        Table.RowCount(
            Table.SelectRows(
                Sheets,
                each
                    [Status]
                    = "MISSING FROM FOLDER 1"
                    or
                    [Status]
                    = "MISSING FROM FOLDER 2"
            )
        ),

    MatchedSheets =
        Table.RowCount(
            Table.SelectRows(
                Sheets,
                each
                    [Status] = "MATCHED"
            )
        ),

    //===========================================================
    // Cell metrics
    //===========================================================

    CellDifferences =
        Table.RowCount(Cells),

    AffectedSheets =
        Table.RowCount(
            Table.Distinct(
                Table.SelectColumns(
                    Cells,
                    {
                        "FileKey",
                        "SheetKey"
                    }
                )
            )
        ),

    AffectedFiles =
        Table.RowCount(
            Table.Distinct(
                Table.SelectColumns(
                    Cells,
                    {
                        "FileKey"
                    }
                )
            )
        ),

    //===========================================================
    // Overall status
    //===========================================================

    OverallStatus =
        if MissingFiles = 0
            and DuplicateFiles = 0
            and MissingSheets = 0
            and CellDifferences = 0
        then
            "PASS"
        else
            "FAIL",

    Output =
        #table(
            {
                "Metric",
                "Value"
            },
            {
                {
                    "Overall Status",
                    OverallStatus
                },
                {
                    "Total Workbooks",
                    TotalFiles
                },
                {
                    "Matched Workbooks",
                    MatchedFiles
                },
                {
                    "Missing Workbooks",
                    MissingFiles
                },
                {
                    "Duplicate Workbooks",
                    DuplicateFiles
                },
                {
                    "Total Sheets",
                    TotalSheets
                },
                {
                    "Matched Sheets",
                    MatchedSheets
                },
                {
                    "Missing Sheets",
                    MissingSheets
                },
                {
                    "Cell Differences",
                    CellDifferences
                },
                {
                    "Files With Differences",
                    AffectedFiles
                },
                {
                    "Sheets With Differences",
                    AffectedSheets
                }
            }
        )
in
    Output
```

---

## Configuration behavior

The parameters now have clear operational semantics.

### `CompareHiddenSheets`

```text
TRUE
```

Compares visible **and hidden** worksheets.

```text
FALSE
```

Ignores hidden worksheets.

---

### `IgnoreBlankCells`

```text
TRUE
```

This will not report:

```text
Folder 1: null
Folder 2: null
```

or, depending on `TreatNullAndBlankAsEqual`:

```text
Folder 1: null
Folder 2: ""
```

This is particularly valuable for large worksheets with thousands of unused cells.

---

### `TreatNullAndBlankAsEqual`

Controls whether:

```text
null
""
"   "
```

are treated as equivalent.

With:

```text
TrimText = TRUE
TreatNullAndBlankAsEqual = TRUE
```

the comparison is deliberately tolerant of insignificant whitespace.

---

### `CaseSensitive`

Controls:

```text
ABC
```

versus:

```text
abc
```

With `FALSE`, they match.

With `TRUE`, they don't.

---

### `TrimText`

Controls:

```text
"ABC"
```

versus:

```text
" ABC "
```

With `TRUE`, they match.

---

### `CompareValues`

This is the actual cell-level reconciliation switch.

If:

```text
CompareValues = FALSE
```

the workbook/sheet structure is still analyzed, but cell values aren't compared.

---

### `InferSheetDimensions`

Set this to:

```text
TRUE
```

when you are dealing with Excel files generated by other systems that sometimes have incorrect worksheet dimension metadata. Microsoft specifically provides this option for that scenario. ([Microsoft Learn][2])

Don't turn it on by default for very large workbooks unless necessary, because inspecting worksheet dimensions can increase processing time.

---

# Formula comparison: an important limitation

I would **not** implement `CompareFormulas = TRUE` using the existing `Excel.Workbook` pipeline.

The reason is architectural: Power Query's Excel connector gives you the worksheet's imported data, while formula inspection is an Excel object-model operation. Microsoft documents `Excel.Workbook` as returning workbook contents and worksheet data; it doesn't provide a cell-level formula property. ([Microsoft Learn][1])

Consequently, this:

```text
Workbook A

B5 = A5 * 1.05
```

and:

```text
Workbook B

B5 = A5 + (A5 * 5%)
```

can produce the same value:

```text
105
```

while having different formulas.

The current engine correctly reports:

```text
VALUE: MATCH
```

It **cannot legitimately report**:

```text
FORMULA: DIFFERENT
```

using only `Excel.Workbook`.

Likewise, standard Power Query ingestion doesn't expose Excel's cell formatting object model, so `CompareFormatting` is intentionally a reserved parameter rather than pretending to perform a formatting comparison.

---

# Recommended production configuration

For an actuarial/model reconciliation use case, I'd start with:

| Parameter                  | Recommended |
| -------------------------- | ----------: |
| `CompareHiddenSheets`      |      `TRUE` |
| `IgnoreBlankCells`         |      `TRUE` |
| `CompareValues`            |      `TRUE` |
| `CompareFormulas`          |     `FALSE` |
| `CompareFormatting`        |     `FALSE` |
| `CaseSensitive`            |     `FALSE` |
| `TrimText`                 |      `TRUE` |
| `TreatNullAndBlankAsEqual` |      `TRUE` |
| `InferSheetDimensions`     |     `FALSE` |

This provides a relatively high-signal reconciliation without generating millions of false positives from blank cells or insignificant text differences.

The final workbook then has this control model:

```text
                    tblParameters
                         │
                         ▼
                    Parameters
                         │
          ┌──────────────┴──────────────┐
          ▼                             ▼
   Folder 1 files                 Folder 2 files
          │                             │
          └──────────────┬──────────────┘
                         ▼
                  File Comparison
                         │
                         ▼
                  Sheet Inventory
                         │
                         ▼
                  Sheet Comparison
                         │
                         ▼
                 fxCompareSheets
                         │
                         ▼
                  Cell Differences
                         │
                         ▼
                      Summary
```

This is substantially more scalable than the original version because **file discovery, workbook extraction, sheet matching, and cell comparison are separated into distinct stages**, while all comparison behavior is controlled from `tblParameters`.

[1]: https://learn.microsoft.com/en-us/powerquery-m/excel-workbook?utm_source=chatgpt.com "Excel.Workbook - PowerQuery M | Microsoft Learn"
[2]: https://learn.microsoft.com/en-us/power-query/connectors/excel?utm_source=chatgpt.com "Power Query Excel connector - Power Query | Microsoft Learn"


# Updates
Remove both parameters entirely. The streamlined configuration is:

| Parameter                | Value                |
| ------------------------ | -------------------- |
| Folder1                  | `C:\Model\Version_A` |
| Folder2                  | `C:\Model\Version_B` |
| CompareHiddenSheets      | `FALSE`              |
| IgnoreBlankCells         | `TRUE`               |
| CompareValues            | `TRUE`               |
| CaseSensitive            | `FALSE`              |
| TrimText                 | `TRUE`               |
| TreatNullAndBlankAsEqual | `TRUE`               |
| InferSheetDimensions     | `FALSE`              |
| FileExtensions           | `.xlsx,.xlsm,.xlsb`  |

### Updated `Parameters`

Use this version in place of the previous one:

```powerquery id="6n1p4v"
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
                {
                    "Parameter",
                    each Text.Trim(_),
                    type text
                },
                {
                    "Value",
                    each Text.Trim(_),
                    type text
                }
            }
        ),

    GetText =
        (ParameterName as text, optional Default as nullable text) as text =>
        let
            Matches =
                Table.SelectRows(
                    Clean,
                    each
                        Text.Upper([Parameter])
                        =
                        Text.Upper(ParameterName)
                )
        in
            if Table.RowCount(Matches) = 0 then
                if Default <> null then
                    Default
                else
                    error
                        "Missing parameter: "
                        & ParameterName
            else
                Matches{0}[Value],

    GetLogical =
        (
            ParameterName as text,
            optional Default as nullable logical
        ) as logical =>
        let
            Raw =
                GetText(
                    ParameterName,
                    if Default = null
                    then null
                    else if Default
                    then "TRUE"
                    else "FALSE"
                ),

            Upper =
                Text.Upper(Raw)
        in
            if List.Contains(
                {"TRUE", "YES", "Y", "1"},
                Upper
            )
            then
                true

            else if List.Contains(
                {"FALSE", "NO", "N", "0"},
                Upper
            )
            then
                false

            else
                error
                    "Parameter '"
                    & ParameterName
                    & "' must be TRUE or FALSE.",

    Folder1 =
        GetText("Folder1"),

    Folder2 =
        GetText("Folder2"),

    CompareHiddenSheets =
        GetLogical(
            "CompareHiddenSheets",
            false
        ),

    IgnoreBlankCells =
        GetLogical(
            "IgnoreBlankCells",
            true
        ),

    CompareValues =
        GetLogical(
            "CompareValues",
            true
        ),

    CaseSensitive =
        GetLogical(
            "CaseSensitive",
            false
        ),

    TrimText =
        GetLogical(
            "TrimText",
            true
        ),

    TreatNullAndBlankAsEqual =
        GetLogical(
            "TreatNullAndBlankAsEqual",
            true
        ),

    InferSheetDimensions =
        GetLogical(
            "InferSheetDimensions",
            false
        ),

    ExtensionText =
        GetText(
            "FileExtensions",
            ".xlsx,.xlsm,.xlsb"
        ),

    FileExtensions =
        List.Transform(
            Text.Split(
                ExtensionText,
                ","
            ),
            each
                Text.Lower(
                    Text.Trim(_)
                )
        ),

    Output =
        [
            Folder1 = Folder1,
            Folder2 = Folder2,
            CompareHiddenSheets =
                CompareHiddenSheets,
            IgnoreBlankCells =
                IgnoreBlankCells,
            CompareValues =
                CompareValues,
            CaseSensitive =
                CaseSensitive,
            TrimText =
                TrimText,
            TreatNullAndBlankAsEqual =
                TreatNullAndBlankAsEqual,
            InferSheetDimensions =
                InferSheetDimensions,
            FileExtensions =
                FileExtensions
        ]
in
    Output
```

### Updated `Configuration`

The visible configuration query should also remove them:

```powerquery id="ypj6cw"
let
    P =
        Parameters,

    Output =
        #table(
            {
                "Parameter",
                "Value"
            },
            {
                {
                    "Folder 1",
                    P[Folder1]
                },
                {
                    "Folder 2",
                    P[Folder2]
                },
                {
                    "Compare Hidden Sheets",
                    Text.From(
                        P[CompareHiddenSheets]
                    )
                },
                {
                    "Ignore Blank Cells",
                    Text.From(
                        P[IgnoreBlankCells]
                    )
                },
                {
                    "Compare Values",
                    Text.From(
                        P[CompareValues]
                    )
                },
                {
                    "Case Sensitive",
                    Text.From(
                        P[CaseSensitive]
                    )
                },
                {
                    "Trim Text",
                    Text.From(
                        P[TrimText]
                    )
                },
                {
                    "Null = Blank",
                    Text.From(
                        P[TreatNullAndBlankAsEqual]
                    )
                },
                {
                    "Infer Sheet Dimensions",
                    Text.From(
                        P[InferSheetDimensions]
                    )
                },
                {
                    "File Extensions",
                    Text.Combine(
                        P[FileExtensions],
                        ", "
                    )
                }
            }
        )
in
    Output
```

### Updated `tblParameters`

The Excel table should now contain only:

```text
Parameter                  Value
-------------------------  -------------------------
Folder1                    C:\Model\Version_A
Folder2                    C:\Model\Version_B
CompareHiddenSheets        FALSE
IgnoreBlankCells           TRUE
CompareValues              TRUE
CaseSensitive              FALSE
TrimText                   TRUE
TreatNullAndBlankAsEqual   TRUE
InferSheetDimensions       FALSE
FileExtensions             .xlsx,.xlsm,.xlsb
```

No other queries need to reference `CompareFormulas` or `CompareFormatting`, so the remaining `File Inventory`, `File Comparison`, `Sheet Inventory`, `Sheet Comparison`, `fxNormalizeValue`, `fxCompareSheets`, `Cell Differences`, and `Summary` queries can remain as previously provided.


