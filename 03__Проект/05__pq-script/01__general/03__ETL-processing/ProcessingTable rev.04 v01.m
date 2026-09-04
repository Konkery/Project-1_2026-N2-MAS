// =============================================================================================
//
// ETL-ПРОЦЕССОР 'ProcessingTable'
//
// версия rev.04 v01
// =============================================================================================

// ---------------------------------------------------------------------------------------------
// Процессор 'ProcessingTable' выполняет комплексную обработку одной или нескольких таблиц Excel
// на основе правил, описанных в JSON конфигурации. Поддерживает слияние и объединение таблиц.
//
// аргументы:
// _configJSONIndex as text        - Текстовый идентификатор JSON-конфигуратора в таблице 'TabConfigJSON'. Обязательный.
// optional _pathExcelBook_1 as text - Полный путь к файлу Excel для ПЕРВОЙ таблицы в JSON.
// optional _pathExcelBook_2 as text - Полный путь к файлу Excel для ВТОРОЙ таблицы в JSON.
// optional _pathExcelBook_3 as text - Полный путь к файлу Excel для ТРЕТЬЕЙ таблицы в JSON.
// optional _pathExcelBook_4 as text - Полный путь к файлу Excel для ЧЕТВЕРТОЙ таблицы в JSON.
//                                     Если путь не указан, поиск таблиц производится в текущей книге.
//
// возвращаемое значение:
// table                           - Обработанная и/или объединенная таблица.
// ---------------------------------------------------------------------------------------------
(
    _configJSONIndex as text,
    optional _pathExcelBook_1 as text,
    optional _pathExcelBook_2 as text,
    optional _pathExcelBook_3 as text,
    optional _pathExcelBook_4 as text
) as table =>
let
    // =============================================================================================
    // ШАГ 1: ПОИСК И ЗАГРУЗКА JSON-КОНФИГУРАТОРА
    // =============================================================================================

    ConfigTableName = "TabConfigJSON",
    ConfigLookupColumn = "IDconfig",
    ConfigJsonColumn = "CONFIG_JSON",
    ConfigWorkbook = Excel.CurrentWorkbook(),
    configTableSource = try ConfigWorkbook{[Name=ConfigTableName]}[Content]
                        otherwise error "Таблица с конфигурациями '" & ConfigTableName & "' не найдена в ТЕКУЩЕЙ книге.",

    validatedConfigTable = if not Table.HasColumns(configTableSource, {ConfigLookupColumn, ConfigJsonColumn}) then
                               error "В таблице конфигурации '" & ConfigTableName & "' отсутствуют обязательные столбцы '" & ConfigLookupColumn & "' и/или '" & ConfigJsonColumn & "'."
                           else
                               configTableSource,
    MatchingConfigRow = Table.SelectRows(validatedConfigTable, each Record.Field(_, ConfigLookupColumn) = _configJSONIndex),
    CheckUniqueness = if Table.RowCount(MatchingConfigRow) = 0 then
        error "В таблице '" & ConfigTableName & "' не найдена конфигурация с ID '" & _configJSONIndex & "'."
    else if Table.RowCount(MatchingConfigRow) > 1 then
        error "В таблице '" & ConfigTableName & "' найдено несколько конфигураций с ID '" & _configJSONIndex & "'. Требуется одна уникальная запись."
    else
        MatchingConfigRow,
    JsonConfigText = Record.Field(CheckUniqueness{0}, ConfigJsonColumn),

    // =============================================================================================
    // ШАГ 2: ПАРСИНГ JSON И ПРЕДВАРИТЕЛЬНАЯ ВАЛИДАЦИЯ КОНФИГУРАЦИИ
    // =============================================================================================

    ConfigRecord = Json.Document(JsonConfigText),
    TableNamesInJson = Record.FieldNames(ConfigRecord),
    TableCount = List.Count(TableNamesInJson),
    
    ConfigValidation = if TableCount > 4 then error "Конфигурация не может содержать более 4 таблиц." else
        let
            JoinCommands = {"RunLeftJOINstructure", "RunRightJOINstructure", "RunJOINdata"},
            FirstTableRulesRecord = Record.Field(ConfigRecord, TableNamesInJson{0}),
            AllMacroBlocksFirstTable = List.Transform(Record.FieldValues(FirstTableRulesRecord), each Record.FieldNames(_)),
            FoundJoinsInFirstTable = List.Combine(List.Transform(AllMacroBlocksFirstTable, each List.Intersect({_, JoinCommands}))),
            
            OtherTableRules = if TableCount > 1 then List.Transform(List.Range(TableNamesInJson, 1), (tableName) => Record.Field(ConfigRecord, tableName)) else {},
            AllMacroBlocksOtherTables = if TableCount > 1 then List.Combine(List.Transform(List.Transform(OtherTableRules, each Record.FieldValues(_)), (allBlocks) => List.Transform(allBlocks, each Record.FieldNames(_)))) else {},
            ForbiddenCommands = JoinCommands & {"ValidationData"},
            FoundForbiddenInOthers = List.Intersect({AllMacroBlocksOtherTables, ForbiddenCommands})
        in
            if TableCount > 1 and List.Count(FoundJoinsInFirstTable) = 0 then error "Многотабличная конфигурация должна содержать ровно одну команду слияния (Run...JOIN...) в секциях первой таблицы."
            else if TableCount > 1 and List.Count(FoundJoinsInFirstTable) > 1 then error "Многотабличная конфигурация не может содержать более одной команды слияния."
            else if TableCount = 1 and List.Count(FoundJoinsInFirstTable) > 0 then error "Однотабличная конфигурация не может содержать команды слияния."
            else if List.IsEmpty(FoundForbiddenInOthers) = false then error "Команды слияния и 'ValidationData' могут быть указаны только в секциях ПЕРВОЙ таблицы в JSON."
            else if List.ContainsAny(FoundJoinsInFirstTable, {"RunLeftJOINstructure", "RunRightJOINstructure"}) and Record.HasFields(FirstTableRulesRecord, "ValidationData") then error "Команда 'ValidationData' не может использоваться совместно с 'RunLeftJOINstructure' или 'RunRightJOINstructure'."
            else "OK",

    // =============================================================================================
    // ШАГ 3: ЗАГРУЗКА И ИНДИВИДУАЛЬНАЯ ОБРАБОТКА КАЖДОЙ ТАБЛИЦЫ
    // =============================================================================================

    PathsList = {_pathExcelBook_1, _pathExcelBook_2, _pathExcelBook_3, _pathExcelBook_4},
    
    ProcessedTablesList = List.Transform({0..TableCount-1}, (i) =>
        let
            TargetTableName = TableNamesInJson{i},
            _pathExcelBook = PathsList{i},
            
            SourceWorkbook = if _pathExcelBook = null or _pathExcelBook = "" then Excel.CurrentWorkbook() else Excel.Workbook(File.Contents(_pathExcelBook)), 
            InitialSourceTable = let Source = if _pathExcelBook = null or _pathExcelBook = "" then try SourceWorkbook{[Name=TargetTableName]}[Content] otherwise null else try SourceWorkbook{[Name=TargetTableName]}[Data] otherwise null in if Source = null then error "Целевая таблица с именем '" & TargetTableName & "' не найдена." else Source,
            
            MacroBlocksRecord = Record.Field(ConfigRecord, TargetTableName),
            MacroBlockNames = Record.FieldNames(MacroBlocksRecord),
            SortedMacroBlockNames = List.Sort(MacroBlockNames, (a, b) => Value.Compare(Number.FromText(a), Number.FromText(b))),

            ProcessedTable = List.Accumulate(SortedMacroBlockNames, InitialSourceTable, (tableState, macroBlockName) =>
                let
                    RulesRecord = Record.Field(MacroBlocksRecord, macroBlockName),

                    TableAfterSetType = let SetTypeRulesList = Record.FieldOrDefault(RulesRecord, "ColumnSetType", {}), Result = if List.IsEmpty(SetTypeRulesList) then tableState else let TypeTransformations = List.Transform(SetTypeRulesList, (rule) => let colName = Record.FieldNames(rule){0}, typeValue = Record.Field(rule, colName), pqType = if typeValue = "text" then type text else if typeValue = "int" then Int64.Type else if typeValue = "float" then type number else if typeValue = "bool" then type logical else null in if pqType = null then error "ColumnSetType: Указан неизвестный тип '" & typeValue & "'." else {colName, pqType}), currentCols = Table.ColumnNames(tableState), colsToTransform = List.Transform(TypeTransformations, each _{0}), missingCols = List.Select(colsToTransform, each not List.Contains(currentCols, _)) in if List.Count(missingCols) > 0 then error "ColumnSetType: В таблице отсутствуют столбцы для установки типа: " & Text.Combine(missingCols, ", ") else Table.TransformColumnTypes(tableState, TypeTransformations) in Result,
                    TableAfterFilterFree = let FilterRules = Record.FieldOrDefault(RulesRecord, "ColumnFilterFreeData", {}), Result = if List.IsEmpty(FilterRules) then TableAfterSetType else List.Accumulate(FilterRules, TableAfterSetType, (subTableState, currentRule) => let colName = Record.FieldNames(currentRule){0}, filterValue = Record.Field(currentRule, colName), validatedTable = if not Table.HasColumns(subTableState, {colName}) then error "ColumnFilterFreeData: В таблице отсутствует столбец для фильтрации: " & colName else subTableState in Table.SelectRows(validatedTable, each Text.From(Record.Field(_, colName)) = Text.From(filterValue))) in Result,
                    TableAfterFilterUnique = let FilterRules = Record.FieldOrDefault(RulesRecord, "ColumnFilterUniqueData", {}), Result = if List.IsEmpty(FilterRules) then TableAfterFilterFree else List.Accumulate(FilterRules, TableAfterFilterFree, (subTableState, columnName) => let validatedTable = if not Table.HasColumns(subTableState, {columnName}) then error "ColumnFilterUniqueData: В таблице отсутствует столбец для фильтрации: " & columnName else subTableState in Table.Distinct(validatedTable, {columnName})) in Result,
                    TableAfterRemoveOneRow = let RemoveRules = Record.FieldOrDefault(RulesRecord, "RemoveRowOne", {}), Result = if List.IsEmpty(RemoveRules) then TableAfterFilterUnique else List.Accumulate(RemoveRules, TableAfterFilterUnique, (subTableState, currentRule) => let ruleRecord = Record.Field(currentRule, Record.FieldNames(currentRule){0}), searchCol = ruleRecord[IndexSearchColumn], searchVal = Text.From(ruleRecord[IndexSearchRow]), columnAsTextList = List.Transform(Table.Column(subTableState, searchCol), each Text.From(_)), rowIndex = List.PositionOf(columnAsTextList, searchVal) in if rowIndex = -1 then error "RemoveRowOne: Не удалось найти строку, где '" & searchCol & "' = '" & searchVal & "'." else Table.RemoveRows(subTableState, rowIndex, 1)) in Result,
                    TableAfterRemoveAllRows = let RemoveRules = Record.FieldOrDefault(RulesRecord, "RemoveRowAll", {}), Result = if List.IsEmpty(RemoveRules) then TableAfterRemoveOneRow else List.Accumulate(RemoveRules, TableAfterRemoveOneRow, (subTableState, currentRule) => let ruleRecord = Record.Field(currentRule, Record.FieldNames(currentRule){0}), searchCol = ruleRecord[IndexSearchColumn], searchVal = Text.From(ruleRecord[IndexSearchRow]) in Table.SelectRows(subTableState, each Text.From(Record.Field(_, searchCol)) <> searchVal)) in Result,
                    TableAfterInsert = let InsertRules = Record.FieldOrDefault(RulesRecord, "InsertColumn", {}), Result = if List.IsEmpty(InsertRules) then TableAfterRemoveAllRows else List.Accumulate(InsertRules, TableAfterRemoveAllRows, (subTableState, newColumnName) => if Table.HasColumns(subTableState, {newColumnName}) then error "InsertColumn: Столбец с именем '" & newColumnName & "' уже существует." else Table.AddColumn(subTableState, newColumnName, each null)) in Result,
                    TableAfterFill = let FillRules = Record.FieldOrDefault(RulesRecord, "ColumnFillData", {}), Result = if List.IsEmpty(FillRules) then TableAfterInsert else List.Accumulate(FillRules, TableAfterInsert, (subTableState, currentRule) => let colName = Record.FieldNames(currentRule){0}, fillValue = Record.Field(currentRule, colName), validatedTable = if not Table.HasColumns(subTableState, {colName}) then error "ColumnFillData: В таблице отсутствует столбец для заполнения: " & colName else subTableState in Table.ReplaceValue(validatedTable, each Record.Field(_, colName), fillValue, Replacer.ReplaceValue, {colName})) in Result,
                    TableAfterModify = let ModifyRules = Record.FieldOrDefault(RulesRecord, "ModifyData", {}), Result = if List.IsEmpty(ModifyRules) then TableAfterFill else List.Accumulate(ModifyRules, TableAfterFill, (subTableState, currentRule) => let ruleRecord = Record.Field(currentRule, Record.FieldNames(currentRule){0}), searchCol = ruleRecord[IndexSearchColumn], searchVal = Text.From(ruleRecord[IndexSearchRow]), newData = Record.RemoveFields(ruleRecord[NewData], "COMMENT", MissingField.Ignore), columnAsTextList = List.Transform(Table.Column(subTableState, searchCol), each Text.From(_)), rowIndex = List.PositionOf(columnAsTextList, searchVal), FinalTable = if rowIndex = -1 then error "ModifyData: Не удалось найти строку, где '" & searchCol & "' = '" & searchVal & "'." else let oldRowRecord = subTableState{rowIndex}, newRowRecord = Record.Combine({oldRowRecord, newData}) in Table.ReplaceRows(subTableState, rowIndex, 1, {newRowRecord}) in FinalTable) in Result,
                    TableAfterRemoveXor = let KeepColRules = Record.FieldOrDefault(RulesRecord, "RemoveXorColumn", {}), Result = if List.IsEmpty(KeepColRules) then TableAfterModify else let currentCols = Table.ColumnNames(TableAfterModify), missingCols = List.Select(KeepColRules, each not List.Contains(currentCols, _)), validatedTable = if List.Count(missingCols) > 0 then error "RemoveXorColumn: В таблице отсутствуют столбцы для сохранения: " & Text.Combine(missingCols, ", ") else TableAfterModify in Table.SelectColumns(validatedTable, KeepColRules) in Result,
                    TableAfterRemoveCol = let RemoveColRules = Record.FieldOrDefault(RulesRecord, "RemoveColumn", {}), Result = if List.IsEmpty(RemoveColRules) then TableAfterRemoveXor else let currentCols = Table.ColumnNames(TableAfterRemoveXor), missingCols = List.Select(RemoveColRules, each not List.Contains(currentCols, _)), validatedTable = if List.Count(missingCols) > 0 then error "RemoveColumn: В таблице отсутствуют столбцы для удаления: " & Text.Combine(missingCols, ", ") else TableAfterRemoveXor in Table.RemoveColumns(validatedTable, RemoveColRules) in Result,
                    TableAfterClearCol = let ClearColRules = Record.FieldOrDefault(RulesRecord, "ClearColumn", {}), Result = if List.IsEmpty(ClearColRules) then TableAfterRemoveCol else let currentCols = Table.ColumnNames(TableAfterRemoveCol), missingCols = List.Select(ClearColRules, each not List.Contains(currentCols, _)), transformations = if List.Count(missingCols) > 0 then error "ClearColumn: В таблице отсутствуют столбцы для очистки: " & Text.Combine(missingCols, ", ") else List.Transform(ClearColRules, (colName) => {colName, each null, type any}) in Table.TransformColumns(TableAfterRemoveCol, transformations) in Result,
                    TableAfterRename = let RenameColRules = Record.FieldOrDefault(RulesRecord, "RenameColumn", {}), transformations = List.Transform(RenameColRules, each {Record.FieldNames(_){0}, Record.Field(_, Record.FieldNames(_){0})}), Result = if List.IsEmpty(transformations) then TableAfterClearCol else let currentCols = Table.ColumnNames(TableAfterClearCol), oldNames = List.Transform(transformations, each _{0}), missingCols = List.Select(oldNames, each not List.Contains(currentCols, _)), validatedTable = if List.Count(missingCols) > 0 then error "RenameColumn: В таблице отсутствуют столбцы для переименования: " & Text.Combine(missingCols, ", ") else TableAfterClearCol in Table.RenameColumns(validatedTable, transformations) in Result,
                    TableAfterReorder = let ReorderColRules = Record.FieldOrDefault(RulesRecord, "ReorderColumn", {}), Result = if List.IsEmpty(ReorderColRules) then TableAfterRename else let currentCols = Table.ColumnNames(TableAfterRename), missingCols = List.Select(ReorderColRules, each not List.Contains(currentCols, _)), otherCols = List.RemoveItems(currentCols, ReorderColRules), finalOrder = ReorderColRules & otherCols in if List.Count(missingCols) > 0 then error "ReorderColumn: В таблице отсутствуют столбцы для переупорядочивания: " & Text.Combine(missingCols, ", ") else Table.ReorderColumns(TableAfterRename, finalOrder) in Result
                in
                    TableAfterReorder
            )
        in
            ProcessedTable
    ),

    // =============================================================================================
    // ШАГ 5: ВЫПОЛНЕНИЕ ОПЕРАЦИИ СЛИЯНИЯ
    // =============================================================================================
    
    TableAfterJoin = if TableCount = 1 then ProcessedTablesList{0} else
        let
            FirstTableRules = Record.Field(ConfigRecord, TableNamesInJson{0}),
            AllMacroBlocksFirstTable = List.Transform(Record.FieldValues(FirstTableRules), each Record.FieldNames(_)),
            JoinCommand = List.Combine(List.Transform(AllMacroBlocksFirstTable, each List.Intersect({_, {"RunLeftJOINstructure", "RunRightJOINstructure", "RunJOINdata"}}))){0},
            
            Result = if JoinCommand = "RunJOINdata" then
                let
                    firstTableCols = Table.ColumnNames(ProcessedTablesList{0}),
                    validation = List.AllTrue(List.Transform(List.Range(ProcessedTablesList, 1), (tbl) => List.MatchesAll(Table.ColumnNames(tbl), each List.Contains(firstTableCols, _)) and List.Count(Table.ColumnNames(tbl)) = List.Count(firstTableCols)))
                in
                    if not validation then error "RunJOINdata: Структура столбцов в объединяемых таблицах не совпадает." else Table.Combine(ProcessedTablesList)
            
            else if JoinCommand = "RunLeftJOINstructure" then
                let
                    validation = List.AllTrue(List.Transform(ProcessedTablesList, each Table.HasColumns(_, {"HASH COLLECTION"})))
                in
                    if not validation then error "RunLeftJOINstructure: Во всех таблицах должен присутствовать столбец 'HASH COLLECTION'." else
                    let
                        JoinedTable = List.Accumulate(List.Range(ProcessedTablesList, 1), ProcessedTablesList{0}, (state, current) => 
                            let
                                joined = Table.NestedJoin(state, {"HASH COLLECTION"}, current, {"HASH COLLECTION"}, "temp", JoinKind.LeftOuter),
                                colsToExpand = List.RemoveItems(Table.ColumnNames(current), {"HASH COLLECTION"}),
                                expanded = Table.ExpandTableColumn(joined, "temp", colsToExpand)
                            in
                                expanded)
                    in
                        JoinedTable
            
            else // RunRightJOINstructure
                let
                    validation = List.AllTrue(List.Transform(ProcessedTablesList, each Table.HasColumns(_, {"HASH COLLECTION"})))
                in
                    if not validation then error "RunRightJOINstructure: Во всех таблицах должен присутствовать столбец 'HASH COLLECTION'." else
                    let
                        JoinedTable = List.Accumulate(List.Reverse(List.RemoveLastN(ProcessedTablesList, 1)), List.Last(ProcessedTablesList), (state, current) => 
                            let
                                joined = Table.NestedJoin(current, {"HASH COLLECTION"}, state, {"HASH COLLECTION"}, "temp", JoinKind.RightOuter),
                                colsToExpand = List.RemoveItems(Table.ColumnNames(current), {"HASH COLLECTION"}),
                                expanded = Table.ExpandTableColumn(joined, "temp", colsToExpand)
                            in
                                expanded)
                    in
                        JoinedTable
        in
            Result,

    // =============================================================================================
    // ШАГ 6: ФИНАЛЬНАЯ ВАЛИДАЦИЯ (ЕСЛИ ПРИМЕНИМО)
    // =============================================================================================
    
    FirstTableRulesRecord = Record.Field(ConfigRecord, TableNamesInJson{0}),
    
    FinalTable = if not Record.HasFields(FirstTableRulesRecord, "ValidationData") then
        TableAfterJoin
    else
        let
            CleanedTable = if Table.HasColumns(TableAfterJoin, {"VALIDATION"}) then Table.RemoveColumns(TableAfterJoin, {"VALIDATION"}) else TableAfterJoin,
            BufferedSourceTable = Table.Buffer(CleanedTable),
            ValidationRulesList = Record.Field(FirstTableRulesRecord, "ValidationData"),
            ColumnRulesRecord = if List.IsEmpty(ValidationRulesList) then [] else Record.Combine(ValidationRulesList),
            ColumnsInJson = Record.FieldNames(ColumnRulesRecord),
            SourceTableColumns = Table.ColumnNames(BufferedSourceTable),
            PreValidationResults = let IdColumns = List.Select(ColumnsInJson, each Record.HasFields(Record.Field(ColumnRulesRecord, _), "fvCheckID")), ResultsRecord = List.Accumulate(IdColumns, [], (state, currentColumn) => let errorList = UtilsFunctionVector[fvCheckID](BufferedSourceTable, currentColumn) in state & Record.FromList({errorList}, {currentColumn})) in ResultsRecord,
            SourceWithIndex = Table.AddIndexColumn(BufferedSourceTable, "__Index__", 0, 1),
            AddValidationColumn = Table.AddColumn(SourceWithIndex, "VALIDATION", each
                let
                    currentRow = _,
                    currentRowIndex = currentRow[__Index__],
                    rowErrorCodes = List.Transform(SourceTableColumns, (currentColumnName) =>
                        let
                            isColumnInRules = Record.HasFields(ColumnRulesRecord, currentColumnName),
                            resultCode = if not isColumnInRules then 0
                            else
                                let
                                    columnIndexForError = List.PositionOf(SourceTableColumns, currentColumnName) + 1,
                                    preValidationError = if Record.HasFields(PreValidationResults, currentColumnName) then Record.Field(PreValidationResults, currentColumnName){currentRowIndex} else -1,
                                    finalResult = if preValidationError > 0 then columnIndexForError
                                    else
                                        let
                                            rulesForColumn = Record.Field(ColumnRulesRecord, currentColumnName),
                                            functionsToRun = if Record.HasFields(rulesForColumn, "typeColumn") then error "ValidationData: Ключ 'typeColumn' не должен использоваться в блоке ValidationData." else Record.FieldNames(rulesForColumn),
                                            cellValue = Record.Field(currentRow, currentColumnName),
                                            finalErrorCode = List.Accumulate(functionsToRun, 0, (state, currentFunctionName) =>
                                                if state <> 0 then state
                                                else
                                                    let
                                                        isIdCheck = currentFunctionName = "fvCheckID",
                                                        errorCode = if isIdCheck then 0 else
                                                            let
                                                                scalarFunctionName = currentFunctionName & "_scalar",
                                                                optionalArgs = Record.Field(rulesForColumn, currentFunctionName),
                                                                allArgs = {cellValue} & optionalArgs,
                                                                InvokeResult = try Function.Invoke(Record.Field(UtilsFunctionScalar, scalarFunctionName), allArgs),
                                                                isValid = if InvokeResult[HasError] then error "Ошибка вызова функции '" & scalarFunctionName & "'. Проверьте аргументы в JSON для столбца '" & currentColumnName & "'." else InvokeResult[Value]
                                                            in
                                                                if not isValid then columnIndexForError else 0
                                                    in errorCode )
                                        in finalErrorCode
                                in finalResult
                        in resultCode ),
                    validationString = Text.Combine(List.Transform(rowErrorCodes, Text.From), "-")
                in validationString, type text),
            FinalResultTable = Table.RemoveColumns(AddValidationColumn, {"__Index__"})
        in
            FinalResultTable
in
    FinalTable