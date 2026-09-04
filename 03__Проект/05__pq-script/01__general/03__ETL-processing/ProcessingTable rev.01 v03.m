// =============================================================================================
//
// ФУНКЦИЯ-ЗАПРОС 'ProcessingTable'
//
// версия rev.01 v03
// =============================================================================================

// ---------------------------------------------------------------------------------------------
// Функция-запрос 'ProcessingTable'
// Выполняет комплексную обработку (трансформацию и валидацию) таблицы Excel на основе
// правил, найденных в централизованной таблице конфигурации.
//
// аргументы:
// _configJSONIndex as text    - Текстовый идентификатор JSON-конфигуратора, который необходимо
//                               найти в столбце 'IDconfig' таблицы 'TabConfigJSON'. Аргумент
//                               является обязательным.
// _pathExcelBook as text      - Полный путь к файлу Excel, в котором находятся исходные данные.
//                               Аргумент является необязательным. Если он не указан,
//                               поиск исходной таблицы будет производиться в текущей активной
//                               книге Excel, в которой запущен этот запрос.
//
// возвращаемое значение:
// table                       - Обработанная таблица с выполненными трансформациями и,
//                               если указано в JSON, с добавленным столбцом 'VALIDATION'.
// ---------------------------------------------------------------------------------------------
(
    _configJSONIndex as text,
    optional _pathExcelBook as text
) as table =>
let
    // =============================================================================================
    // ШАГ 1: ПОИСК И ЗАГРУЗКА JSON-КОНФИГУРАТОРА
    // =============================================================================================

    // Жестко заданные имена для таблицы и столбцов с конфигурациями
    ConfigTableName = "TabConfigJSON",
    ConfigLookupColumn = "IDconfig",
    ConfigJsonColumn = "CONFIG_JSON",

    // Источник для конфигурации - ВСЕГДА текущая активная книга Excel
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
    // ШАГ 2: ПАРСИНГ JSON И ЗАГРУЗКА ИСХОДНОЙ ТАБЛИЦЫ
    // =============================================================================================

    ConfigRecord = Json.Document(JsonConfigText),
    TargetTableName = Record.FieldNames(ConfigRecord){0},
    RulesRecord = Record.Field(ConfigRecord, TargetTableName),

    SourceWorkbook = if _pathExcelBook = null or _pathExcelBook = "" then
                         Excel.CurrentWorkbook() 
                     else
                         Excel.Workbook(File.Contents(_pathExcelBook)), 

    InitialSourceTable = 
        let
            Source = if _pathExcelBook = null or _pathExcelBook = "" then
                try SourceWorkbook{[Name=TargetTableName]}[Content] otherwise null
            else
                try SourceWorkbook{[Name=TargetTableName]}[Data] otherwise null
        in
            if Source = null then error "Целевая таблица с именем '" & TargetTableName & "', указанная в конфигураторе, не найдена." else Source,

    // =============================================================================================
    // ШАГ 3: ПОСЛЕДОВАТЕЛЬНОЕ ПРИМЕНЕНИЕ ТРАНСФОРМАЦИЙ
    // =============================================================================================
    
    // -- ЭТАП 1: МОДИФИКАЦИЯ ДАННЫХ --
    TableAfterModify = 
        let
            ModifyRules = Record.FieldOrDefault(RulesRecord, "ModifyData", {}),
            Result = if List.IsEmpty(ModifyRules) then InitialSourceTable else
                List.Accumulate(ModifyRules, InitialSourceTable, (tableState, currentRule) =>
                    let
                        ruleRecord = Record.Field(currentRule, Record.FieldNames(currentRule){0}),
                        searchCol = ruleRecord[IndexSearchColumn],
                        searchVal = ruleRecord[IndexSearchRow],
                        newData = Record.RemoveFields(ruleRecord[NewData], "COMMENT", MissingField.Ignore),
                        rowIndex = Table.PositionOf(tableState, Record.SelectFields(tableState{0}, {searchCol}), {[Value = searchVal]}, Occurrence.First),
                        FinalTable = if rowIndex = -1 then error "ModifyData: Не удалось найти строку, где '" & searchCol & "' = '" & Text.From(searchVal) & "'." else let oldRowRecord = tableState{rowIndex}, newRowRecord = Record.Combine({oldRowRecord, newData}) in Table.ReplaceRows(tableState, rowIndex, 1, {newRowRecord})
                    in FinalTable)
        in
            Result,

    // -- ЭТАП 2: УДАЛЕНИЕ СТРОК --
    TableAfterRemoveRow =
        let
            RemoveRowRules = Record.FieldOrDefault(RulesRecord, "RemoveRow", {}),
            Result = if List.IsEmpty(RemoveRowRules) then TableAfterModify else
                let
                    rowIndexesToRemove = List.Transform(RemoveRowRules, (rule) =>
                        let
                            ruleRecord = Record.Field(rule, Record.FieldNames(rule){0}),
                            searchCol = ruleRecord[IndexSearchColumnData],
                            searchVal = ruleRecord[IndexSearchRowData],
                            rowIndex = Table.PositionOf(TableAfterModify, Record.SelectFields(TableAfterModify{0}, {searchCol}), {[Value = searchVal]}, Occurrence.First)
                        in
                            if rowIndex = -1 then error "RemoveRow: Не удалось найти строку, где '" & searchCol & "' = '" & Text.From(searchVal) & "'." else rowIndex),
                    sortedIndexes = List.Sort(rowIndexesToRemove, Order.Descending)
                in
                    Table.RemoveRows(TableAfterModify, sortedIndexes)
        in
            Result,

    // -- ЭТАП 3: УДАЛЕНИЕ СТОЛБЦОВ --
    TableAfterRemoveCol = 
        let
            RemoveColRules = Record.FieldOrDefault(RulesRecord, "RemoveColumn", {}),
            Result = if List.IsEmpty(RemoveColRules) then TableAfterRemoveRow else
                let
                    currentCols = Table.ColumnNames(TableAfterRemoveRow),
                    missingCols = List.Select(RemoveColRules, each not List.Contains(currentCols, _)),
                    validatedTable = if List.Count(missingCols) > 0 then error "RemoveColumn: В таблице отсутствуют столбцы для удаления: " & Text.Combine(missingCols, ", ") else TableAfterRemoveRow
                in
                    Table.RemoveColumns(validatedTable, RemoveColRules)
        in
            Result,

    // -- ЭТАП 4: ОЧИСТКА СТОЛБЦОВ --
    TableAfterClearCol =
        let
            ClearColRules = Record.FieldOrDefault(RulesRecord, "ClearColumn", {}),
            Result = if List.IsEmpty(ClearColRules) then TableAfterRemoveCol else
                let
                    currentCols = Table.ColumnNames(TableAfterRemoveCol),
                    missingCols = List.Select(ClearColRules, each not List.Contains(currentCols, _)),
                    transformations = if List.Count(missingCols) > 0 then error "ClearColumn: В таблице отсутствуют столбцы для очистки: " & Text.Combine(missingCols, ", ") else List.Transform(ClearColRules, (colName) => {colName, each null})
                in
                    Table.TransformColumns(TableAfterRemoveCol, transformations)
        in
            Result,

    // -- ЭТАП 5: ПЕРЕИМЕНОВАНИЕ СТОЛБЦОВ --
    TableAfterRename =
        let
            RenameColRules = Record.FieldOrDefault(RulesRecord, "RenameColumn", {}),
            transformations = List.Transform(RenameColRules, each {Record.FieldNames(_){0}, Record.Field(_, Record.FieldNames(_){0})}),
            Result = if List.IsEmpty(transformations) then TableAfterClearCol else
                let
                    currentCols = Table.ColumnNames(TableAfterClearCol),
                    oldNames = List.Transform(transformations, each _{0}),
                    missingCols = List.Select(oldNames, each not List.Contains(currentCols, _)),
                    validatedTable = if List.Count(missingCols) > 0 then error "RenameColumn: В таблице отсутствуют столбцы для переименования: " & Text.Combine(missingCols, ", ") else TableAfterClearCol
                in
                    Table.RenameColumns(validatedTable, transformations)
        in
            Result,

    // -- ЭТАП 6: ПЕРЕУПОРЯДОЧИВАНИЕ СТОЛБЦОВ --
    TableAfterReorder =
        let
            ReorderColRules = Record.FieldOrDefault(RulesRecord, "ReorderColumn", {}),
            Result = if List.IsEmpty(ReorderColRules) then TableAfterRename else
                let
                    currentCols = Table.ColumnNames(TableAfterRename),
                    missingCols = List.Select(ReorderColRules, each not List.Contains(currentCols, _)),
                    otherCols = List.RemoveItems(currentCols, ReorderColRules),
                    finalOrder = ReorderColRules & otherCols
                in
                    if List.Count(missingCols) > 0 then error "ReorderColumn: В таблице отсутствуют столбцы для переупорядочивания: " & Text.Combine(missingCols, ", ") else Table.ReorderColumns(TableAfterRename, finalOrder)
        in
            Result,

    // =============================================================================================
    // ШАГ 4 и 5: ВАЛИДАЦИЯ И ПРИМЕНЕНИЕ ТИПОВ
    // =============================================================================================
    
    FinalTable = if not Record.HasFields(RulesRecord, "ValidationData") then
        TableAfterReorder
    else
        let
            CleanedTable = if Table.HasColumns(TableAfterReorder, {"VALIDATION"}) then Table.RemoveColumns(TableAfterReorder, {"VALIDATION"}) else TableAfterReorder,
            BufferedSourceTable = Table.Buffer(CleanedTable),
            ValidationRulesList = Record.Field(RulesRecord, "ValidationData"),
            ColumnRulesRecord = if List.IsEmpty(ValidationRulesList) then [] else Record.Combine(ValidationRulesList),
            ColumnsInJson = Record.FieldNames(ColumnRulesRecord),
            SourceTableColumns = Table.ColumnNames(BufferedSourceTable),

            PreValidationResults =
                let
                    IdColumns = List.Select(ColumnsInJson, each Record.HasFields(Record.Field(ColumnRulesRecord, _), "fvCheckID")),
                    ResultsRecord = List.Accumulate(IdColumns, [], (state, currentColumn) =>
                        let
                            // --- ИСПРАВЛЕНИЕ: Передаем ОБРАБОТАННУЮ таблицу ---
                            errorList = QueryUtilsFunction[fvCheckID](BufferedSourceTable, currentColumn)
                        in
                            state & Record.FromList({errorList}, {currentColumn}))
                in
                    ResultsRecord,
            
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
                                            functionsToRun = List.RemoveItems(Record.FieldNames(rulesForColumn), {"typeColumn"}),
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
                                                                InvokeResult = try Function.Invoke(Record.Field(QueryUtilsFunction_scalar, scalarFunctionName), allArgs),
                                                                isValid = if InvokeResult[HasError] then error "Ошибка вызова функции '" & scalarFunctionName & "'. Проверьте аргументы в JSON для столбца '" & currentColumnName & "'." else InvokeResult[Value]
                                                            in
                                                                if not isValid then columnIndexForError else 0
                                                    in errorCode )
                                        in finalErrorCode
                                in finalResult
                        in resultCode ),
                    validationString = Text.Combine(List.Transform(rowErrorCodes, Text.From), "-")
                in validationString, type text),
            TableWithValidation = Table.RemoveColumns(AddValidationColumn, {"__Index__"}),

            TypeTransformations = List.Transform(ColumnsInJson, (colName) =>
                let
                    ruleSet = Record.Field(ColumnRulesRecord, colName),
                    typeValue = Record.FieldOrDefault(ruleSet, "typeColumn"),
                    pqType = if typeValue = "text" then type text else if typeValue = "int" then Int64.Type else if typeValue = "float" then type number else if typeValue = "bool" then type logical else null
                in {colName, pqType} ),
            FinalTypeTransformations = List.Select(TypeTransformations, each _{1} <> null),

            result = if List.IsEmpty(FinalTypeTransformations) then TableWithValidation else Table.TransformColumnTypes(TableWithValidation, FinalTypeTransformations)
        in
            result
in
    FinalTable