// =============================================================================================
//
// ФУНКЦИЯ-ЗАПРОС 'ProcessingTable'
//
// версия rev.01 v05
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
            // Извлекаем блок правил "ModifyData" из JSON, если он есть, иначе - пустой список.
            ModifyRules = Record.FieldOrDefault(RulesRecord, "ModifyData", {}),
            // Выполняем операцию только если правила для модификации существуют.
            Result = if List.IsEmpty(ModifyRules) then InitialSourceTable else
                // Итерируем по каждому правилу модификации, последовательно применяя изменения к таблице.
                List.Accumulate(ModifyRules, InitialSourceTable, (tableState, currentRule) =>
                    let
                        // Извлекаем запись с параметрами для текущего правила (data_1, data_2, ...).
                        ruleRecord = Record.Field(currentRule, Record.FieldNames(currentRule){0}),
                        // Получаем имя столбца для поиска.
                        searchCol = ruleRecord[IndexSearchColumn],
                        // Получаем значение для поиска и принудительно преобразуем его в текст для надежного сравнения.
                        searchVal = Text.From(ruleRecord[IndexSearchRow]),
                        // Извлекаем запись с новыми данными, игнорируя поле "COMMENT".
                        newData = Record.RemoveFields(ruleRecord[NewData], "COMMENT", MissingField.Ignore),

                        // Преобразуем весь столбец поиска в текст для корректного сравнения типов.
                        columnAsTextList = List.Transform(Table.Column(tableState, searchCol), each Text.From(_)),
                        // Находим индекс (позицию) первой строки, соответствующей критерию поиска.
                        rowIndex = List.PositionOf(columnAsTextList, searchVal),
                        
                        // Если строка не найдена (rowIndex = -1), генерируем ошибку.
                        FinalTable = if rowIndex = -1 then 
                                        error "ModifyData: Не удалось найти строку, где '" & searchCol & "' = '" & searchVal & "'." 
                                     else
                                        // Если строка найдена:
                                        let 
                                            // Получаем текущую запись (строку) по найденному индексу.
                                            oldRowRecord = tableState{rowIndex}, 
                                            // Объединяем старую запись с новыми данными (новые данные перезапишут старые).
                                            newRowRecord = Record.Combine({oldRowRecord, newData})
                                        in 
                                            // Заменяем в таблице старую строку на новую по тому же индексу.
                                            Table.ReplaceRows(tableState, rowIndex, 1, {newRowRecord})
                    in FinalTable)
        in
            Result,

    // -- ЭТАП 2: УДАЛЕНИЕ СТРОК --
    TableAfterRemoveRow =
        let
            // Извлекаем блок правил "RemoveRow" из JSON, если он есть, иначе - пустой список.
            RemoveRowRules = Record.FieldOrDefault(RulesRecord, "RemoveRow", {}),
            // Выполняем операцию только если правила для удаления существуют.
            Result = if List.IsEmpty(RemoveRowRules) then TableAfterModify else
                let
                    // Преобразуем таблицу в буфер для повышения производительности при многократном поиске.
                    bufferedTable = Table.Buffer(TableAfterModify),
                    // Создаем список индексов строк, которые нужно удалить.
                    rowIndexesToRemove = List.Transform(RemoveRowRules, (rule) =>
                        let
                            // Извлекаем запись с параметрами для текущего правила (row_1, row_2, ...).
                            ruleRecord = Record.Field(rule, Record.FieldNames(rule){0}),
                            // Получаем имя столбца для поиска.
                            searchCol = ruleRecord[IndexSearchColumn],
                            // Получаем значение для поиска и принудительно преобразуем его в текст.
                            searchVal = Text.From(ruleRecord[IndexSearchRow]),
                            
                            // Преобразуем весь столбец поиска в текст для корректного сравнения.
                            columnAsTextList = List.Transform(Table.Column(bufferedTable, searchCol), each Text.From(_)),
                            // Находим индекс (позицию) первой строки, соответствующей критерию поиска.
                            rowIndex = List.PositionOf(columnAsTextList, searchVal)
                        in
                            // Если строка не найдена, генерируем ошибку, иначе возвращаем ее индекс.
                            if rowIndex = -1 then error "RemoveRow: Не удалось найти строку, где '" & searchCol & "' = '" & searchVal & "'." else rowIndex),
                    
                    // --- ИСПРАВЛЕНИЕ: Используем правильный метод для удаления строк по списку индексов ---
                    // Отбираем только уникальные индексы на случай дублирования правил.
                    distinctIndexes = List.Distinct(rowIndexesToRemove),
                    // Добавляем к таблице временный столбец с индексом от 0 до N-1.
                    indexedTable = Table.AddIndexColumn(TableAfterModify, "__IndexToRemove__", 0, 1),
                    // Фильтруем таблицу, оставляя только те строки, индекс которых НЕ содержится в списке на удаление.
                    filteredTable = Table.SelectRows(indexedTable, each not List.Contains(distinctIndexes, [__IndexToRemove__])),
                    // Удаляем временный столбец индекса.
                    finalTable = Table.RemoveColumns(filteredTable, {"__IndexToRemove__"})
                in
                    finalTable
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
                    transformations = if List.Count(missingCols) > 0 then error "ClearColumn: В таблице отсутствуют столбцы для очистки: " & Text.Combine(missingCols, ", ") else List.Transform(ClearColRules, (colName) => {colName, each null, type any})
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
                            errorList = QueryUtilsFunction_vector[fvCheckID](BufferedSourceTable, currentColumn)
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