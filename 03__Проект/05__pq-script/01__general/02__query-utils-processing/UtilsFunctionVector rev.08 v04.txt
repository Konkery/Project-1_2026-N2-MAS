let
    // =============================================================================================
    //
    // ФУНКЦИИ ТРАНСФОРМАЦИИ И ВАЛИДАЦИИ - ВЕКТОРНЫЕ
    // версия rev.08 v04
    //
    // Данный агрегатный Query содержит набор функций выполняющих:
    // 1. валидацию данных;
    // 2. трансформацию данных.
    //
    // =============================================================================================
     
    // ---------------------------------------------------------------------------------------------
    // Функция трансформации 'ftCreateFullNameColumn'
    // Создает столбец "FULL_NAME' из трех исходных столбцов ФИО в формате Петров С.Н.
    // и размещает его в указанной позиции.
    //
    // аргументы:
    //    _selectTableName as text  - текстовое имя таблицы
    //    _lastNameCol as text      - имя столбца с фамилией
    //    _firstNameCol as text     - имя столбца с именем
    //    _patronymicCol as text    - имя столбца с отчеством
    //    _position as number       - позиция для размещения нового столбца (начиная с 0)
    //
    // возвращаемое значение:
    //    table                      - исходная таблица с вновь созданным столбцом ФИО
    //
    // ---------------------------------------------------------------------------------------------
    ftCreateFullNameColumn = (
        _selectTableName as text, 
        _lastNameCol as text, 
        _firstNameCol as text, 
        _patronymicCol as text,
        _position as number
    ) as table =>
    let
        // Получить все таблицы текущей Excel книги
        allTables = Excel.CurrentWorkbook(),

        // Фильтруем список таблиц по имени
        filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),

        // Проверка существования таблицы
        selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content]
                      else error "Таблица с именем '" & _selectTableName & "' не найдена в текущей книге.",

        // Проверка наличия обязательных столбцов
        validatedTable = if not Table.HasColumns(selectTable, {_lastNameCol, _firstNameCol, _patronymicCol}) then
            error "В таблице '" & _selectTableName & "' отсутствуют обязательные столбцы для создания ФИО."
        else
            selectTable,

        // Проверка, что столбец FULL_NAME еще не существует
        ColumnNames = Table.ColumnNames(validatedTable),
        FullNameExists = List.Contains(ColumnNames, "FULL_NAME"),

        // Если столбец уже существует, возвращаем ошибку
        result = if FullNameExists then 
            error "Столбец 'FULL_NAME' уже существует в таблице '" & _selectTableName & "'."
        else
            let
                TableWithFullName = Table.AddColumn(
                    validatedTable,
                    "FULL_NAME",
                    each 
                        let
                            lastName = if Record.Field(_, _lastNameCol) = null then "" else Text.From(Record.Field(_, _lastNameCol)),
                            firstName = if Record.Field(_, _firstNameCol) = null then "" else Text.From(Record.Field(_, _firstNameCol)),
                            patronymic = if Record.Field(_, _patronymicCol) = null then "" else Text.From(Record.Field(_, _patronymicCol)),
                            firstInitial = if firstName = "" then "" else Text.Upper(Text.Start(firstName, 1)) & ".",
                            patronymicInitial = if patronymic = "" then "" else Text.Upper(Text.Start(patronymic, 1)) & ".",
                            fullName = Text.Trim(lastName & " " & firstInitial & patronymicInitial)
                        in
                            fullName,
                    type text
                ),
                CurrentColumns = Table.ColumnNames(TableWithFullName),
                ColumnsWithoutFullName = List.RemoveItems(CurrentColumns, {"FULL_NAME"}),
                InsertPosition = if _position > List.Count(ColumnsWithoutFullName) then List.Count(ColumnsWithoutFullName) else _position,
                NewColumnOrder = List.InsertRange(ColumnsWithoutFullName, InsertPosition, {"FULL_NAME"}),
                result = Table.ReorderColumns(TableWithFullName, NewColumnOrder)
            in
                result
    in
        result,

    // ---------------------------------------------------------------------------------------------
    // Функция трансформации 'ftGetFullNameColumn'
    // Возвращает список значений FULL_NAME, сформированный из трех исходных столбцов ФИО 
    // в формате Петров С.Н.
    //
    // аргументы:
    //    _selectTableName as text  - текстовое имя таблицы
    //    _lastNameCol as text      - имя столбца с фамилией
    //    _firstNameCol as text     - имя столбца с именем
    //    _patronymicCol as text    - имя столбца с отчеством
    //
    // возвращаемое значение:
    //    list                     - список значений ФИО
    //
    // ---------------------------------------------------------------------------------------------
    ftGetFullNameColumn = (
        _selectTableName as text, 
        _lastNameCol as text, 
        _firstNameCol as text, 
        _patronymicCol as text
    ) as list =>
    let
        // Получить все таблицы текущей Excel книги
        allTables = Excel.CurrentWorkbook(),

        // Фильтруем список таблиц по имени
        filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),

        // Проверка существования таблицы
        selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content]
                     else error "Таблица с именем '" & _selectTableName & "' не найдена в текущей книге.",

        // Проверка наличия обязательных столбцов
        validatedTable = if not Table.HasColumns(selectTable, {_lastNameCol, _firstNameCol, _patronymicCol}) then
            error "В таблице '" & _selectTableName & "' отсутствуют обязательные столбцы для создания ФИО."
        else
            selectTable,

        CreateFullName = (row as record) as text =>
            let
                lastName = if Record.Field(row, _lastNameCol) = null then "" else Text.Trim(Text.From(Record.Field(row, _lastNameCol))),
                firstName = if Record.Field(row, _firstNameCol) = null then null else Text.Trim(Text.From(Record.Field(row, _firstNameCol))),
                patronymic = if Record.Field(row, _patronymicCol) = null then null else Text.Trim(Text.From(Record.Field(row, _patronymicCol))),
                firstInitial = if firstName <> null and Text.Length(firstName) > 0 then Text.Upper(Text.Start(firstName, 1)) & "." else "",
                patronymicInitial = if patronymic <> null and Text.Length(patronymic) > 0 then Text.Upper(Text.Start(patronymic, 1)) & "." else "",
                fullName = lastName & (if Text.Length(firstInitial & patronymicInitial) > 0 then " " & firstInitial & patronymicInitial else "")
            in
                fullName,

        fullNameList = List.Transform(Table.ToRecords(validatedTable), each CreateFullName(_)),
        result = fullNameList
    in
        result,

    // ---------------------------------------------------------------------------------------------
    // Функция трансформации 'ftReorderColumnsStrict'
    // Переупорядочивает столбцы таблицы по заданному списку.
    //
    // аргументы:
    // _selectTableName as text - текстовое имя таблицы
    // _columnList as list      - список имен столбцов в необходимом порядке
    //
    // возвращаемое значение:
    // table                    - таблица
    //
    // ---------------------------------------------------------------------------------------------
    ftReorderColumnsStrict = (
        _selectTableName as text,
        _columnList as list
    ) as table =>
    let
        // Получить все таблицы текущей Excel книги
        allTables = Excel.CurrentWorkbook(),

        // Фильтруем список таблиц по имени
        filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),

        // Проверка существования таблицы
        selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content]
                      else error "Таблица с именем '" & _selectTableName & "' не найдена в текущей книге.",
        
        // Проверка наличия обязательных столбцов
        validatedTable = 
            let
                tableColumns = Table.ColumnNames(selectTable),
                missingColumns = List.Select(_columnList, each not List.Contains(tableColumns, _))
            in
                if List.Count(missingColumns) > 0 then
                    error "В таблице '" & _selectTableName & "' отсутствуют столбцы для переупорядочивания: " & Text.Combine(missingColumns, ", ")
                else
                    selectTable,
        
        result = Table.ReorderColumns(validatedTable, _columnList)
    in
        result,
    
    // ---------------------------------------------------------------------------------------------
    // Функция трансформации 'ftRemoveColumnsStrict'
    // Удаляет указанные столбцы из таблицы.
    //
    // аргументы:
    // _selectTableName as text - текстовое имя таблицы
    // _columnList as list      - список имен столбцов которые необходимо удалить
    //
    // возвращаемое значение:
    // table                    - таблица
    //
    // ---------------------------------------------------------------------------------------------
    ftRemoveColumnsStrict = (
        _selectTableName as text,
        _columnList as list
    ) as table =>
    let
        // Получить все таблицы текущей Excel книги
        allTables = Excel.CurrentWorkbook(),

        // Фильтруем список таблиц по имени
        filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),

        // Проверка существования таблицы
        selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content]
                      else error "Таблица с именем '" & _selectTableName & "' не найдена в текущей книге.",

        // Проверка наличия обязательных столбцов
        validatedTable = 
            let
                tableColumns = Table.ColumnNames(selectTable),
                missingColumns = List.Select(_columnList, each not List.Contains(tableColumns, _))
            in
                if List.Count(missingColumns) > 0 then
                    error "В таблице '" & _selectTableName & "' нельзя удалить несуществующие столбцы: " & Text.Combine(missingColumns, ", ")
                else
                    selectTable,
        
        result = Table.RemoveColumns(validatedTable, _columnList)
    in
        result,

    // ---------------------------------------------------------------------------------------------
    // Функция трансформации 'ftCapitalizeWords'
    // Делает первую букву каждого слова в тексте заглавной.
    //
    // аргументы:
    // _selectTableName as text - текстовое имя таблицы
    // _columnList as list      - список имен столбцов у которых необходимо изменить регистр значений
    //
    // возвращаемое значение:
    // table                    - таблица
    //
    // ---------------------------------------------------------------------------------------------
    ftCapitalizeWords = (
        _selectTableName as text,
        _columnList as list
    ) as table =>
    let
        // Получить все таблицы текущей Excel книги
        allTables = Excel.CurrentWorkbook(),

        // Фильтруем список таблиц по имени
        filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),

        // Проверка существования таблицы
        selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content]
                     else error "Таблица с именем '" & _selectTableName & "' не найдена в текущей книге.",

        // Проверка наличия обязательных столбцов
        validatedTable = 
            let
                tableColumns = Table.ColumnNames(selectTable),
                missingColumns = List.Select(_columnList, each not List.Contains(tableColumns, _))
            in
                if List.Count(missingColumns) > 0 then
                    error "В таблице '" & _selectTableName & "' отсутствуют столбцы: " & Text.Combine(missingColumns, ", ")
                else
                    selectTable,

        // Применяем преобразование
        Transformations = List.Transform(
            _columnList,
            each {_, (cell) => if cell <> null then Text.Proper(Text.From(cell)) else null, type text}
        ),
        result = Table.TransformColumns(validatedTable, Transformations)
    in
        result,

    // ---------------------------------------------------------------------------------------------
    // Функция трансформации 'ftToUpperCase'
    // Преобразует текст в верхний регистр.
    //
    // аргументы:
    // _selectTableName as text - текстовое имя таблицы
    // _columnList as list      - список имен столбцов у которых необходимо изменить регистр значений
    //
    // возвращаемое значение:
    // table                    - таблица
    //
    // ---------------------------------------------------------------------------------------------
    ftToUpperCase = (
        _selectTableName as text,
        _columnList as list
    ) as table =>
    let
        // Получить все таблицы текущей Excel книги
        allTables = Excel.CurrentWorkbook(),

        // Фильтруем список таблиц по имени
        filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),

        // Проверка существования таблицы
        selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content]
                      else error "Таблица с именем '" & _selectTableName & "' не найдена в текущей книге.",

        // Проверка наличия обязательных столбцов
        validatedTable = 
            let
                tableColumns = Table.ColumnNames(selectTable),
                missingColumns = List.Select(_columnList, each not List.Contains(tableColumns, _))
            in
                if List.Count(missingColumns) > 0 then
                    error "В таблице '" & _selectTableName & "' отсутствуют столбцы: " & Text.Combine(missingColumns, ", ")
                else
                    selectTable,

        Transformations = List.Transform(
            _columnList,
            each {_, (cell) => if cell <> null then Text.Upper(Text.From(cell)) else null, type text}
        ),
        result = Table.TransformColumns(validatedTable, Transformations)
    in
        result,

    // ---------------------------------------------------------------------------------------------
    // Функция трансформации 'ftToLowerCase'
    // Преобразует текст в нижний регистр.
    //
    // аргументы:
    // _selectTableName as text - текстовое имя таблицы
    // _columnList as list      - список имен столбцов у которых необходимо изменить регистр значений
    //
    // возвращаемое значение:
    // table                    - таблица
    //
    // ---------------------------------------------------------------------------------------------
    ftToLowerCase = (
        _selectTableName as text,
        _columnList as list
    ) as table =>
    let
        // Получить все таблицы текущей Excel книги
        allTables = Excel.CurrentWorkbook(),

        // Фильтруем список таблиц по имени
        filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),

        // Проверка существования таблицы
        selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content]
                      else error "Таблица с именем '" & _selectTableName & "' не найдена в текущей книге.",

        // Проверка наличия обязательных столбцов
        validatedTable = 
            let
                tableColumns = Table.ColumnNames(selectTable),
                missingColumns = List.Select(_columnList, each not List.Contains(tableColumns, _))
            in
                if List.Count(missingColumns) > 0 then
                    error "В таблице '" & _selectTableName & "' отсутствуют столбцы: " & Text.Combine(missingColumns, ", ")
                else
                    selectTable,

        Transformations = List.Transform(
            _columnList,
            each {_, (cell) => if cell <> null then Text.Lower(Text.From(cell)) else null, type text}
        ),
        result = Table.TransformColumns(validatedTable, Transformations)
    in
        result,

    // ---------------------------------------------------------------------------------------------
    // Функция трансформации 'ftTrimTextColumns'
    // Удаляет ведущие и завершающие пробелы в текстовых столбцах.
    //
    // аргументы:
    // _selectTableName as text - текстовое имя таблицы
    // _columnList as list      - список имен столбцов у которых необходимо удалить ведущие и завершающие пробелы
    //
    // возвращаемое значение:
    // table                    - таблица
    //
    // ---------------------------------------------------------------------------------------------
    ftTrimTextColumns = (
        _selectTableName as text,
        _columnList as list
    ) as table =>
    let
        // Получить все таблицы текущей Excel книги
        allTables = Excel.CurrentWorkbook(),

        // Фильтруем список таблиц по имени
        filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),

        // Проверка существования таблицы
        selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content]
                      else error "Таблица с именем '" & _selectTableName & "' не найдена в текущей книге.",

        // Проверка наличия обязательных столбцов
        validatedTable = 
            let
                tableColumns = Table.ColumnNames(selectTable),
                missingColumns = List.Select(_columnList, each not List.Contains(tableColumns, _))
            in
                if List.Count(missingColumns) > 0 then
                    error "В таблице '" & _selectTableName & "' отсутствуют столбцы: " & Text.Combine(missingColumns, ", ")
                else
                    selectTable,

        Transformations = List.Transform(
            _columnList,
            each {_, (cell) => 
                if Value.Type(cell) = type text 
                then Text.Trim(cell) 
                else cell, Value.Type(cell)}
        ),
        result = Table.TransformColumns(validatedTable, Transformations)
    in
        result,
    
    // ---------------------------------------------------------------------------------------------
    // Функция трансформации 'ftNormalizeDates'
    // Нормализует даты в указанном столбце.
    //
    // аргументы:
    // _selectTableName as text         - текстовое имя таблицы
    // _columnName as text              - имя столбца с датой
    //
    // возвращаемое значение:
    // table                            - таблица
    //
    // ---------------------------------------------------------------------------------------------
    ftNormalizeDates = (
        _selectTableName as text,
        _columnName as text
    ) as table =>
    let
        // Получить все таблицы текущей Excel книги
        allTables = Excel.CurrentWorkbook(),

        // Фильтруем список таблиц по имени
        filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),

        // Проверка существования таблицы
        selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content]
                      else error "Таблица с именем '" & _selectTableName & "' не найдена в текущей книге.",

        // Проверка наличия обязательного столбца
        validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then
            error "В таблице '" & _selectTableName & "' отсутствует обязательный столбец: " & _columnName
        else
            selectTable,
        
        // Трансформируем столбец дат
        result = Table.TransformColumns(
            validatedTable,
            {
            { _columnName, (dateValue) =>
                let
                    result =
                        if dateValue = null or Text.Trim(Text.From(dateValue)) = "" 
                        then dateValue
                        else if Value.Is(dateValue, type date) or Value.Is(dateValue, type datetime) 
                        then Date.From(dateValue)
                        else
                            let
                                textValue = Text.Trim(Text.From(dateValue)),
                                parts = Text.Split(textValue, "."),
                                normalizedDate =
                                    if List.Count(parts) = 3 then
                                        let
                                            day = Number.FromText(parts{0}),
                                            month = Number.FromText(parts{1}),
                                            yearRaw = parts{2},
                                            year = 
                                                if Text.Length(yearRaw) = 2 then 
                                                    if Number.FromText(yearRaw) >= 31 then 1900 + Number.FromText(yearRaw)
                                                    else 2000 + Number.FromText(yearRaw)
                                                else if Text.Length(yearRaw) = 4 then 
                                                    Number.FromText(yearRaw)
                                                else
                                                    error "Некорректный год в дате: " & textValue,
                                            dateCheck = try #date(year, month, day) otherwise error "Неверная дата: " & textValue
                                        in dateCheck
                                    else error "Неверный формат даты: " & textValue
                            in
                                if Value.Is(normalizedDate, type date) then normalizedDate
                                else error "Некорректное значение: " & textValue
                in
                    result
            }, type date
            }
        )
    in
        result,
    
    // =============================================================================================
    // ФУНКЦИИ ВАЛИДАЦИИ (ВЕКТОРНЫЕ)
    // =============================================================================================

    // ---------------------------------------------------------------------------------------------
    // Вспомогательная функция для стандартизации вызова скалярных функций
    // ---------------------------------------------------------------------------------------------
    InvokeScalarValidation = (scalarFunction as function, table as table, columnName as text, optionalArgs as list) as list =>
        let
            WithIndex = Table.AddIndexColumn(table, "__RowIndex__", 0, 1),
            AddErrorCode = Table.AddColumn(WithIndex, "ErrorCode", each
                let
                    value = Record.Field(_, columnName),
                    allArgs = {value} & optionalArgs,
                    isValid = Function.Invoke(scalarFunction, allArgs),
                    resultCode = if isValid then 0 else [__RowIndex__] + 1
                in
                    resultCode, Int64.Type),
            result = AddErrorCode[ErrorCode]
        in
            result,

    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckNonEmpty' (Векторная)
    // ---------------------------------------------------------------------------------------------
    fvCheckNonEmpty = (_selectTableName as text, _columnName as text) as list =>
        let
            allTables = Excel.CurrentWorkbook(),
            filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
            selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
            validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
            result = InvokeScalarValidation(QueryUtilsFunction_scalar[fvCheckNonEmpty_scalar], validatedTable, _columnName, {})
        in
            result,

    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckInteger' (Векторная)
    // ---------------------------------------------------------------------------------------------
    fvCheckInteger = (_selectTableName as text, _columnName as text, optional _verificationMode as logical) as list =>
        let
            allTables = Excel.CurrentWorkbook(),
            filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
            selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
            validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
            result = InvokeScalarValidation(QueryUtilsFunction_scalar[fvCheckInteger_scalar], validatedTable, _columnName, {_verificationMode})
        in
            result,

    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckFloat' (Векторная)
    // ---------------------------------------------------------------------------------------------
    fvCheckFloat = (_selectTableName as text, _columnName as text, optional _verificationMode as logical) as list =>
        let
            allTables = Excel.CurrentWorkbook(),
            filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
            selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
            validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
            result = InvokeScalarValidation(QueryUtilsFunction_scalar[fvCheckFloat_scalar], validatedTable, _columnName, {_verificationMode})
        in
            result,
            
    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckDate' (Векторная)
    // ---------------------------------------------------------------------------------------------
    fvCheckDate = (_selectTableName as text, _columnName as text, optional _verificationMode as logical) as list =>
        let
            allTables = Excel.CurrentWorkbook(),
            filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
            selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
            validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
            result = InvokeScalarValidation(QueryUtilsFunction_scalar[fvCheckDate_scalar], validatedTable, _columnName, {_verificationMode})
        in
            result,
            
    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckINN_LLC' (Векторная)
    // ---------------------------------------------------------------------------------------------
    fvCheckINN_LLC = (_selectTableName as text, _columnName as text, optional _verificationMode as logical) as list =>
        let
            allTables = Excel.CurrentWorkbook(),
            filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
            selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
            validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
            result = InvokeScalarValidation(QueryUtilsFunction_scalar[fvCheckINN_LLC_scalar], validatedTable, _columnName, {_verificationMode})
        in
            result,
            
    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckINN_Personal' (Векторная)
    // ---------------------------------------------------------------------------------------------
    fvCheckINN_Personal = (_selectTableName as text, _columnName as text, optional _verificationMode as logical) as list =>
        let
            allTables = Excel.CurrentWorkbook(),
            filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
            selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
            validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
            result = InvokeScalarValidation(QueryUtilsFunction_scalar[fvCheckINN_Personal_scalar], validatedTable, _columnName, {_verificationMode})
        in
            result,
            
    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckSNILS' (Векторная)
    // ---------------------------------------------------------------------------------------------
    fvCheckSNILS = (_selectTableName as text, _columnName as text, optional _verificationMode as logical) as list =>
        let
            allTables = Excel.CurrentWorkbook(),
            filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
            selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
            validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
            result = InvokeScalarValidation(QueryUtilsFunction_scalar[fvCheckSNILS_scalar], validatedTable, _columnName, {_verificationMode})
        in
            result,
            
    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckKPP' (Векторная)
    // ---------------------------------------------------------------------------------------------
    fvCheckKPP = (_selectTableName as text, _columnName as text, optional _verificationMode as logical) as list =>
        let
            allTables = Excel.CurrentWorkbook(),
            filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
            selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
            validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
            result = InvokeScalarValidation(QueryUtilsFunction_scalar[fvCheckKPP_scalar], validatedTable, _columnName, {_verificationMode})
        in
            result,
            
    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckPhone' (Векторная)
    // ---------------------------------------------------------------------------------------------
    fvCheckPhone = (_selectTableName as text, _columnName as text, optional _verificationMode as logical) as list =>
        let
            allTables = Excel.CurrentWorkbook(),
            filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
            selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
            validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
            result = InvokeScalarValidation(QueryUtilsFunction_scalar[fvCheckPhone_scalar], validatedTable, _columnName, {_verificationMode})
        in
            result,
            
    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckEmail' (Векторная)
    // ---------------------------------------------------------------------------------------------
    fvCheckEmail = (_selectTableName as text, _columnName as text, optional _verificationMode as logical) as list =>
        let
            allTables = Excel.CurrentWorkbook(),
            filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
            selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
            validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
            result = InvokeScalarValidation(QueryUtilsFunction_scalar[fvCheckEmail_scalar], validatedTable, _columnName, {_verificationMode})
        in
            result,
            
    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckBoolean' (Векторная)
    // ---------------------------------------------------------------------------------------------
    fvCheckBoolean = (_selectTableName as text, _columnName as text, optional _verificationMode as logical) as list =>
        let
            allTables = Excel.CurrentWorkbook(),
            filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
            selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
            validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
            result = InvokeScalarValidation(QueryUtilsFunction_scalar[fvCheckBoolean_scalar], validatedTable, _columnName, {_verificationMode})
        in
            result,
            
    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckPassportSeries' (Векторная)
    // ---------------------------------------------------------------------------------------------
    fvCheckPassportSeries = (_selectTableName as text, _columnName as text, optional _verificationMode as logical) as list =>
        let
            allTables = Excel.CurrentWorkbook(),
            filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
            selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
            validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
            result = InvokeScalarValidation(QueryUtilsFunction_scalar[fvCheckPassportSeries_scalar], validatedTable, _columnName, {_verificationMode})
        in
            result,
            
    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckPassportNumber' (Векторная)
    // ---------------------------------------------------------------------------------------------
    fvCheckPassportNumber = (_selectTableName as text, _columnName as text, optional _verificationMode as logical) as list =>
        let
            allTables = Excel.CurrentWorkbook(),
            filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
            selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
            validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
            result = InvokeScalarValidation(QueryUtilsFunction_scalar[fvCheckPassportNumber_scalar], validatedTable, _columnName, {_verificationMode})
        in
            result,
            
    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckPassportUnitCode' (Векторная)
    // ---------------------------------------------------------------------------------------------
    fvCheckPassportUnitCode = (_selectTableName as text, _columnName as text, optional _verificationMode as logical) as list =>
        let
            allTables = Excel.CurrentWorkbook(),
            filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
            selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
            validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
            result = InvokeScalarValidation(QueryUtilsFunction_scalar[fvCheckPassportUnitCode_scalar], validatedTable, _columnName, {_verificationMode})
        in
            result,

    // ---------------------------------------------------------------------------------------------
    // Функции с несколькими аргументами (требуют отдельной обертки)
    // ---------------------------------------------------------------------------------------------
    fvCheckTextLength = (_selectTableName as text, _columnName as text, _minLength as number, _maxLength as number, optional _verificationMode as logical) as list =>
        let
            allTables = Excel.CurrentWorkbook(),
            filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
            selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
            validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
            result = InvokeScalarValidation(QueryUtilsFunction_scalar[fvCheckTextLength_scalar], validatedTable, _columnName, {_minLength, _maxLength, _verificationMode})
        in
            result,
        
    fvCheckIntegerRange = (_selectTableName as text, _columnName as text, _minValue as number, _maxValue as number, optional _verificationMode as logical) as list =>
        let
            allTables = Excel.CurrentWorkbook(),
            filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
            selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
            validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
            result = InvokeScalarValidation(QueryUtilsFunction_scalar[fvCheckIntegerRange_scalar], validatedTable, _columnName, {_minValue, _maxValue, _verificationMode})
        in
            result,
        
    fvCheckFloatRange = (_selectTableName as text, _columnName as text, _minValue as number, _maxValue as number, optional _verificationMode as logical) as list =>
        let
            allTables = Excel.CurrentWorkbook(),
            filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
            selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
            validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
            result = InvokeScalarValidation(QueryUtilsFunction_scalar[fvCheckFloatRange_scalar], validatedTable, _columnName, {_minValue, _maxValue, _verificationMode})
        in
            result,
        
    fvCheckDateRange = (_selectTableName as text, _columnName as text, _minDate as date, _maxDate as date, optional _verificationMode as logical) as list =>
        let
            allTables = Excel.CurrentWorkbook(),
            filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
            selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
            validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
            result = InvokeScalarValidation(QueryUtilsFunction_scalar[fvCheckDateRange_scalar], validatedTable, _columnName, {_minDate, _maxDate, _verificationMode})
        in
            result,

    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckID'
    // Проверяет столбец-идентификатор (ID) на соответствие правилам первичного ключа:
    // - Значения должны быть целыми, положительными числами.
    // - Первое значение в столбце должно быть равно 1.
    // - Каждое последующее значение должно быть на единицу больше предыдущего.
    // - Пустые значения не допускаются.
    //
    // Вариант специально разработанный для работы в таких запросах как в ETL-процессора 'ProcessingTableQuery',
    // для совместимости у данной функции изменена сигнатура, а именно первый аргумент не ТЕКСТОВОЕ ИМЯ 
    // таблицы у которой будет производиться валидация столбца, а непосредственно ТАБЛИЦА.
    // Это делает невозможным вызов данной функции из автономных функций-запросов т.к. пользователь не может в UI 
    // передать объект power query - ТАБЛИЦУ. 
    // Для автономных функций нужно использовать другую версию функции проверки цифровых ID столбцов: fvCheckID_2
    //
    // аргументы:
    // _sourceTable as table    - таблица (объект Power Query), в которой нужно проверить столбец
    // _columnName as text      - имя столбца для проверки
    //
    // возвращаемое значение:
    // list                     - список с кодами ошибок (0 или номер строки)
    // ---------------------------------------------------------------------------------------------
    fvCheckID = (_sourceTable as table, _columnName as text) as list =>
    let
        // Проверка наличия обязательного столбца в переданной таблице
        validatedTable = if not Table.HasColumns(_sourceTable, {_columnName}) then 
                             error "Функция fvCheckID: В переданной таблице отсутствует обязательный столбец '" & _columnName & "'."
                         else
                             _sourceTable,

        // Если таблица пуста, возвращаем пустой список
        result = if Table.IsEmpty(validatedTable) then {} else
            let
                // Шаг 1: Буферизируем столбец в список для эффективной обработки
                columnAsList = Table.Column(validatedTable, _columnName),

                // Шаг 2: Используем List.Accumulate для итерации с сохранением состояния
                finalState = List.Accumulate(
                    // Итерируем по списку индексов от 0 до N-1
                    {0..List.Count(columnAsList)-1},
                    
                    // Начальное состояние "аккумулятора"
                    [
                        lastValidValue = 0,      // Ожидаемое значение для первой строки - 1 (0+1)
                        errorHasOccurred = false, // Флаг, который "включается" при первой ошибке
                        errorCodes = {}          // Список, в который собираем результат
                    ],
                    
                    // Функция, которая выполняется на каждом шаге и обновляет состояние
                    (state as record, currentIndex as number) =>
                        let
                            // Если ошибка уже была найдена на предыдущих шагах,
                            // все последующие значения автоматически невалидны.
                            newState = if state[errorHasOccurred] then
                                [
                                    lastValidValue = state[lastValidValue],
                                    errorHasOccurred = true,
                                    errorCodes = state[errorCodes] & {currentIndex + 1}
                                ]
                            // Если ошибок еще не было, выполняем полную проверку
                            else
                                let
                                    currentValue = columnAsList{currentIndex},
                                    currentNumber = try Number.From(currentValue) otherwise null,
                                    
                                    // Проверяем, является ли значение целым положительным числом
                                    isPositiveInteger = currentNumber <> null and currentNumber > 0 and Number.Round(currentNumber) = currentNumber,
                                    
                                    // Проверяем, соблюдается ли последовательность
                                    isSequenceCorrect = isPositiveInteger and (currentNumber = state[lastValidValue] + 1)
                                in
                                    // Если все проверки пройдены
                                    if isSequenceCorrect then
                                        // Обновляем состояние: ошибка не найдена, запоминаем последнее валидное значение
                                        [
                                            lastValidValue = currentNumber,
                                            errorHasOccurred = false,
                                            errorCodes = state[errorCodes] & {0}
                                        ]
                                    // Если хотя бы одна проверка не пройдена
                                    else
                                        // Обновляем состояние: "включаем" флаг ошибки
                                        [
                                            lastValidValue = state[lastValidValue],
                                            errorHasOccurred = true,
                                            errorCodes = state[errorCodes] & {currentIndex + 1}
                                        ]
                        in
                            newState
                ),
                // Шаг 3: Извлекаем из конечного состояния собранный список кодов ошибок
                errorList = finalState[errorCodes]
            in
                errorList
    in
        result,

    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckID_2'
    // Проверяет столбец-идентификатор (ID) на соответствие правилам первичного ключа:
    // - Значения должны быть целыми, положительными числами.
    // - Первое значение в столбце должно быть равно 1.
    // - Каждое последующее значение должно быть на единицу больше предыдущего.
    // - Пустые значения не допускаются.
    //
    // Это версия только для автономных валидирующих функций, которые пользователи запускают
    // через UI Excel
    //
    // аргументы:
    // _selectTableName as text - текстовое имя таблицы
    // _columnName as text      - имя столбца для проверки
    //
    // возвращаемое значение:
    // list                     - список с кодами ошибок (0 или номер строки)
    // ---------------------------------------------------------------------------------------------
    fvCheckID_2 = (_selectTableName as text, _columnName as text) as list =>
    let
        allTables = Excel.CurrentWorkbook(),
        filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),
        selectTable = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица '" & _selectTableName & "' не найдена.",
        validatedTable = if not Table.HasColumns(selectTable, {_columnName}) then error "В таблице '" & _selectTableName & "' столбец '" & _columnName & "' не найден." else selectTable,
        result = if Table.IsEmpty(validatedTable) then {} else
            let
                columnAsList = Table.Column(validatedTable, _columnName),
                finalState = List.Accumulate({0..List.Count(columnAsList)-1}, [lastValidValue = 0, errorHasOccurred = false, errorCodes = {}], (state, currentIndex) =>
                        let
                            newState = if state[errorHasOccurred] then
                                [lastValidValue = state[lastValidValue], errorHasOccurred = true, errorCodes = state[errorCodes] & {currentIndex + 1}]
                            else
                                let
                                    currentValue = columnAsList{currentIndex},
                                    currentNumber = try Number.From(currentValue) otherwise null,
                                    isPositiveInteger = currentNumber <> null and currentNumber > 0 and Number.Round(currentNumber) = currentNumber,
                                    isSequenceCorrect = isPositiveInteger and (currentNumber = state[lastValidValue] + 1)
                                in
                                    if isSequenceCorrect then
                                        [lastValidValue = currentNumber, errorHasOccurred = false, errorCodes = state[errorCodes] & {0}]
                                    else
                                        [lastValidValue = state[lastValidValue], errorHasOccurred = true, errorCodes = state[errorCodes] & {currentIndex + 1}]
                        in newState ),
                errorList = finalState[errorCodes]
            in errorList
    in
        result,


// =============================================================================================
// ОБЪЕДИНЕНИЕ ВСЕХ ФУНКЦИЙ В RECORD
// =============================================================================================
FunctionRecord = [
    // Функции трансформации
    ftCapitalizeWords       = ftCapitalizeWords,        // 1
    ftCreateFullNameColumn  = ftCreateFullNameColumn,   // 2
    ftGetFullNameColumn     = ftGetFullNameColumn,      // 3
    ftNormalizeDates        = ftNormalizeDates,         // 4
    ftRemoveColumnsStrict   = ftRemoveColumnsStrict,    // 5
    ftReorderColumnsStrict  = ftReorderColumnsStrict,   // 6
    ftToLowerCase           = ftToLowerCase,            // 7
    ftToUpperCase           = ftToUpperCase,            // 8
    ftTrimTextColumns       = ftTrimTextColumns,        // 9
    
    // Функции валидации
    fvCheckBoolean          = fvCheckBoolean,           // 10
    fvCheckDate             = fvCheckDate,              // 11
    fvCheckDateRange        = fvCheckDateRange,         // 12
    fvCheckEmail            = fvCheckEmail,             // 13
    fvCheckFloat            = fvCheckFloat,             // 14
    fvCheckINN_LLC          = fvCheckINN_LLC,           // 15
    fvCheckINN_Personal     = fvCheckINN_Personal,      // 16
    fvCheckSNILS            = fvCheckSNILS,             // 17
    fvCheckInteger          = fvCheckInteger,           // 18
    fvCheckKPP              = fvCheckKPP,               // 19
    fvCheckNonEmpty         = fvCheckNonEmpty,          // 20
    fvCheckIntegerRange     = fvCheckIntegerRange,      // 21
    fvCheckFloatRange       = fvCheckFloatRange,        // 22
    fvCheckPhone            = fvCheckPhone,             // 23
    fvCheckTextLength       = fvCheckTextLength,        // 24
    fvCheckPassportNumber   = fvCheckPassportNumber,    // 25
    fvCheckPassportSeries   = fvCheckPassportSeries,    // 26
    fvCheckPassportUnitCode = fvCheckPassportUnitCode,  // 27
    fvCheckID               = fvCheckID,                // 28
    fvCheckID_2             = fvCheckID_2               // 29
]

in
    FunctionRecord