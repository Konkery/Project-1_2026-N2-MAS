let
    // =============================================================================================
    //
    // УТИЛИТАРНЫЙ ЗАПРОС 'UtilsGetPathTable'
    // версия rev.01 v01
    //
    // Данный запрос-функция предназначен для получения полного пути к файлу-источнику
    // на основе его табличного имени. Поиск производится в централизованном файле-справочнике,
    // содержащем список всех таблиц-справочников первого типа.
    //
    // =============================================================================================
    // ---------------------------------------------------------------------------------------------
    // Функция 'UtilsGetPathTable'
    // Находит в таблице 'TabTableList' запись, соответствующую искомой таблице, и
    // возвращает полный путь к файлу, в котором эта таблица хранится.
    //
    // аргументы:
    //    _sourceTable as text  - текстовое имя таблицы-справочника первого типа,
    //                            путь к которой необходимо найти.
    //
    // возвращаемое значение:
    //    text                  - полный путь к файлу Excel (например, "C:\...\FileName.xlsx").
    //
    // ---------------------------------------------------------------------------------------------
    UtilsGetPathTable = (_sourceTable as text) as text =>
        let
            // --- ШАГ 1: Определение констант для доступа к файлу-справочнику ---
            // Задаем статический путь, имя файла и имя таблицы, где хранится список всех таблиц
            masterFilePath = "c:\OneDrive\01__LLC_MAS_2025\КИТ\01__КБД\03__Excel\TableList\",
            masterFileName = "Таблицы-КБД список #2025.xlsx",
            masterTableName = "TabTableList",
            fullPathToMasterFile = masterFilePath & masterFileName,
            // --- ШАГ 2: Загрузка и валидация таблицы-справочника 'TabTableList' ---
            // Загружаем книгу-справочник. В случае ошибки выводим информативное сообщение.
            masterWorkbook =
                try
                    Excel.Workbook(File.Contents(fullPathToMasterFile), null, true)
                otherwise
                    error
                        "Не удалось найти или открыть файл-справочник по пути: '"
                            & fullPathToMasterFile
                            & "'. Проверьте путь и доступ к файлу.",
            // Извлекаем из книги таблицу-список. В случае ошибки также выводим сообщение.
            sourceTable =
                try
                    masterWorkbook{[Item = masterTableName, Kind = "Table"]}[Data]
                otherwise
                    error
                        "В файле '"
                            & masterFileName
                            & "' не найдена таблица-справочник с именем '"
                            & masterTableName
                            & "'.",
            // Проверяем наличие обязательных столбцов, необходимых для построения пути
            requiredColumns = {"TABLE NAME", "PATH", "FILE NAME"},
            validatedTable =
                if not Table.HasColumns(sourceTable, requiredColumns) then
                    error
                        "В таблице '"
                            & masterTableName
                            & "' отсутствуют один или несколько обязательных столбцов: "
                            & Text.Combine(requiredColumns, ", ")
                else
                    sourceTable,
            // --- ШАГ 3: Поиск строки и валидация ее содержимого ---
            // Фильтруем таблицу-справочник, чтобы найти строку с искомым именем таблицы
            filteredRow = Table.SelectRows(validatedTable, each Record.Field(_, "TABLE NAME") = _sourceTable),
            // Проверяем, что найдена ровно одна строка, чтобы избежать неоднозначности
            checkedRow =
                if Table.RowCount(filteredRow) = 0 then
                    error
                        "В справочнике '"
                            & masterTableName
                            & "' не найдена запись для таблицы с именем '"
                            & _sourceTable
                            & "'."
                else if Table.RowCount(filteredRow) > 1 then
                    error
                        "В справочнике '"
                            & masterTableName
                            & "' найдено несколько записей для таблицы '"
                            & _sourceTable
                            & "'. Имя таблицы должно быть уникальным."
                else
                    filteredRow{0},
            // Получаем запись (record) найденной строки
            // Извлекаем значение пути и проверяем его на пустоту
            pathPart = Record.Field(checkedRow, "PATH"),
            validatedPath =
                if pathPart = null or Text.Trim(Text.From(pathPart)) = "" then
                    error
                        "Для таблицы '"
                            & _sourceTable
                            & "' в справочнике '"
                            & masterTableName
                            & "' не указан путь в столбце 'PATH'."
                else
                    pathPart,
            // Извлекаем значение имени файла и проверяем его на пустоту
            fileNamePart = Record.Field(checkedRow, "FILE NAME"),
            validatedFileName =
                if fileNamePart = null or Text.Trim(Text.From(fileNamePart)) = "" then
                    error
                        "Для таблицы '"
                            & _sourceTable
                            & "' в справочнике '"
                            & masterTableName
                            & "' не указано имя файла в столбце 'FILE NAME'."
                else
                    fileNamePart,
            // --- ШАГ 4: Формирование результата ---
            // Формируем и возвращаем итоговый полный путь путем конкатенации
            result = validatedPath & validatedFileName
        in
            result
in
    UtilsGetPathTable
