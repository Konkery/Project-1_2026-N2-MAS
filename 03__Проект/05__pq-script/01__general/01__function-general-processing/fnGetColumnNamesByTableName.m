// fnGetColumnNamesByTableName - функция принимает имя таблицы (текст) и возвращает список имён столбцов
// ---------------------------------------------------------------------------------------------
// Функция информационная 'fnGetColumnNamesByTableName'
// Возвращает список имён столбцов заданной таблицы.
// 
// аргументы:
// _selectTableName as text - текстовое имя таблицы
//
// возвращаемое значение:
// list                     - список столбцов
//
// ---------------------------------------------------------------------------------------------

(
    _selectTableName as text
) as list =>
let
    // Получаем все таблицы текущей книги Excel с помощью встроенной функции Excel.CurrentWorkbook()
    allTables = Excel.CurrentWorkbook(),

    // Фильтруем список таблиц, оставляя только ту, имя которой совпадает с переданным параметром tableName
    filteredTable = Table.SelectRows(allTables, each [Name] = _selectTableName),

    // Извлекаем содержимое таблицы из столбца Content (это объект таблицы Power Query)
    tableContent = if Table.RowCount(filteredTable) = 1 then filteredTable{0}[Content] else error "Таблица с таким именем не найдена",

    // Получаем список имён столбцов из выбранной таблицы
    result = Table.ColumnNames(tableContent)
in
    result