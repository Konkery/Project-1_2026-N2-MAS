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
// ВАЖНО: 
// данная функция предназначена только (!) для создания и использования в автономном режиме
// т.е. она должна быть создана в Excel в виде автономной, самостоятельно вызываемой функции,
// и не предназначена для вызова из query.
// Вариант данной функции для вызова из другого query находиться в агрегатном
// файле 'QueryUtilsFunction.m'
// ---------------------------------------------------------------------------------------------
(
    _selectTableName as text, 
    _lastNameCol as text, 
    _firstNameCol as text, 
    _patronymicCol as text
) as list =>
let
    // Получение модульной функции из 'QueryUtilsFunction'
    UF = QueryUtilsFunction,
    
    result = UF[ftGetFullNameColumn](
        _selectTableName,
        _lastNameCol,
        _firstNameCol,
        _patronymicCol
    )
in
    result