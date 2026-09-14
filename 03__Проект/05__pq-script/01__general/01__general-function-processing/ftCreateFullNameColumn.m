// ---------------------------------------------------------------------------------------------
// Функция трансформации 'ftCreateFullNameColumn'
// Создает столбец "FULL_NAME' из трех исходных столбцов ФИО в формате Петров С.Н.
// и размещает его в указанной позиции.
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
// ВАЖНО: 
// данная функция предназначена только (!) для создания и использования в автономном режиме
// т.е. она должна быть создана в Excel в виде автономной, самостоятельно вызываемой функции,
// и не предназначена для вызова из query.
// Вариант данной функции для вызова из другого query находиться в агрегатном
// файле 'QueryUtilsFunction.m'
//
// ---------------------------------------------------------------------------------------------
(
    _selectTableName as text, 
    _lastNameCol as text, 
    _firstNameCol as text, 
    _patronymicCol as text,
    _position as number
) as table =>
let
    // Получение модульной функции из 'QueryUtilsFunction'
    UF = QueryUtilsFunction,

    result = UF[ftCreateFullNameColumn](
        _selectTableName,
        _lastNameCol,
        _firstNameCol,
        _patronymicCol,
        _position
    )
in
    result