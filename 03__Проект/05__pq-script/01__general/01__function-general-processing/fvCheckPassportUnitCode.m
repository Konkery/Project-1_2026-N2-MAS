// ---------------------------------------------------------------------------------------------
// Функция валидации 'fvCheckPassportUnitCode'
// Проверяет код подразделения паспорта РФ на соответствие формату (XXX-XXX).
//
// Функция поддерживает два режима проверки данных: "мягкий" и "жесткий".
//
// аргументы:
// _selectTableName as text              - текстовое имя таблицы
// _columnName as text                   - имя столбца для проверки
// optional _verificationMode as logical - режим валидации (true/null = жесткий, false = мягкий)
//
// возвращаемое значение:
// list                                  - список с кодами ошибок (0 или номер строки)
// ---------------------------------------------------------------------------------------------
(
    _selectTableName as text,
    _columnName as text,
    optional _verificationMode as logical
) as list =>
let
    UF = QueryUtilsFunction,
    result = UF[fvCheckPassportUnitCode](
        _selectTableName,
        _columnName,
        _verificationMode
    )
in
    result