// ---------------------------------------------------------------------------------------------
// Функция валидации 'fvCheckPassportSeries'
// Проверяет серию паспорта РФ на соответствие формату (4 цифры).
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
    result = UF[fvCheckPassportSeries](
        _selectTableName,
        _columnName,
        _verificationMode
    )
in
    result