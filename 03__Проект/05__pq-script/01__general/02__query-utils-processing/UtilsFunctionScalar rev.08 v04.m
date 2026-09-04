let
    // =============================================================================================
    //
    // ФУНКЦИИ ВАЛИДАЦИИ ДАННЫХ - СКАЛЯРНЫЕ
    // версия rev.08 v04
    //
    // Данный агрегатный Query содержит набор функций выполняющих:
    // 1. валидацию данных.
    //
    // =============================================================================================

    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckNonEmpty_scalar'
    // Скалярная функция. Проверяет одно значение на пустоту.
    //
    // аргументы:
    // value as any - проверяемое значение
    //
    // возвращаемое значение:
    // logical      - true (валидно) или false (невалидно)
    // ---------------------------------------------------------------------------------------------
    fvCheckNonEmpty_scalar = (value as any) as logical =>
        not (value = null or (Value.Is(value, type text) and Text.Trim(Text.From(value)) = "")),

    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckInteger_scalar'
    // Скалярная функция. Проверяет, является ли одно значение целым числом.
    //
    // аргументы:
    // value as any                          - проверяемое значение
    // optional _verificationMode as logical - режим валидации (true/null = жесткий, false = мягкий)
    //
    // возвращаемое значение:
    // logical                               - true (валидно) или false (невалидно)
    // ---------------------------------------------------------------------------------------------
    fvCheckInteger_scalar = (value as any, optional _verificationMode as logical) as logical =>
        let
            isHardMode = if _verificationMode = false then false else true,
            isEmpty = value = null or (Value.Is(value, type text) and Text.Trim(Text.From(value)) = ""),
            result = if isEmpty then not isHardMode else
                let
                    numberValue = try Number.From(value) otherwise null
                in
                    numberValue <> null and Number.Round(numberValue) = numberValue
        in
            result,
    
    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckFloat_scalar'
    // Скалярная функция. Проверяет, является ли одно значение числом (целым или с плавающей точкой).
    //
    // аргументы:
    // value as any                          - проверяемое значение
    // optional _verificationMode as logical - режим валидации (true/null = жесткий, false = мягкий)
    //
    // возвращаемое значение:
    // logical                               - true (валидно) или false (невалидно)
    // ---------------------------------------------------------------------------------------------
    fvCheckFloat_scalar = (value as any, optional _verificationMode as logical) as logical =>
        let
            isHardMode = if _verificationMode = false then false else true,
            isEmpty = value = null or (Value.Is(value, type text) and Text.Trim(Text.From(value)) = ""),
            result = if isEmpty then not isHardMode else
                let
                    // Просто проверяем, можно ли значение преобразовать в число.
                    // Number.From корректно обработает и целые, и дробные значения.
                    isNumber = (try Number.From(value) otherwise null) <> null
                in
                    isNumber
        in
            result,

    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckDate_scalar'
    // Скалярная функция. Проверяет одно значение на соответствие типу ДАТА.
    //
    // аргументы:
    // value as any                          - проверяемое значение
    // optional _verificationMode as logical - режим валидации (true/null = жесткий, false = мягкий)
    //
    // возвращаемое значение:
    // logical                               - true (валидно) или false (невалидно)
    // ---------------------------------------------------------------------------------------------
    fvCheckDate_scalar = (value as any, optional _verificationMode as logical) as logical =>
        let
            isHardMode = if _verificationMode = false then false else true,
            isEmpty = value = null or (Value.Type(value) = type text and Text.Trim(Text.From(value)) = ""),
            result = if isEmpty then not isHardMode else
                let
                    parsedDate = try (
                        if Value.Is(value, type date) or Value.Is(value, type datetime) then Date.From(value)
                        else let
                            textValue = Text.Trim(Text.From(value)),
                            parts = Text.Split(textValue, "."),
                            dateCheck = if List.Count(parts) = 3 then let
                                day = Number.FromText(parts{0}),
                                month = Number.FromText(parts{1}),
                                yearRaw = parts{2},
                                year = if Text.Length(yearRaw) = 2 then if Number.FromText(yearRaw) >= 31 then 1900 + Number.FromText(yearRaw) else 2000 + Number.FromText(yearRaw)
                                    else if Text.Length(yearRaw) = 4 then Number.FromText(yearRaw)
                                    else error "Некорректный год",
                                checked = #date(year, month, day)
                            in checked else error "Неверный формат"
                        in dateCheck
                    ) otherwise null
                in
                    parsedDate <> null
        in
            result,

    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckINN_LLC_scalar'
    // Скалярная функция. Проверяет ИНН юридического лица (10 цифр) с контрольной суммой.
    //
    // аргументы:
    // value as any                          - проверяемое значение
    // optional _verificationMode as logical - режим валидации (true/null = жесткий, false = мягкий)
    //
    // возвращаемое значение:
    // logical                               - true (валидно) или false (невалидно)
    // ---------------------------------------------------------------------------------------------
    fvCheckINN_LLC_scalar = (value as any, optional _verificationMode as logical) as logical =>
        let
            isHardMode = if _verificationMode = false then false else true,
            strValue = if value = null then "" else Text.Trim(Text.From(value)),
            isEmpty = strValue = "",
            result = if isEmpty then not isHardMode else
                if Text.Length(strValue) <> 10 or Text.Length(Text.Select(strValue, {"0".."9"})) <> 10 then false
                else
                    let
                        digitList = List.Transform(Text.ToList(strValue), each Number.FromText(_)),
                        coeffs = {2, 4, 10, 3, 5, 9, 4, 6, 8},
                        sum = List.Sum(List.Transform({0..8}, (i) => digitList{i} * coeffs{i})),
                        controlDigit = Number.Mod(Number.Mod(sum, 11), 10)
                    in
                        controlDigit = digitList{9}
        in
            result,

    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckINN_Personal_scalar'
    // Скалярная функция. Проверяет ИНН физического лица (12 цифр) с контрольными суммами.
    //
    // аргументы:
    // value as any                          - проверяемое значение
    // optional _verificationMode as logical - режим валидации (true/null = жесткий, false = мягкий)
    //
    // возвращаемое значение:
    // logical                               - true (валидно) или false (невалидно)
    // ---------------------------------------------------------------------------------------------
    fvCheckINN_Personal_scalar = (value as any, optional _verificationMode as logical) as logical =>
        let
            isHardMode = if _verificationMode = false then false else true,
            strValue = if value = null then "" else Text.Trim(Text.From(value)),
            isEmpty = strValue = "",
            result = if isEmpty then not isHardMode else
                if Text.Length(strValue) <> 12 or Text.Length(Text.Select(strValue, {"0".."9"})) <> 12 then false
                else
                    let
                        digitList = List.Transform(Text.ToList(strValue), each Number.FromText(_)),
                        coeffs1 = {7, 2, 4, 10, 3, 5, 9, 4, 6, 8},
                        sum1 = List.Sum(List.Transform({0..9}, (i) => digitList{i} * coeffs1{i})),
                        controlDigit1 = Number.Mod(Number.Mod(sum1, 11), 10),
                        isChecksum1Valid = (controlDigit1 = digitList{10}),
                        coeffs2 = {3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8},
                        sum2 = List.Sum(List.Transform({0..10}, (i) => digitList{i} * coeffs2{i})),
                        controlDigit2 = Number.Mod(Number.Mod(sum2, 11), 10),
                        isChecksum2Valid = (controlDigit2 = digitList{11})
                    in
                        isChecksum1Valid and isChecksum2Valid
        in
            result,

    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckSNILS_scalar'
    // Скалярная функция. Проверяет СНИЛС (формат и контрольная сумма).
    //
    // аргументы:
    // value as any                          - проверяемое значение
    // optional _verificationMode as logical - режим валидации (true/null = жесткий, false = мягкий)
    //
    // возвращаемое значение:
    // logical                               - true (валидно) или false (невалидно)
    // ---------------------------------------------------------------------------------------------
    fvCheckSNILS_scalar = (value as any, optional _verificationMode as logical) as logical =>
        let
            isHardMode = if _verificationMode = false then false else true,
            strValue = if value = null then "" else Text.Trim(Text.From(value)),
            isEmpty = strValue = "",
            result = if isEmpty then not isHardMode else
                if Text.Length(strValue) <> 14 or Text.Middle(strValue, 3, 1) <> "-" or Text.Middle(strValue, 7, 1) <> "-" or Text.Middle(strValue, 11, 1) <> " " then false
                else
                    let
                        numberPart = Text.Remove(strValue, {"-", " "}),
                        isAllDigits = Text.Length(Text.Select(numberPart, {"0".."9"})) = 11
                    in
                        if not isAllDigits then false
                        else
                            let
                                digitList = List.Transform(Text.ToList(Text.Start(numberPart, 9)), each Number.FromText(_)),
                                sum = List.Sum(List.Transform({0..8}, (i) => digitList{i} * (9 - i))),
                                calculatedChecksum = if sum < 100 then sum else if sum = 100 or sum = 101 then 0 else let remainder = Number.Mod(sum, 101) in if remainder = 100 then 0 else remainder,
                                actualChecksum = Number.FromText(Text.End(numberPart, 2))
                            in
                                calculatedChecksum = actualChecksum
        in
            result,

    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckKPP_scalar'
    // Скалярная функция. Проверяет КПП (9 цифр).
    //
    // аргументы:
    // value as any                          - проверяемое значение
    // optional _verificationMode as logical - режим валидации (true/null = жесткий, false = мягкий)
    //
    // возвращаемое значение:
    // logical                               - true (валидно) или false (невалидно)
    // ---------------------------------------------------------------------------------------------
    fvCheckKPP_scalar = (value as any, optional _verificationMode as logical) as logical =>
        let
            isHardMode = if _verificationMode = false then false else true,
            strValue = if value = null then "" else Text.Trim(Text.From(value)),
            isEmpty = strValue = "",
            result = if isEmpty then not isHardMode else
                let
                    isAllDigits = Text.Length(Text.Select(strValue, {"0".."9"})) = Text.Length(strValue),
                    hasCorrectLength = Text.Length(strValue) = 9
                in
                    isAllDigits and hasCorrectLength
        in
            result,

    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckTextLength_scalar'
    // Скалярная функция. Проверяет длину текста.
    //
    // аргументы:
    // value as any                          - проверяемое значение
    // _minLength as number                  - минимальная длина
    // _maxLength as number                  - максимальная длина
    // optional _verificationMode as logical - режим валидации (true/null = жесткий, false = мягкий)
    //
    // возвращаемое значение:
    // logical                               - true (валидно) или false (невалидно)
    // ---------------------------------------------------------------------------------------------
    fvCheckTextLength_scalar = (value as any, _minLength as number, _maxLength as number, optional _verificationMode as logical) as logical =>
        let
            isHardMode = if _verificationMode = false then false else true,
            strValue = if value = null then "" else Text.Trim(Text.From(value)),
            isEmpty = strValue = "",
            result = if isEmpty then not isHardMode else
                let
                    length = Text.Length(strValue)
                in
                    length >= _minLength and length <= _maxLength
        in
            result,

    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckIntegerRange_scalar'
    // Скалярная функция. Проверяет, входит ли целое число в диапазон.
    //
    // аргументы:
    // value as any                          - проверяемое значение
    // _minValue as number                   - минимальное значение
    // _maxValue as number                   - максимальное значение
    // optional _verificationMode as logical - режим валидации (true/null = жесткий, false = мягкий)
    //
    // возвращаемое значение:
    // logical                               - true (валидно) или false (невалидно)
    // ---------------------------------------------------------------------------------------------
    fvCheckIntegerRange_scalar = (value as any, _minValue as number, _maxValue as number, optional _verificationMode as logical) as logical =>
        let
            isHardMode = if _verificationMode = false then false else true,
            isEmpty = value = null or (Value.Is(value, type text) and Text.Trim(Text.From(value)) = ""),
            result = if isEmpty then not isHardMode else
                let
                    numberValue = try Number.From(value) otherwise null
                in
                    numberValue <> null and numberValue >= _minValue and numberValue <= _maxValue
        in
            result,
        
    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckFloatRange_scalar'
    // Скалярная функция. Проверяет, входит ли число (целое или с плавающей точкой) в диапазон.
    //
    // аргументы:
    // value as any                          - проверяемое значение
    // _minValue as number                   - минимальное значение
    // _maxValue as number                   - максимальное значение
    // optional _verificationMode as logical - режим валидации (true/null = жесткий, false = мягкий)
    //
    // возвращаемое значение:
    // logical                               - true (валидно) или false (невалидно)
    // ---------------------------------------------------------------------------------------------
    fvCheckFloatRange_scalar = (value as any, _minValue as number, _maxValue as number, optional _verificationMode as logical) as logical =>
        let
            isHardMode = if _verificationMode = false then false else true,
            isEmpty = value = null or (Value.Is(value, type text) and Text.Trim(Text.From(value)) = ""),
            result = if isEmpty then not isHardMode else
                let
                    // Пытаемся преобразовать значение в число
                    numberValue = try Number.From(value) otherwise null,
                    
                    // Проверяем, что значение является числом И входит в заданный диапазон
                    isInRange = if numberValue <> null then
                                    numberValue >= _minValue and numberValue <= _maxValue
                                else
                                    false
                in
                    isInRange
        in
            result,

    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckDateRange_scalar'
    // Скалярная функция. Проверяет, входит ли дата в диапазон.
    //
    // аргументы:
    // value as any                          - проверяемое значение
    // _minDate as date                      - минимальная дата
    // _maxDate as date                      - максимальная дата
    // optional _verificationMode as logical - режим валидации (true/null = жесткий, false = мягкий)
    //
    // возвращаемое значение:
    // logical                               - true (валидно) или false (невалидно)
    // ---------------------------------------------------------------------------------------------
    fvCheckDateRange_scalar = (value as any, _minDate as date, _maxDate as date, optional _verificationMode as logical) as logical =>
        let
            isHardMode = if _verificationMode = false then false else true,
            isEmpty = value = null or (Value.Is(value, type text) and Text.Trim(Text.From(value)) = ""),
            result = if isEmpty then not isHardMode else
                let
                    parsedDate = try ( if Value.Is(value, type date) or Value.Is(value, type datetime) then Date.From(value) else let textValue = Text.Trim(Text.From(value)), parts = Text.Split(textValue, "."), dateCheck = if List.Count(parts) = 3 then let day = Number.FromText(parts{0}), month = Number.FromText(parts{1}), yearRaw = parts{2}, year = if Text.Length(yearRaw) = 2 then if Number.FromText(yearRaw) >= 31 then 1900 + Number.FromText(yearRaw) else 2000 + Number.FromText(yearRaw) else if Text.Length(yearRaw) = 4 then Number.FromText(yearRaw) else error "e", checked = #date(year, month, day) in checked else error "e" in dateCheck ) otherwise null
                in
                    parsedDate <> null and parsedDate >= _minDate and parsedDate <= _maxDate
        in
            result,

    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckPhone_scalar'
    // Скалярная функция. Проверяет номер телефона на соответствие формату +7 XXX XXX-XX-XX.
    //
    // аргументы:
    // value as any                          - проверяемое значение
    // optional _verificationMode as logical - режим валидации (true/null = жесткий, false = мягкий)
    //
    // возвращаемое значение:
    // logical                               - true (валидно) или false (невалидно)
    // ---------------------------------------------------------------------------------------------
    fvCheckPhone_scalar = (value as any, optional _verificationMode as logical) as logical =>
        let
            isHardMode = if _verificationMode = false then false else true,
            strValue = if value = null then "" else Text.Trim(Text.From(value)),
            isEmpty = strValue = "",
            result = if isEmpty then not isHardMode else
                let
                    isValidFormat = if Text.Length(strValue) <> 16 then false else if not Text.StartsWith(strValue, "+7 ") then false else if Text.Middle(strValue, 6, 1) <> " " then false else if Text.Middle(strValue, 10, 1) <> "-" then false else if Text.Middle(strValue, 13, 1) <> "-" then false else true,
                    allDigits = if isValidFormat then let part1 = Text.Select(Text.Middle(strValue, 3, 3), {"0".."9"}), part2 = Text.Select(Text.Middle(strValue, 7, 3), {"0".."9"}), part3 = Text.Select(Text.Middle(strValue, 11, 2), {"0".."9"}), part4 = Text.Select(Text.Middle(strValue, 14, 2), {"0".."9"}) in Text.Length(part1) = 3 and Text.Length(part2) = 3 and Text.Length(part3) = 2 and Text.Length(part4) = 2 else false
                in
                    isValidFormat and allDigits
        in
            result,
    
    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckEmail_scalar'
    // Скалярная функция. Проверяет email-адрес.
    //
    // аргументы:
    // value as any                          - проверяемое значение
    // optional _verificationMode as logical - режим валидации (true/null = жесткий, false = мягкий)
    //
    // возвращаемое значение:
    // logical                               - true (валидно) или false (невалидно)
    // ---------------------------------------------------------------------------------------------
    fvCheckEmail_scalar = (value as any, optional _verificationMode as logical) as logical =>
        let
            isHardMode = if _verificationMode = false then false else true,
            trimmed = if value = null then "" else Text.Trim(Text.From(value)),
            isEmpty = trimmed = "",
            result = if isEmpty then not isHardMode else
                let
                    atCount = List.Count(Text.PositionOf(trimmed, "@", Occurrence.All)), atPos = if atCount = 1 then Text.PositionOf(trimmed, "@") else -1, isValidStructure = atCount = 1 and atPos > 0 and atPos < Text.Length(trimmed) - 1 and not Text.Contains(trimmed, " "), localPart = if isValidStructure then Text.Start(trimmed, atPos) else "", domainPart = if isValidStructure then Text.Range(trimmed, atPos + 1) else "", isValidLocalPart = isValidStructure and Text.Length(localPart) > 0 and Text.Length(localPart) <= 64 and not Text.StartsWith(localPart, ".") and not Text.EndsWith(localPart, ".") and not Text.Contains(localPart, ".."), isValidDomainPart = isValidStructure and Text.Length(domainPart) > 0 and Text.Length(domainPart) <= 255 and not Text.StartsWith(domainPart, "-") and not Text.EndsWith(domainPart, "-") and not Text.StartsWith(domainPart, ".") and not Text.EndsWith(domainPart, ".") and Text.Contains(domainPart, ".")
                in
                    isValidLocalPart and isValidDomainPart
        in
            result,
    
    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckBoolean_scalar'
    // Скалярная функция. Проверяет одно значение на соответствие булеву типу.
    //
    // аргументы:
    // value as any                          - проверяемое значение
    // optional _verificationMode as logical - режим валидации (true/null = жесткий, false = мягкий)
    //
    // возвращаемое значение:
    // logical                               - true (валидно) или false (невалидно)
    // ---------------------------------------------------------------------------------------------
    fvCheckBoolean_scalar = (value as any, optional _verificationMode as logical) as logical =>
        let
            isHardMode = if _verificationMode = false then false else true,
            isEmpty = value = null,
            result = if isEmpty then not isHardMode else
                if Value.Is(value, type logical) then true
                else if Value.Is(value, type number) then value = 0 or value = 1
                else if Value.Is(value, type text) then let lowerText = Text.Lower(Text.Trim(Text.From(value))) in lowerText = "true" or lowerText = "false" or lowerText = "1" or lowerText = "0"
                else false
        in
            result,
        
    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckPassportSeries_scalar'
    // Скалярная функция. Проверяет серию паспорта РФ (4 цифры).
    //
    // аргументы:
    // value as any                          - проверяемое значение
    // optional _verificationMode as logical - режим валидации (true/null = жесткий, false = мягкий)
    //
    // возвращаемое значение:
    // logical                               - true (валидно) или false (невалидно)
    // ---------------------------------------------------------------------------------------------
    fvCheckPassportSeries_scalar = (value as any, optional _verificationMode as logical) as logical =>
        let
            isHardMode = if _verificationMode = false then false else true,
            strValue = if value = null then "" else Text.Trim(Text.From(value)),
            isEmpty = strValue = "",
            result = if isEmpty then not isHardMode else
                let
                    isAllDigits = Text.Length(Text.Select(strValue, {"0".."9"})) = Text.Length(strValue),
                    hasCorrectLength = Text.Length(strValue) = 4
                in
                    isAllDigits and hasCorrectLength
        in
            result,
        
    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckPassportNumber_scalar'
    // Скалярная функция. Проверяет номер паспорта РФ (6 цифр).
    //
    // аргументы:
    // value as any                          - проверяемое значение
    // optional _verificationMode as logical - режим валидации (true/null = жесткий, false = мягкий)
    //
    // возвращаемое значение:
    // logical                               - true (валидно) или false (невалидно)
    // ---------------------------------------------------------------------------------------------
    fvCheckPassportNumber_scalar = (value as any, optional _verificationMode as logical) as logical =>
        let
            isHardMode = if _verificationMode = false then false else true,
            strValue = if value = null then "" else Text.Trim(Text.From(value)),
            isEmpty = strValue = "",
            result = if isEmpty then not isHardMode else
                let
                    isAllDigits = Text.Length(Text.Select(strValue, {"0".."9"})) = Text.Length(strValue),
                    hasCorrectLength = Text.Length(strValue) = 6
                in
                    isAllDigits and hasCorrectLength
        in
            result,

    // ---------------------------------------------------------------------------------------------
    // Функция валидации 'fvCheckPassportUnitCode_scalar'
    // Скалярная функция. Проверяет код подразделения паспорта РФ (XXX-XXX).
    //
    // аргументы:
    // value as any                          - проверяемое значение
    // optional _verificationMode as logical - режим валидации (true/null = жесткий, false = мягкий)
    //
    // возвращаемое значение:
    // logical                               - true (валидно) или false (невалидно)
    // ---------------------------------------------------------------------------------------------
    fvCheckPassportUnitCode_scalar = (value as any, optional _verificationMode as logical) as logical =>
        let
            isHardMode = if _verificationMode = false then false else true,
            strValue = if value = null then "" else Text.Trim(Text.From(value)),
            isEmpty = strValue = "",
            result = if isEmpty then not isHardMode else
                if Text.Length(strValue) <> 7 or Text.Middle(strValue, 3, 1) <> "-" then false
                else
                    let
                        numberPart = Text.Remove(strValue, {"-"}),
                        isAllDigits = Text.Length(Text.Select(numberPart, {"0".."9"})) = 6
                    in
                        isAllDigits
        in
            result,

// =============================================================================================
// ОБЪЕДИНЕНИЕ ВСЕХ ФУНКЦИЙ В RECORD
// =============================================================================================
FunctionRecord = [
    fvCheckNonEmpty_scalar         = fvCheckNonEmpty_scalar,
    fvCheckInteger_scalar          = fvCheckInteger_scalar,
    fvCheckFloat_scalar            = fvCheckFloat_scalar,
    fvCheckDate_scalar             = fvCheckDate_scalar,
    fvCheckINN_LLC_scalar          = fvCheckINN_LLC_scalar,
    fvCheckINN_Personal_scalar     = fvCheckINN_Personal_scalar,
    fvCheckSNILS_scalar            = fvCheckSNILS_scalar,
    fvCheckKPP_scalar              = fvCheckKPP_scalar,
    fvCheckTextLength_scalar       = fvCheckTextLength_scalar,
    fvCheckIntegerRange_scalar     = fvCheckIntegerRange_scalar,
    fvCheckFloatRange_scalar       = fvCheckFloatRange_scalar,
    fvCheckDateRange_scalar        = fvCheckDateRange_scalar,
    fvCheckPhone_scalar            = fvCheckPhone_scalar,
    fvCheckEmail_scalar            = fvCheckEmail_scalar,
    fvCheckBoolean_scalar          = fvCheckBoolean_scalar,
    fvCheckPassportSeries_scalar   = fvCheckPassportSeries_scalar,
    fvCheckPassportNumber_scalar   = fvCheckPassportNumber_scalar,
    fvCheckPassportUnitCode_scalar = fvCheckPassportUnitCode_scalar
]

in
    FunctionRecord