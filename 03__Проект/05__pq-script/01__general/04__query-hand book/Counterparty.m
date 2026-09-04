// =============================================================================================
//
// ЗАПРОС ФОРМИРОВАНИЯ СПРАВОЧНИКА 'HandbookCounterparty' 2-го ТИПА 
//
// версия rev.01 v01
// =============================================================================================

// ---------------------------------------------------------------------------------------------
// Запрос извлекает необходимые для работы столбцы из справочника 1-го типа и возвращает 
// сокращенный датасет для использования в прикладных таблицах.
// Запрос работает со справочником хранящим список контрагентов, эти данные применяются в
// различных документов компании.
//
// источник:
//    TabCounterpartyList   - справочник 1-го типа
//
// возвращаемое значение:
//    table                 - обработанный датасет - справочник 2-го типа
//
// ---------------------------------------------------------------------------------------------
let
    // извлечь объекты из файла-справочника
    sourceWorkbook = Excel.Workbook(Web.Contents(sourceWorkbook_TabCounterpartyList), null, true),
    // извлечь датасет из таблицы-справочника
    sourceTable = sourceWorkbook{[Item=sourceTable_TabCounterpartyList,Kind="Table"]}[Data],
    
    // задать типы данных столбцов которые будут возвращены в результирующем датасете
    setType = Table.TransformColumnTypes(    sourceTable,
                                            {
                                                {"IDnum", Int64.Type},
                                                {"IDcounterparty", Int64.Type},
                                                {"COUNTERPARTY", type text},
                                                {"FORM OWNERSHIP", type text}
                                            }),
    
    //удалить избыточные столбцы из результирующего датасета
    result = Table.RemoveColumns(    setType,
                                    {
                                        "CORPORATION",
                                        "COUNTERPARTY TYPE FIRST",
                                        "COUNTERPARTY TYPE SECOND",
                                        "EMPLOYEE",
                                        "INN",
                                        "KPP",
                                        "WORK MAIL",
                                        "PHONE WORK",
                                        "PHONE MOBILE",
                                        "ADDRESS",
                                        "WEBSITE",
                                        "COMPANY CARD",
                                        "COMMENT"
                                    })
in
    result