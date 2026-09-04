// =============================================================================================
//
// ЗАПРОС ФОРМИРОВАНИЯ СПРАВОЧНИКА 'HandbookProject' 2-го ТИПА 
//
// версия rev.01 v01
// =============================================================================================

// ---------------------------------------------------------------------------------------------
// Запрос извлекает необходимые для работы столбцы из справочника 1-го типа и возвращает 
// сокращенный датасет для использования в прикладных таблицах.
// Запрос работает со справочником хранящим список проектов, эти данные применяются в
// различных документов компании.
//
// источник:
//    TabProjectList    - справочник 1-го типа
//
// возвращаемое значение:
//    table             - обработанный датасет - справочник 2-го типа
//
// ---------------------------------------------------------------------------------------------
let
    // извлечь объекты из файла-справочника
    sourceWorkbook = Excel.Workbook(Web.Contents(sourceWorkbook_TabProjectList), null, true),
    // извлечь датасет из таблицы-справочника
    sourceTable = sourceWorkbook{[Item=sourceTable_TabProjectList,Kind="Table"]}[Data],
    
    //удалить избыточные столбцы из результирующего датасета
    removeColumnTable = Table.RemoveColumns(     sourceTable,
                                            {
                                                "IDprojectllc",
                                                "IDproject",
                                                "IDsubproject",
                                                "YEAR",
                                                "IDprojectlite",
                                                "IDprojectmidle",
                                                "RELATED PROJECT",
                                                "IDprojectext",
                                                "FRAME PROJECT",
                                                "PURPOSE PROJECT",
                                                "COMMENT",
                                                "BUYER",
                                                "END CUSTOMER",
                                                "EXECUTOR",
                                                "TYPE PROJECT",
                                                "TYPE WORK",
                                                "START DATE",
                                                "END DATE",
                                                "PROJECT DIRECTOR",
                                                "OFFICIAL  PROJECT STATUS",
                                                "UNOFFICIAL  PROJECT STATUS",
                                                "LINK TO FOLDER",
                                                "WHO CHANGED",
                                                "WHO CREATED"
                                            }),

    // откорректировать типы данных столбцов результирующего датасета
    result = Table.TransformColumnTypes(    removeColumnTable,
                                            {{"IDnum", Int64.Type},
                                            {"IDprojectfull", type text}})

in
    result