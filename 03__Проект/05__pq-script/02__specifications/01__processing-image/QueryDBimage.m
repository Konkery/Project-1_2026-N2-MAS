/*
<QueryDBimage>
Данный запрос возвращает датасет содержащий значение артикула номенклатуры и ссылки на изображение
номеклатуры, плюс некоторые дополнительные технические поля.
Данные беруться из БД номенклатора.
*/
let
    Источник = Excel.Workbook(File.Contents("C:\OneDrive\LLC MAS\Номенклатор\Вендоры и бренды\DB Nomenclature img rev.01 v01 #2023.xlsx"), null, true),
    TableDBimage_Table = Источник{[Item="TableDBimage",Kind="Table"]}[Data],
    #"Измененный тип" = Table.TransformColumnTypes(TableDBimage_Table,{{"№ п/п", Int64.Type}, {"ID", type text}, {"FileName", type text}, {"LocalPathFile", type text}, {"InternetPathFile", type text}}),
    #"Удаленные столбцы" = Table.RemoveColumns(#"Измененный тип",{"FileName", "LocalPathFile", "ИЗОБРАЖЕНИЕ ТМЦ"})
in
    #"Удаленные столбцы"