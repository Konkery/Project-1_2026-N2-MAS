/*
<FuncSpecResultImage>
Данная функция выполняет теже действия, что и запрос QuerySpecResultImage (см описание), т.е. 
подтягивает в переданную ей спецификацию данные по изображению номеклатуры спецификации.
отличие в том, что это функция, а не запрос, и у данной фукнции есть два аргумента, что 
позволяет выбрать динамически подтянуть источник спецификации и источник ссылок на изображения.
*/
let
    Источник = (LeftTable as table, RightTable as table) =>
let
    Источник = Table.NestedJoin(LeftTable, {"АРТИКУЛ МАС"}, RightTable, {"ID"}, "QuertyDBimage", JoinKind.LeftOuter),
    #"Развернутый элемент QuerytDBimage" = Table.ExpandTableColumn(Источник, "QuertyDBimage", {"InternetPathFile"}, {"InternetPathFile"}),
    #"Добавлен пользовательский объект" = Table.AddColumn(#"Развернутый элемент QuerytDBimage", "ИЗОБРАЖЕНИЕ ТМЦ", each "=ИЗОБРАЖЕНИЕ(" & """" & [InternetPathFile]  & """" & ";""книга"";3;320)"),
    #"Удаленные столбцы1" = Table.RemoveColumns(#"Добавлен пользовательский объект",{"InternetPathFile"})
in
    #"Удаленные столбцы1"
in
    Источник