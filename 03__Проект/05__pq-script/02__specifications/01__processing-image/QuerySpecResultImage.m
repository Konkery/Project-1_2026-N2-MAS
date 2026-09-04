/*
<QuerySpecResultImage>
Данный объединяющий запрос генерирует датасет который является по сути разновидностью
спецификации (рабочая, для заказчика или иная) в которой "подтянуты" данные изображений
номенклатуры данной спецификации (привязка осуществляется по артикулу).
*/
let
    Источник = Table.NestedJoin(QuerySpecResultNoImage, {"АРТИКУЛ МАС"}, QueryDBimage, {"ID"}, "QuertyDBimage", JoinKind.LeftOuter),
    #"Развернутый элемент QuerytDBimage" = Table.ExpandTableColumn(Источник, "QuertyDBimage", {"InternetPathFile"}, {"InternetPathFile"}),
    #"Добавлен пользовательский объект" = Table.AddColumn(#"Развернутый элемент QuerytDBimage", "ИЗОБРАЖЕНИЕ ТМЦ", each "=ИЗОБРАЖЕНИЕ(" & """" & [InternetPathFile]  & """" & ";""книга"";3;320)"),
    #"Удаленные столбцы1" = Table.RemoveColumns(#"Добавлен пользовательский объект",{"InternetPathFile"})
in
    #"Удаленные столбцы1"