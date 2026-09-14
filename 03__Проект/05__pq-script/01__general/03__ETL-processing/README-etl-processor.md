# РУКОВОДСТВО ПО ЭКСПЛУАТАЦИИ: ETL-ПРОЦЕССОР «PROCESSING TABLE»

---

## 1. ВВЕДЕНИЕ И ПАСПОРТ СИСТЕМЫ

### 1.1. Назначение и бизнес-роль в ГК «Передовые решения»

**ETL-процессор `ProcessingTable`** — это центральное программное ядро экосистемы **HorizonBI**, реализованное на функциональном языке Power Query (M). Процессор предназначен для автоматизации полного жизненного цикла обработки табличных массивов данных (Master Data) в среде Microsoft Excel: от извлечения сырых данных из разнородных файлов до сложных трансформаций, слияний и сквозной верификации качества данных.

В операционной модели ГК «Передовые решения» процессор выполняет роль **шлюза изоляции и трансформации данных**:

- Обеспечивает создание производных спецификаций (`to_customer_tech`, `to_customer_fin`, `to_supplier`, `purchase`, `shipment`) из единого источника правды — мастер-спецификации `work_spec` (`TabSpecificationWork_1` + `TabSpecificationWork_2`).
- Гарантирует абсолютную защиту коммерческой тайны (автоматическое и безошибочное удаление закупочных цен, внутренних наценок и данных о реальных поставщиках при формировании клиентских отгрузочных документов).
- Устраняет человеческий фактор при рутинных операциях нормализации номенклатуры, типизации, переименования столбцов и фильтрации.

### 1.2. Паспорт компонента

| Параметр                     | Значение                                                                                          |
| :--------------------------- | :------------------------------------------------------------------------------------------------ |
| **Идентификатор компонента** | `ProcessingTable`                                                                                 |
| **Текущая версия ядра**      | `rev.04 v01`                                                                                      |
| **Среда выполнения**         | Power Query Engine (Excel 2016+, Excel M365, Power BI Desktop)                                    |
| **Язык разработки**          | Power Query (M)                                                                                   |
| **Язык управления (DSL)**    | `HorizonBI` (на базе JSON-конфигураторов)                                                         |
| **Тип вызова**               | Параметризованная M-функция                                                                       |
| **Входные контракты**        | `_configJSONIndex as text`, до 4-х необязательных путей к файлам `_pathExcelBook_1..4 as text`    |
| **Выходной контракт**        | Результирующий объект `table` с примененными трансформациями и опциональным столбцом `VALIDATION` |
| **Зависимости**              | Внешние библиотеки `UtilsFunctionVector`, `UtilsFunctionScalar`, таблица `TabConfigJSON`          |

### 1.3. Архитектурная схема взаимодействия компонентов

```mermaid
flowchart TD
    classDef storage fill:#E3F2FD,stroke:#1565C0,stroke-width:2px,color:#0D47A1;
    classDef engine fill:#FFF3E0,stroke:#E65100,stroke-width:2px,color:#BF360C;
    classDef lib fill:#E8F5E9,stroke:#2E7D32,stroke-width:2px,color:#1B5E20;
    classDef output fill:#F3E5F5,stroke:#7B1FA2,stroke-width:2px,color:#4A148C;

    subgraph ConfigStorage ["Конфигурационный слой (Текущая книга Excel)"]
        TabConfig["Таблица TabConfigJSON<br/>(Колонки: IDconfig, CONFIG_JSON)"]:::storage
    end

    subgraph ProcessorCore ["Ядро ETL: ProcessingTable (rev.04 v01)"]
        ArgResolver["1. Парсер аргументов и резолвер путей"]:::engine
        ConfigValidator["2. Ранняя валидация структуры JSON"]:::engine
        TableLoader["3. Экстрактор таблиц 1..4 (Current vs External)"]:::engine
        MacroEngine["4. Конвейер МакроБлоков (Секции '1', '2'...)<br/>13 этапов трансформации"]:::engine
        JoinEngine["5. Движок многотабличного слияния<br/>(Left/Right/Combine + Expand)"]:::engine
        ValidationEngine["6. Модуль верификации качества данных"]:::engine
    end

    subgraph DataSources ["Источники данных (До 4 таблиц Excel)"]
        Src1["Источник 1: Local / External XLSX"]:::storage
        Src2["Источник 2: Local / External XLSX"]:::storage
        Src3["Источник 3: Local / External XLSX"]:::storage
        Src4["Источник 4: Local / External XLSX"]:::storage
    end

    subgraph Libraries ["Библиотеки M-функций"]
        ScalarLib["UtilsFunctionScalar<br/>(18 атомарных проверок)"]:::lib
        VectorLib["UtilsFunctionVector<br/>(fvCheckID, трансформации)"]:::lib
    end

    subgraph TargetOutput ["Итоговый артефакт"]
        ResultTable["Результирующий датасет Excel<br/>(+ столбец VALIDATION)"]:::output
    end

    TabConfig -->|JSON по IDconfig| ArgResolver
    ArgResolver --> ConfigValidator
    ConfigValidator --> TableLoader
    Src1 & Src2 & Src3 & Src4 --> TableLoader
    TableLoader --> MacroEngine
    MacroEngine --> JoinEngine
    JoinEngine --> ValidationEngine
    VectorLib -.->|Векторная предвалидация fvCheckID| ValidationEngine
    ScalarLib -.->|Построчная скалярная проверка| ValidationEngine
    ValidationEngine --> ResultTable
```

---

## 2. ИНСТРУКЦИЯ ДЛЯ ОПЕРАТОРА (БЫСТРЫЙ СТАРТ)

### 2.1. Подготовка конфигурации в таблице `TabConfigJSON`

Перед запуском обработки необходимо убедиться, что в текущей книге Excel присутствует смарт-таблица с именем **`TabConfigJSON`**. Таблица должна содержать обязательные столбцы:

1.  **`IDconfig`** (Текстовый) — уникальный код конфигурации (например, `config_40_to_customer_fin_OSNO` или `join_work1_work2`).
2.  **`CONFIG_JSON`** (Текстовый) — валидный JSON-документ на языке `HorizonBI`.
3.  **`COMMENT`** (Текстовый, опционально) — служебное описание назначения сценария.

```
+----+------------------------------------+---------------------------------------------+------------------------------------+
| N  | IDconfig                           | CONFIG_JSON                                 | COMMENT                            |
+----+------------------------------------+---------------------------------------------+------------------------------------+
| 1  | 40_to_customer_fin_HORIZON         | {"TabSpecificationWork": {"1": {...}}}     | Генерация ТКП для ОСНО (с НДС 22%) |
+----+------------------------------------+---------------------------------------------+------------------------------------+
```

### 2.2. Запуск процессора

#### Вариант А: Вызов через пользовательский интерфейс Power Query

1. В редакторе Power Query выберите функцию **`ProcessingTable`**.
2. В появившейся форме ввода параметров заполните:
   - **`_configJSONIndex`**: введите точный идентификатор конфигурации (например, `40_to_customer_fin_HORIZON`).
   - **`_pathExcelBook_1`** (опционально): полный абсолютный путь к файлу-источнику для первой таблицы (например, `C:\OneDrive\Project_1\20__work_spec\spec.xlsx`). Оставьте пустым, если таблица находится в текущем файле.
   - **`_pathExcelBook_2..4`** (опционально): пути к файлам для остальных таблиц многотабличной конфигурации.
3. Нажмите кнопку **«Вызвать» (Invoke)**.

#### Вариант Б: Вызов через формулу языка M в новом запросе

Создайте пустой запрос и вставьте код вызова:

```m
let
    Source = ProcessingTable(
        "40_to_customer_fin_HORIZON",
        "C:\OneDrive\02__Спецификации\20__work_spec\22__Спецификация-Горизонт - work spec rev.22 v03 #1_2024.xlsx"
    )
in
    Source
```

### 2.3. Правила разрешения путей к файлам (Workbook Resolver)

```mermaid
flowchart TD
    classDef check fill:#FFF3E0,stroke:#E65100,stroke-width:2px,color:#BF360C;
    classDef success fill:#E8F5E9,stroke:#2E7D32,stroke-width:2px,color:#1B5E20;
    classDef err fill:#FFEBEE,stroke:#C62828,stroke-width:2px,color:#B71C1C;

    Start["Запуск ProcessingTable для Таблицы N"] --> CheckArg{"Аргумент _pathExcelBook_N<br/>передан и не пуст?"}:::check
    CheckArg -- НЕТ / null --> UseCurrent["Режим: Локальный<br/>Excel.CurrentWorkbook()<br/>Столбец доступа: [Content]"]:::success
    CheckArg -- ДА (Строка пути) --> CheckFile{"Файл существует на диске?"}:::check
    CheckFile -- ДА --> UseExternal["Режим: Внешний файл<br/>Excel.Workbook(File.Contents(path))<br/>Столбец доступа: [Data]"]:::success
    CheckFile -- НЕТ --> ThrowPathError["Критическая ошибка M:<br/>Файл не найден по указанному пути"]:::err
    UseCurrent --> LoadTable["Поиск целевой смарт-таблицы по имени"]
    UseExternal --> LoadTable
    LoadTable --> TableFound{"Таблица найдена?"}:::check
    TableFound -- ДА --> ReturnTable["Возврат объекта Table в конвейер"]:::success
    TableFound -- НЕТ --> ThrowTableError["Критическая ошибка M:<br/>Целевая таблица 'X' не найдена"]:::err
```

---

## 3. АРХИТЕКТУРА МАКРОБЛОКОВ И ПРАВИЛА ВАЛИДАЦИИ JSON

### 3.1. Концепция нумерованных секций (Управление состоянием)

Язык `HorizonBI` организует обработку данных через **МакроБлоки (секции)**. МакроБлок — это атомарный контейнер операций, гарантирующий последовательное изменение состояния датасета.

```json
{
  "Таблица_1": {
    "1": {
      /* Шаг 1: Первичная очистка и фильтрация */
    },
    "2": {
      /* Шаг 2: Модификация значений */
    },
    "3": {
      /* Шаг 3: Финальное переименование и типизация */
    }
  }
}
```

**Принцип работы конвейера состояний:**
$$\text{Table}_{\text{Initial}} \xrightarrow{\text{МакроБлок "1"}} \text{Table}_{\text{State 1}} \xrightarrow{\text{МакроБлок "2"}} \text{Table}_{\text{State 2}} \dots \xrightarrow{\text{МакроБлок "N"}} \text{Table}_{\text{Final}}$$

### 3.2. Системные правила и ограничения структуры JSON

Нарушение любого из следующих правил приводит к аварийной остановке процессора на этапе **ранней валидации**:

1. **Числовые ключи секций:** Имена МакроБлоков обязаны быть строковыми представлениями положительных целых чисел (`"1"`, `"2"`, `"10"`, `"77"`). Произвольные текстовые имена (`"Step1"`, `"Clean"`) **запрещены**.
2. **Сортировка по числовому значению:** Порядок выполнения секций определяется строго числовым возрастанием: секция `"2"` всегда выполняется раньше секции `"10"` (в отличие от стандартной алфавитной сортировки строк, где `"10"` идет раньше `"2"`).
3. **Уникальность команд в секции:** Внутри одного МакроБлока каждая BI-команда может быть вызвана **строго один раз**. Если требуется применить команду повторно (например, выполнить еще одну фильтрацию после добавления столбца), оператор обязан создать следующую по номеру секцию.
4. **Ограничение многотабличности:** Корневой JSON-объект может содержать от **1 до 4 целевых таблиц**.

---

## 4. СПРАВОЧНИК КОМАНД DSL «HORIZONBI» (14 ЭТАПОВ ТРАНСФОРМАЦИИ)

Внутри каждого МакроБлока процессор выполняет команды строго по системному конвейеру приоритетов:

```mermaid
flowchart TD
    classDef step fill:#E1F5FE,stroke:#0288D1,stroke-width:2px,color:#01579B;
    classDef final fill:#E8F5E9,stroke:#2E7D32,stroke-width:2px,color:#1B5E20;

    S1["1. ColumnSetType (Приведение типов)"]:::step --> S2["2. ColumnFilterFreeData (Фильтр по значению)"]:::step
    S2 --> S3["3. ColumnFilterUniqueData (Дедупликация)"]:::step
    S3 --> S4["4. RemoveRowOne (Удаление 1-й строки)"]:::step
    S4 --> S5["5. RemoveRowAll (Удаление всех совпавших строк)"]:::step
    S5 --> S6["6. InsertColumn (Вставка пустых колонок)"]:::step
    S6 --> S7["7. ColumnFillData (Заполнение константой)"]:::step
    S7 --> S8["8. ModifyData (Точечная замена ячеек)"]:::step
    S8 --> S9["9. RemoveXorColumn (Изоляция / Белый список)"]:::step
    S9 --> S10["10. RemoveColumn (Удаление списка колонок)"]:::step
    S10 --> S11["11. ClearColumn (Очистка значений в NULL)"]:::step
    S11 --> S12["12. RenameColumn (Переименование)"]:::step
    S12 --> S13["13. ReorderColumn (Упорядочивание)"]:::step
    S13 --> S14["14. ValidationData (Верификация качества данных)"]:::final
```

---

### Этап 1: `ColumnSetType` (Установка типов данных)

- **Назначение:** Принудительное назначение системных типов Power Query (`text`, `int`, `float`, `bool`) указанным столбцам.
- **JSON-сигнатура:**
  ```json
  "ColumnSetType": [
    { "ИмяСтолбца1": "int" },
    { "ИмяСтолбца2": "text" },
    { "ИмяСтолбца3": "float" },
    { "ИмяСтолбца4": "bool" }
  ]
  ```

---

### Этап 2: `ColumnFilterFreeData` (Фильтрация по значению)

- **Назначение:** Фильтрация строк таблицы по строгому равенству значения в указанном столбце (`WHERE column = 'value'`).
- **JSON-сигнатура:**
  ```json
  "ColumnFilterFreeData": [
    { "ИмяСтолбца": "ИскомоеЗначение" }
  ]
  ```

---

### Этап 3: `ColumnFilterUniqueData` (Фильтрация по уникальности)

- **Назначение:** Дедупликация. Удаляет все повторяющиеся строки, оставляя только **первое вхождение**.
- **JSON-сигнатура:**
  ```json
  "ColumnFilterUniqueData": [
    "ИмяСтолбца1",
    "ИмяСтолбца2"
  ]
  ```

---

### Этап 4: `RemoveRowOne` (Удаление первой найденной строки)

- **Назначение:** Поиск по ключевому полю и удаление **только одной (первой по порядку)** строки.
- **JSON-сигнатура:**
  ```json
  "RemoveRowOne": [
    {
      "rule_1": {
        "IndexSearchColumn": "HASH COLLECTION",
        "IndexSearchRow": "1/2024-N1-ГОРИЗОНТ-ver_1-set_1.1-true"
      }
    }
  ]
  ```

---

### Этап 5: `RemoveRowAll` (Удаление всех совпавших строк)

- **Назначение:** Удаление **всех** строк датасета, в которых значение в столбце совпадает с заданным.
- **JSON-сигнатура:**
  ```json
  "RemoveRowAll": [
    {
      "rule_1": {
        "IndexSearchColumn": "isSHIPPED",
        "IndexSearchRow": "false"
      }
    }
  ]
  ```

---

### Этап 6: `InsertColumn` (Вставка новых столбцов)

- **Назначение:** Добавление новых столбцов со значением по умолчанию `null`.
- **JSON-сигнатура:**
  ```json
  "InsertColumn": [
    "STATUS",
    "APPROVER_NAME"
  ]
  ```

---

### Этап 7: `ColumnFillData` (Массовое заполнение константой)

- **Назначение:** Замена всех значений в указанном столбце на фиксированное значение.
- **JSON-сигнатура:**
  ```json
  "ColumnFillData": [
    { "STATUS": "ВЫПОЛНЕНО" }
  ]
  ```

---

### Этап 8: `ModifyData` (Точечная модификация ячеек)

- **Назначение:** Адресное изменение значений в ячейках конкретной строки (поле `COMMENT` игнорируется).
- **JSON-сигнатура:**
  ```json
  "ModifyData": [
    {
      "data_1": {
        "IndexSearchColumn": "HASH COLLECTION",
        "IndexSearchRow": "1/2024-N1-ГОРИЗОНТ-ver_1-set_3.1-true",
        "NewData": {
          "QUANTITY TOTAL": 14,
          "COST $TOTAL": 15290.10,
          "COMMENT": "Служебный комментарий"
        }
      }
    }
  ]
  ```

---

### Этап 9: `RemoveXorColumn` (Изоляция структуры / Белый список)

- **Назначение:** Удаляет абсолютно все столбцы, **кроме перечисленных в списке**.
- **JSON-сигнатура:**
  ```json
  "RemoveXorColumn": [
    "NUM FULL SET",
    "NAME PRODUCT",
    "QUANTITY TOTAL",
    "PRICE VAT UNIT $TOTAL",
    "COST VAT $TOTAL"
  ]
  ```

---

### Этап 10: `RemoveColumn` (Удаление списка столбцов)

- **Назначение:** Удаление явно заданного перечня столбцов (черный список).
- **JSON-сигнатура:**
  ```json
  "RemoveColumn": [
    "SUPPLIER",
    "PRICE UNIT $PURCHASE",
    "COST PRICE $PURCHASE"
  ]
  ```

---

### Этап 11: `ClearColumn` (Очистка значений столбцов)

- **Назначение:** Сохраняет структуру столбца, но заменяет все ячейки на `null`.
- **JSON-сигнатура:**
  ```json
  "ClearColumn": [
    "COMMENT WORK",
    "TCP SOURCE"
  ]
  ```

---

### Этап 12: `RenameColumn` (Переименование столбцов)

- **Назначение:** Локализация и переименование заголовков.
- **JSON-сигнатура:**
  ```json
  "RenameColumn": [
    { "NAME PRODUCT": "НАИМЕНОВАНИЕ" },
    { "PRICE VAT UNIT $TOTAL": "ЦЕНА ЗА ЕД., С НДС" }
  ]
  ```

---

### Этап 13: `ReorderColumn` (Переупорядочивание столбцов)

- **Назначение:** Выстраивание последовательности столбцов слева направо.
- **JSON-сигнатура:**
  ```json
  "ReorderColumn": [
    "№ п/п",
    "НАИМЕНОВАНИЕ",
    "КОЛИЧЕСТВО",
    "ЦЕНА ЗА ЕД., С НДС",
    "СТОИМОСТЬ, С НДС"
  ]
  ```

---

### Этап 14: `ValidationData` (Комплексная валидация датасета)

- **Назначение:** Формирует сервисный столбец **`VALIDATION`** с кодами ошибок по каждой строке (`0` — валидно, `$N$` — номер столбца с ошибкой).
- **JSON-сигнатура:**
  ```json
  "ValidationData": [
    {
      "IDnum": {
        "fvCheckID": []
      }
    },
    {
      "НАИМЕНОВАНИЕ": {
        "fvCheckNonEmpty": [],
        "fvCheckTextLength_scalar": [3, 200]
      }
    }
  ]
  ```

---

## 5. МНОГОТАБЛИЧНЫЙ ДВИЖОК СЛИЯНИЯ (MULTI-TABLE JOIN ENGINE)

### 5.1. Архитектурные принципы

Процессор поддерживает работу с **1, 2, 3 или 4 таблицами** в рамках одной конфигурации. В многотабличном режиме процессор:

1. Загружает каждую таблицу по отдельности из локального или внешнего файла.
2. Прогоняет каждую таблицу через **ее собственный набор МакроБлоков** (Этапы 1..13).
3. Применяет **глобальную операцию слияния**, указанную в конфигурации первой таблицы.

```mermaid
flowchart TD
    classDef check fill:#FFF3E0,stroke:#E65100,stroke-width:2px,color:#BF360C;
    classDef join fill:#E1F5FE,stroke:#0288D1,stroke-width:2px,color:#01579B;
    classDef err fill:#FFEBEE,stroke:#C62828,stroke-width:2px,color:#B71C1C;
    classDef success fill:#E8F5E9,stroke:#2E7D32,stroke-width:2px,color:#1B5E20;

    StartJoin["Старт блока Join Engine"] --> CountCheck{"Количество таблиц в JSON?"}:::check
    CountCheck -- "1 Таблица" --> BypassJoin["Пропуск слияния -> Переход к финализации"]:::success
    CountCheck -- "2-4 Таблицы" --> FindJoin{"Поиск команды слияния в Секциях Таблицы 1"}:::check

    FindJoin -- "0 команд или >1 команды" --> ErrJoinCount["Ошибка: Требуется ровно 1 команда слияния"]:::err
    FindJoin -- "Найдена ровно 1 команда" --> CheckOthers{"Есть команды слияния/валидации в Таблицах 2..4?"}:::check

    CheckOthers -- ДА --> ErrForbidden["Ошибка: Слияние и валидация разрешены только в Таблице 1"]:::err
    CheckOthers -- НЕТ --> RouteJoin{"Тип команды слияния?"}:::check

    RouteJoin -- "RunJOINdata" --> CheckCols{"Структура столбцов во всех таблицах идентична?"}:::check
    CheckCols -- ДА --> DoCombine["Table.Combine(ProcessedTablesList)"]:::join
    CheckCols -- НЕТ --> ErrCols["Ошибка: Несовпадение структуры колонок"]:::err

    RouteJoin -- "RunLeftJOINstructure" --> CheckHashL{"Столбец 'HASH COLLECTION' есть во всех таблицах?"}:::check
    CheckHashL -- ДА --> DoLeftJoin["Последовательный Table.NestedJoin (LeftOuter)<br/>+ Table.ExpandTableColumn"]:::join
    CheckHashL -- НЕТ --> ErrHashL["Ошибка: Отсутствует ключ HASH COLLECTION"]:::err

    RouteJoin -- "RunRightJOINstructure" --> CheckHashR{"Столбец 'HASH COLLECTION' есть во всех таблицах?"}:::check
    CheckHashR -- ДА --> DoRightJoin["Последовательный Table.NestedJoin (RightOuter)<br/>+ Table.ExpandTableColumn"]:::join
    CheckHashR -- НЕТ --> ErrHashR["Ошибка: Отсутствует ключ HASH COLLECTION"]:::err

    DoCombine --> ValGateway{"В Таблице 1 объявлен блок ValidationData?"}:::check
    DoLeftJoin --> CheckNoVal{"В Таблице 1 объявлен блок ValidationData?"}:::check
    DoRightJoin --> CheckNoVal

    CheckNoVal -- ДА --> ErrValJoin["Ошибка: ValidationData запрещена при структурных слияниях"]:::err
    CheckNoVal -- НЕТ --> EndJoin["Возврат объединенного датасета"]:::success

    ValGateway -- ДА --> RunVal["Выполнение валидации объединенных строк"]:::join
    ValGateway -- НЕТ --> EndJoin
    RunVal --> EndJoin
```

### 5.2. Спецификация команд слияния

1.  **`RunLeftJOINstructure` (Левое структурное слияние):** Слияние структур по горизонтали. Таблица №1 — базовая («левая»), к ней последовательно присоединяются остальные по ключу `HASH COLLECTION` с автоматическим раскрытием колонок через `Table.ExpandTableColumn`.  
    _Синтаксис:_ `"RunLeftJOINstructure": []`
2.  **`RunRightJOINstructure` (Правое структурное слияние):** Базовой («правой») становится последняя таблица в JSON, к ней присоединяются предыдущие по ключу `HASH COLLECTION`.  
    _Синтаксис:_ `"RunRightJOINstructure": []`
3.  **`RunJOINdata` (Вертикальная склейка датасетов):** Добавление строк (`UNION ALL`) для таблиц идентичной структуры.  
    _Синтаксис:_ `"RunJOINdata": []`

---

## 6. СВЯЗЬ С ВНЕШНИМИ БИБЛИОТЕКАМИ И КАТАЛОГ ССЫЛОК

- [Головное руководство проекта (README.md)](../../../../../README.md)
- [Руководство по функциям валидации и трансформации (UtilsFunction)](../02__utils-function-processing/README-utils-function.md)
- [Руководство по системным сервисным функциям (GeneralFunction)](../01__general-function-processing/README-general-function.md)
- [Руководство по функциям справочников (HandbookFunction)](../04__handbook-function-processing/README-handbook-function.md)
- [Спецификация Master Data шаблона (README-work_spec)](../../../07__template-excel/20__work_spec/README-work_spec.md)

---

## 7. ПРАКТИКУМ: РЕАЛЬНЫЕ BI-СЦЕНАРИИ ИЗ ПРАКТИКИ ГК

### Сценарий 1: Генерация финансовой спецификации `to_customer_fin` для ООО «ГОРИЗОНТ» (ОСНО, 22% НДС)

```json
{
  "TabSpecificationWork": {
    "1": {
      "ColumnSetType": [
        { "NUM PROJECT": "text" },
        { "QUANTITY TOTAL": "int" },
        { "PRICE VAT UNIT $TOTAL": "float" },
        { "COST VAT $TOTAL": "float" }
      ],
      "ColumnFilterFreeData": [{ "NUM PROJECT": "1/2024-N1-ГОРИЗОНТ" }]
    },
    "2": {
      "RemoveXorColumn": [
        "NUM FULL SET",
        "NAME PRODUCT",
        "DESCRIPTION PRODUCT",
        "QUANTITY TOTAL",
        "UNITS",
        "PRICE VAT UNIT $TOTAL",
        "COST VAT $TOTAL",
        "DURATION SALE $TOTAL",
        "COMMENT CUSTOMER"
      ]
    },
    "3": {
      "RenameColumn": [
        { "NUM FULL SET": "№ п/п" },
        { "NAME PRODUCT": "Наименование оборудования и услуг" },
        { "DESCRIPTION PRODUCT": "Технические характеристики" },
        { "QUANTITY TOTAL": "Кол-во" },
        { "UNITS": "Ед. изм." },
        { "PRICE VAT UNIT $TOTAL": "Цена за ед., руб. (с НДС 22%)" },
        { "COST VAT $TOTAL": "Стоимость, руб. (с НДС 22%)" },
        { "DURATION SALE $TOTAL": "Срок поставки (дней)" },
        { "COMMENT CUSTOMER": "Примечание" }
      ],
      "ReorderColumn": [
        "№ п/п",
        "Наименование оборудования и услуг",
        "Технические характеристики",
        "Кол-во",
        "Ед. изм.",
        "Цена за ед., руб. (с НДС 22%)",
        "Стоимость, руб. (с НДС 22%)",
        "Срок поставки (дней)",
        "Примечание"
      ]
    }
  }
}
```

---

## 8. ДИАГНОСТИКА, ОШИБКИ И ИХ УСТРАНЕНИЕ (TROUBLESHOOTING)

| Текст сообщения об ошибке                                                             | Причина возникновения                                                               | Регламент устранения                                                                                             |
| :------------------------------------------------------------------------------------ | :---------------------------------------------------------------------------------- | :--------------------------------------------------------------------------------------------------------------- |
| `Таблица с конфигурациями 'TabConfigJSON' не найдена в ТЕКУЩЕЙ книге.`                | В активном файле Excel отсутствует смарт-таблица `TabConfigJSON`.                   | Создать лист `ConfigJSON`, создать смарт-таблицу с именем `TabConfigJSON` и столбцами `IDconfig`, `CONFIG_JSON`. |
| `В таблице 'TabConfigJSON' не найдена конфигурация с ID 'X'.`                         | Опечатка в параметре `_configJSONIndex` или отсутствует строка в конфигураторе.     | Проверить соответствие имени в вызове и столбце `IDconfig`. Проверить регистр букв.                              |
| `Целевая таблица с именем 'X', указанная в конфигураторе, не найдена.`                | В файле Excel нет листа/смарт-таблицы с именем, заданным в корне JSON.              | Проверить имя умной таблицы в исходном файле. Проверить правильность пути `_pathExcelBook`.                      |
| `Ошибка конфигурации: Имена МакроБлоков в JSON должны быть только числами.`           | Секция названа текстом (напр. `"Step1"` вместо `"1"`).                              | Переименовать ключи секций в строковые целые числа: `"1"`, `"2"`, `"3"`.                                         |
| `Ошибка конфигурации: В МакроБлоке "X" обнаружены повторяющиеся BI-команды.`          | Внутри одной секции дважды вызвана одна команда (напр. две `ColumnFilterFreeData`). | Разнести повторяющиеся команды по разным секциям (МакроБлок `"1"` и МакроБлок `"2"`).                            |
| `RemoveColumn: В таблице отсутствуют столбцы для удаления: X, Y`                      | Попытка удалить столбцы, которых нет в текущем состоянии таблицы.                   | Проверить, не были ли эти столбцы переименованы или удалены на более ранних шагах/секциях.                       |
| `ModifyData: Не удалось найти строку, где 'Col' = 'Val'`                              | Значение поиска не существует в указанном столбце.                                  | Проверить точность значения в `IndexSearchRow` (пробелы, регистр, опечатки в хэшах).                             |
| `Команда 'ValidationData' не может использоваться совместно с 'RunLeftJOINstructure'` | Нарушение правил совместимости многотабличного слияния.                             | Удалить блок `ValidationData` из многотабличной конфигурации структурного слияния.                               |

---

## 9. РЕГЛАМЕНТ ТЕХНИЧЕСКОГО ОБСЛУЖИВАНИЯ И ВЕРСИОНИРОВАНИЯ

1. Любое изменение функционала процессора сопровождается обязательным инкрементом версии в шапке кода и документации:
   - **Major (`rev.0X`):** Изменение архитектуры, сигнатуры параметров или структуры языка `HorizonBI`.
   - **Minor (`v0X`):** Исправление внутренних ошибок, оптимизация производительности M-кода.
2. Перед выпуском новой ревизии код процессора проходит обязательное тестирование на тестовом стенде `03__Проект/06__example-query/`.
