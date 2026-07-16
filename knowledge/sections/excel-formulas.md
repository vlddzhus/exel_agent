# Excel Formulas Reference (EN/RU)

## Cell References
- Relative: `A1` — changes when copied
- Absolute: `$A$1` — stays fixed when copied
- Mixed: `$A1` or `A$1` — one dimension fixed

## Operators
| Operator | Meaning | Example |
|----------|---------|---------|
| `+` | Addition | =A1+B1 |
| `-` | Subtraction | =A1-B1 |
| `*` | Multiplication | =A1*B1 |
| `/` | Division | =A1/B1 |
| `^` | Exponentiation | =A1^2 |
| `&` | Concatenation | =A1&" "&B1 |

## Math & Trig
- `=SUM(range)` / `=СУММ(диапазон)` — Sum
- `=AVERAGE(range)` / `=СРЗНАЧ(диапазон)` — Average
- `=COUNT(range)` / `=СЧЁТ(диапазон)` — Count numbers
- `=COUNTA(range)` / `=СЧЁТЗ(диапазон)` — Count non-empty
- `=MAX(range)` / `=МАКС(диапазон)` — Maximum
- `=MIN(range)` / `=МИН(диапазон)` — Minimum
- `=PRODUCT(range)` / `=ПРОИЗВЕД(диапазон)` — Multiply
- `=ROUND(num, places)` / `=ОКРУГЛ(число; разряд)` — Round
- `=ROUNDUP(num, places)` / `=ОКРУГЛВВЕРХ(число; разряд)` — Round up
- `=ROUNDDOWN(num, places)` / `=ОКРУГЛВНИЗ(число; разряд)` — Round down
- `=SUMIF(range, criteria, sum_range)` / `=СУММЕСЛИ(диапазон; критерий; сумма_диапазон)` — Conditional sum
- `=SUMPRODUCT(arr1, arr2)` / `=СУММПРОИЗВ(массив1; массив2)` — Sum product

## Logical
- `=IF(test, true_val, false_val)` / `=ЕСЛИ(тест; да; нет)` — Conditional
- `=IFERROR(val, default)` / `=ЕСЛИОШИБКА(знач; по_умолч)` — Handle errors
- `=IFNA(val, default)` / `=ЕСЛИНД(знач; по_умолч)` — Handle #N/A
- `=AND(cond1, cond2)` / `=И(усл1; усл2)` — All true
- `=OR(cond1, cond2)` / `=ИЛИ(усл1; усл2)` — Any true
- `=NOT(cond)` / `=НЕ(усл)` — Negate
- `=SWITCH(val, case1, res1, ...)` / `=ВЫБОР(знач; вариант1; рез1)` — Multi-condition

## Lookup
- `=VLOOKUP(val, table, col, exact)` / `=ВПР(знач; табл; номер; ложь)` — Vertical lookup
- `=HLOOKUP(val, table, row, exact)` / `=ГПР(знач; табл; номер; ложь)` — Horizontal lookup
- `=XLOOKUP(val, arr, ret)` / `=ПРОСМОТРХ(знач; массив; возврат)` — Modern lookup
- `=INDEX(range, row, col)` / `=ИНДЕКС(диапазон; строка; столбец)` — Value at position
- `=MATCH(val, range, type)` / `=ПОИСКПОЗ(знач; диапазон; тип)` — Find position
- `=CHOOSE(idx, val1, val2, ...)` / `=ВЫБОР(индекс; знач1; знач2)` — Select by index

## Date & Time
- `=TODAY()` / `=СЕГОДНЯ()` — Current date
- `=NOW()` / `=ТДАТА()` — Current date and time
- `=DATE(year, month, day)` / `=ДАТА(год; месяц; день)` — Create date
- `=DATEDIF(start, end, unit)` / `=РАЗНДАТ(нач; кон; единица)` — Date difference (unit: "d", "m", "y")
- `=DAY(date)` / `=ДЕНЬ(дата)` — Day of month
- `=MONTH(date)` / `=МЕСЯЦ(дата)` — Month number
- `=YEAR(date)` / `=ГОД(дата)` — Year
- `=WEEKDAY(date, type)` / `=ДЕНЬНЕД(дата; тип)` — Day of week (type 2 = Mon=1..Sun=7)
- `=EOMONTH(date, months)` / `=КОНМЕСЯЦА(дата; месяцы)` — End of month
- `=NETWORKDAYS(start, end)` / `=ЧИСТРАБДНИ(нач; кон)` — Work days count

## Text
- `=CONCATENATE(t1, t2)` / `=СЦЕПИТЬ(т1; т2)` — Join text
- `=TEXT(val, format)` / `=ТЕКСТ(знач; формат)` — Format as text
- `=LEFT(text, n)` / `=ЛЕВСИМВ(текст; n)` — First n characters
- `=RIGHT(text, n)` / `=ПРАВСИМВ(текст; n)` — Last n characters
- `=MID(text, start, n)` / `=ПСТР(текст; нач; n)` — Middle characters
- `=LEN(text)` / `=ДЛСТР(текст)` — Text length
- `=FIND(find, text)` / `=НАЙТИ(что; текст)` — Find position
- `=REPLACE(text, start, n, new)` / `=ЗАМЕНИТЬ(текст; нач; n; новый)` — Replace
- `=SUBSTITUTE(text, old, new)` / `=ПОДСТАВИТЬ(текст; старый; новый)` — Substitute all
- `=TRIM(text)` / `=СЖПРОБЕЛЫ(текст)` — Remove extra spaces
- `=UPPER(text)` / `=ПРОПИСН(текст)` — Uppercase
- `=LOWER(text)` / `=СТРОЧН(текст)` — Lowercase

## Statistical
- `=COUNTIF(range, criteria)` / `=СЧЁТЕСЛИ(диапазон; критерий)` — Count if
- `=COUNTIFS(range1, crit1, ...)` / `=СЧЁТЕСЛИМН(диап1; крит1)` — Count multiple
- `=AVERAGEIF(range, crit, avg_range)` / `=СРЗНАЧЕСЛИ(диап; крит; ср_диап)` — Average if
- `=MEDIAN(range)` / `=МЕДИАНА(диапазон)` — Median
- `=MODE(range)` / `=МОДА(диапазон)` — Mode
- `=STDEV(range)` / `=СТАНДОТКЛОН(диапазон)` — Standard deviation
- `=VAR(range)` / `=ДИСП(диапазон)` — Variance
- `=RANK(val, range, order)` / `=РАНГ(знач; диапазон; порядок)` — Rank
- `=LARGE(range, k)` / `=НАИБОЛЬШИЙ(диапазон; k)` — K-th largest
- `=SMALL(range, k)` / `=НАИМЕНЬШИЙ(диапазон; k)` — K-th smallest

## Formula Construction Rules

### IMPERATIVE: Cell references adjacent MUST have an operator
```
CORRECT:  =B8*B6
WRONG:    =B8B6        (treated as text, #NAME? error)
CORRECT:  =B8*B6+B9*B7
WRONG:    =B8B6+B9B7   (both pairs missing *)
CORRECT:  =(B8-B6)*B9/B7
```

### Semicolons as separators
Russian locale Excel uses semicolons (;) not commas (,) in function arguments.
```
=IF(A1>10; "High"; "Low")     — Russian
=IF(A1>10, "High", "Low")     — English
```

### Absolute references with $
```
=B1*$A$1   — A1 is absolute, B1 is relative
=$B1       — Column B is absolute, row is relative
=B$1       — Row 1 is absolute, column is relative
```
