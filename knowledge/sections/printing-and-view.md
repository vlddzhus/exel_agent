# ПЕЧАТЬ, ВИД ЛИСТА И ГРУППИРОВКА — Template Reference

Рецепты для отчётов, которые нужно печатать или показывать в чистом виде.
Источник истины для `manageSheetView`, `managePageSetup`, `manageGrouping`.

---

## Вид листа (manageSheetView)

### Чистый отчёт (без сетки)
```
manageSheetView: showGridlines=false, showHeadings=false
```
Применяйте для дашбордов, финальных отчётов, презентаций — когда таблица
оформлена вручную (через `applyAutoDesign`) и сетка только мешает.

### Скрыть нули
```
manageSheetView: showZeros=false
```
Пустые ячейки вместо `0` — чище вид, особенно в финансовых моделях.

### Масштаб
```
manageSheetView: zoom=80   // уменьшить (чтобы поместилось больше)
manageSheetView: zoom=120  // увеличить (для презентации)
```
Диапазон 10-400%.

---

## Настройка печати (managePageSetup)

### Стандартный отчёт A4 альбомный
```
managePageSetup:
  orientation: "landscape"
  paperSize: "a4"
  margins: { top: 36, bottom: 36, left: 36, right: 36 }   // ~0.5 inch
  printTitleRows: "1:1"    // заголовок на каждой странице
  centerHorizontally: true
```

### Вписать в одну страницу по ширине
```
managePageSetup:
  fitToWidth: 1      // 1 страница по ширине (высота авто)
  fitToHeight: 0     // не ограничивать по высоте
```
Типично для широких таблиц. `fitToPage` включается автоматически.

### Вписать в одну страницу целиком
```
managePageSetup:
  fitToWidth: 1
  fitToHeight: 1
```

### Область печати
```
managePageSetup: printArea="A1:F50"
```
Печатать только указанный диапазон, игнорируя остальное на листе.

### Размеры бумаги (paperSize)
| Имя | Office.js | Размер |
|---|---|---|
| `a4` | 9 | 210×297 мм (стандарт Europe) |
| `a3` | 8 | 297×420 мм |
| `letter` | 1 | 8.5×11 inch (US) |
| `legal` | 5 | 8.5×14 inch |

### Поля (margins) — в пунктах
1 inch = 72 pt. Стандартные поля Excel ~0.75 inch = 54 pt.
Узкие поля для плотных отчётов: ~0.25 inch = 18 pt.

Доступные ключи: `top`, `bottom`, `left`, `right`, `header`, `footer`.

---

## Группировка и структура (manageGrouping)

### Сворачиваемые секции в финмодели
```
manageGrouping:
  action: "groupRows"
  address: "A5:A20"          // детали (строки 5-20)
  summaryBelow: true         // итог в строке 21
```
Появятся кнопки +/- слева — пользователь сворачивает детали, оставляя итоги.

### Промежуточные итоги (паттерн)
1. Сгруппируйте строки детализации через `manageGrouping` action="groupRows".
2. В строке под группой добавьте формулу `=SUBTOTAL(9;A5:A20)` (или `=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(9;A5:A20)`).
3. `SUBTOTAL` автоматически учитывает только видимые строки — при сворачивании группы пересчитывается.

### Колонки-детали
```
manageGrouping:
  action: "groupColumns"
  address: "C1:F1"          // колонки C-F
  summaryRight: true        // итог в колонке G
```

### Очистить всю структуру
```
manageGrouping: action="clearOutline"
```

---

## Workflow: «Подготовь к печати»

1. `applyAutoDesign` — оформить таблицу.
2. `manageSheetView` showGridlines=false — чистый вид.
3. `managePageSetup` orientation="landscape", fitToWidth=1, printTitleRows="1:1".
4. (Опц.) `manageGrouping` — спрятать детальные строки под итогами.
5. `manageNamedRanges` — назвать ключевые диапазоны (для формул и навигации).

---

## Anti-patterns

❌ **Не печатайте с сеткой.** Для оформленных отчётов — showGridlines=false.

❌ **Не задавайте fitToWidth=1 и fitToHeight=1 одновременно** для длинных таблиц — текст станет нечитаемым. Лучше fitToHeight=0.

❌ **Не группируйте строку с итогом.** Итог должен быть вне группы, иначе он скроется вместе с деталями.

❌ **Поля в пунктах, не в сантиметрах.** 1 cm ≈ 28.3 pt. Указывайте margins в pt.
