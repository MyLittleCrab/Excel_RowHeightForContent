# Excel VBA: Auto-Resize Merged Cells

## English

### Description

This VBA function automatically adjusts the height and width of merged cells in Excel based on their content. It's particularly useful when working with merged cells that contain varying amounts of text and need to be properly sized for optimal display.

### Features

- ✅ Automatically adjusts row height for merged cells
- ✅ Automatically adjusts column width for merged cells  
- ✅ Handles both single cells and merged cell ranges
- ✅ Preserves original formatting and structure
- ✅ Easy integration into existing Excel workbooks
- ✅ Supports both manual calls and event-driven automation

### Function Signature

```vb
Function RowColHeightForContent(rc As Range, Optional bRowHeight As Boolean = True)
```

### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `rc` | Range | Yes | - | The cell or range of cells to resize |
| `bRowHeight` | Boolean | No | True | `True` = adjust row height, `False` = adjust column width |

### Installation

1. Open your Excel workbook
2. Press `Alt + F11` to open the VBA Editor
3. Insert a new module (`Insert` → `Module`)
4. Copy and paste the `RowColHeightForContent` function from `RowHeightForContent.vb`
5. Save your workbook as `.xlsm` (macro-enabled)

### Usage Examples

#### Basic Usage

```vb
' Resize row height for cell A1
Call RowColHeightForContent(Range("A1"))

' Resize column width for cell B2
Call RowColHeightForContent(Range("B2"), False)
```

#### Auto-resize on Workbook Open

```vb
Private Sub Workbook_Open()
    RowColHeightForContent ActiveWorkbook.Sheets(1).Range("A1")
    RowColHeightForContent ActiveWorkbook.Sheets(1).Range("D2")
    RowColHeightForContent ActiveWorkbook.Sheets(1).Range("G2")
End Sub
```

#### Auto-resize on Worksheet Change

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("A1:G10")) Is Nothing Then
        RowColHeightForContent Target, True  ' Adjust height
    End If
End Sub
```

### How It Works

1. **Detection**: Checks if the target cell is part of a merged range
2. **Analysis**: Calculates the total dimensions of the merged area
3. **Temporary Unmerge**: Temporarily unmerges cells to measure content
4. **AutoFit**: Uses Excel's AutoFit functionality on the unmerged content
5. **Restoration**: Remerges cells and applies the calculated optimal size

### Requirements

- Microsoft Excel with VBA support
- Macro-enabled workbook (`.xlsm` format)

### Authors & Credits

- **Original Author**: The_Prist (Щербаков Дмитрий) - [excel-vba.ru](http://www.excel-vba.ru)
- **Enhanced by**: Носов Роман (Github: [JustAddAcid](https://github.com/JustAddAcid))

### License

This project is licensed under CC0 1.0 Universal - see the [LICENSE](LICENSE) file for details.

---

## Русский

### Описание

Эта VBA-функция автоматически подстраивает высоту и ширину объединённых ячеек в Excel в соответствии с их содержимым. Особенно полезна при работе с объединёнными ячейками, содержащими различное количество текста и требующими правильного размера для оптимального отображения.

### Возможности

- ✅ Автоматическая подстройка высоты строк для объединённых ячеек
- ✅ Автоматическая подстройка ширины столбцов для объединённых ячеек
- ✅ Работа как с отдельными ячейками, так и с диапазонами объединённых ячеек
- ✅ Сохранение исходного форматирования и структуры
- ✅ Простая интеграция в существующие книги Excel
- ✅ Поддержка как ручных вызовов, так и автоматизации по событиям

### Сигнатура функции

```vb
Function RowColHeightForContent(rc As Range, Optional bRowHeight As Boolean = True)
```

### Параметры

| Параметр | Тип | Обязательный | По умолчанию | Описание |
|----------|-----|--------------|--------------|----------|
| `rc` | Range | Да | - | Ячейка или диапазон ячеек для изменения размера |
| `bRowHeight` | Boolean | Нет | True | `True` = подстроить высоту строки, `False` = подстроить ширину столбца |

### Установка

1. Откройте книгу Excel
2. Нажмите `Alt + F11` для открытия редактора VBA
3. Вставьте новый модуль (`Вставка` → `Модуль`)
4. Скопируйте и вставьте функцию `RowColHeightForContent` из файла `RowHeightForContent.vb`
5. Сохраните книгу как `.xlsm` (с поддержкой макросов)

### Примеры использования

#### Базовое использование

```vb
' Изменить высоту строки для ячейки A1
Call RowColHeightForContent(Range("A1"))

' Изменить ширину столбца для ячейки B2
Call RowColHeightForContent(Range("B2"), False)
```

#### Автоматическое изменение размера при открытии книги

```vb
Private Sub Workbook_Open()
    RowColHeightForContent ActiveWorkbook.Sheets(1).Range("A1")
    RowColHeightForContent ActiveWorkbook.Sheets(1).Range("D2")
    RowColHeightForContent ActiveWorkbook.Sheets(1).Range("G2")
End Sub
```

#### Автоматическое изменение размера при изменении листа

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("A1:G10")) Is Nothing Then
        RowColHeightForContent Target, True  ' Подстроить высоту
    End If
End Sub
```

### Принцип работы

1. **Обнаружение**: Проверяет, является ли целевая ячейка частью объединённого диапазона
2. **Анализ**: Вычисляет общие размеры объединённой области
3. **Временное разъединение**: Временно разъединяет ячейки для измерения содержимого
4. **Автоподбор**: Использует функцию AutoFit Excel для разъединённого содержимого
5. **Восстановление**: Повторно объединяет ячейки и применяет вычисленный оптимальный размер

### Требования

- Microsoft Excel с поддержкой VBA
- Книга с поддержкой макросов (формат `.xlsm`)

### Авторы и благодарности

- **Оригинальный автор**: The_Prist (Щербаков Дмитрий) - [excel-vba.ru](http://www.excel-vba.ru)
- **Доработал**: Носов Роман (Github: [JustAddAcid](https://github.com/JustAddAcid))

### Лицензия

Проект лицензирован под CC0 1.0 Universal - подробности в файле [LICENSE](LICENSE).