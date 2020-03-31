## How to use

1. Copy&paste function in your excel file macros
2. Call function with cell ranges, which need to be resized

# Example:

```vb

    Private Sub Workbook_Open()
     RowColHeightForContent ActiveWorkbook.Sheets(1).Range("A1")
     RowColHeightForContent ActiveWorkbook.Sheets(1).Range("D2")
     RowColHeightForContent ActiveWorkbook.Sheets(1).Range("G2")
    End Sub

```

< -- this script will resize cells A1, D2, G2 on first sheet, when user opens the file.