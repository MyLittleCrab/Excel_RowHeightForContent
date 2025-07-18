'---------------------------------------------------------------------------------------
' Procedure : RowHeightForContent
' Author    : The_Prist(Щербаков Дмитрий)
'             http://www.excel-vba.ru
'             Enhanced by: Носов Роман (Github: MyLittleCrab)
'             Доработано: Носов Роман (Github: MyLittleCrab)
' Purpose   : Function adjusts row height/column width of merged cells based on content
'           : Функция подбирает высоту строки/ширину столбца объединенных ячеек по содержимому
'---------------------------------------------------------------------------------------

'rc -         cell whose row height or column width needs to be adjusted
'             ячейка, высоту строки или ширину столбца которой необходимо подобрать
'bRowHeight - True - if row height adjustment is needed
'             True - если необходимо подобрать высоту строки
'             False - if column width adjustment is needed
'             False - если необходимо подобрать ширину столбца
Function RowColHeightForContent(rc As Range, Optional bRowHeight As Boolean = True)

    Dim OldR_Height As Single, OldC_Width As Single ' Original width/height of the first cell in range | Ширина/высота первой ячейки рейнджа
    Dim MergedR_Height As Single, MergedC_Width As Single ' Width/height of the entire range | Ширина/высота всего рейнджа
    Dim CellPaddings As Single ' Internal paddings (should be considered for large ranges) | Внутренние отступы (при большом рейндже нужно учитывать)
    Dim CurrCell As Range
    Dim ih As Integer
    Dim iw As Integer
    Dim NewR_Height As Single, NewC_Width As Single ' Calculated height and width based on content | Вычисленные высота и ширина по контенту
    Dim ActiveCellHeight As Single
 
    CellPaddings = 0.3

    If rc.MergeCells Then
        With rc.MergeArea 'if cell is merged | если ячейка объединена
            'remember the number of columns | запоминаем кол-во столбцов
            iw = .Columns(.Columns.Count).Column - rc.Column + 1
            'remember the number of rows | запоминаем кол-во строк.
            ih = .Rows(.Rows.Count).Row - rc.Row + 1
            'Determine height and width of merged cells | Определяем высоту и ширину объединения ячеек
            MergedR_Height = 0
            For Each CurrCell In .Rows
                MergedR_Height = CurrCell.RowHeight + MergedR_Height + CellPaddings
            Next
            MergedC_Width = 0
            For Each CurrCell In .Columns
                MergedC_Width = CurrCell.ColumnWidth + MergedC_Width + CellPaddings
            Next
            'remember height and width of the first cell from merged ones | запоминаем высоту и ширину первой ячейки из объединенных
            OldR_Height = .Cells(1, 1).RowHeight
            OldC_Width = .Cells(1, 1).ColumnWidth
            'unmerge cells | отменяем объединение ячеек
            .MergeCells = False
            'assign new height and width for the first cell | назначаем новую высоту и ширину для первой ячейки
            .Cells(1).RowHeight = MergedR_Height
            .Cells(1, 1).EntireColumn.ColumnWidth = MergedC_Width
            'if row height needs to be changed | если необходимо изменить высоту строк
            If bRowHeight Then
                .EntireRow.AutoFit
                NewR_Height = .Cells(1).RowHeight    'remember row height | запоминаем высоту строки
                .MergeCells = True
                If OldR_Height < (NewR_Height / ih) Then
                    .RowHeight = NewR_Height / ih
                Else
                    .RowHeight = OldR_Height
                End If
                'restore column width of the first cell | возвращаем ширину столбца первой ячейки
                .Cells(1, 1).EntireColumn.ColumnWidth = OldC_Width
            Else 'if column width needs to be changed | если необходимо изменить ширину столбца
                .EntireColumn.AutoFit
                NewC_Width = .Cells(1).EntireColumn.ColumnWidth    'remember column width | запоминаем ширину столбца
                .MergeCells = True
                If OldC_Width < (NewC_Width / iw) Then
                    .ColumnWidth = NewC_Width / iw
                Else
                    .ColumnWidth = OldC_Width
                End If
                'restore row height of the first cell | возвращаем высоту строки первой ячейки
                .Cells(1, 1).RowHeight = OldR_Height
            End If
        End With
    End If
End Function