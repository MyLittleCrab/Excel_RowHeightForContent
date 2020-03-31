'---------------------------------------------------------------------------------------
' Procedure : RowHeightForContent
' Author    : The_Prist(Щербаков Дмитрий)
'             http://www.excel-vba.ru
'             Доработано: Носов Роман (Github: JustAddAcid)
' Purpose   : Функция подбирает высоту строки/ширину столбца объединенных ячеек по содержимому
'---------------------------------------------------------------------------------------

'rc -         ячейка, высоту строки или ширину столбца которой необходимо подобрать
'bRowHeight - True - если необходимо подобрать высоту строки
'             False - если необходимо подобрать ширину столбца
Function RowColHeightForContent(rc As Range, Optional bRowHeight As Boolean = True)

    Dim OldR_Height As Single, OldC_Width As Single ' Ширина/высота первой ячейки рейнджа
    Dim MergedR_Height As Single, MergedC_Width As Single ' Ширина/высота всего рейнджа
    Dim CellPaddings As Single ' Внутренние отступы (при большом рейндже нужно учитывать)
    Dim CurrCell As Range
    Dim ih As Integer
    Dim iw As Integer
    Dim NewR_Height As Single, NewC_Width As Single ' Вычисленные высота и ширина по контенту
    Dim ActiveCellHeight As Single
 
    CellPaddings = 0.3

    If rc.MergeCells Then
        With rc.MergeArea 'если ячейка объединена
            'запоминаем кол-во столбцов
            iw = .Columns(.Columns.Count).Column - rc.Column + 1
            'запоминаем кол-во строк.
            ih = .Rows(.Rows.Count).Row - rc.Row + 1
            'Определяем высоту и ширину объединения ячеек
            MergedR_Height = 0
            For Each CurrCell In .Rows
                MergedR_Height = CurrCell.RowHeight + MergedR_Height + CellPaddings
            Next
            MergedC_Width = 0
            For Each CurrCell In .Columns
                MergedC_Width = CurrCell.ColumnWidth + MergedC_Width + CellPaddings
            Next
            'запоминаем высоту и ширину первой ячейки из объединенных
            OldR_Height = .Cells(1, 1).RowHeight
            OldC_Width = .Cells(1, 1).ColumnWidth
            'отменяем объединение ячеек
            .MergeCells = False
            'назначаем новую высоту и ширину для первой ячейки
            .Cells(1).RowHeight = MergedR_Height
            .Cells(1, 1).EntireColumn.ColumnWidth = MergedC_Width
            'если необходимо изменить высоту строк
            If bRowHeight Then
                .EntireRow.AutoFit
                NewR_Height = .Cells(1).RowHeight    'запоминаем высоту строки
                .MergeCells = True
                If OldR_Height < (NewR_Height / ih) Then
                    .RowHeight = NewR_Height / ih
                Else
                    .RowHeight = OldR_Height
                End If
                'возвращаем ширину столбца первой ячейки
                .Cells(1, 1).EntireColumn.ColumnWidth = OldC_Width
            Else 'если необходимо изменить ширину столбца
                .EntireColumn.AutoFit
                NewC_Width = .Cells(1).EntireColumn.ColumnWidth    'запоминаем ширину столбца
                .MergeCells = True
                If OldC_Width < (NewC_Width / iw) Then
                    .ColumnWidth = NewC_Width / iw
                Else
                    .ColumnWidth = OldC_Width
                End If
                'возвращаем высоту строки первой ячейки
                .Cells(1, 1).RowHeight = OldR_Height
            End If
        End With
    End If
End Function