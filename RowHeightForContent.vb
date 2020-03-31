'---------------------------------------------------------------------------------------
' Procedure : RowHeightForContent
' Author    : The_Prist(Щербаков Дмитрий)
'             http://www.excel-vba.ru
' Purpose   : Функция подбирает высоту строки/ширину столбца объединенных ячеек по содержимому
'---------------------------------------------------------------------------------------
Function RowColHeightForContent(rc As Range, Optional bRowHeight As Boolean = True)
'rc -         ячейка, высоту строки или ширину столбца которой необходимо подобрать
'bRowHeight - True - если необходимо подобрать высоту строки
'             False - если необходимо подобрать ширину столбца
    Dim OldR_Height As Single, OldC_Widht As Single
    Dim MergedR_Height As Single, MergedC_Widht As Single
    Dim CurrCell As Range
    Dim ih As Integer
    Dim iw As Integer
    Dim NewR_Height As Single, NewC_Widht As Single
    Dim ActiveCellHeight As Single
 
    If rc.MergeCells Then
        With rc.MergeArea 'если ячейка объединена
            'запоминаем кол-во столбцов
            iw = .Columns(.Columns.Count).Column - rc.Column + 1
            'запоминаем кол-во строк.
            ih = .Rows(.Rows.Count).Row - rc.Row + 1
            'Определяем высоту и ширину объединения ячеек
            MergedR_Height = 0
            For Each CurrCell In .Rows
                MergedR_Height = CurrCell.RowHeight + MergedR_Height
            Next
            MergedC_Widht = 0
            For Each CurrCell In .Columns
                MergedC_Widht = CurrCell.ColumnWidth + MergedC_Widht
            Next
            'запоминаем высоту и ширину первой ячейки из объединенных
            OldR_Height = .Cells(1, 1).RowHeight
            OldC_Widht = .Cells(1, 1).ColumnWidth
            'отменяем объединение ячеек
            .MergeCells = False
            'назначаем новую высоту и ширину для первой ячейки
            .Cells(1).RowHeight = MergedR_Height
            .Cells(1, 1).EntireColumn.ColumnWidth = MergedC_Widht
            'если необходимо изменить высоту строк
            If bRowHeight Then
                '.WrapText = True 'раскомментировать, если необходимо принудительно выставлять перенос текста
                .EntireRow.AutoFit
                NewR_Height = .Cells(1).RowHeight    'запоминаем высоту строки
                .MergeCells = True
                If OldR_Height < (NewR_Height / ih) Then
                    .RowHeight = NewR_Height / ih
                Else
                    .RowHeight = OldR_Height
                End If
                'возвращаем ширину столбца первой ячейки
                .Cells(1, 1).EntireColumn.ColumnWidth = OldC_Widht
            Else 'если необходимо изменить ширину столбца
                .EntireColumn.AutoFit
                NewC_Widht = .Cells(1).EntireColumn.ColumnWidth    'запоминаем ширину столбца
                .MergeCells = True
                If OldC_Widht < (NewC_Widht / iw) Then
                    .ColumnWidth = NewC_Widht / iw
                Else
                    .ColumnWidth = OldC_Widht
                End If
                'возвращаем высоту строки первой ячейки
                .Cells(1, 1).RowHeight = OldR_Height
            End If
        End With
    End If
End Function