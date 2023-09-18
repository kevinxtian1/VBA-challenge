Attribute VB_Name = "Module1"
Sub Stock()
Dim WS_Count As Integer
Dim I As Integer
Dim Counter As Integer
Dim X As Integer
Dim Ticker As String
Dim Open1 As Double
Dim Close1 As Double
Dim Change As Double
Dim Det As Double
Dim TotalS As Double
Dim Counter2 As Integer
Dim Z As Integer
WS_Count = ActiveWorkbook.Worksheets.Count
For I = 1 To WS_Count
    ActiveWorkbook.Worksheets(I).Cells(1, 9) = "Ticker"
    ActiveWorkbook.Worksheets(I).Cells(1, 10) = "Yearly Change"
    ActiveWorkbook.Worksheets(I).Cells(1, 11) = "Percent Change"
    ActiveWorkbook.Worksheets(I).Cells(1, 12) = "Total Stock Volume"
    ActiveWorkbook.Worksheets(I).Cells(2, 15) = "Greatest % Increase"
    ActiveWorkbook.Worksheets(I).Cells(3, 15) = "Greatest % Decrease"
    ActiveWorkbook.Worksheets(I).Cells(4, 15) = "Greatest Total Volume"
    ActiveWorkbook.Worksheets(I).Cells(1, 16) = "Ticker"
    ActiveWorkbook.Worksheets(I).Cells(1, 17) = "Value"
    ActiveWorkbook.Worksheets(I).Cells(2, 17) = 0
    ActiveWorkbook.Worksheets(I).Cells(3, 17) = 10000
    ActiveWorkbook.Worksheets(I).Cells(4, 17) = 0
    
    

    RowCount = ActiveWorkbook.Worksheets(I).UsedRange.Rows.Count
    Counter = 2
    Open1 = ActiveWorkbook.Worksheets(I).Cells(2, 3).Value
    TotalS = 0
    Counter2 = 0
    For X = 2 To RowCount
        TotalS = TotalS + ActiveWorkbook.Worksheets(I).Cells(X, 7)

        If ActiveWorkbook.Worksheets(I).Cells(X + 1, 1).Value <> ActiveWorkbook.Worksheets(I).Cells(X, 1).Value Then
            Counter2 = Counter2 + 1
            ActiveWorkbook.Worksheets(I).Cells(Counter, 9) = ActiveWorkbook.Worksheets(I).Cells(X, 1).Value
            ActiveWorkbook.Worksheets(I).Cells(Counter + 1, 9) = ActiveWorkbook.Worksheets(I).Cells(X + 1, 1).Value
            Close1 = ActiveWorkbook.Worksheets(I).Cells(X, 6).Value
            Change = Close1 - Open1
            If Change < 0 Then
                ActiveWorkbook.Worksheets(I).Cells(Counter, 10).Interior.ColorIndex = 3
                Det = (Close1 - Open1) / Open1
                ActiveWorkbook.Worksheets(I).Cells(Counter, 11) = FormatPercent(Det)
                ActiveWorkbook.Worksheets(I).Cells(Counter, 11).Interior.ColorIndex = 3
            Else
                Det = (Close1 - Open1) / Open1
                ActiveWorkbook.Worksheets(I).Cells(Counter, 11) = FormatPercent(Det)
                ActiveWorkbook.Worksheets(I).Cells(Counter, 10).Interior.ColorIndex = 4
                ActiveWorkbook.Worksheets(I).Cells(Counter, 11).Interior.ColorIndex = 4
            End If
            ActiveWorkbook.Worksheets(I).Cells(Counter, 10) = Change
            Open1 = ActiveWorkbook.Worksheets(I).Cells(X + 1, 3).Value
            ActiveWorkbook.Worksheets(I).Cells(Counter, 12) = TotalS
            If TotalS > ActiveWorkbook.Worksheets(I).Cells(4, 17) Then
                ActiveWorkbook.Worksheets(I).Cells(4, 16) = ActiveWorkbook.Worksheets(I).Cells(X, 1)
                ActiveWorkbook.Worksheets(I).Cells(4, 17) = TotalS
            End If
            TotalS = 0
            Counter = Counter + 1
        End If
    Next X
    For Z = 2 To Counter2
        If ActiveWorkbook.Worksheets(I).Cells(Z, 11).Value > ActiveWorkbook.Worksheets(I).Cells(2, 17).Value Then
            ActiveWorkbook.Worksheets(I).Cells(2, 17) = FormatPercent(ActiveWorkbook.Worksheets(I).Cells(Z, 11).Value)
            ActiveWorkbook.Worksheets(I).Cells(2, 16) = ActiveWorkbook.Worksheets(I).Cells(Z, 9)
        End If
        If ActiveWorkbook.Worksheets(I).Cells(Z, 11).Value < ActiveWorkbook.Worksheets(I).Cells(3, 17).Value Then
            ActiveWorkbook.Worksheets(I).Cells(3, 17) = FormatPercent(ActiveWorkbook.Worksheets(I).Cells(Z, 11).Value)
            ActiveWorkbook.Worksheets(I).Cells(3, 16) = ActiveWorkbook.Worksheets(I).Cells(Z, 9)
        End If
    Next Z
Next I
End Sub

