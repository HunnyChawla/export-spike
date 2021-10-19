
Sub worksheet_activate()
    Dim i As Integer
    Dim j As Integer
    Dim asc As Integer

    i = 0
    j = 0
    asc = 66
    For ascii = 67 To 81
        ActiveSheet.Shapes.AddChart2(217, xlBarClustered).Select
        With ActiveChart
        .SetSourceData Source:=Range("=Sheet1!$B$5:$B$8,Sheet1!$" & Chr(ascii) & "$5:$" & Chr(ascii) & "$8")
        '.ApplyLayout 9, xlBarClustered
        If asc + (5 * i) > 88 Then
            asc = 66
            i = 0
            j = j + 12
        End If
        .Parent.Top = Range(Chr(asc + (5 * i)) & (12 + j)).Top
        .Parent.Left = Range(Chr(asc + (5 * i)) & (12 + j)).Left
        .Parent.Height = Range(Chr(asc + (5 * i)) & (12 + j) & ":" & Chr(asc + (5 * i)) & (21 + j)).Height
        .Parent.Width = Range(Chr(asc + (5 * i)) & (12) & ":" & Chr(asc + 4 + (5 * i)) & (21)).Width
        i = i + 1
        asc = asc + 1
        End With


    Next
    
    For Each objCO In ActiveSheet.ChartObjects
        ' Make each one visible
        objCO.Visible = True
        ' If the chart is empty make it not visible
        If IsChartEmpty(objCO.Chart) Then objCO.Visible = False
        
    Next objCO
End Sub

Private Function IsChartEmpty(chtAnalyse As Chart) As Boolean

Dim i As Integer
Dim j As Integer
Dim objSeries As Series
    ' Loop through all series of data within the chart
    For i = 1 To chtAnalyse.SeriesCollection.Count
        Set objSeries = chtAnalyse.SeriesCollection(i)

        ' Loop through each value of the series
        For j = 1 To UBound(objSeries.Values)

            ' If we have a non-zero value then the chart is not deemed to be empty
            If objSeries.Values(j) <> 0 Then
                ' Set return value and quit function
                IsChartEmpty = False
                Exit Function
            End If

        Next j

    Next i

    IsChartEmpty = True

End Function
