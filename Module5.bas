Attribute VB_Name = "Module5"
Sub CreateLiquidRatioChart()

    Dim ws As Worksheet
    Dim chtObj As ChartObject
    Dim chartRange As Range
    Dim LastUsedCol As Long

    Set ws = ThisWorkbook.Worksheets(" Liquidity Ratios Over Time")

    LastUsedCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Set chartRange = ws.Range(ws.Cells(1, 1), ws.Cells(5, LastUsedCol))

    For Each chtObj In ws.ChartObjects
        chtObj.Delete
    Next chtObj

    Set chtObj = ws.ChartObjects.Add(Left:=500, Top:=50, Width:=400, Height:=250)
    With chtObj.Chart
        .SetSourceData Source:=chartRange
        .ChartType = xlColumnClustered

        .HasTitle = True
        .ChartTitle.Text = "5-Year Liquidity Ratio Analysis"

        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Year"

        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Ratio"

        .Legend.Position = xlLegendPositionBottom
    End With
End Sub


