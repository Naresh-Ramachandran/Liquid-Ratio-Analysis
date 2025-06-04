Attribute VB_Name = "Module4"
Sub LiquidRatioCalculation()
    Dim LRCWorksheet As Worksheet
    Dim LastUsedCol As Long
    Dim ILoop As Long
    Dim CurrentRatio As Double
    Dim quickRatio As Double
    Dim CashRatio As Double
    Dim OpCashFlowRatio As Double

    Set LRCWorksheet = ThisWorkbook.Worksheets(" Liquidity Ratios Over Time")
    LastUsedCol = LRCWorksheet.Cells(1, LRCWorksheet.Columns.Count).End(xlToLeft).Column

    For ILoop = 2 To LastUsedCol
        CurrentRatio = ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(8, ILoop).Value / _
                       ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(14, ILoop).Value
        quickRatio = ( _
            ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(3, ILoop).Value + _
            ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(4, ILoop).Value + _
            ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(5, ILoop).Value _
        ) / ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(13, ILoop).Value

        CashRatio = ( _
            ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(3, ILoop).Value + _
            ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(4, ILoop).Value _
        ) / ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(13, ILoop).Value

        OpCashFlowRatio = ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(15, ILoop).Value / _
                          ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(14, ILoop).Value

        ThisWorkbook.Worksheets(" Liquidity Ratios Over Time").Cells(2, ILoop).Value = CurrentRatio
        ThisWorkbook.Worksheets(" Liquidity Ratios Over Time").Cells(3, ILoop).Value = quickRatio
        ThisWorkbook.Worksheets(" Liquidity Ratios Over Time").Cells(4, ILoop).Value = CashRatio
        ThisWorkbook.Worksheets(" Liquidity Ratios Over Time").Cells(5, ILoop).Value = OpCashFlowRatio
    Next ILoop
End Sub

