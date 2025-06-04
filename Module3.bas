Attribute VB_Name = "Module3"
Sub TakingValues()
    Dim TakingValuesworkbook As Workbook
    Dim TakingValuesWorksheet As Worksheet
    Dim Workbookname As String
    Dim Worksheetname As String
    Dim LastRow As Long
    Dim i As Long

    'Cashandcashequivalents Declaration
    Dim Cashandcashequivalents As Range
    Dim Cashandcashequivalentsvalue As Long

    'Marketablesecurities Declaration
    Dim MarketableSecuritieskeywords As Variant
    Dim MarketableSecuritiesRange As Range
    Dim MarketableSecuritiesLoop As Long
    Dim MarketableSecuritiesValue As Long
    Dim FoundMarketableSecurities As Boolean

    'Accounts Receivable Declaration
    Dim Accountsreceivable As Range
    Dim AccountsreceivableValue As Long

    'TotalInventory Declaration
    Dim TotalInventory As Range
    Dim TotalInventoryValue As Long

    'Other Current Assets Declaration
    Dim Othercurrentassets As Range
    Dim Othercurrentassetsvalue As Long

    'Acccountspayable Declaration
    Dim Accountspayable As Range
    Dim Accountspayablevalue As Long

    'Short Term Declaration
    Dim Shorttermdebt As Range
    Dim Shorttermdebtvalue As Long

    'Accrued Expenses dictionary
    Dim AESheet As Worksheet
    
    Dim AEdic As Object
    Set AEdic = CreateObject("Scripting.Dictionary")
    AEdic.Add Key:=LCase("Marketing and promotion"), Item:=0
    AEdic.Add Key:=LCase("compensation expenses"), Item:=0
    AEdic.Add Key:=LCase("Accrued marketing And promotion"), Item:=0
    AEdic.Add Key:=LCase("Accrued compensation"), Item:=0
    AEdic.Add Key:=LCase("Accrued interest"), Item:=0
    
    'Other Current Liabilites
    Dim Othercurrentliabilites As Range
    Dim OthercurrentliabilitesValue As Long
    
    'Op Cash Flow Ratio
    Dim OperatingCashFlowRatio As Range
    Dim OperatingCashFlowValue As Long
    
    LastRow = ThisWorkbook.Worksheets("WB NAMES").Cells(Rows.Count, 1).End(xlUp).Row

    For i = 1 To LastRow
        Workbookname = ThisWorkbook.Worksheets("WB NAMES").Cells(i, 1).Value
        Worksheetname = Trim(ThisWorkbook.Worksheets("WB NAMES").Cells(i, 2).Value)

        Set TakingValuesworkbook = Workbooks(Workbookname)

        On Error Resume Next
        Set TakingValuesWorksheet = TakingValuesworkbook.Worksheets(Worksheetname)
        On Error GoTo 0

        With TakingValuesWorksheet.Range("A4:A49")

            'Cash and cash equivalents
            Set Cashandcashequivalents = .Find(What:="Cash and cash equivalents", LookIn:=xlValues, LookAt:=xlWhole)
            If Not Cashandcashequivalents Is Nothing Then
                Cashandcashequivalentsvalue = Cashandcashequivalents.Offset(0, 1).Value
                ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(3, i + 1).Value = Cashandcashequivalentsvalue
            Else
                MsgBox "Cash and cash equivalents not found in " & Workbookname
                ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(3, i + 1).Value = 0
            End If

            'Marketable securities
            FoundMarketableSecurities = False
            MarketableSecuritieskeywords = Array("Marketable Securities", "Available-for-sale investment securities")
            For MarketableSecuritiesLoop = LBound(MarketableSecuritieskeywords) To UBound(MarketableSecuritieskeywords)
                Set MarketableSecuritiesRange = .Find(What:=MarketableSecuritieskeywords(MarketableSecuritiesLoop), LookIn:=xlValues, LookAt:=xlWhole)
                If Not MarketableSecuritiesRange Is Nothing Then
                    MarketableSecuritiesValue = MarketableSecuritiesRange.Offset(0, 1).Value
                    ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(4, i + 1).Value = MarketableSecuritiesValue
                    FoundMarketableSecurities = True
                    Exit For
                End If
            Next MarketableSecuritiesLoop
            If Not FoundMarketableSecurities Then
                MsgBox "Marketable Securities not Found in " & Workbookname
                ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(4, i + 1).Value = 0
            End If

            'Accounts Receivable
            Set Accountsreceivable = .Find(What:="Accounts receivable", LookIn:=xlValues, LookAt:=xlWhole)
            If Not Accountsreceivable Is Nothing Then
                AccountsreceivableValue = Accountsreceivable.Offset(0, 1).Value
                ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(5, i + 1).Value = AccountsreceivableValue
            Else
                MsgBox "Accounts Receivable not Found in " & Workbookname
                ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(5, i + 1).Value = 0
            End If

            'Total Inventory
            Set TotalInventory = .Find(What:="Total inventories", LookIn:=xlValues, LookAt:=xlWhole)
            If Not TotalInventory Is Nothing Then
                TotalInventoryValue = TotalInventory.Offset(0, 1).Value
                ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(6, i + 1).Value = TotalInventoryValue
            Else
                MsgBox "Total Inventory not found in " & Workbookname
                ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(6, i + 1).Value = 0
            End If

            'Other Current Assets
            Set Othercurrentassets = .Find(What:="Prepaid expenses and other current assets", LookIn:=xlValues, LookAt:=xlWhole)
            If Not Othercurrentassets Is Nothing Then
                Othercurrentassetsvalue = Othercurrentassets.Offset(0, 1).Value
                ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(7, i + 1).Value = Othercurrentassetsvalue
            Else
                MsgBox "Prepaid expenses and other current assets not found in " & Workbookname
                ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(7, i + 1).Value = 0
            End If

            'Accounts Payable
            Set Accountspayable = .Find(What:="Accounts payable", LookIn:=xlValues, LookAt:=xlWhole)
            If Not Accountspayable Is Nothing Then
                Accountspayablevalue = Accountspayable.Offset(0, 1).Value
                ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(10, i + 1).Value = Accountspayablevalue
            Else
                MsgBox "Accounts Payable not found in " & Workbookname
                ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(10, i + 1).Value = 0
            End If

            'Short term Debt
            Set Shorttermdebt = .Find(What:="Debt due within one year", LookIn:=xlValues, LookAt:=xlWhole)
            If Not Shorttermdebt Is Nothing Then
                Shorttermdebtvalue = Shorttermdebt.Offset(0, 1).Value
                ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(11, i + 1).Value = Shorttermdebtvalue
            Else
                MsgBox "Short term Debt not found in " & Workbookname
                ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(11, i + 1).Value = 0
            End If
        End With

        'Accrued Expenses Calculation
        On Error Resume Next
        Set AESheet = TakingValuesworkbook.Worksheets("Sales Data")
        On Error GoTo 0
        
        If Not AESheet Is Nothing Then
            Dim AERow As Long
            Dim AEName As String
            Dim AEValue As Long
            Dim AEtotal As Long
            AEtotal = 0
            
            For AccrRow = 2 To AESheet.Cells(AESheet.Rows.Count, "M").End(xlUp).Row
                AEName = LCase(Trim(AESheet.Cells(AccrRow, "M").Value))
                AEValue = AESheet.Cells(AccrRow, "N").Value

                If AEdic.exists(AEName) Then
                    AEtotal = AEtotal + AEValue
                End If
            Next AccrRow
            ThisWorkbook.Sheets("Liquidity Ratio Analysis ").Cells(12, i + 1).Value = AEtotal
        Else
            MsgBox "Accured Expenses Not Found " & Workbookname
            ThisWorkbook.Sheets("Liquidity Ratio Analysis ").Cells(12, i + 1).Value = 0
        End If
        
        'Other Current Liabilites
        With AESheet.Range("O1:O10")
            Set Othercurrentliabilites = .Find("Total Other Current Liabilities", LookIn:=xlValues, LookAt:=xlWhole)
            If Not Othercurrentliabilites Is Nothing Then
                OthercurrentliabilitesValue = Othercurrentliabilites.Offset(0, 1).Value
                ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(13, i + 1).Value = OthercurrentliabilitesValue
            Else
                MsgBox "Other Current Liabilites Value Not Found in  " & Workbookname
                ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(13, i + 1).Value = 0
            End If
        End With

        ' TOTAL OPERATING ACTIVITIES on DIFFERENT SHEET ***
        Dim CashFlowWorksheetName As String
        Dim CashFlowWorksheet As Worksheet
        Dim TotalOperatingActivities As Range
        CashFlowWorksheetName = Trim(ThisWorkbook.Worksheets("WB NAMES").Cells(i, 3).Value)
        On Error Resume Next
        Set CashFlowWorksheet = TakingValuesworkbook.Worksheets(CashFlowWorksheetName)
        On Error GoTo 0

        If Not CashFlowWorksheet Is Nothing Then
            With CashFlowWorksheet.Range("A:A")
                Set TotalOperatingActivities = .Find(What:="TOTAL OPERATING ACTIVITIES", LookIn:=xlValues, LookAt:=xlWhole)
                If Not TotalOperatingActivities Is Nothing Then
                    ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(15, i + 1).Value = TotalOperatingActivities.Offset(0, 1).Value
                Else
                    MsgBox "TOTAL OPERATING ACTIVITIES not found in " & Workbookname & " sheet " & CashFlowWorksheetName
                    ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(15, i + 1).Value = 0
                End If
            End With
        Else
            MsgBox "Cash Flow Worksheet '" & CashFlowWorksheetName & "' not found in " & Workbookname
            ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(15, i + 1).Value = 0
        End If
    Next i
    
    'Total Current Assets and Liabilites Calculation
    Dim TotalCurrentAssets As Long
    Dim TotalCurrentLiabilites As Long
    Dim TCACCount As Long
    Dim TCACCountLoop As Long
    TCACCount = ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(7, Columns.Count).End(xlToLeft).Column

    For TCACCountLoop = 2 To TCACCount

        TotalCurrentAssets = Application.Sum(ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Range( _
            ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(3, TCACCountLoop), _
            ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(7, TCACCountLoop)))

        TotalCurrentLiabilites = Application.Sum(ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Range( _
            ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(10, TCACCountLoop), _
            ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(13, TCACCountLoop)))

        ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(8, TCACCountLoop).Value = TotalCurrentAssets
        ThisWorkbook.Worksheets("Liquidity Ratio Analysis ").Cells(14, TCACCountLoop).Value = TotalCurrentLiabilites
    Next TCACCountLoop

End Sub


