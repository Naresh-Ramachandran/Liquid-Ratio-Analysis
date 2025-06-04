Attribute VB_Name = "Module2"
Sub Takingsheets()
    Dim OpenedWorkbook As Workbook
    Dim BalanceSheet As Worksheet
    Dim WbName As String
    Dim i As Long
    Dim WBnamesnumber As Long
    Dim FoundSheet As Boolean
    Dim FoundCashSheet As Boolean
    Dim wsNames As Worksheet
    Dim ws As Worksheet
    Dim Cell As Range
    
    Set wsNames = ThisWorkbook.Worksheets("WB NAMES")
    
    WBnamesnumber = wsNames.Cells(wsNames.Rows.Count, 1).End(xlUp).Row

    For i = 1 To WBnamesnumber
        WbName = wsNames.Cells(i, 1).Value
        
        On Error Resume Next
        Set OpenedWorkbook = Workbooks(WbName)
        On Error GoTo 0

        If OpenedWorkbook Is Nothing Then
            MsgBox WbName & " Workbook Closed"
            Exit Sub
        Else
            FoundSheet = False
            FoundCashSheet = False
            
            ' Balance Sheet search & formatting
            For Each BalanceSheet In OpenedWorkbook.Worksheets
                If BalanceSheet.Name Like "*Balance Sheet*" Then
                    wsNames.Cells(i, 2).Value = BalanceSheet.Name
                    For Each Cell In BalanceSheet.Range("B7:B44")
                        Cell.NumberFormat = "0.00"
                    Next Cell
                    FoundSheet = True
                    Exit For
                End If
            Next BalanceSheet
            
            ' Statement of Cash sheet & format
            For Each ws In OpenedWorkbook.Worksheets
                If ws.Name Like "*Statement of Cash*" Then
                    wsNames.Cells(i, 3).Value = ws.Name
                    For Each Cell In ws.Range("B7:B44")
                        Cell.NumberFormat = "0.00"
                    Next Cell
                    FoundCashSheet = True
                    Exit For
                End If
            Next ws

            If Not FoundSheet Then
                MsgBox "No sheet with name containing 'Balance Sheet' found in " & WbName
                wsNames.Cells(i, 2).ClearContents
            End If
            
            If Not FoundCashSheet Then
                MsgBox "No sheet with name containing 'Statement of Cash' found in " & WbName
                wsNames.Cells(i, 3).ClearContents
            End If
        End If
    Next i
End Sub




