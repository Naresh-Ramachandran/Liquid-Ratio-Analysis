Attribute VB_Name = "Module1"
Public RelatedWorkbooks(1 To 5) As Workbook

Sub LiquidRatioanalysismine()
   Dim LRAInput As String
    LRAInput = InputBox("Enter the last two digits of the starting year From 2001 to 2099 (e.g., '20' for 2020) to open Excel sheets for the next five years.")
    If StrPtr(LRAInput) = 0 Then Exit Sub
    LRAInput = Trim(LRAInput)

    If Not IsNumeric(LRAInput) Or LRAInput = "" Then
        MsgBox "Please enter a numeric value without spaces or special characters."
        Exit Sub
    ElseIf InStr(LRAInput, ".") > 0 Or Val(LRAInput) < 1 Or Val(LRAInput) > 99 Then
        MsgBox "Please enter a whole number between 01 and 99."
        Exit Sub
    End If

    Dim LRAyear As Long
    LRAyear = 2000 + CLng(LRAInput)

    Dim WorkbookPath As String
    Dim Workbookname As String
    Dim Workbookopen As Workbook
    Dim WorkbookTrue As Boolean
    Dim WorkbookopenFiveyears As Long

    For WorkbookopenFiveyears = 0 To 4
        Workbookname = "P&G-" & (LRAyear + WorkbookopenFiveyears) & ".xlsm"
        WorkbookPath = "C:\Users\NARESH\Desktop\VBA PROJECTS\Naresh Vba\" & Workbookname
        
        WorkbookTrue = False

        For Each Workbookopen In Workbooks
            If Workbookopen.Name = Workbookname Then
                MsgBox Workbookname & " is already opened"
                Set RelatedWorkbooks(WorkbookopenFiveyears + 1) = Workbookopen
                WorkbookTrue = True
                Exit For
            End If
        Next Workbookopen

        If Not WorkbookTrue Then
            If Dir(WorkbookPath) = "" Then
                MsgBox Workbookname & " does not exist in the path"
            Else
                Set RelatedWorkbooks(WorkbookopenFiveyears + 1) = Workbooks.Open(WorkbookPath)
            End If
        End If
    Next WorkbookopenFiveyears

    Call SaveWorkbookNames
End Sub

Sub SaveWorkbookNames()
    Dim i As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("WB NAMES")

    For i = 1 To 5
        If Not RelatedWorkbooks(i) Is Nothing Then
            ws.Cells(i, 1).Value = RelatedWorkbooks(i).Name
        Else
            ws.Cells(i, 1).ClearContents
        End If
    Next i
End Sub




