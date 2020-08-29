Attribute VB_Name = "Module1"
Option Explicit

Const RowTitleResult_1 = 8
Const FirstColumnResult_1 = 10 'J

Dim arrInput
Dim B_Error As Boolean

Dim i As Byte

Private Sub ProcessScript1_AllSheets() 'generate total table
    
    On Error GoTo Err_Execute1
    
    Application.ScreenUpdating = False
    
    For i = 1 To Thisworkbook.Sheets.Count
        Call ProcessScript1_OneSheet(Thisworkbook.Sheets(i).Name, B_Error)
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "Done!"
    Exit Sub

Err_Execute1:
    Application.ScreenUpdating = True
    MsgBox "An error occurred. Exiting", vbCritical
End Sub

Private Sub ProcessScript1_Activesheet() 'generate total table
    
    On Error GoTo Err_Execute2
    
    Application.ScreenUpdating = False
    
    Call ProcessScript1_OneSheet(ActiveSheet.Name, B_Error)
    If B_Error = True Then Exit Sub
    
    
    Application.ScreenUpdating = True
    MsgBox "Done!"
    Exit Sub

Err_Execute2:
    Application.ScreenUpdating = True
    MsgBox "An error occurred. Exiting", vbCritical
End Sub

Private Sub ProcessScript1_OneSheet(strSheetName_ As String, B_Error_ As Boolean)   'generate total table
Dim ii As Long, jj As Long
Dim TickerSymbol As String
Dim FirstValue, LastValue, TotalStockVolume
Dim FirstColumn_Letter As String, LastColumn_Letter As String

    FirstColumn_Letter = Split(Cells(1, FirstColumnResult_1).Address, "$")(1)
    LastColumn_Letter = Split(Cells(1, FirstColumnResult_1 + 3).Address, "$")(1)
    
    B_Error_ = False
    Thisworkbook.Sheets(strSheetName_).Columns(FirstColumn_Letter & ":" & LastColumn_Letter).Clear
    If Thisworkbook.Sheets(strSheetName_).AutoFilterMode = True Then Thisworkbook.Sheets(strSheetName_).Cells.AutoFilter
    
    If Thisworkbook.Sheets(strSheetName_).Cells(1, 1) <> "<ticker>" Or Thisworkbook.Sheets(strSheetName_).Cells(2, 1) = "" Then
        B_Error_ = True
        Application.ScreenUpdating = True
        MsgBox "Sheet """ & strSheetName_ & """ has wrong dataset!" & vbCrLf & "Check it and try again!", vbCritical
        Exit Sub
    End If
    With Thisworkbook.Sheets(strSheetName_)
        arrInput = .Range("A1:G" & .Cells(.Rows.Count, 1).End(xlUp).Row)
        .Cells(RowTitleResult_1, FirstColumnResult_1) = "Ticker Symbol"
        .Cells(RowTitleResult_1, FirstColumnResult_1 + 1) = "Yearly Change"
        .Cells(RowTitleResult_1, FirstColumnResult_1 + 2) = "Percent Change"
        .Cells(RowTitleResult_1, FirstColumnResult_1 + 3) = "Total Stock Volume"
        
        jj = RowTitleResult_1 + 1
        FirstValue = arrInput(2, 3)
        LastValue = arrInput(2, 6)
        TotalStockVolume = arrInput(2, 7)
        TickerSymbol = arrInput(2, 1)
        For ii = LBound(arrInput, 1) + 2 To UBound(arrInput, 1)
            If StrComp(arrInput(ii, 1), arrInput(ii - 1, 1), vbTextCompare) <> 0 Then
                
                    .Cells(jj, FirstColumnResult_1) = TickerSymbol
                    .Cells(jj, FirstColumnResult_1 + 1) = LastValue - FirstValue
                    If FirstValue <> 0 Then
                            .Cells(jj, FirstColumnResult_1 + 2) = (LastValue - FirstValue) / FirstValue
                        ElseIf FirstValue = 0 And LastValue <> 0 Then
                            .Cells(jj, FirstColumnResult_1 + 2) = 1
                        ElseIf FirstValue = 0 And LastValue = 0 Then
                            .Cells(jj, FirstColumnResult_1 + 2) = 0
                    End If
                    .Cells(jj, FirstColumnResult_1 + 3) = TotalStockVolume
                    jj = jj + 1
                     
                    TickerSymbol = arrInput(ii, 1)
                    FirstValue = arrInput(ii, 3)
                    LastValue = arrInput(ii, 6)
                    TotalStockVolume = arrInput(ii, 7)
                Else
                    LastValue = arrInput(ii, 6)
                    TotalStockVolume = TotalStockVolume + arrInput(ii, 7)
                 
            End If
        Next ii
        .Cells(jj, FirstColumnResult_1) = TickerSymbol
        .Cells(jj, FirstColumnResult_1 + 1) = LastValue - FirstValue
        .Cells(jj, FirstColumnResult_1 + 2) = (LastValue - FirstValue) / FirstValue
        .Cells(jj, FirstColumnResult_1 + 3) = TotalStockVolume
        .Columns(FirstColumnResult_1 + 2).NumberFormat = "0.00%"
        'conditional formatiing
        With .Range(.Cells(RowTitleResult_1 + 1, FirstColumnResult_1 + 1), .Cells(.Cells(.Rows.Count, FirstColumnResult_1 + 1).End(xlUp).Row, FirstColumnResult_1 + 2))
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 5296274
                .TintAndShade = 0
            End With
            .FormatConditions(1).StopIfTrue = False
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
            End With
            .FormatConditions(1).StopIfTrue = False
        End With

        .Range(FirstColumn_Letter & RowTitleResult_1 & ":" & LastColumn_Letter & RowTitleResult_1).Font.Bold = True
        .Range(FirstColumn_Letter & RowTitleResult_1 & ":" & LastColumn_Letter & RowTitleResult_1).HorizontalAlignment = xlCenter
        .Range(FirstColumn_Letter & RowTitleResult_1 & ":" & LastColumn_Letter & RowTitleResult_1).VerticalAlignment = xlCenter
        
        .Range(FirstColumn_Letter & RowTitleResult_1 & ":" & LastColumn_Letter & RowTitleResult_1).AutoFilter
        
        .Columns(FirstColumn_Letter & ":" & LastColumn_Letter).AutoFit
        
    End With
End Sub
'-----------------------------------------------------------------
