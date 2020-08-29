Attribute VB_Name = "Module2"
Option Explicit

Const RowTitleResult_2 = 8
Const FirstColumnResult_2 = 15 'J

Dim arrInput
Dim B_Error As Boolean

Dim i As Byte

Private Sub ProcessScript2_AllSheets() 'generate Greatest table
    
    On Error GoTo Err_Execute1
    
    Application.ScreenUpdating = False
    
    For i = 1 To Thisworkbook.Sheets.Count
        Call ProcessScript2_OneSheet(Thisworkbook.Sheets(i).Name, B_Error)
    Next i
    
    Application.ScreenUpdating = True
    MsgBox "Done!"
    Exit Sub

Err_Execute1:
    Application.ScreenUpdating = True
    MsgBox "An error occurred. Exiting", vbCritical
End Sub

Private Sub ProcessScript2_Activesheet() 'generate Greatest table
    
    On Error GoTo Err_Execute2
    
    Application.ScreenUpdating = False
    
    Call ProcessScript2_OneSheet(ActiveSheet.Name, B_Error)
    If B_Error = True Then Exit Sub
    
    Application.ScreenUpdating = True
    MsgBox "Done!"
    Exit Sub

Err_Execute2:
    Application.ScreenUpdating = True
    MsgBox "An error occurred. Exiting", vbCritical
End Sub

Private Sub ProcessScript2_OneSheet(strSheetName_ As String, B_Error_ As Boolean)   'generate Greatest table
Dim ii As Long
Dim TickerSymbol As String
Dim FirstValue, LastValue, TotalStockVolume
Dim FirstColumn_Letter As String, LastColumn_Letter As String
Dim currentDifference
Dim GreatestInc, GreatestDec, GreatestVolume
Dim TickerSymbol_Inc, TickerSymbol_Dec, TickerSymbol_Total

    FirstColumn_Letter = Split(Cells(1, FirstColumnResult_2).Address, "$")(1)
    LastColumn_Letter = Split(Cells(1, FirstColumnResult_2 + 2).Address, "$")(1)

    B_Error_ = False
    Thisworkbook.Sheets(strSheetName_).Columns(FirstColumn_Letter & ":" & LastColumn_Letter).Clear
    If Thisworkbook.Sheets(strSheetName_).AutoFilterMode = True Then Thisworkbook.Sheets(strSheetName_).AutoFilter.ShowAllData
    
    If Thisworkbook.Sheets(strSheetName_).Cells(1, 1) <> "<ticker>" Or Thisworkbook.Sheets(strSheetName_).Cells(2, 1) = "" Then
        B_Error_ = True
        Application.ScreenUpdating = True
        MsgBox "Sheet """ & strSheetName_ & """ has wrong dataset!" & vbCrLf & "Check it and try again!", vbCritical
        Exit Sub
    End If
    With Thisworkbook.Sheets(strSheetName_)
        arrInput = .Range("A1:G" & .Cells(.Rows.Count, 1).End(xlUp).Row)
        .Cells(RowTitleResult_2, FirstColumnResult_2) = ""
        .Cells(RowTitleResult_2, FirstColumnResult_2 + 1) = "Ticker"
        .Cells(RowTitleResult_2, FirstColumnResult_2 + 2) = "Value"
        .Cells(RowTitleResult_2 + 1, FirstColumnResult_2) = "Greatest % increase"
        .Cells(RowTitleResult_2 + 2, FirstColumnResult_2) = "Greatest % decrease"
        .Cells(RowTitleResult_2 + 3, FirstColumnResult_2) = "Greatest total volume"

        
        TickerSymbol_Inc = ""
        TickerSymbol_Dec = ""
        TickerSymbol_Total = ""
        GreatestInc = 0
        GreatestDec = 0
        GreatestVolume = 0
        
        FirstValue = arrInput(2, 3)
        LastValue = arrInput(2, 6)
        TotalStockVolume = arrInput(2, 7)
        TickerSymbol = arrInput(2, 1)
        For ii = LBound(arrInput, 1) + 2 To UBound(arrInput, 1)
            If StrComp(arrInput(ii, 1), arrInput(ii - 1, 1), vbTextCompare) <> 0 Then
            
                    If FirstValue <> 0 Then
                            currentDifference = (LastValue - FirstValue) / FirstValue
                        ElseIf FirstValue = 0 And LastValue <> 0 Then
                            currentDifference = 1
                        ElseIf FirstValue = 0 And LastValue = 0 Then
                            currentDifference = 0
                    End If
                    '------------------------------------------------------------
                    If StrComp(TickerSymbol_Inc, "", vbTextCompare) <> 0 And currentDifference = GreatestInc Then
                        TickerSymbol_Inc = TickerSymbol_Inc & ", " & TickerSymbol
                    End If
                    
                    If StrComp(TickerSymbol_Inc, "", vbTextCompare) = 0 Or currentDifference > GreatestInc Then
                        TickerSymbol_Inc = TickerSymbol
                        GreatestInc = currentDifference
                    End If
                    '-------------------------------------------------------------
                    If StrComp(TickerSymbol_Dec, "", vbTextCompare) <> 0 And currentDifference = GreatestDec Then
                        TickerSymbol_Dec = TickerSymbol_Dec & ", " & TickerSymbol
                    End If
                    
                    If StrComp(TickerSymbol_Dec, "", vbTextCompare) = 0 Or currentDifference < GreatestDec Then
                        TickerSymbol_Dec = TickerSymbol
                        GreatestDec = currentDifference
                    End If
                    '---------------------------------------------------------------
                    If StrComp(TickerSymbol_Total, "", vbTextCompare) <> 0 And TotalStockVolume = GreatestVolume Then
                        TickerSymbol_Total = TickerSymbol_Total & ", " & TickerSymbol
                    End If
                    
                    If StrComp(TickerSymbol_Total, "", vbTextCompare) = 0 Or TotalStockVolume > GreatestVolume Then
                        TickerSymbol_Total = TickerSymbol
                        GreatestVolume = TotalStockVolume
                    End If
                    '----------------------------------------------------------------
                    '----------------------------------------------------------------
                    TickerSymbol = arrInput(ii, 1)
                    FirstValue = arrInput(ii, 3)
                    LastValue = arrInput(ii, 6)
                    TotalStockVolume = arrInput(ii, 7)
                Else
                    LastValue = arrInput(ii, 6)
                    TotalStockVolume = TotalStockVolume + arrInput(ii, 7)

            End If
        Next ii
        
        .Cells(RowTitleResult_2 + 1, FirstColumnResult_2 + 1) = TickerSymbol_Inc
        .Cells(RowTitleResult_2 + 1, FirstColumnResult_2 + 2) = GreatestInc
        .Cells(RowTitleResult_2 + 2, FirstColumnResult_2 + 1) = TickerSymbol_Dec
        .Cells(RowTitleResult_2 + 2, FirstColumnResult_2 + 2) = GreatestDec
        .Cells(RowTitleResult_2 + 3, FirstColumnResult_2 + 1) = TickerSymbol_Total
        .Cells(RowTitleResult_2 + 3, FirstColumnResult_2 + 2) = GreatestVolume
        
        .Cells(RowTitleResult_2 + 1, FirstColumnResult_2 + 2).NumberFormat = "0.00%"
        .Cells(RowTitleResult_2 + 2, FirstColumnResult_2 + 2).NumberFormat = "0.00%"
        
        .Range(FirstColumn_Letter & RowTitleResult_2 & ":" & LastColumn_Letter & RowTitleResult_2).Font.Bold = True
        .Range(FirstColumn_Letter & RowTitleResult_2 & ":" & FirstColumn_Letter & RowTitleResult_2 + 3).Font.Bold = True
        .Range(FirstColumn_Letter & RowTitleResult_2 & ":" & LastColumn_Letter & RowTitleResult_2).HorizontalAlignment = xlCenter
        .Range(FirstColumn_Letter & RowTitleResult_2 & ":" & LastColumn_Letter & RowTitleResult_2).VerticalAlignment = xlCenter

        .Columns(FirstColumn_Letter & ":" & LastColumn_Letter).AutoFit
        
        
    End With
End Sub
'-----------------------------------------------------------------

