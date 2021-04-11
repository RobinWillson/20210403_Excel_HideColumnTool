Public xWB As Workbook
Public xWS As Worksheet

Sub SheetProtect()
    ActiveSheet.Protect "1234"
End Sub
Sub SheetUnProtect()
    ActiveSheet.Unprotect "1234"
End Sub
Sub select_WB()

    Dim xWBName
    With ThisWorkbook.Sheets("001")
    For i = 5 To 15
        If .Cells(i, 3) = "V" Then
            xWBName = .Cells(i, 2)
            'MsgBox "The active workbook is " & vbCrLf & xWBName
            'Workbooks(xWBName).Activate
            GoTo SelectWorkBookFinish
        End If
    Next
    End With
SelectWorkBookFinish:
    
    Set xWB = Workbooks(xWBName)
    
End Sub
Sub select_WS()

    Dim xShName
    With ThisWorkbook.Sheets("001")
    For i = 21 To 35
        If .Cells(i, 3) = "V" Then
            xShName = .Cells(i, 2)
            GoTo SelectWorkSheetFinish
        End If
    Next
    End With
SelectWorkSheetFinish:

    Set xWS = xWB.Sheets(xShName)
    
End Sub

Sub List_All_Opened_XLS()
    Dim xWBName As String
    Dim xWB As Workbook
    Dim xSelect As String
    Dim iRow, iColumn
    Call SheetUnProtect
    iColumn = 2
    iRow = 5
    
    ThisWorkbook.Sheets("001").Range(Cells(5, 2), Cells(15, 3)).ClearContents
    For Each xWB In Application.Workbooks
        xWBName = xWB.Name
        ThisWorkbook.Sheets("001").Cells(iRow, iColumn) = xWBName
        iRow = iRow + 1
    Next
    Call SheetProtect
End Sub

Sub List_Opened_XLS_Sheets()
    Dim xWBName As String
    Dim xWSName As String
    Dim xSelect As String
    Dim iRow, iColumn
    Dim i, j, k

On Error GoTo ErrorHandler   ' 啟用錯誤處理機制
    Call SheetUnProtect
    Call select_WB
    
    iColumn = 2
    iRow = 21
    ThisWorkbook.Sheets("001").Range(Cells(21, 2), Cells(35, 3)).ClearContents
    'Set xWB = Workbooks(xWBName)
    For i = 1 To xWB.Sheets.Count
        ThisWorkbook.Sheets("001").Cells(iRow, iColumn) = xWB.Sheets(i).Name
        iRow = iRow + 1
    Next
SetWorkSheetFinish:

Call SheetProtect
Exit Sub
ErrorHandler:       ' 錯誤處理用的程式碼
MsgBox "錯誤 請確認檔名選擇正確"

End Sub

'Sub QuickClearUp()
'    Dim mColor As Long
'    Dim xEndRow
'    Dim i, j, k
'On Error GoTo ErrorHandler   ' 啟用錯誤處理機制
'
'    '--選擇檔名--------------
'    Call select_WB
'
'    '--選擇分頁--------------
'    Call select_WS
'
'    '--設定LV2背景顏色---------
'    mColor = RGB(255, 242, 204)
'    xWB.Activate
'    xWS.Activate
'    xEndRow = xWS.Cells(65535, 1).End(xlUp).Row
'    xWS.Range(Cells(9, 1), Cells(xEndRow, 132)).Interior.Color = xlNone
'    For j = 9 To xEndRow
'        If xWS.Cells(j, 65) = 2 Then
'            xWS.Range(Cells(j, 1), Cells(j, 132)).Interior.Color = mColor
'        End If
'    Next
'    '--設定全部列高為25---------
'    xWS.Cells.Select
'    Selection.RowHeight = 25
'    '--隱藏1~6列---------
'    xWS.Rows("1:6").Select
'    Selection.EntireRow.Hidden = True
'    '--啟動篩選功能------
'    xWS.AutoFilterMode = False
'    xWS.Rows("8:8").Select
'    Selection.AutoFilter
'    '--凍結窗格第9列------
'    ActiveWindow.FreezePanes = False
'    xWS.Rows("9:9").Select
'    ActiveWindow.FreezePanes = True
'
'Exit Sub
'ErrorHandler:       ' 錯誤處理用的程式碼
'MsgBox "錯誤 請確認檔名選擇正確"
'
'End Sub
Sub ListTitle()
    Dim xEndRow
    Dim TitleRow, TitleEndCol
    Dim TitleWord, TmpTitleWord, Letter
    Dim i, j, k
On Error GoTo ErrorHandler   ' 啟用錯誤處理機制
Call SheetUnProtect

    '--選擇檔名--------------
    Call select_WB

    '--選擇分頁--------------
    Call select_WS
    
    '--展開所有欄位----------
    xWB.Activate
    xWS.Activate
    Columns.EntireColumn.Hidden = False

    
    '--List Title ----
    TitleRow = ThisWorkbook.Sheets("001").Cells(3, 8)
    TitleEndCol = xWS.Cells(TitleRow, 1999).End(xlToLeft).Column
    ThisWorkbook.Activate
    ThisWorkbook.Sheets("001").Range(Cells(8, 7), Cells(1999, 8)).ClearContents
    For i = 1 To TitleEndCol
        TitleWord = xWS.Cells(TitleRow, i).Text
        TmpTitleWord = Replace(TitleWord, Chr(10), "")
        ThisWorkbook.Sheets("001").Cells(7 + i, 7) = TmpTitleWord
        ThisWorkbook.Sheets("001").Cells(7 + i, 8) = i
    Next
    
'---------

Call SheetProtect
Exit Sub
ErrorHandler:       ' 錯誤處理用的程式碼
MsgBox "錯誤 請確認檔名選擇正確"

End Sub

Sub ManagerColumn(ByRef InputColumn As Integer)
    Dim i, j, k
    Dim tCol
    
On Error GoTo ErrorHandler   ' 啟用錯誤處理機制


    '--選擇檔名--------------
    Call select_WB

    '--選擇分頁--------------
    Call select_WS

    '-------------------
    xWB.Activate
    xWS.Activate
    Cells.EntireColumn.Hidden = False
    xWS.Range(Cells(1, 1), Cells(1, 999)).EntireColumn.Hidden = True
    
    '-------------------
    With ThisWorkbook.Sheets("002")
        For i = 5 To 50
            tmpbool = (.Cells(i, InputColumn) = "Y")
            If .Cells(i, InputColumn) = "Y" Then
                tCol = .Cells(i, InputColumn - 1)
                If tCol <= 0 Or IsNumeric(tCol) = False Then GoTo SkipHidden
                xWS.Activate
                xWS.Cells(1, tCol).EntireColumn.Hidden = False
            End If
SkipHidden:
        Next
        
        ActiveWindow.ScrollColumn = 1
    End With
'---------

Exit Sub
ErrorHandler:       ' 錯誤處理用的程式碼
MsgBox "錯誤" & Chr(10) & "請確認檔名選擇正確" & Chr(10) & "確認是否有欄位數值為0"

End Sub

Sub UnHideAllColumn()
    Dim i, j, k
    
On Error GoTo ErrorHandler   ' 啟用錯誤處理機制
'On Error Resume Next
    

    '--選擇檔名--------------
    Call select_WB

    '--選擇分頁--------------
    Call select_WS
    
    '-------------------
    xWB.Activate
    xWS.Activate
    Columns.EntireColumn.Hidden = False

'-------------------
Exit Sub
ErrorHandler:       ' 錯誤處理用的程式碼
MsgBox "錯誤 請確認檔名選擇正確"

End Sub


Public Function xFind(Target As Variant, FindRange As Range, ReturnColumn As Integer) As Variant
    Set myfind = FindRange.Find(What:=Target, LookIn:=xlValues, lookat:=xlWhole)
    
    xFind = myfind.Cells(1).Offset(0, ReturnColumn).Text
    MsgBox "find target = " & Target & Chr(10) & "result = " & myfind
    '==> find target=* Part number , result = Parent Part Number
    'xFind = result
End Function
Public Function xFind2(Target As Variant, FindRange As Range, ReturnColumn As Integer) As Variant

    Dim str1, str2, Audience, result
    Dim i, j, k
    '--------
    If Target = "" Then
        Find2 = CVErr(xlErrNA) '=> returen "N/A"
        Exit Function
    End If
    '--------
    For i = 1 To FindRange.Rows.Count
       For j = 1 To Len(Target)
           str1 = Mid(Target, j, 1)
           result = FindRange.Cells(1).Offset(i, 0)
           str2 = Mid(result, j, 1)
           If str1 <> str2 Then GoTo FindNothing
       Next
       GoTo FindResult
FindNothing:
    Next
FindResult:
    'MsgBox result
    'MsgBox FindRange.Cells(1).Offset(i, ReturnColumn)
    xFind2 = FindRange.Cells(1).Offset(i, ReturnColumn)
End Function
