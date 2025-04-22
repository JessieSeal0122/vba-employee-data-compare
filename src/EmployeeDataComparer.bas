Attribute VB_Name = "Module_EmployeeDataComparer"
Sub EmployeeDataComparer()

 Dim ws1 As Worksheet, ws2 As Worksheet, wsResult As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, lastRowResult As Long
    Dim cell As Range, foundCell As Range
    Dim employeeID As String
    Dim repeatedIDs1 As String, repeatedIDs2 As String
    Dim duplicateFound As Boolean
    Dim colsToCompare As Variant
    Dim i As Integer, j As Integer
    Dim hasDifferences As Boolean
    
    ' 初始化變數
    hasDifferences = False         ' 用於追蹤是否有差異
    duplicateFound = False         ' 用於追蹤是否有重複員工編號
    repeatedIDs1 = ""              ' 用於儲存 Raw data 1 中的重複員工編號
    repeatedIDs2 = ""              ' 用於儲存 Raw data 2 中的重複員工編號
    
    ' 設定工作表
    Set ws1 = ThisWorkbook.Sheets("Raw data 1")
    Set ws2 = ThisWorkbook.Sheets("Raw data 2")
    
    ' 找出工作表的最後一列
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    
    ' 檢查 Raw data 1 中的重複員工編號
    For i = 2 To lastRow1
        employeeID = ws1.Cells(i, "A").Value
        For j = i + 1 To lastRow1
            If ws1.Cells(j, "A").Value = employeeID Then
                duplicateFound = True
                repeatedIDs1 = repeatedIDs1 & employeeID & ", "
                ws1.Cells(i, "A").Interior.Color = RGB(255, 0, 0)  ' 將重複的儲存格標記為紅色
                ws1.Cells(j, "A").Interior.Color = RGB(255, 0, 0)
            End If
        Next j
    Next i

    ' 檢查 Raw data 2 中的重複員工編號
    For i = 2 To lastRow2
        employeeID = ws2.Cells(i, "A").Value
        For j = i + 1 To lastRow2
            If ws2.Cells(j, "A").Value = employeeID Then
                duplicateFound = True
                repeatedIDs2 = repeatedIDs2 & employeeID & ", "
                ws2.Cells(i, "A").Interior.Color = RGB(255, 0, 0)  ' 將重複的儲存格標記為紅色
                ws2.Cells(j, "A").Interior.Color = RGB(255, 0, 0)
            End If
        Next j
    Next i

    ' 如果發現重複的員工編號，顯示訊息並退出子程序
    If duplicateFound Then
        Dim message As String
        message = "重複的員工編號:" & vbCrLf
        If repeatedIDs1 <> "" Then
            repeatedIDs1 = Left(repeatedIDs1, Len(repeatedIDs1) - 2) ' 去掉最後一個逗號和空格
            message = message & "Raw data 1: " & repeatedIDs1 & vbCrLf
        End If
        If repeatedIDs2 <> "" Then
            repeatedIDs2 = Left(repeatedIDs2, Len(repeatedIDs2) - 2) ' 去掉最後一個逗號和空格
            message = message & "Raw data 2: " & repeatedIDs2
        End If
        MsgBox message
        Exit Sub
    End If

    ' 檢查Results工作表是否存在，若不存在則新增一個名為Results的工作表，
    ' 若存在則清除原本的Results工作表的內容，確保每次執行時VBA時看到的都是最新的比對結果
    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("Results")
    On Error GoTo 0
    If wsResult Is Nothing Then
        Set wsResult = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsResult.Name = "Results"
    End If
    wsResult.Cells.Clear
    
    ' 設定比對欄位和標題列
    colsToCompare = Array("O", "Y", "R", "BL")
    wsResult.Range("B1:M1").Value = Array("Employee ID", "Worker", "Dept (Raw1)", "Cost Center (Raw1)", "Company Code (Raw1)", "Location (Raw1)", _
                                          "AR (Raw2)", "Dept (Raw2)", "Cost Center (Raw2)", "Company Code (Raw2)", "Location (Raw2)", "AS (Raw2)")
    lastRowResult = 2

    ' 從 Raw data 1 逐行檢查
    For i = 2 To lastRow1
        employeeID = ws1.Cells(i, "A").Value
        Set foundCell = ws2.Range("A2:A" & lastRow2).Find(employeeID, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not foundCell Is Nothing Then
            Dim mismatch As Boolean
            mismatch = False
            
            ' 比對每個指定欄位
            For j = LBound(colsToCompare) To UBound(colsToCompare)
                If ws1.Cells(i, colsToCompare(j)).Value <> foundCell.Offset(0, ws2.Columns(colsToCompare(j)).Column - 1).Value Then
                    mismatch = True
                    hasDifferences = True ' 發現不一致
                End If
            Next j
            
            ' 若有差異，將資料複製到 Results 工作表
            If mismatch Then
                wsResult.Cells(lastRowResult, "B").Value = employeeID
                wsResult.Cells(lastRowResult, "C").Value = ws1.Cells(i, "B").Value
                wsResult.Cells(lastRowResult, "D").Value = ws1.Cells(i, "Y").Value
                wsResult.Cells(lastRowResult, "E").Value = ws1.Cells(i, "R").Value
                wsResult.Cells(lastRowResult, "F").Value = ws1.Cells(i, "O").Value
                wsResult.Cells(lastRowResult, "G").Value = ws1.Cells(i, "BL").Value
                
                wsResult.Cells(lastRowResult, "H").Value = foundCell.Offset(0, ws2.Columns("AR").Column - 1).Value
                wsResult.Cells(lastRowResult, "I").Value = foundCell.Offset(0, ws2.Columns("Y").Column - 1).Value
                wsResult.Cells(lastRowResult, "J").Value = foundCell.Offset(0, ws2.Columns("R").Column - 1).Value
                wsResult.Cells(lastRowResult, "K").Value = foundCell.Offset(0, ws2.Columns("O").Column - 1).Value
                wsResult.Cells(lastRowResult, "L").Value = foundCell.Offset(0, ws2.Columns("BL").Column - 1).Value
                wsResult.Cells(lastRowResult, "M").Value = foundCell.Offset(0, ws2.Columns("AS").Column - 1).Value
                
                lastRowResult = lastRowResult + 1
            End If
        End If
    Next i

    ' 檢查是否有差異，顯示訊息
    If hasDifferences Then
        wsResult.Columns("B:M").AutoFit
        MsgBox "比對完成！結果已儲存到 'Results' 工作表。"
    Else
        MsgBox "無變更資料"
    End If

    ' 比對 Results 工作表中的欄位內容
    For i = 2 To lastRowResult - 1
        If wsResult.Cells(i, "D").Value <> wsResult.Cells(i, "I").Value Then
            wsResult.Cells(i, "D").Interior.Color = RGB(255, 255, 0) ' 黃色
            wsResult.Cells(i, "I").Interior.Color = RGB(255, 255, 0) ' 黃色
        End If
        If wsResult.Cells(i, "E").Value <> wsResult.Cells(i, "J").Value Then
            wsResult.Cells(i, "E").Interior.Color = RGB(255, 255, 0) ' 黃色
            wsResult.Cells(i, "J").Interior.Color = RGB(255, 255, 0) ' 黃色
        End If
        If wsResult.Cells(i, "F").Value <> wsResult.Cells(i, "K").Value Then
            wsResult.Cells(i, "F").Interior.Color = RGB(255, 255, 0) ' 黃色
            wsResult.Cells(i, "K").Interior.Color = RGB(255, 255, 0) ' 黃色
        End If
        If wsResult.Cells(i, "G").Value <> wsResult.Cells(i, "L").Value Then
            wsResult.Cells(i, "G").Interior.Color = RGB(255, 255, 0) ' 黃色
            wsResult.Cells(i, "L").Interior.Color = RGB(255, 255, 0) ' 黃色
        End If
    Next i

    ' 設定 A1 儲存格為黃色
    wsResult.Range("A1").Interior.Color = RGB(255, 255, 0)

    ' 設定 A1 儲存格的內容為 "Effective Date"
    wsResult.Range("A1").Value = "Effective Date"

    ' 設定 B1 到 G1 儲存格的 RGB 顏色為 (142, 169, 219)
    wsResult.Range("B1:G1").Interior.Color = RGB(142, 169, 219)

    ' 設定 H1 到 M1 儲存格的 RGB 顏色為 (198, 224, 180)
    wsResult.Range("H1:M1").Interior.Color = RGB(198, 224, 180)

    ' 切換到標準視圖模式
    With wsResult
        .Activate
        ActiveWindow.View = xlNormalView
    End With

End Sub

