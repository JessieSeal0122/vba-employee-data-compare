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
    
    ' ��l���ܼ�
    hasDifferences = False         ' �Ω�l�ܬO�_���t��
    duplicateFound = False         ' �Ω�l�ܬO�_�����ƭ��u�s��
    repeatedIDs1 = ""              ' �Ω��x�s Raw data 1 �������ƭ��u�s��
    repeatedIDs2 = ""              ' �Ω��x�s Raw data 2 �������ƭ��u�s��
    
    ' �]�w�u�@��
    Set ws1 = ThisWorkbook.Sheets("Raw data 1")
    Set ws2 = ThisWorkbook.Sheets("Raw data 2")
    
    ' ��X�u�@���̫�@�C
    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    
    ' �ˬd Raw data 1 �������ƭ��u�s��
    For i = 2 To lastRow1
        employeeID = ws1.Cells(i, "A").Value
        For j = i + 1 To lastRow1
            If ws1.Cells(j, "A").Value = employeeID Then
                duplicateFound = True
                repeatedIDs1 = repeatedIDs1 & employeeID & ", "
                ws1.Cells(i, "A").Interior.Color = RGB(255, 0, 0)  ' �N���ƪ��x�s��аO������
                ws1.Cells(j, "A").Interior.Color = RGB(255, 0, 0)
            End If
        Next j
    Next i

    ' �ˬd Raw data 2 �������ƭ��u�s��
    For i = 2 To lastRow2
        employeeID = ws2.Cells(i, "A").Value
        For j = i + 1 To lastRow2
            If ws2.Cells(j, "A").Value = employeeID Then
                duplicateFound = True
                repeatedIDs2 = repeatedIDs2 & employeeID & ", "
                ws2.Cells(i, "A").Interior.Color = RGB(255, 0, 0)  ' �N���ƪ��x�s��аO������
                ws2.Cells(j, "A").Interior.Color = RGB(255, 0, 0)
            End If
        Next j
    Next i

    ' �p�G�o�{���ƪ����u�s���A��ܰT���ðh�X�l�{��
    If duplicateFound Then
        Dim message As String
        message = "���ƪ����u�s��:" & vbCrLf
        If repeatedIDs1 <> "" Then
            repeatedIDs1 = Left(repeatedIDs1, Len(repeatedIDs1) - 2) ' �h���̫�@�ӳr���M�Ů�
            message = message & "Raw data 1: " & repeatedIDs1 & vbCrLf
        End If
        If repeatedIDs2 <> "" Then
            repeatedIDs2 = Left(repeatedIDs2, Len(repeatedIDs2) - 2) ' �h���̫�@�ӳr���M�Ů�
            message = message & "Raw data 2: " & repeatedIDs2
        End If
        MsgBox message
        Exit Sub
    End If

    ' �ˬdResults�u�@��O�_�s�b�A�Y���s�b�h�s�W�@�ӦW��Results���u�@��A
    ' �Y�s�b�h�M���쥻��Results�u�@�����e�A�T�O�C�������VBA�ɬݨ쪺���O�̷s����ﵲ�G
    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("Results")
    On Error GoTo 0
    If wsResult Is Nothing Then
        Set wsResult = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsResult.Name = "Results"
    End If
    wsResult.Cells.Clear
    
    ' �]�w������M���D�C
    colsToCompare = Array("O", "Y", "R", "BL")
    wsResult.Range("B1:M1").Value = Array("Employee ID", "Worker", "Dept (Raw1)", "Cost Center (Raw1)", "Company Code (Raw1)", "Location (Raw1)", _
                                          "AR (Raw2)", "Dept (Raw2)", "Cost Center (Raw2)", "Company Code (Raw2)", "Location (Raw2)", "AS (Raw2)")
    lastRowResult = 2

    ' �q Raw data 1 �v���ˬd
    For i = 2 To lastRow1
        employeeID = ws1.Cells(i, "A").Value
        Set foundCell = ws2.Range("A2:A" & lastRow2).Find(employeeID, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not foundCell Is Nothing Then
            Dim mismatch As Boolean
            mismatch = False
            
            ' ���C�ӫ��w���
            For j = LBound(colsToCompare) To UBound(colsToCompare)
                If ws1.Cells(i, colsToCompare(j)).Value <> foundCell.Offset(0, ws2.Columns(colsToCompare(j)).Column - 1).Value Then
                    mismatch = True
                    hasDifferences = True ' �o�{���@�P
                End If
            Next j
            
            ' �Y���t���A�N��ƽƻs�� Results �u�@��
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

    ' �ˬd�O�_���t���A��ܰT��
    If hasDifferences Then
        wsResult.Columns("B:M").AutoFit
        MsgBox "��粒���I���G�w�x�s�� 'Results' �u�@��C"
    Else
        MsgBox "�L�ܧ���"
    End If

    ' ��� Results �u�@������줺�e
    For i = 2 To lastRowResult - 1
        If wsResult.Cells(i, "D").Value <> wsResult.Cells(i, "I").Value Then
            wsResult.Cells(i, "D").Interior.Color = RGB(255, 255, 0) ' ����
            wsResult.Cells(i, "I").Interior.Color = RGB(255, 255, 0) ' ����
        End If
        If wsResult.Cells(i, "E").Value <> wsResult.Cells(i, "J").Value Then
            wsResult.Cells(i, "E").Interior.Color = RGB(255, 255, 0) ' ����
            wsResult.Cells(i, "J").Interior.Color = RGB(255, 255, 0) ' ����
        End If
        If wsResult.Cells(i, "F").Value <> wsResult.Cells(i, "K").Value Then
            wsResult.Cells(i, "F").Interior.Color = RGB(255, 255, 0) ' ����
            wsResult.Cells(i, "K").Interior.Color = RGB(255, 255, 0) ' ����
        End If
        If wsResult.Cells(i, "G").Value <> wsResult.Cells(i, "L").Value Then
            wsResult.Cells(i, "G").Interior.Color = RGB(255, 255, 0) ' ����
            wsResult.Cells(i, "L").Interior.Color = RGB(255, 255, 0) ' ����
        End If
    Next i

    ' �]�w A1 �x�s�欰����
    wsResult.Range("A1").Interior.Color = RGB(255, 255, 0)

    ' �]�w A1 �x�s�檺���e�� "Effective Date"
    wsResult.Range("A1").Value = "Effective Date"

    ' �]�w B1 �� G1 �x�s�檺 RGB �C�⬰ (142, 169, 219)
    wsResult.Range("B1:G1").Interior.Color = RGB(142, 169, 219)

    ' �]�w H1 �� M1 �x�s�檺 RGB �C�⬰ (198, 224, 180)
    wsResult.Range("H1:M1").Interior.Color = RGB(198, 224, 180)

    ' ������зǵ��ϼҦ�
    With wsResult
        .Activate
        ActiveWindow.View = xlNormalView
    End With

End Sub

