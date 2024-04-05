Private Sub CommandButton1_Click()
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim password As String
    Dim inputPassword As String
    
    ' 작업할 시트 지정
    Set ws = ThisWorkbook.Sheets("Student_Database")
    
    password = "gomsupak12"
    
    ' 셀 보호 해제
    ws.Unprotect password:=password
    
    ' abcdef 열의 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' abcdef 열의 마지막 행 비우기
    ws.Range("A" & lastRow & ":F" & lastRow).ClearContents
    
    ' 셀 보호 재설정
    ws.Protect password:=password
    
    MsgBox ("마지막 데이터를 제거하였습니다")

End Sub



Sub RemoveDuplicates()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim foundMatch As Boolean
    
    ' 작업할 워크시트 선택
    Set ws = ThisWorkbook.Sheets("Student_Database") ' 원하는 시트 이름으로 변경
    
    password = "gomsupak12"
    
    ' 셀 보호 해제
    ws.Unprotect password:=password
    
    ' 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' ABCD 열을 기준으로 중복된 행 확인 및 제거
    For i = lastRow To 2 Step -1
        foundMatch = False
        For j = i - 1 To 1 Step -1
            If ws.Cells(i, "A").Value = ws.Cells(j, "A").Value And _
               ws.Cells(i, "B").Value = ws.Cells(j, "B").Value And _
               ws.Cells(i, "C").Value = ws.Cells(j, "C").Value And _
               ws.Cells(i, "D").Value = ws.Cells(j, "D").Value Then
                foundMatch = True
                Exit For
            End If
        Next j
        If foundMatch Then
            ws.Rows(i).Delete
            'ws.Cells(i, "A").Value = ""
            'ws.Cells(i, "B").Value = ""
            'ws.Cells(i, "C").Value = ""
            'ws.Cells(i, "D").Value = ""
            'ws.Cells(i, "E").Value = ""
            'ws.Cells(i, "F").Value = ""
            
        End If
    Next i
    
    ' 셀 보호 재설정
    ws.Protect password:=password

    MsgBox ("중복된 데이터를 삭제하였습니다")
End Sub




Private Sub CommandButton2_Click()
    RemoveDuplicates
End Sub

