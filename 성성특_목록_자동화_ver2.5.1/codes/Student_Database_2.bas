Private Sub CommandButton1_Click()
    ProcessStudentDatabase

End Sub

Sub ProcessStudentDatabase()
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim studentCourses As Object
    Dim studentInfo As Variant ' 문자열로 선언
    Dim lastRow As Long
    Dim i As Long
    Dim studentNames() As Variant
    Dim index As Long
    
    ' Student_Database 시트 선택
    Set ws = ThisWorkbook.Sheets("Student_Database")
    Set ws2 = ThisWorkbook.Sheets("Student_Database_2")
    
    password = "gomsupak12"
    
    ' 셀 보호 해제
    ws.Unprotect password:=password
    ws2.Unprotect password:=password
    
    
    
    ' 초기화
    For i = 2 To 999
        ws2.Range("A" & i).Value = ""
        ws2.Range("B" & i).Value = ""
        ws2.Range("C" & i).Value = ""
        ws2.Range("D" & i).Value = ""
    Next i
    
    ' 학생 수강 강의 수를 저장할 딕셔너리 초기화
    Set studentCourses = CreateObject("Scripting.Dictionary")
    
    ' 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 각 학생의 수강 강의 수 세기
    For i = 2 To lastRow
        studentID = ws.Cells(i, 2).Value ' 학번
        studentInfo = ws.Cells(i, 1).Value & "|" & studentID & "|" & ws.Cells(i, 3).Value ' 이름|학번|학과
        If studentCourses.Exists(studentInfo) Then
            ' 이미 해당 학생의 정보가 딕셔너리에 존재하는 경우 수강한 강의 수 증가
            studentCourses(studentInfo) = studentCourses(studentInfo) + 1
        Else
            ' 해당 학생의 정보가 딕셔너리에 존재하지 않는 경우 새로운 키로 추가
            studentCourses.Add studentInfo, 1
        End If
    Next i
    
    ' 딕셔너리의 키를 임시 배열에 복사
    ReDim studentNames(1 To studentCourses.Count)
    index = 1
    For Each studentInfo In studentCourses.Keys
        studentNames(index) = studentInfo
        index = index + 1
    Next studentInfo
    
    ' 임시 배열을 반복하여 결과를 기록
    For i = 1 To UBound(studentNames)
        Dim studentInfoParts() As String
        studentInfoParts = Split(studentNames(i), "|")
        ws2.Cells(i + 1, 1).Value = studentInfoParts(0) ' 이름
        ws2.Cells(i + 1, 2).Value = studentInfoParts(1) ' 학번
        ws2.Cells(i + 1, 3).Value = studentInfoParts(2) ' 학과
        ws2.Cells(i + 1, 4).Value = studentCourses(studentNames(i)) ' 수강한 강의 개수
    Next i
    
    ' 셀 보호 재설정
    ws.Protect password:=password
    ws2.Protect password:=password
    
    MsgBox ("새로고침이 완료되었습니다")
    
End Sub





