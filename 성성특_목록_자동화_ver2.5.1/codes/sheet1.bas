Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim i As Integer
    Dim lastRow As Integer
    Dim strPythonPath As String
    Dim strScriptPath As String
    Dim strCommand As String
    Dim strResult As String
    Dim objShell As Object
    Dim objFSO As Object
    Dim objFile As Object
    
    ' 작업할 시트 선택
    Set ws = ThisWorkbook.Sheets("Sheet1") ' 시트명에 맞게 수정
    password = "gomsupak12"
    
    ' 셀 보호 해제
    ws.Unprotect password:=password
    
    
    
    ' F2부터 아래로 모든 내용 지우기
    ws.Range("F2:F" & ws.Cells(ws.Rows.Count, "F").End(xlUp).Row).ClearContents
    
    ' E2부터 시작하여 각 셀에 대해 YoutubeCheck.py 실행
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 2 To lastRow
        ' 유튜브 링크가 있는 셀만 처리
        If ws.Range("E" & i).Value <> "" Then
            ' 파이썬 경로 설정
            strPythonPath = "C:\Users\SOSC근로\AppData\Local\Microsoft\WindowsApps\PythonSoftwareFoundation.Python.3.7_qbz5n2kfra8p0\python.exe" ' 파이썬 설치 경로에 맞게 수정
            
            ' 파이썬 스크립트 경로 설정
            strScriptPath = "C:\Users\SOSC근로\Desktop\근로학생\강성철\성성특_목록_자동화_ver2.5.1\codes\YoutubeCheck.py" ' 파이썬 스크립트 경로에 맞게 수정
            
            ' HWP 파일 경로
            strHWPFilePath = ws.Range("E" & i).Value
            
            ' 파이썬 스크립트 실행 명령 생성
            strCommand = strPythonPath & " " & strScriptPath & " """ & strHWPFilePath & """"
            
            ' 파이썬 스크립트 실행하여 결과 받아오기
            Set objShell = CreateObject("WScript.Shell")
            strResult = objShell.Exec(strCommand).StdOut.ReadAll
            
            ' 결과를 셀에 넣어주기
            ws.Range("F" & i).Value = strResult
        End If
    Next i
    ' 셀 보호 재설정
    ws.Protect password:=password
    
    MsgBox ("코드업데이트가 완료되었습니다")
    
End Sub

