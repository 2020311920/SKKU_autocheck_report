



Private Sub CommandButton1_Click()

    Dim filePath As Variant
    Dim strHWPFilePath As String
    Dim objFSO As Object
    Dim objFile As Object
    
    Dim objShell As Object
    Dim strPythonPath As String
    Dim strScriptPath As String
    Dim strCommand As String
    Dim strResult As String

    Dim resultLines As Variant
    Dim i As Integer
    Dim ws As Worksheet
    
    ' 작업할 시트 선택
    Set ws = ThisWorkbook.Sheets("macro") ' 시트명에 맞게 수정
    
    password = "gomsupak12"
    
    ' 셀 보호 해제
    ws.Unprotect password:=password
    
    
    
    
    '초기화
    For i = 7 To 199
        ws.Range("B" & i).Value = ""
    Next i
    
    
    
    For i = 14 To 199
        ws.Range("D" & i).Value = ""
    Next i
    
    For i = 14 To 199
        ws.Range("E" & i).Value = ""
    Next i
    For i = 14 To 199
        ws.Range("F" & i).Value = ""
    Next i
    
    ' 파일 다이얼로그 열기
    filePath = Application.GetOpenFilename("All Files (*.*), *.*")

    ' 사용자가 파일을 선택한 경우
    If filePath <> False Then
        ' 파일 경로 출력
        strHWPFilePath = filePath
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFile = objFSO.GetFile(strHWPFilePath)
        ws.Cells(2, 3).Value = objFile.Name ' 파일 이름만 셀에 넣기
        
        ' 파일 형식에 따라 실행할 파이썬 스크립트 경로 설정
        If Right(strHWPFilePath, 4) = "docx" Then
            ws.Cells(2, 2).Value = "docx"
            strScriptPath = "C:\Users\SOSC근로\Desktop\근로학생\강성철\성성특_목록_자동화_ver2.5.1\codes\DocxToXls.py"
        ElseIf Right(strHWPFilePath, 3) = "doc" Then
            ws.Cells(2, 2).Value = "doc"
            strScriptPath = "C:\Users\SOSC근로\Desktop\근로학생\강성철\성성특_목록_자동화_ver2.5.1\codes\DocToXls.py"
        ElseIf Right(strHWPFilePath, 3) = "hwp" Then
            ws.Cells(2, 2).Value = "hwp"
            strScriptPath = "C:\Users\SOSC근로\Desktop\근로학생\강성철\성성특_목록_자동화_ver2.5.1\codes\HwpToXls.py"
        ElseIf Right(strHWPFilePath, 3) = "pdf" Then
            ws.Cells(2, 2).Value = "pdf"
            strScriptPath = "C:\Users\SOSC근로\Desktop\근로학생\강성철\성성특_목록_자동화_ver2.5.1\codes\PdfToXls2.py"
        Else
            MsgBox "지원하지 않는 파일 형식입니다."
            Exit Sub
        End If
        
    Else
        ' 사용자가 취소를 선택한 경우
        ws.Cells(2, 3).Value = "파일 선택을 취소하셨군요 쌤"
        Exit Sub
    End If
    
    
    
    
    
    
    ' 파이썬 경로 설정
    ' strPythonPath = "C:\\Users\\kangs\\AppData\\Local\\Programs\\Python\\Python311\\python.exe" ' 파이썬 설치 경로에 맞게 수정
    strPythonPath = "C:\Users\SOSC근로\AppData\Local\Microsoft\WindowsApps\PythonSoftwareFoundation.Python.3.7_qbz5n2kfra8p0\python.exe" ' 파이썬 설치 경로에 맞게 수정
    
    
    ' 파이썬 스크립트 경로 설정
    ' strScriptPath = "C:\Users\kangs\Desktop\성성특_목록_자동화\성성특_목록_자동화_ver1.3.1\HwpToXls.py" ' 파이썬 스크립트 경로에 맞게 수정
    ' strScriptPath = "C:\Users\SOSC근로\Desktop\근로학생\강성철\성성특_목록_자동화_ver2.3\HwpToXls.py" ' 파이썬 스크립트 경로에 맞게 수정
    
    
    ' HWP 파일 이름 입력 받기
    ' strHWPFileName = ws.Range("E3").Value ' A1 셀에 HWP 파일 이름이 입력되도록 설정
    ' HWP 파일 경로 생성
    ' strHWPFilePath = "C:\Users\kangs\Desktop\성성특_목록_자동화\서성특_목록_자동화\" & strHWPFileName & ".hwp"
    
    ' 파이썬 스크립트 실행 명령 생성
    strCommand = strPythonPath & " " & strScriptPath & " """ & strHWPFilePath & """"
    
    ' 파이썬 스크립트 실행하여 결과 받아오기
    Set objShell = CreateObject("WScript.Shell")
    strResult = objShell.Exec(strCommand).StdOut.ReadAll

    
    ' 개행문자를 기준으로 분리하여 각 셀에 출력
    resultLines = Split(strResult, vbCrLf)
    For i = LBound(resultLines) To UBound(resultLines)
        ws.Cells(8 + i, 2).Value = resultLines(i)
    Next i
    ''''''
    
    RunYoutubeCheck
    
    CommandButton2_Click

    ' FindDuplicatesAndOutputRows 'commandbutton2에서 시행함
    
    ' 셀 보호 재설정
    ws.Protect password:=password
    
    
End Sub
Sub RunYoutubeCheck()
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
    Set ws = ThisWorkbook.Sheets("macro") ' 시트명에 맞게 수정
    
    password = "gomsupak12"
    
    ' 셀 보호 해제
    ws.Unprotect password:=password
    
    
    
    ' B14부터 시작하여 각 셀에 대해 YoutubeCheck.py 실행
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 14 To lastRow
        ' 유튜브 링크가 있는 셀만 처리
        If ws.Range("B" & i).Value <> "" Then
            ' 파이썬 경로 설정
            strPythonPath = "C:\Users\SOSC근로\AppData\Local\Microsoft\WindowsApps\PythonSoftwareFoundation.Python.3.7_qbz5n2kfra8p0\python.exe" ' 파이썬 설치 경로에 맞게 수정
            
            ' 파이썬 스크립트 경로 설정
            strScriptPath = "C:\Users\SOSC근로\Desktop\근로학생\강성철\성성특_목록_자동화_ver2.5.1\codes\YoutubeCheck.py" ' 파이썬 스크립트 경로에 맞게 수정
            
            ' HWP 파일 경로
            strHWPFilePath = ws.Range("B" & i).Value
            
            ' 파이썬 스크립트 실행 명령 생성
            strCommand = strPythonPath & " " & strScriptPath & " """ & strHWPFilePath & """"
            
            ' 파이썬 스크립트 실행하여 결과 받아오기
            Set objShell = CreateObject("WScript.Shell")
            strResult = objShell.Exec(strCommand).StdOut.ReadAll
            
            ' 결과를 셀에 넣어주기
            ws.Range("B" & i).Value = strResult
        End If
    Next i
    ' 셀 보호 재설정
    ws.Protect password:=password
    
End Sub

Sub FindDuplicatesAndOutputRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim checkRange As Range
    Dim cell As Range
    Dim dict As Object
    Dim duplicates As String
    Dim colorIndex As Long
    
    ' 작업할 시트 선택
    Set ws = ThisWorkbook.Sheets("macro") ' 시트명에 맞게 수정
    
    password = "gomsupak12"
    
    ' 셀 보호 해제
    ws.Unprotect password:=password
    
    
    
    ' 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' 체크할 범위 설정
    Set checkRange = ws.Range("B14:B" & lastRow)
    
    ' 중복된 값 체크를 위한 Dictionary 객체 생성
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 중복된 값들을 저장할 변수 초기화
    duplicates = ""
    
    ' 이전에 적용된 색상 초기화
    ws.Range("A14:A99").Interior.colorIndex = xlNone
    
    ' 다음에 사용할 색상 인덱스 초기화
    colorIndex = 2
    
    ' 중복 여부 확인
    For Each cell In checkRange
        ' 셀이 완전한 공백이 아닌 경우에만 처리
        If Trim(cell.Value) <> "" Then
            If dict.Exists(cell.Value) Then
                ' 중복된 값이 발견되면 해당 행 번호를 문자열에 추가
                duplicates = duplicates & ", " & cell.Row
                ' 중복된 셀에 색상 적용
                cell.Offset(0, -1).Interior.colorIndex = colorIndex
            Else
                ' 중복된 값이 없으면 Dictionary에 추가
                dict(cell.Value) = cell.Row
            End If
            ' 다음 색상으로 변경
            colorIndex = colorIndex + 1
        End If
    Next cell
    
    ' 셀 보호 재설정
    ws.Protect password:=password
    
    
    ' 중복된 값들이 있는지 확인 후 결과 출력
    If Len(duplicates) > 0 Then
        ' 중복된 값들이 있으면 MsgBox를 통해 출력
        MsgBox "중복된 값이 발견되었습니다. 각 행 번호: " & Mid(duplicates, 3) ' 앞의 ", "를 제거하여 출력
    Else
        ' 중복된 값이 없으면 메시지 출력
        MsgBox "중복된 값 없음"
    End If
End Sub




Sub CopyAndPaste(ByVal rowNum As Long)
    Dim ws As Worksheet
    Dim db_ws As Worksheet
    Dim lastRow As Long
    Dim pasteRow As Long
    
    ' 작업할 시트 선택
    Set ws = ThisWorkbook.Sheets("macro") ' 시트명에 맞게 수정
    Set db_ws = ThisWorkbook.Sheets("Student_Database") ' 학생 데이터베이스 시트 선택
    
    password = "gomsupak12"
    
    ' 셀 보호 해제
    ws.Unprotect password:=password
    db_ws.Unprotect password:=password
    
    
    
    ' 복사할 범위 지정
    Dim copyRange As Range
    Set copyRange = ws.Range("G" & rowNum & ":L" & rowNum)
    
    ' 붙여넣을 행 지정 (Student_Database 시트의 마지막 행 + 1)
    pasteRow = db_ws.Cells(db_ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ' 해당 행의 데이터 복사하여 붙여넣기
    copyRange.Copy
    db_ws.Cells(pasteRow, "A").PasteSpecial Paste:=xlPasteValues ' Student_Database 시트에 붙여넣기
    Application.CutCopyMode = False ' 복사 모드 해제
    
    ' 셀 보호 재설정
    ws.Protect password:=password
    db_ws.Protect password:=password
    
    MsgBox ("반영이 완료되었습니다")
End Sub




Private Sub CommandButton11_Click()
    CopyAndPaste 20
    
End Sub

Private Sub CommandButton12_Click()
    CopyAndPaste 21
End Sub

Private Sub CommandButton13_Click()
    CopyAndPaste 22
End Sub

Private Sub CommandButton14_Click()
    CopyAndPaste 23
End Sub

Private Sub CommandButton15_Click() ' 한번에 반영
    Dim ws As Worksheet
    Dim db_ws As Worksheet
    Dim lastRow As Long
    Dim pasteRow As Long
    Dim cell As Range
    
    ' 작업할 시트 선택
    Set ws = ThisWorkbook.Sheets("macro") ' 시트명에 맞게 수정
    Set db_ws = ThisWorkbook.Sheets("Student_Database") ' 학생 데이터베이스 시트 선택
    
    password = "gomsupak12"
    
    ' 셀 보호 해제
    ws.Unprotect password:=password
    db_ws.Unprotect password:=password
    
    
    ' 복사할 범위 지정
    Dim copyRange As Range
    
    ' Student_Database 시트의 마지막 행 찾기
    pasteRow = db_ws.Cells(db_ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ' D열에서 "아직 듣지 않음"인 행만 처리
    For Each cell In ws.Range("D14:D" & ws.Cells(ws.Rows.Count, "D").End(xlUp).Row)
        If cell.Value = "아직 듣지 않음" Then
            ' 복사할 범위 지정
            Set copyRange = ws.Range("G" & cell.Row & ":L" & cell.Row)
            
            ' 해당 행의 데이터 복사하여 붙여넣기
            copyRange.Copy
            db_ws.Cells(pasteRow, "A").PasteSpecial Paste:=xlPasteValues ' Student_Database 시트에 붙여넣기
            pasteRow = pasteRow + 1 ' 다음 행으로 이동
        End If
    Next cell
    
    Application.CutCopyMode = False ' 복사 모드 해제
    
    ' 셀 보호 재설정
    ws.Protect password:=password
    db_ws.Protect password:=password
    
    MsgBox ("반영이 완료되었습니다. 재확인 버튼을 눌러서 확인해주세요")
End Sub

Private Sub CommandButton2_Click()
    Dim ws As Worksheet
    Dim db_ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim studentName As String
    Dim studentID As String
    Dim courseName As String
    Dim isCompleted As Boolean
    Dim completionYear As String
    Dim completionMonth As String
    
    ' 작업할 시트 선택
    Set ws = ThisWorkbook.Sheets("macro") ' 시트명에 맞게 수정
    Set db_ws = ThisWorkbook.Sheets("Student_Database") ' 학생 데이터베이스 시트 선택
    
    password = "gomsupak12"
    
    ' 셀 보호 해제
    ws.Unprotect password:=password
    db_ws.Unprotect password:=password
    
    
    
    For i = 14 To 199
        ws.Range("D" & i).Value = ""
    Next i
    
    For i = 14 To 199
        ws.Range("E" & i).Value = ""
    Next i
    For i = 14 To 199
        ws.Range("F" & i).Value = ""
    Next i
    
    ' 이수 강의명이 입력된 셀 범위
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    
    
    ' 각 이수 강의명에 대해 이수 여부 확인 (강좌 코드로 변경)
    For i = 14 To lastRow ' 이수 강의명이 있는 첫 번째 셀의 행 번호부터 시작
        ' 이수 강의명 및 학번 가져오기
        courseName = ws.Cells(i, "B").Value
        studentName = ws.Range("B10").Value
        studentID = ws.Range("B9").Value
        
        ' 이수 여부 및 이수 년도/월 초기화
        isCompleted = False
        completionYear = ""
        completionMonth = ""
        
        ' Student_Database 시트에서 해당 학생 이름, 학번과 강의명이 일치하는 데이터 찾기
        lastRowDB = db_ws.Cells(db_ws.Rows.Count, "A").End(xlUp).Row
        For j = 2 To lastRowDB
            If db_ws.Cells(j, "B").Value = studentID And InStr(db_ws.Cells(j, "D").Value, courseName) > 0 Then
                ' 해당 학생 학번과 강의명이 일치하는 데이터가 있으면 이수 여부를 True로 설정하고 이수 년도/월 가져오기
                isCompleted = True
                completionYear = db_ws.Cells(j, "E").Value
                completionMonth = db_ws.Cells(j, "F").Value
                Exit For
            End If
        Next j
        
        ' 결과 출력
        If isCompleted Then
            ws.Cells(i, "D").Value = "이미 들음"
            ws.Cells(i, "E").Value = completionYear
            ws.Cells(i, "F").Value = completionMonth
        ElseIf ws.Cells(i, "B").Value = "유효하지 않은 유튜브 링크입니다." Then
            ws.Cells(i, "D").Value = "유효x 링크"
            ws.Cells(i, "E").Value = ""
            ws.Cells(i, "F").Value = ""
        
        Else
            ws.Cells(i, "D").Value = "아직 듣지 않음"
            ws.Cells(i, "E").Value = ""
            ws.Cells(i, "F").Value = ""
        End If
    Next i

    FindDuplicatesAndOutputRows
    
    ' 셀 보호 재설정
    ws.Protect password:=password
    db_ws.Protect password:=password
    
    
End Sub


Private Sub CommandButton3_Click()
    CopyAndPaste 14
End Sub


Private Sub CommandButton4_Click()
    CopyAndPaste 15
End Sub

Private Sub CommandButton5_Click()
    CopyAndPaste 16
End Sub

Private Sub CommandButton6_Click()
    CopyAndPaste 17
End Sub

Private Sub CommandButton7_Click()
    CopyAndPaste 18
End Sub

Private Sub CommandButton8_Click()
    CopyAndPaste 19
End Sub


