Option Explicit


Public Sub ExtractVars()
    
    Dim fnr As Integer
    Dim BASFileName As String
    
    Dim RecData As String
    Dim TRecData As String
    Dim SRecData() As String
    Dim RecDataSaved As String
    
    Dim i As Long
    Dim j As Long
    Dim k As Long, l As Long
    Dim h As Long
    
    Dim idx As Long
    ReDim VarArray(1) As String
    Dim UArray As Variant
    Dim ArrCount As Long
    Dim Flag As Long
    
    Dim strFileToOpen As Variant
    
    Dim KeywordRec(2) As String
    ReDim KeyWord(0) As String
    ReDim KeyWordLen(0) As Long
    Dim ExceptFlag As Boolean
    
    'File Dialog
    strFileToOpen = Application.GetOpenFilename _
    (Title:="Please choose a file to open", _
    FileFilter:="VBA Code Files *.bas (*.bas),")
    
    '선택된 파일(경로 포함)명 전달
    If strFileToOpen <> "" And strFileToOpen <> False Then
        BASFileName = strFileToOpen
    Else
        Exit Sub
    End If
    
    Worksheets("Temp").Range("A:XFD").ClearContents
    Worksheets("Temp2").Range("A:XFD").ClearContents
    
    fnr = FreeFile()
    k = 1
    idx = 0
    
    i = 0
    
    Open ThisWorkbook.Path & "\VBAKeywords.txt" For Input As #fnr
    
    Do While Not EOF(fnr)
        
        ReDim Preserve KeyWord(i)
        ReDim Preserve KeyWordLen(i)
        
        Line Input #fnr, RecData
        
        KeyWord(i) = RecData
        KeyWordLen(i) = Len(KeyWord(i))
        
        i = i + 1
        
    Loop
    
    Close #fnr
    
    i = 0
    
    fnr = FreeFile()
    Open BASFileName For Input As #fnr
    
    Do While Not EOF(fnr)
        
        Line Input #fnr, RecData
        
        '명령어 뒷부분(오른쪽)에 위치한 주석 내용을 삭제 처리 - 단, 큰 따옴표 사이에 들어가는 작은 따옴표의 경우, 주석이 아니기 때문에 예외 처리
        If InStr(RecData, "'") > 0 Then
            If InStr(RecData, """") > 0 And InStr(RecData, "'") > InStr(RecData, """") Then
                TRecData = RecData
                Do While InStr(TRecData, """") > 0
                    TRecData = Right(TRecData, Len(TRecData) - InStr(InStr(TRecData, """") + 1, TRecData, """"))
                Loop
                If InStr(TRecData, "'") > 0 Then
                    RecData = Replace(RecData, Right(TRecData, Len(TRecData) - InStr(TRecData, "'") + 1), "")
                End If
            Else
                RecData = Left(RecData, InStr(RecData, "'") - 1)
            End If
        End If
        
        '문장 내에서 큰 따옴표에 해당하는 부분을 모두 삭제 처리
        Do While InStr(RecData, """") > 0
            RecData = Left(RecData, InStr(RecData, """") - 1) & Right(RecData, Len(RecData) - InStr(InStr(RecData, """") + 1, RecData, """"))
        Loop
        
        'tab 문자가 있는 경우 trim 처리가 의도한 대로 되지 않기 때문에 tab 문자를 모두 제거
        RecData = Replace(RecData, vbTab, "")
        
        '문장의 양 옆에 있는 blank 를 모두 제거
        RecData = Trim(RecData)
        
        'trim 처리 후 문장 끝에 underline 이 있는 경우에 다음 처리로 넘김
        If Right(RecData, 1) = "_" Then
            RecDataSaved = RecDataSaved & Left(RecData, Len(RecData) - 1)
        Else
            RecData = RecDataSaved & RecData

            i = 2
            
            '======================================================================================================================================================
            '* 배열 파라미터 체크 시작
            '======================================================================================================================================================
            '열린 괄호 있을 때만 수행
            If InStr(RecData, "(") > 0 Then
                
                '진행 문자 위치 이후에 열린 괄호가 있는 동안 계속 수행
                Do While InStr(i, RecData, "(") > 1
                    
                    '열린 괄호 앞의 문자가 공백이 아닌 경우, 즉, 배열(하지만 함수나 프로시저가 될 수도 있음)로 판단되는 경우에 수행
                    If Mid(RecData, i, 1) = "(" And Mid(RecData, i - 1, 1) <> " " Then
                        
                        ArrCount = 0
                        Flag = 1
                        
                        '열린 괄호 다음 위치에서 처음으로 콤마가 나오는 위치까지 반복
                        j = i + 1
                        Do While Flag > 0
                            If Mid(RecData, j, 1) = "(" Then
                                Flag = Flag + 1
                            ElseIf Mid(RecData, j, 1) = ")" Then
                                Flag = Flag - 1
                            ElseIf Mid(RecData, j, 1) = "," And Flag = 1 Then
                                ArrCount = ArrCount + 1
                            End If
                            j = j + 1
                        Loop
                        
                        'MsgBox ArrCount + 1 & "개의 인자를 가지고 있습니다."
                        RecData = Left(RecData, i - 1) & "_" & CStr(ArrCount + 1) & "차원배열 " & Right(RecData, Len(RecData) - i)
                        
                    End If
                    
                    i = i + 1
                    
                Loop
                
            End If
            '======================================================================================================================================================
            '* 배열 파라미터 체크 끝
            '======================================================================================================================================================
            
            TRecData = RecData
        
            '문장의 맨 앞에 올 수 있는 명령어 기준으로 파싱 대상 제외
            If TRecData <> "" And _
            Left(TRecData, 1) <> "'" And _
            Left(TRecData, 9) <> "Attribute" And _
            Left(TRecData, 6) <> "Option" And _
            Left(TRecData, 3) <> "Sub" And _
            Left(TRecData, 6) <> "Public" And _
            Left(TRecData, 4) <> "Call" Then
                
                i = i + 1
                
                '콤마 기준 구분
                TRecData = Replace(TRecData, ", ", " ")
                
                '한 줄에 두 개 이상 명령어 입력된 것 구분
                TRecData = Replace(TRecData, ":", "")
                
                '큰 따옴표 공백으로 치환 -> 큰 따옴표 내의 내용은 모두 삭제 처리 필요
                'TRecData = Replace(TRecData, """", "")
                
                '열린 괄호 앞에 공백이 있는 경우는 배열 괄호가 아니라서 공백으로 치환
                TRecData = Replace(TRecData, " (", " ")
                
                '닫힌 괄호는 삭제
                TRecData = Replace(TRecData, ")", "")
                
                '남은 열린 괄호는 배열 괄호로 간주하여 공백으로 치환 - 이렇게 처리하면 배열 차원 구분 불가
                'TRecData = Replace(TRecData, "(", "[Array] ")
                
                '배열 선언에서 사용되는 "To" 부분을 공백으로 치환
                TRecData = Replace(TRecData, " To ", " ")
                
                'TRecData = Replace(TRecData, " ", "")
                
                '공백이 두 개 이상 존재하는 경우 하나로 치환
                Do While InStr(TRecData, "  ") > 0
                    TRecData = Replace(TRecData, "  ", " ")
                Loop
                
                '다시 한번 양쪽 끝 공백 제거
                TRecData = Trim(TRecData)
                
                '공백 기준으로 문자열 분해
                SRecData = Split(TRecData, " ")
                
                j = 0
                l = 0
                
                Do While j <= UBound(SRecData)
                    
                    ExceptFlag = False
                    j = j + 1
                    
                    For h = 0 To UBound(KeyWord)
                        If UCase(Left(SRecData(j - 1), KeyWordLen(h))) = UCase(KeyWord(h)) Then
                            ExceptFlag = True
                        End If
                    Next h
                    
                    If ExceptFlag = False And _
                    ((Asc(Left(SRecData(j - 1), 1)) >= 65 And Asc(Left(SRecData(j - 1), 1)) <= 90) Or (Asc(Left(SRecData(j - 1), 1)) >= 97 And Asc(Left(SRecData(j - 1), 1)) <= 122)) Then
                        l = l + 1
                        Worksheets("Temp").Cells(k, l) = SRecData(j - 1)
                        ReDim Preserve VarArray(idx + 1)
                        VarArray(idx) = SRecData(j - 1)
                        idx = idx + 1
                    End If
                    
                Loop
                
                If l > 0 Then k = k + 1
                
            End If
            RecDataSaved = ""
        End If
        
    Loop
    
    Close #fnr
    
    For i = 1 To UBound(VarArray)
        Worksheets("Temp2").Cells(i, 1) = VarArray(i - 1)
    Next i
    
    UArray = ArrayUnique(VarArray)
    
    For i = 1 To UBound(UArray)
        Worksheets("Temp2").Cells(i, 2) = UArray(i - 1)
    Next i
    
End Sub
