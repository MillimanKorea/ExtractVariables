Attribute VB_Name = "M02_UDF"
Option Explicit



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SplitMultiDelims by alainbryden
' This function splits Text into an array of substrings, each substring
' delimited by any character in DelimChars. Only a single character
' may be a delimiter between two substrings, but DelimChars may
' contain any number of delimiter characters. It returns a single element
' array containing all of text if DelimChars is empty, or a 1 or greater
' element array if the Text is successfully split into substrings.
' If IgnoreConsecutiveDelimiters is true, empty array elements will not occur.
' If Limit greater than 0, the function will only split Text into 'Limit'
' array elements or less. The last element will contain the rest of Text.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function SplitMultiDelims(ByRef TEXT As String, ByRef DelimChars As String, _
    Optional ByVal IgnoreConsecutiveDelimiters As Boolean = False, _
    Optional ByVal Limit As Long = -1) As String()
    Dim ElemStart As Long, N As Long, M As Long, Elements As Long
    Dim lDelims As Long, lText As Long
    Dim Arr() As String
    Dim index_start As Boolean, index_start2 As Boolean
    lText = Len(TEXT)
    lDelims = Len(DelimChars)
    If lDelims = 0 Or lText = 0 Or Limit = 1 Then
        ReDim Arr(0 To 0)
        Arr(0) = TEXT
        SplitMultiDelims = Arr
        Exit Function
    End If
    ReDim Arr(0 To IIf(Limit = -1, lText - 1, Limit))
    
    Elements = 0: ElemStart = 1: index_start = False: index_start2 = False:
    For N = 1 To lText
        If N > 1 Then
            If Mid(TEXT, N, 1) = "(" Then
                If index_start = True Then
                    index_start2 = True
                ElseIf Mid(TEXT, N - 1, 1) <> " " Then
                    index_start = True
                End If
            End If
        End If
        
        
        If index_start2 = True And Mid(TEXT, N, 1) = ")" Then
            index_start2 = False
            GoTo p
        ElseIf index_start = True And Mid(TEXT, N, 1) = ")" Then
            index_start = False
            GoTo p
        End If
        
        If index_start = True Then GoTo p
        
        If InStr(DelimChars, Mid(TEXT, N, 1)) Then
            Arr(Elements) = Mid(TEXT, ElemStart, N - ElemStart)
            If IgnoreConsecutiveDelimiters Then
                If Len(Arr(Elements)) > 0 Then Elements = Elements + 1
            Else
                Elements = Elements + 1
            End If
            ElemStart = N + 1
            If Elements + 1 = Limit Then Exit For
        End If
p:
    Next N
    'Get the last token terminated by the end of the string into the array
    If ElemStart <= lText Then Arr(Elements) = Mid(TEXT, ElemStart)
    'Since the end of string counts as the terminating delimiter, if the last character
    'was also a delimiter, we treat the two as consecutive, and so ignore the last elemnent
    If IgnoreConsecutiveDelimiters Then If Len(Arr(Elements)) = 0 Then Elements = Elements - 1
    
    ReDim Preserve Arr(0 To Elements) 'Chop off unused array elements
    SplitMultiDelims = Arr
End Function



' 배열 내의 중복된 내용을 제거 (Remove duplicate items)
Function ArrayUnique(ByVal aArrayIn As Variant) As Variant
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ArrayUnique
    ' This function removes duplicated values from a single dimension array
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim aArrayOut() As Variant
    Dim bFlag As Boolean
    Dim vIn As Variant
    Dim vOut As Variant
    Dim i%, j%, k%
    
    ReDim aArrayOut(LBound(aArrayIn) To UBound(aArrayIn))
    i = LBound(aArrayIn)
    j = i
    
    For Each vIn In aArrayIn
        For k = j To i - 1
            If vIn = aArrayOut(k) Then bFlag = True: Exit For
        Next
        
        If Not bFlag Then aArrayOut(i) = vIn: i = i + 1
        bFlag = False
    Next
    
    If i <> UBound(aArrayIn) Then ReDim Preserve aArrayOut(LBound(aArrayIn) To i - 1)
    ArrayUnique = aArrayOut
    
End Function



' 배열 내의 빈 요소를 제거 (Remove empty items)
Function Clean_Array(Arr As Variant) As Variant
    
    Dim i As Integer, j As Integer
    
    ReDim NewArr(LBound(Arr) To UBound(Arr))
    
    For i = LBound(Arr) To UBound(Arr)
        If Arr(i) <> "" And IsNumeric(Arr(i)) = False Then
            NewArr(j) = Trim(Arr(i))
            j = j + 1
        End If
    Next i
    
    ReDim Preserve NewArr(LBound(Arr) To j - 1)
    
    Clean_Array = NewArr
    
End Function




Public Function MyMin(ByVal a As Double, ByVal b As Double) As Double
    MyMin = a
    If a > b Then MyMin = b
End Function
