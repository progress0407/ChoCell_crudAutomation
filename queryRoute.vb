Public Function numToChar(num) As String
    num = CInt(num) - 1
    '몫, 나머지
    Dim quotient As Integer, remainder As Integer
    If num < 26 Then
        numToChar = Chr(65 + num)
        Exit Function
    ElseIf num >= 26 Then
        'num+1을 해 주어야 52를 입력했을시 B@같은 상황이 발생하지 않아
        quotient = (num) \ 26
        remainder = (num) Mod 26
        numToChar = numToChar(quotient) & Chr(65 + remainder)
    End If
    
End Function

Public Function getColor(src As Variant As Variant, Optional row As Variant) As Variant
 
    Dim colorVal As Variant
    Dim rng As Range

    'Range(numToChar(leftCol + 1) & r & ":" & numToChar(rightCol) & r)
    IF TypeName(src) = "Integer" Or TypeName(src) = "Double" Or TypeName(src) = "Long" Then
        rng = Range(numToChar() & r &":"& numToChar() & r)

    ElseIf TypeName(src) = "Range" Then
        rng = src;

    Else
        getColor = "Wrong Argument"
        Exit Function

    End If


    If rng.Columns.Count <= 1 And rng.Rows.Count <= 1 Then
        colorVal = rng.Interior.Color
        getColor = Hex(colorVal)
    
    Else
            getColor = CVErr(xlErrValue)
    End If
 
End Function


Public Function getTableDic()
    Dim tableDicList As Object
    Set tableDicList = CreateObject("Scripting.Dictionary")
    Dim title As String
    Dim desc As String
    Dim tableRow, tableCol

    Dim flg As Boolean: flg = False
    For r = 1 To 10
        For c = 1 To 5
            If getColor(r, c) = "FFFF" Then
                tableRow = r
                tableCol = c
                flg = True
                Exit For
            End If
        Next c

        If flg = True Then
                Exit For
        End If    
    Next r

    flg = False

    title = Cells(tableRow, tableCol).Value
    desc = Cells(tableRow, tableCol + 1).Value

    Dim tempObj As Object
    Set tempObj As CreateObject("Scripting.Dictionary")

    tempObj.Add "desc", desc

    Dim r, c
    r = leftRow
    c = leftCol
    Do while(StrCmp(Cells(r, c).Value, "") = 0)
        tempObj.Add Cells(r, c).Value, Cells(r, c + 1).Value
        r = r + 1
    Loop

    tableDicList.Add title, tempObj
    tempObj = Nothing

End Function
