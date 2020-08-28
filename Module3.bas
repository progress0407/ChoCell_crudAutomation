Attribute VB_Name = "Module1"
Sub 헬로()
Attribute 헬로.VB_ProcData.VB_Invoke_Func = "M\n14"
    '출력해보자.
    Sheets("test").Range("A1:b2").Value = "hello"
    MsgBox ("Hello World")
End Sub



Sub SquareClink()
    ActiveCell.Offset(0, 0).Value = "현재 셀 내용 입력"
End Sub

Function getValue(cell)
    Dim val
    val = Range(cell, cell).Value
    getValue = val
End Function

Function hasWord(cell, texToFind)
    Dim val
    val = Range(cell, cell).Value
    ' 기본 내장 함수
    idx = InStr(1, val, texToFind, vbTextCompare)
    
    If idx > 0 Then
        hasWord = "O"
    Else
        hasWord = "X"
    End If
End Function
        
