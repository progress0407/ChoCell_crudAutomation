Attribute VB_Name = "Module1"
Sub ���()
Attribute ���.VB_ProcData.VB_Invoke_Func = "M\n14"
    '����غ���.
    Sheets("test").Range("A1:b2").Value = "hello"
    MsgBox ("Hello World")
End Sub



Sub SquareClink()
    ActiveCell.Offset(0, 0).Value = "���� �� ���� �Է�"
End Sub

Function getValue(cell)
    Dim val
    val = Range(cell, cell).Value
    getValue = val
End Function

Function hasWord(cell, texToFind)
    Dim val
    val = Range(cell, cell).Value
    ' �⺻ ���� �Լ�
    idx = InStr(1, val, texToFind, vbTextCompare)
    
    If idx > 0 Then
        hasWord = "O"
    Else
        hasWord = "X"
    End If
End Function
        
