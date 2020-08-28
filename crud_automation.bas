Attribute VB_Name = "Module1"

'ù ���ڷ� ������ �Ǵ�����.. �ް� �Ǵ� �� ������.
' ù��° ���ڴ� �ǵ��� Optional�� ���� ����
' ù��° ���ں��� Variant�� �ʼ��� ������ Integer, String, Range�� ����� ��������
Public Function hasC(src As Variant, Optional c As Variant) As Boolean
    
    Dim str As String
    
    'Integer�� �ƴ� Double�� return�� ��!
    If TypeName(src) = "Integer" Or TypeName(src) = "Double" Or TypeName(src) = "Double" Then
        ' �̰�� src�� row��
        str = Cells(src, c).Value
        
    ElseIf TypeName(src) = "String" Then
        str = src
        
    ElseIf TypeName(src) = "Range" Then
        str = src.Value
        
    Else
        has = "Missing Argument"
        hasC = has
        Exit Function
    
    End If

    hasC = InStr(str, "C") + InStr(str, "c") > 0
   
End Function

Public Function hasR(src As Variant, Optional c As Variant) As Boolean
    
    Dim str As String
    
    If TypeName(src) = "Integer" Or TypeName(src) = "Double" Or TypeName(src) = "Double" Then
        str = Cells(src, c).Value
        
    ElseIf TypeName(src) = "String" Then
        str = src
        
    ElseIf TypeName(src) = "Range" Then
        str = src.Value
        
    Else
        hasR = "Missing Argument"
        Exit Function
    
    End If
    
    hasR = InStr(str, "R") + InStr(str, "r") > 0
    
End Function

Public Function hasU(src As Variant, Optional c As Variant) As Boolean
    
    Dim str As String
    
    If TypeName(src) = "Integer" Or TypeName(src) = "Double" Or TypeName(src) = "Double" Then
        str = Cells(src, c).Value
        
    ElseIf TypeName(src) = "String" Then
        str = src
        
    ElseIf TypeName(src) = "Range" Then
        str = src.Value
        
    Else
        hasU = "Missing Argument"
        Exit Function
    
    End If
    
    hasU = InStr(str, "U") + InStr(str, "u") > 0
    
End Function

Public Function hasD(src As Variant, Optional c As Variant) As Variant
    
    Dim str As String
    
    If TypeName(src) = "Integer" Or TypeName(src) = "Double" Or TypeName(src) = "Double" Then
        str = Cells(src, c).Value
        
    ElseIf TypeName(src) = "String" Then
        str = src
        
    ElseIf TypeName(src) = "Range" Then
        str = src.Value
        
    Else
        hasD = "Missing Argument"
        Exit Function
    
    End If
    
    hasD = InStr(str, "D") + InStr(str, "d") > 0
    
End Function

' C, R, U, D �� ��� �ϳ��� ������ ������ true�� ��ȯ�Ѵ�.
Public Function hasCRUD(src As Variant, Optional c As Variant) As Boolean
    
    Dim str As String

    If TypeName(src) = "Integer" Or TypeName(src) = "Double" Or TypeName(src) = "Double" Then
        str = Cells(src, c).Value
        
    ElseIf TypeName(src) = "String" Then
        str = src
        
    ElseIf TypeName(src) = "Range" Then
        str = src.Value
        
    Else
        hasCRUD = "Missing Argument"
        Exit Function
    
    End If
    
    hasCRUD = hasC(str) Or hasR(str) Or hasU(str) Or hasD(str)
    
End Function

' CRUD ������ �°� ������ִ� �Լ���
Function sortCRUD(toSortStr As String) As String
    Dim resultStr As String: resultStr = ""
    
    If hasC(toSortStr) Then
        resultStr = resultStr + "C"
    End If
    If hasR(toSortStr) Then
        resultStr = resultStr + "R"
    End If
    If hasU(toSortStr) Then
        resultStr = resultStr + "U"
    End If
    If hasD(toSortStr) Then
        resultStr = resultStr + "D"
    End If
    
    sortCRUD = resultStr
    
End Function

Public Function GetColor(rng As Range, Optional return_type As Integer = 0) As Variant
 
    Dim colorVal As Variant
 
    If rng.Columns.Count <= 1 And rng.Rows.Count <= 1 Then
        colorVal = rng.Interior.Color
        GetColor = Hex(colorVal)
    
    Else
            GetColor = CVErr(xlErrValue)
    End If
 
End Function


Public Function numToChar(num As Integer) As String
    num = num - 1
    '��, ������
    Dim quotient, remainder As Integer
    If num < 26 Then
        numToChar = Chr(65 + num)
        Exit Function
    ElseIf num >= 26 Then
        'num+1�� �� �־�� 52�� �Է������� B@���� ��Ȳ�� �߻����� �ʾ�
        quotient = (num) \ 26 - 1
        remainder = (num) Mod 26 + 1
        numToChar = Chr(65 + quotient) & numToChar(remainder)
    End If
    
End Function

' ���� ������ �� ��ǩ���� ��ȯ���ִ� �Լ���
Public Function getAreaCoord(cell As Range) As Object
'    MsgBox ("getAreaCoord ȣ���")
    ' $A$1:$B$3 ó�� ��ȯ�Ǵ� cell������ $�� �����Ѵ�
    Dim addr
    addr = cell.Address
    addr = Replace(addr, "$", "")
    
    Dim colIdx
    colIdx = InStr(addr, ":")
    
    Dim leftAddr, leftCol, leftRow As Integer
    Dim rightAddr, rightCol, rightRow As Integer
    
    ' ���� ���ڷ� ��� �� ���� ������ �ϳ�(A1)��� A1:A1 �� ����� �ش�
    If colIdx = 0 Then
        leftAddr = addr
        rightAddr = addr
'        MsgBox ("���̰� 1�� ���� �������� �ʽ��ϴ�. A1:A10 ���·� �Է¹ٶ��ϴ�")
'        Exit Function
    Else
        leftAddr = Mid(addr, 1, colIdx - 1)
        rightAddr = Mid(addr, colIdx + 1, Len(addr) - colIdx)
    End If
    
    leftRow = Range(leftAddr).row
    leftCol = Range(leftAddr).Column
    
    rightRow = Range(rightAddr).row
    rightCol = Range(rightAddr).Column
    
    ' Pointer���� ���� ��, �ϳ��� Dictionary �ڷ����� ���� ������.
    Dim Coord As Object: Set Coord = CreateObject("Scripting.Dictionary")
        
    Coord.Add "leftRow", leftRow
    Coord.Add "leftCol", leftCol
    Coord.Add "rightRow", rightRow
    Coord.Add "rightCol", rightCol
    
'    For Each k In Coord.keys
'        MsgBox (k & " :  " & Coord.Item(k))
'        Next
    
    Set getAreaCoord = Coord

End Function

' ������ �߿��� CRUD�� �������·� �����Ͽ� �����մϴ� (�� : R, C => CR  , U RC => CRU)
Public Function maxText(cell As Range) As String
    Dim resText As String: resText = ""
    
    Dim Coord As Object: Set Coord = getAreaCoord(cell)
    Dim leftRow, leftCol, rightRow, rightCol As Integer
    
    leftRow = Coord.Item("leftRow")
    leftCol = Coord.Item("leftCol")
    rightRow = Coord.Item("rightRow")
    rightCol = Coord.Item("rightCol")
    
    Dim fixedRow As Integer: fixedRow = leftRow
    
    For c = leftCol To rightCol
        'MsgBox (TypeName(c))
        If Not hasC(resText) And hasC(fixedRow, c) Then
            resText = resText & "C"
        End If

        If Not hasR(resText) And hasR(fixedRow, c) Then
            resText = resText & "R"
        End If

        If Not hasU(resText) And hasU(fixedRow, c) Then
            resText = resText & "U"
        End If

        If Not hasD(resText) And hasD(fixedRow, c) Then
            resText = resText & "D"
        End If
        
        Next c
    
    resText = sortCRUD(resText)
    
    maxText = resText
    
End Function


' �� ���̺� ������ �����ִ� �Լ�
Function getTotTable(cell As Range)
    
    Dim Coord As Object: Set Coord = getAreaCoord(cell)
'    Dim Coord As Object: Set Coord = tDic(cell)
    
    Dim leftRow, leftCol, rightRow, rightCol As Integer
    
    leftRow = Coord.Item("leftRow")
    leftCol = Coord.Item("leftCol")
    rightRow = Coord.Item("rightRow")
    rightCol = Coord.Item("rightCol")

'    MsgBox (": " & Coord.Item("leftRow"))
'    MsgBox ("leftR: " & leftRow)
    
    'CRUD �� ����ϳ��� ������ �ִ°�
    Dim isHas
    
    Dim totCnt As Integer: totCnt = 0
    For r = leftRow To rightRow
        For c = leftCol To rightCol
            If hasCRUD(r, c) Then
                totCnt = totCnt + 1
            End If
            Next c
        Next r
        
    getTotTable = totCnt

End Function

' �÷� ������ ����
'���� ��ġ�� ���� �� �ֱ� ������, ����� Column�� ��ġ�� �����η� ��������
Function getColSum(cell As Range)
    
    Dim Coord As Object: Set Coord = getAreaCoord(cell)
    Dim leftRow, leftCol, rightRow, rightCol As Integer
    
    leftRow = Coord.Item("leftRow")
    leftCol = Coord.Item("leftCol")
    rightRow = Coord.Item("rightRow")
    rightCol = Coord.Item("rightCol")
    
    Dim colSum, r As Integer: colSum = 0
    fixedRow = leftRow
    For c = leftCol To rightCol
        If hasCRUD(fixedRow, c) Then
            colSum = colSum + Cells(1, c).Value
        End If
        Next c

    colSum = Fix(colSum / 2)
    
    getColSum = colSum
    
End Function
' ���� ����ų� ������ �����ϴ� ��츦 ���� �ԾƳ�����
Function hasTrig(cell As Range)
    Dim addr
    addr = cell.Address
    addr = Replace(addr, "$", "")
    
    Dim colIdx
    colIdx = InStr(addr, ":")
    
    If colIdx = 0 Then
        MsgBox ("���̰� 1�� ���� �������� �ʽ��ϴ�. A1:A10 ���·� �Է¹ٶ��ϴ�")
        Exit Function
    End If
    
    Dim leftAddr, leftCol As Integer
    leftAddr = Mid(addr, 1, colIdx - 1)
    
    leftCol = Range(leftAddr).Column
    
    Dim rightAddr, rightCol
    rightAddr = Mid(addr, colIdx + 1, Len(addr) - colIdx)
    rightCol = Range(rightAddr).Column
    
    Dim r As Integer
    r = Range(leftAddr).row
    For c = leftCol To rightCol
        If GetColor(Cells(r, c)) = "FFFF" And (InStr(Cells(r, c).Value, "C") Or InStr(Cells(r, c).Value, "R") Or InStr(Cells(r, c).Value, "U") Or InStr(Cells(r, c).Value, "D")) Then
            hasTrig = "trig"
            Exit Function
        End If
        Next c
    hasTrig = "noTrig"
End Function






Sub addRows()
Attribute addRows.VB_ProcData.VB_Invoke_Func = "m\n14"
    
    Application.ScreenUpdating = False
    
    '�� ȭ���� ������ ã�ƾ� ��
    Dim leftRow, leftCol, rightRow, rightCol As Integer
    
    'Find Start Point : name "progNa"
    Dim flg As Boolean: flg = False
    For r = 1 To (26)
        If flg = True Then
            Exit For
        End If
        For c = 1 To (26)
            If StrComp(Cells(r, c).Value, "progNa") = 0 Then
                leftRow = r
                leftCol = c
                flg = True
                
                'MsgBox ("leftRow : " & leftRow)
                'MsgBox ("leftCol : " & leftCol)
                
                Exit For
            End If
        Next c
    Next r
    
    If flg = False Then
        MsgBox ("���̺� �̸��� 'progNa' �־��ּ���!")
        Exit Sub
    End If
    
    'Find End Point : ������ ��������
    For r = leftRow To (26 * 26 * 26)
        If StrComp(Cells(r, leftCol).Value, "") = 0 Then
            rightRow = r - 1
            MsgBox ("rightRow : " & rightRow)
            Exit For
        End If
    Next r
        
    For c = leftCol To (26 * 26 * 26)
        If StrComp(Cells(leftRow, c).Value, "") = 0 Then
            rightCol = c - 1
            MsgBox ("rightCol : " & rightCol)
            Exit For
        End If
    Next c
    
    ' Header�� �����ϰ� ��������
    leftRow = leftRow + 1

    Dim topicCRUD As Object: Set topicCRUD = CreateObject("Scripting.Dictionary")

    For r = leftRow To rightRow
        'title�� crudLIst �� �ӽ� ����Ҿ�. �ٸ� �뵵�� ��Ȱ�� ������
        Dim title, crudList As String
        title = Cells(r, leftCol).Value
        crudList = maxText(Range(numToChar(leftCol + 1) & r & ":" & numToChar(rightCol) & r))
        topicCRUD(title) = crudList
    Next r

    ' ���ϴ� Row ��
    Dim cursorRow As Integer: cursorRow = leftRow + 1
    Dim titleRow, titleCol As Integer
    
    For Each title In topicCRUD
        
        crudList = topicCRUD(title)
        titleRow = cursorRow
        titleCol = leftRow
        
        If hasC(crudList) Then
            Rows(cursorRow).Insert
            Cells(cursorRow, leftCol).Value = title & " ���"
            cursorRow = cursorRow + 1
        End If
        
        If hasR(crudList) Then
            Rows(cursorRow).Insert
            Cells(cursorRow, leftCol).Value = title & " ��ȸ"
            cursorRow = cursorRow + 1
        End If
        
        If hasU(crudList) Then
            Rows(cursorRow).Insert
            Cells(cursorRow, leftCol).Value = title & " ����"
            cursorRow = cursorRow + 1
        End If
        
        If hasD(crudList) Then
            Rows(cursorRow).Insert
            Cells(cursorRow, leftCol).Value = title & " ����"
            cursorRow = cursorRow + 1
        End If
        
        cursorRow = cursorRow + 1
        
    Next

End Sub



Function searchTwoWord(cell As Range)
    Dim str As String: str = ""
    str = cell.Value
    
    Dim lastTwoWord As String: lastTwoWord = Right(str, 2)
    
    If lastTwoWord = "���" Then
        searchTwoWord = "EI"
        
    ElseIf lastTwoWord = "��ȸ" Then
        searchTwoWord = "Select"
        
    ElseIf lastTwoWord = "����" Then
        searchTwoWord = "Update"
    
    ElseIf lastTwoWord = "����" Then
        searchTwoWord = "Delete"
        
    Else
        searchTwoWord = ""
        
    End If
    
End Function










        
        

        
        

        
        


