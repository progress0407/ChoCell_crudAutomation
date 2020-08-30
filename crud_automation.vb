'첫 인자로 무엇이 되던간에.. 받게 되는 게 문제야.
' 첫번째 인자는 되도록 Optional을 쓰지 말자
' 첫번째 인자부터 Variant로 필수로 받은뒤 Integer, String, Range로 나누어서 생각하자
Public Function hasC(src As Variant, Optional c As Variant) As Boolean
    
    Dim str As String
    
    'Integer가 아닌 Double로 return이 돼!
    If TypeName(src) = "Integer" Or TypeName(src) = "Double" Or TypeName(src) = "Long" Then
        ' 이경우 src가 row야
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

' C, R, U, D 중 어느 하나라도 가지고 있으면 true를 반환한다.
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

' CRUD 순서에 맞게 만들어주는 함수야
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

' 셀의 범위를 각 좌푯값을 반환해주는 함수야
Public Function getAreaCoord(cell As Range) As Object
    ' $A$1:$B$3 처럼 반환되는 cell값에서 $를 제외한다
    Dim addr
    addr = cell.Address
    addr = Replace(addr, "$", "")
    
    Dim colIdx
    colIdx = InStr(addr, ":")
    
    Dim leftAddr, leftCol, leftRow As Integer
    Dim rightAddr, rightCol, rightRow As Integer
    
    ' 만일 인자로 들어 온 셀이 오로지 하나(A1)라면 A1:A1 로 만들어 준다
    If colIdx = 0 Then
        leftAddr = addr
        rightAddr = addr
    Else
        leftAddr = Mid(addr, 1, colIdx - 1)
        rightAddr = Mid(addr, colIdx + 1, Len(addr) - colIdx)
    End If
    
    leftRow = Range(leftAddr).row
    leftCol = Range(leftAddr).Column
    
    rightRow = Range(rightAddr).row
    rightCol = Range(rightAddr).Column
    
    ' Pointer변수 같은 애, 하나의 Dictionary 자료형을 만들어서 참조해.
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

' 셀값들 중에서 CRUD를 집합형태로 수집하여 리턴합니다 (예 : R, C => CR  , U RC => CRU)
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


' 총 테이블 갯수를 구해주는 함수
Function getTotTable(cell As Range, Optional selectCRUD As String)
    
    Dim Coord As Object: Set Coord = getAreaCoord(cell)
'    Dim Coord As Object: Set Coord = tDic(cell)
    
    Dim leftRow, leftCol, rightRow, rightCol As Integer
    
    leftRow = Coord.Item("leftRow")
    leftCol = Coord.Item("leftCol")
    rightRow = Coord.Item("rightRow")
    rightCol = Coord.Item("rightCol")

    Dim totCnt As Integer: totCnt = 0
    If StrComp(selectCRUD, "") = 0 Or StrComp(selectCRUD, "CRUD") = 0 Or IsEmpty(selectCRUD) Then
        For r = leftRow To rightRow
            For c = leftCol To rightCol
                If hasCRUD(r, c) Then
                    totCnt = totCnt + 1
                End If
            Next c
        Next r

    ElseIf StrComp(selectCRUD, "C") = 0 Then
        For r = leftRow To rightRow
            For c = leftCol To rightCol
                If hasC(r, c) Then
                    totCnt = totCnt + 1
                End If
            Next c
        Next r

    ElseIf StrComp(selectCRUD, "R") = 0 Then
        For r = leftRow To rightRow
            For c = leftCol To rightCol
                If hasR(r, c) Then
                    totCnt = totCnt + 1
                End If
            Next c
        Next r

    ElseIf StrComp(selectCRUD, "U") = 0 Then
        For r = leftRow To rightRow
            For c = leftCol To rightCol
                If hasU(r, c) Then
                    totCnt = totCnt + 1
                End If
            Next c
        Next r

    ElseIf StrComp(selectCRUD, "D") = 0 Then
        For r = leftRow To rightRow
            For c = leftCol To rightCol
                If hasD(r, c) Then
                    totCnt = totCnt + 1
                End If
            Next c
        Next r

    End If
        
    getTotTable = totCnt

End Function

' 컬럼 갯수의 총합
'열의 위치는 변할 수 있기 때문에, 상단의 Column의 위치는 절대경로로 설정하자
Function getColSum(cell As Range, Optional selectCRUD As String)
    
    Dim Coord As Object: Set Coord = getAreaCoord(cell)
    Dim leftRow, leftCol, rightRow, rightCol As Integer
    
    leftRow = Coord.Item("leftRow")
    leftCol = Coord.Item("leftCol")
    rightRow = Coord.Item("rightRow")
    rightCol = Coord.Item("rightCol")
    
    Dim colSum, r As Integer: colSum = 0
    Dim fixedRow As Integer: fixedRow = leftRow

    If StrComp(selectCRUD, "") = 0 Or StrComp(selectCRUD, "CRUD") = 0 Or IsEmpty(selectCRUD) Then
        For c = leftCol To rightCol
            If hasCRUD(fixedRow, c) Then
                colSum = colSum + Cells(1, c).Value
            End If
        Next c

    ElseIf StrComp(selectCRUD, "C") = 0 Then
        For c = leftCol To rightCol
            If hasC(fixedRow, c) Then
                colSum = colSum + Cells(1, c).Value
            End If
        Next c

    ElseIf StrComp(selectCRUD, "R") = 0 Then
        For c = leftCol To rightCol
            If hasR(fixedRow, c) Then
                colSum = colSum + Cells(1, c).Value
            End If
        Next c
    
    ElseIf StrComp(selectCRUD, "U") = 0 Then
        For c = leftCol To rightCol
            If hasU(fixedRow, c) Then
                colSum = colSum + Cells(1, c).Value
            End If
        Next c

    ElseIf StrComp(selectCRUD, "D") = 0 Then
        For c = leftCol To rightCol
            If hasD(fixedRow, c) Then
                colSum = colSum + Cells(1, c).Value
            End If
        Next c
    
    Else
        MsgBox ("getColSum Error: CRUD 값 중 하나만 입력하셔야 합니다! ")

    End If

    colSum = Fix(colSum / 2)
    getColSum = colSum
    
End Function


' 색이 노랗거나 셀값이 존재하는 경우를 따로 솎아내야해
Function hasTrig(cell As Range)
    
    Dim Coord As Object: Set Coord = getAreaCoord(cell)
    Dim leftRow, leftCol, rightRow, rightCol As Integer
    
    leftRow = Coord.Item("leftRow")
    leftCol = Coord.Item("leftCol")
    rightRow = Coord.Item("rightRow")
    rightCol = Coord.Item("rightCol")
    
    Dim fixedRow As Integer: fixedRow = leftRow
    
    For c = leftCol To rightCol
        If GetColor(Cells(fixedRow, c)) = "FFFF" And hasCRUD(fixedRow, c) Then
            hasTrig = "trig"
            Exit Function
        End If
        Next c
    hasTrig = "noTrig"
End Function


Sub addRows()
    
    Application.ScreenUpdating = False
    
    '이 화면의 영역을 찾아야 해
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
                
                Exit For
            End If
        Next c
    Next r
    
    If flg = False Then
        MsgBox ("테이블 이름에 'progNa' 넣어주세요!")
        Exit Sub
    End If
    
    'Find End Point : 시작점 기준으로
    For r = leftRow To (26 * 26 * 26)
        If StrComp(Cells(r, leftCol).Value, "") = 0 Then
            rightRow = r - 1
            Exit For
        End If
    Next r
        
    For c = leftCol To (26 * 26 * 26)
        If StrComp(Cells(leftRow, c).Value, "") = 0 Then
            rightCol = c - 1
            Exit For
        End If
    Next c
    
    ' Header를 제외하고 생각하자
    leftRow = leftRow + 1

    Dim topicCRUD As Object: Set topicCRUD = CreateObject("Scripting.Dictionary")

    For r = leftRow To rightRow
        'title과 crudLIst 는 임시 저장소야. 다른 용도로 재활용 가능해
        Dim title, crudList As String
        title = Cells(r, leftCol).Value
        crudList = maxText(Range(numToChar(leftCol + 1) & r & ":" & numToChar(rightCol) & r))
        topicCRUD(title) = crudList
    Next r

    ' 변하는 Row 값
    Dim cursorRow As Integer: cursorRow = leftRow + 1
    Dim titleRow, titleCol As Integer
    
    For Each title In topicCRUD
        
        crudList = topicCRUD(title)
        titleRow = cursorRow
        titleCol = leftRow
        
        If hasC(crudList) Then
            Rows(cursorRow).Insert
            Cells(cursorRow, leftCol).Value = title & " 등록"
            cursorRow = cursorRow + 1
        End If
        
        If hasR(crudList) Then
            Rows(cursorRow).Insert
            Cells(cursorRow, leftCol).Value = title & " 조회"
            cursorRow = cursorRow + 1
        End If
        
        If hasU(crudList) Then
            Rows(cursorRow).Insert
            Cells(cursorRow, leftCol).Value = title & " 수정"
            cursorRow = cursorRow + 1
        End If
        
        If hasD(crudList) Then
            Rows(cursorRow).Insert
            Cells(cursorRow, leftCol).Value = title & " 삭제"
            cursorRow = cursorRow + 1
        End If
        
        cursorRow = cursorRow + 1
        
    Next

End Sub


Function replaceWord(cell As Range)
    Dim str As String: str = ""
    str = cell.Value
    
    Dim lastWord As String: lastWord = Right(str, 3)
    
    If lastWord = " 등록" Then
        replaceWord = "Insert"
        
    ElseIf lastWord = " 조회" Then
        replaceWord = "Select"
        
    ElseIf lastWord = " 수정" Then
        replaceWord = "Update"
    
    ElseIf lastWord = " 삭제" Then
        replaceWord = "Delete"
        
    Else
        replaceWord = ""
        
    End If
    
End Function


Sub addRowsAdvanced()
    
    Application.ScreenUpdating = False
    
    '이 화면의 영역을 찾아야 해
    Dim leftRow, leftCol, rightRow, rightCol As Integer
    
    'Find Start Point : name "성우쿤"
    ' 플래그 변수 flg는 2중 for문을 벗어나기 위한 변수야.
    Dim flg As Boolean: flg = False
    For r = 1 To (26)
        If flg = True Then
            Exit For
        End If
        For c = 1 To (26)
            '성우쿤을 찾는 다면, row, col 좌표를 저장 해줘
            If StrComp(Cells(r, c).Value, "성우쿤") = 0 Then
                leftRow = r
                leftCol = c
                flg = True
                Exit For
            End If
        Next c
    Next r
    
    If flg = False Then
        MsgBox ("테이블 이름에 '성우쿤' 넣어주세요!")
        MsgBox ("간장치킨 최고")
        Exit Sub
    End If
    
    'Find End Point : 끝점을 찾자
    For r = leftRow To (26 * 26 * 26)
        If StrComp(Cells(r, leftCol).Value, "") = 0 Then
            rightRow = r - 1
            Exit For
        End If
    Next r
        
    For c = leftCol To (26 * 26 * 26)
        If StrComp(Cells(leftRow, c).Value, "") = 0 Then
            rightCol = c - 1
            Exit For
        End If
    Next c
    
    ' Header를 제외하고 생각하자
    leftRow = leftRow + 1

    ' [make DICTIONARY]
    ' row별 maxText를 수집하여 Dictionary 형태의 정보로 저장한다.
    Dim topicCRUD As Object
    Set topicCRUD = Nothing
    Set topicCRUD = CreateObject("Scripting.Dictionary")

    Dim title As Variant, crudList As Variant, rngStr As Variant

    Dim fixedLeftCol: fixedLeftCol = numToChar(leftCol + 1)
    Dim fixedRightCol: fixedRightCol = numToChar(rightCol)

    Dim isAlreadyHave As Boolean: isAlreadyHave = False

    Dim titleStr As String
    ' 중복이 된 것이 있다면 +1씩 해줄 것이야.
    Dim reDuple As Integer: reDuple = 1

    'table의 갯수, col의 합 넣기
    For r = leftRow To rightRow
        rngStr = fixedLeftCol & r & ":" & fixedRightCol & r

        ' 첫 행이면 그냥 넣자.
        If (r = leftRow) Then
            title = Cells(r, leftCol).Value
        Else
            For Each checkTitle In topicCRUD
                titleStr = CStr(checkTitle)
                '만일 중복이 있다면
                If StrComp(CStr(Cells(r, leftCol).Value), CStr(titleStr)) = 0 Then
                    title = Cells(r, leftCol).Value & "(" & reDuple & ")"
                    reDuple = reDuple + 1
                    Exit For
                Else
                    title = Cells(r, leftCol).Value
                End If
            Next
        End If

        '하위 객체(crudList, getTot, getColSum)들을 담을 임시 객체를 선언한다.
        Dim tempObj As Object
        Set tempObj = Nothing
        Set tempObj = CreateObject("Scripting.Dictionary")

        ' CRUD 정보를 받을 변수 선언 및 임시 객체에 할당
        crudList = maxText(Range(rngStr))
        tempObj.Add "crudList", crudList

        Dim totTable As Object
        Set totTable = Nothing
        Set totTable = CreateObject("Scripting.Dictionary")

        totTable.Add "C", getTotTable(Range(rngStr), "C")
        totTable.Add "R", getTotTable(Range(rngStr), "R")
        totTable.Add "U", getTotTable(Range(rngStr), "U")
        totTable.Add "D", getTotTable(Range(rngStr), "D")

        '임시 객체에 할당한 후 기존  totTable객체를 해제 해야 재 할당이 가능하다.
        tempObj.Add "totTable", totTable
        Set totTable = Nothing

        Dim colSum As Object
        Set colSum = Nothing
        Set colSum = CreateObject("Scripting.Dictionary")
        
        colSum.Add "C", getColSum(Range(rngStr), "C")
        colSum.Add "R", getColSum(Range(rngStr), "R")
        colSum.Add "U", getColSum(Range(rngStr), "U")
        colSum.Add "D", getColSum(Range(rngStr), "D")

        tempObj.Add "colSum", colSum
        Set colNum = Nothing

        topicCRUD.Add title, tempObj
        Set tempObj = Nothing

    Next r


    ' [만든 정보를 바탕으로 행 삽입하기]
    Dim cursorRow As Integer: cursorRow = leftRow
    Dim titleRow As Integer, titleCol As Integer

    '구할 값에 대한 열 삽입하고 관심의 대상을 옮겨야 해. (해당 열 만큼)
    Columns(leftCol).Insert
    Columns(leftCol).Insert

    Cells(leftRow - 1, leftCol).Value = "getTotTable"
    Cells(leftRow - 1, leftCol + 1).Value = "getColSum"

    ' column을 2개 추가해서 관심의 대상은 +2 가 되었어.
    leftCol = leftCol + 2

    'title을 넣었더니.. 조회가 안됀다..

    Dim swCho As Integer: swCho = leftCol + 1
    For Each title In topicCRUD
    
        ' title이라는 variant를 str로 형변환해주지 않으면 만든 자료형에서 읽기가 안돼.
        titleStr = CStr(title)
        crudList = topicCRUD(title)("crudList")
        Rows(cursorRow).Delete
    
        If hasC(crudList) Then
            title = Replace(title, "등록", "")
            Rows(cursorRow).Insert
            Columns(cursorRow).Interior.Color = RGB(255, 255, 255)
    
            Cells(cursorRow, leftCol).Value = title & " 등록"
            Cells(cursorRow, leftCol).Interior.Color = RGB(255, 215, 215)
    
            Cells(cursorRow, leftCol - 2).Value = topicCRUD(titleStr)("totTable")("C")
            Cells(cursorRow, leftCol - 2).Interior.Color = RGB(255, 240, 240)
    
            Cells(cursorRow, leftCol - 1).Value = topicCRUD(titleStr)("colSum")("C")
            Cells(cursorRow, leftCol - 1).Interior.Color = RGB(255, 240, 240)
    
            Cells(cursorRow, swCho).Value = "47"
            Cells(cursorRow, swCho).Interior.Color = RGB(245, 245, 245)
    
            swCho = swCho + 1
            cursorRow = cursorRow + 1
        End If
    
        If hasR(crudList) Then
            title = Replace(title, "조회", "")
            Rows(cursorRow).Insert
            Columns(cursorRow).Interior.Color = RGB(255, 255, 255)
    
            Cells(cursorRow, leftCol).Value = title & " 조회"
            Cells(cursorRow, leftCol).Interior.Color = RGB(255, 255, 175)
    
            Cells(cursorRow, leftCol - 2).Value = topicCRUD(titleStr)("totTable")("R")
            Cells(cursorRow, leftCol - 2).Interior.Color = RGB(255, 255, 220)
    
            Cells(cursorRow, leftCol - 1).Value = topicCRUD(titleStr)("colSum")("R")
            Cells(cursorRow, leftCol - 1).Interior.Color = RGB(255, 255, 220)
    
            Cells(cursorRow, swCho).Value = "47"
            Cells(cursorRow, swCho).Interior.Color = RGB(245, 245, 245)
    
            swCho = swCho + 1
            cursorRow = cursorRow + 1
        End If
        
        If hasU(crudList) Then
            title = Replace(title, "수정", "")
            Rows(cursorRow).Insert
            Columns(cursorRow).Interior.Color = RGB(255, 255, 255)
    
            Cells(cursorRow, leftCol).Value = title & " 수정"
            Cells(cursorRow, leftCol).Interior.Color = RGB(215, 255, 215)
    
            Cells(cursorRow, leftCol - 2).Value = topicCRUD(titleStr)("totTable")("U")
            Cells(cursorRow, leftCol - 2).Interior.Color = RGB(240, 255, 240)
    
            Cells(cursorRow, leftCol - 1).Value = topicCRUD(titleStr)("colSum")("U")
            Cells(cursorRow, leftCol - 1).Interior.Color = RGB(240, 255, 240)
    
            Cells(cursorRow, swCho).Value = "47"
            Cells(cursorRow, swCho).Interior.Color = RGB(245, 245, 245)
    
            swCho = swCho + 1
            cursorRow = cursorRow + 1
        End If
    
        If hasD(crudList) Then
            title = Replace(title, "삭제", "")
            Rows(cursorRow).Insert
            Columns(cursorRow).Interior.Color = RGB(255, 255, 255)
    
            Cells(cursorRow, leftCol).Value = title & " 삭제"
            Cells(cursorRow, leftCol).Interior.Color = RGB(215, 215, 255)
    
            Cells(cursorRow, leftCol - 2).Value = topicCRUD(titleStr)("totTable")("D")
            Cells(cursorRow, leftCol - 2).Interior.Color = RGB(240, 240, 255)
    
            Cells(cursorRow, leftCol - 1).Value = topicCRUD(titleStr)("colSum")("D")
            Cells(cursorRow, leftCol - 1).Interior.Color = RGB(240, 240, 255)
    
            Cells(cursorRow, swCho).Value = "47"
            Cells(cursorRow, swCho).Interior.Color = RGB(245, 245, 245)
    
            swCho = swCho + 1
            cursorRow = cursorRow + 1
        End If
    Next

    ' [TestPage 만들기]
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets.Add(After:= _
             ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    Ws.Name = "testSwCho"

    ' 밑에 중에 취사 선택할 것
    Ws.Select
    'Ws.Cells(1, 1).Value = "hi"
    
    Cells(1, 1).Value = "성우쿤"
    Cells(1, 2).Value = "maxText"
    Cells(1, 1).Value = "Length"
    Cells(1, 1).Value = "tot Summation"

    Dim row As Integer: row = 1
    Dim red As Integer, green As Integer, blue As Integer
    
    For Each title In topicCRUD
        row = row + 1

        titleStr = CStr(title)
        crudList = topicCRUD(title)("crudList")
    
        Cells(row, 1).Value = titleStr
        Cells(row, 2).Value = crudList
        Cells(row, 3).Formula = "=Len(B" & row & ")"

        red = 0
        green = 0
        blue = 0
        
        If hasC(crudList) Then
            green = 15
            blue = 15

        ElseIf hasR(crudList) Then
            blue = 30

        ElseIf hasU(crudList) Then
            red = 15
            blue = 15

        ElseIf hasD(crudList) Then
            red = 15
            green = 15
        End If

        Ws.Cells(row, 1).Interior.Color = RGB(255 - red, 255 - green, 255 - blue)
        Ws.Cells(row, 2).Interior.Color = RGB(255 - red, 255 - green, 255 - blue)
        Ws.Cells(row, 3).Interior.Color = RGB(255 - red, 255 - green, 255 - blue)

    Next

    Cells(2, 4).Formula = "=SUM(C2:C" & row & ")"
    Cells(2, 4).Interior.Color = RGB(245, 245, 245)
    
    Set Ws = Nothing

    ' [모든것의 완료]
    MsgBox ("간장치킨과 4월 7일")

End Sub

Sub copySheet()
    ActiveSheet.Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Copy After:=Sheets(Sheets.Count)
End Sub


Sub copyTest()
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets.Add(After:= _
            ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    Ws.Name = "testSwCho"
End Sub


















