''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'If you see this comment, you have already cracked this tool.I hope this code will help you.
'Submit issue if you have problem.
'Author: zhyongquan
'Email: zhyongquan@gmail.com
'GitHub: https://github.com/zhyongquan/DBC2Excel
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Changed By Xia
'date   : 2020-03-16
'log    : add clm of TxMethod, Multi/G
 ' Update DBC2Excel.vb
 '    Changes:
 '    1. add vClmMsg, vClmSig to loc the colomns easier. in the mean time, change the colomn num to Enum;
 '    2. add colomn: TxMothed, MultiGrp. And, change the cells()
 '    3. add bus tpye to convert excel to dbc.
'''''''''''''''''''''''''''''''''''''''''''''
'Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, source As Any, ByVal Length As Long)

Option Explicit

'clms of Msg and Sig 
Dim vClmMsg, vClmSig As Integer
Public Enum eCLM 
    eMessage =1
    eID =2
    eDLC =3
    eTxMethod =4
    eCycleTime =5
    eSignal =6
    eMultipGrp =7
    eStartbit =8
    eLength =9
    eByteOrder =10
    eValueType =11
    eInitialValue =12
    eFactor =13
    eOffset =14
    eMinimum =15
    eMaximum =16
    eUnit =17
    eValueTable =18
    eComment =19 
End Enum

Dim vErrLog As String

Private Type Message
    Index As Integer
    Name As String
    ID As Double
    DLC As Integer
    TxMethod As String
    Transmitter As String
    CycleTime As Integer
    SignalCount As Integer
End Type

Private Type Signal
    Index As Integer
    ID As Double
    Name As String
    Multiplexing_Group As String
    Startbit As Integer
    Length As Integer
    ByteOrder As String
    ValueType As String
    InitialValue As Double
    Factor As Double
    Offset As Double
    Minimum As Double
    Maximum As Double
    Unit As String
    ValueTable As String
    Comment As String
    Receiver() As String
    Encoding As String
    Range As String
End Type

Private Type SignalComment
    ID As Double
    Name As String
    Comment As String
    OKStart As Boolean
    OKEnd As Boolean
End Type

Dim dicMessage, dicSignal, dicNode, dicAttr As Scripting.Dictionary

Dim countMessage, countSignal As Integer
Dim attrMsgSendType As String


Private Sub dbc_clear()

Dim i, j As Integer

i = 3
While Len(ActiveSheet.Cells(i, vClmMsg+1)) > 0
    i = i + 1
Wend

j = 18
While Len(ActiveSheet.Cells(2, j)) > 0
    j = j + 1
Wend



ActiveWindow.FreezePanes = False
ActiveSheet.Cells(1, 2) = ""
Rows("2:" + CStr(i)).Select
Selection.ClearOutline
Selection.Delete Shift:=xlUp
Columns("R:" + Col_Letter(j)).Select
Selection.Delete Shift:=xlToLeft
    
End Sub
Private Sub dbc_Click()

Application.DisplayAlerts = False
'On Error Resume Next

Dim str, rline, text As String
Dim i, j, Index, start_row, ii As Integer
Dim Filename, k, v
Dim arr
Dim lines() As String
Dim fso As New FileSystemObject
Dim head As String
Dim temp, temp_high, temp_low
Dim isUnix As Boolean
Dim starttime, endtime As Date
Dim elapsed As Double
Dim baudrate, totalbit As Double
Dim emptyMessage As Integer
baudrate = 500000
vClmMsg = 5
vClmSig = 19
Filename = Application.GetOpenFilename("DBC File,*.dbc")

If Filename = False Then
    Exit Sub
End If

starttime = Now
endtime = starttime

'activesheet.ShowAllData
Call dbc_clear

text = GetElapsedTime(endtime, "Clear")
endtime = Now


head = "Message,ID,DLC [Byte],TxMethod,Cycle Time [Ms],Signal,Multip/Grp,Startbit,Length [Bit],Byte Order,Value Type,Initial Value,Factor,Offset,Minimum,Maximum,Unit,Value Table,Comment"
arr = Split(head, ",")
For i = 0 To UBound(arr)
    ActiveSheet.Cells(2, i + 1) = arr(i)
Next i

Set dicNode = New Scripting.Dictionary
Set dicMessage = New Scripting.Dictionary
Set dicSignal = New Scripting.Dictionary
Set dicAttr = New Scripting.Dictionary

countSignal = 0
countMessage = 0

Open Filename For Input As #1
Dim m As Message
Dim S As Signal
Dim sm As SignalComment
start_row = 3

Line Input #1, rline
If EOF(1) Then
    isUnix = True
    rline = Replace(rline, vbLf, vbLf)
    lines = Split(rline, vbLf)
Else
    isUnix = False
    str = rline
    While Not EOF(1)
        Line Input #1, rline
        Index = Index + 1
        str = str + vbLf + rline
    Wend
    lines = Split(str, vbLf)
End If
Close #1

'change from hailing.hu
Columns("A:A").Select
Selection.NumberFormatLocal = "@"

text = text + vbLf + GetElapsedTime(endtime, "Read file")
endtime = Now

For ii = 0 To UBound(lines)
    rline = lines(ii)
    rline = Trim(rline)
'    If isUnix And Len(rline) > 1 Then
'        rline = Mid(rline, 1, Len(rline) - 1)
'    End If
    ' Get the Node, and add to list after vClmSig
    If InStr(1, rline, "BU_: ") = 1 Then
        arr = Split(rline, " ")
        For i = 1 To UBound(arr)
            ActiveSheet.Cells(2, vClmSig + i) = arr(i)
            dicNode.Add arr(i), i
        Next
    ElseIf InStr(1, rline, "BO_ ") = 1 Then
        'move to next row for empty frame
        If m.ID > 0 And m.SignalCount = 0 Then
            emptyMessage = emptyMessage + 1
            start_row = start_row + 1
        End If
        m = GetMessage(start_row, rline)
        ActiveSheet.Cells(start_row, eMessage) = m.Name
        ActiveSheet.Cells(start_row, eID) = m.ID
        ActiveSheet.Cells(start_row, eDLC) = m.DLC
        CheckNode (m.Transmitter)
    ElseIf InStr(1, rline, "SG_ ") = 1 Then
        m.SignalCount = m.SignalCount + 1
        S = GetSignal(start_row, CStr(m.ID), rline)
        ActiveSheet.Cells(start_row, 1) = m.Name
        'ActiveSheet.Cells(start_row, 2) = m.ID
        ActiveSheet.Cells(start_row, vClmMsg+1) = S.Name
        ActiveSheet.Cells(start_row, vClmMsg+2) = S.Multiplexing_Group
        ActiveSheet.Cells(start_row, vClmMsg+3) = S.Startbit
        ActiveSheet.Cells(start_row, vClmMsg+4) = S.Length
        ActiveSheet.Cells(start_row, vClmMsg+5) = S.ByteOrder
        ActiveSheet.Cells(start_row, vClmMsg+6) = S.ValueType
        ActiveSheet.Cells(start_row, vClmMsg+7) = S.InitialValue
        ActiveSheet.Cells(start_row, vClmMsg+8) = S.Factor
        ActiveSheet.Cells(start_row, vClmMsg+9) = S.Offset
        ActiveSheet.Cells(start_row, vClmMsg+10) = S.Minimum
        ActiveSheet.Cells(start_row, vClmMsg+11) = S.Maximum
        ActiveSheet.Cells(start_row, vClmMsg+12) = S.Unit
        'activesheet.Cells(start_row, vClmMsg+13) = s.Encoding
        
        j = dicNode.Item(m.Transmitter)
        ActiveSheet.Cells(start_row, vClmSig + j) = "T"
        Range(Col_Letter(vClmSig + j) + CStr(start_row)).Select
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        For i = 0 To UBound(S.Receiver)
            CheckNode (S.Receiver(i))
            j = dicNode.Item(S.Receiver(i))
            ActiveSheet.Cells(start_row, vClmSig + j) = "R"
            Range(Col_Letter(vClmSig + j) + CStr(start_row)).Select
            Selection.Font.Bold = True
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65280
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next i
        start_row = start_row + 1
    ElseIf InStr(1, rline, "CM_ SG_ ") = 1 Then
        sm = GetComment(rline)
    ElseIf InStr(1, rline, "BA_DEF_ BO_  ") = 1 Then
        If InStr(1, rline, "GenMsgSendType") > 0 Then
            SetMsgSendTypeAttr(rline)
        End If
'    ElseIf InStr(1, rline, "BA_DEF_ SG_  ") = 1 Then
'    ElseIf InStr(1, rline, "BA_DEF_   ") = 1 Then
'    ElseIf InStr(1, rline, "BA_DEF_DEF_  ") = 1 Then
'        SetDBCAttr rline
    ElseIf InStr(1, rline, "BA_ ") = 1 Then
        If InStr(1, rline, "GenSigStartValue") > 0 And InStr(1, rline, "SG_") > 0 Then
            SetInitialValue (rline)
        ElseIf InStr(1, rline, "GenMsgCycleTime") > 0 And InStr(1, rline, "BO_") > 0 Then
            SetCycleTime (rline)
        ElseIf InStr(1, rline, "GenMsgSendType") > 0 And InStr(1, rline, "BO_") > 0 Then
            SetSendType (rline)   
        End If
    ElseIf InStr(1, rline, "VAL_ ") = 1 Then
        SetValueTable (rline)
    ElseIf sm.ID > 0 And sm.OKStart And Not sm.OKEnd Then
        arr = Split(rline, """")
        If UBound(arr) = 1 Then
            sm.Comment = sm.Comment + vbLf + arr(0)
            sm.OKEnd = True
        Else
            sm.Comment = vbLf + rline
        End If
    End If
    If sm.ID > 0 And sm.OKStart And sm.OKEnd Then
        i = dicSignal.Item(CStr(sm.ID) + "-" + sm.Name)
        ActiveSheet.Cells(i, eComment) = sm.Comment
        sm.ID = 0
    End If
    
Next ii

text = text + vbLf + GetElapsedTime(endtime, "Fill table")
endtime = Now

k = dicMessage.Keys
v = dicMessage.Items

For i = 0 To dicMessage.Count - 1
    temp = ActiveSheet.Cells(v(i), eID)
    If temp > 65535 Then
        temp_high = Fix(temp / 65536)
        temp_low = temp - temp_high * 65536
        If temp_high > 32768 Then
            ActiveSheet.Cells(v(i), eID) = "0x" & Right(String(4, "0") & Hex(temp_high - 32768), 4) & Right(String(4, "0") & Hex(temp_low), 4)
        Else
            ActiveSheet.Cells(v(i), eID) = "0x" + Hex(temp_low)
        End If
    Else
        ActiveSheet.Cells(v(i), eID) = "0x" + Hex(temp)
    End If
Next i

text = text + vbLf + GetElapsedTime(endtime, "Format message id")
endtime = Now

sort countSignal + 2, dicNode.Count + vClmSig

text = text + vbLf + GetElapsedTime(endtime, "Sort")
endtime = Now

start_row = 3
For i = 4 To countSignal + 3 + emptyMessage
    If ActiveSheet.Cells(i, 1) <> ActiveSheet.Cells(i - 1, 1) Then
        If i - start_row > 1 Then
            group start_row, i - 1
            combine "A", start_row, i - 1
            combine "B", start_row, i - 1
            combine "C", start_row, i - 1
            combine "D", start_row, i - 1
            combine "E", start_row, i - 1
        End If
        start_row = i
    End If
Next i

'For i = 0 To dicMessage.Count - 2
'    If v(i + 1) - v(i) > 1 Then
'        combine "A", v(i), v(i + 1) - 1
'        combine "B", v(i), v(i + 1) - 1
'        combine "C", v(i), v(i + 1) - 1
'        combine "D", v(i), v(i + 1) - 1
''        group v(i), v(i + 1) - 1
'    End If
'Next i
'If 2 + countSignal > v(i) Then
'    combine "A", v(i), 2 + countSignal
'    combine "B", v(i), 2 + countSignal
'    combine "C", v(i), 2 + countSignal
'    combine "D", v(i), 2 + countSignal
''    group v(i), 2 + countSignal
'End If

text = text + vbLf + GetElapsedTime(endtime, "Format message")
endtime = Now

Range("A2:" + Col_Letter(vClmSig + dicNode.Count) + "2").Select
Selection.Font.Bold = True
Selection.AutoFilter

text = text + vbLf + GetElapsedTime(endtime, "Format title")
endtime = Now

If emptyMessage > 0 Then
    emptyMessage = emptyMessage + 1
End If
Range("A2:" + Col_Letter(vClmSig + dicNode.Count) + CStr(countSignal + 2 + emptyMessage)).Select

Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With Selection.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

text = text + vbLf + GetElapsedTime(endtime, "Format grid")
endtime = Now

Columns("R:" + Col_Letter(vClmSig + dicNode.Count)).Select
Columns("R:" + Col_Letter(vClmSig + dicNode.Count)).EntireColumn.AutoFit

Range("B3").Select
ActiveWindow.FreezePanes = True

text = text + vbLf + GetElapsedTime(endtime, "Format fillter and frezze")
endtime = Now

str = "DBC File= " + fso.GetFileName(Filename) + vbLf
str = str + "ECU Nodes Count= " + CStr(dicNode.Count) + vbLf
str = str + "Messages Count= " + CStr(dicMessage.Count) + vbLf
str = str + "Signals Count= " + CStr(dicSignal.Count)
'str = str + vbLf + "Bus Load= " + Format(totalbit * 100 / baudrate, "0.00") + "%"
ActiveSheet.Cells(1, eMessage+1) = str

ActiveSheet.Cells(1, 2) = "Standard"
    ActiveCell.Range("B1:" & Col_Letter(vClmMsg) & "1").Select
    With Selection
        .Merge Across:=False
        .HorizontalAlignment = xlHAlignCenter
    End With

Set dicMessage = Nothing
Set dicMessage = Nothing
Set dicAttr = Nothing
Set fso = Nothing

text = text + vbLf + GetElapsedTime(endtime, "End")
endtime = Now

MsgBox "Finish, " + GetElapsedTime(starttime, "elapsed time") + vbLf + text + vbLf + vErrLog
' MsgBox "Finish, " + GetElapsedTime(starttime, "elapsed time") + vbLf + vErrLog

End Sub

Private Sub SetDBCAttr(ByVal str As String)
Dim arr
Dim attr_name, attr_value

arr = Split(Mid(str, 1, Len(str) - 1), " ")
attr_name = Mid(arr(2), 3, Len(arr(2)) - 2)
attr_value = arr(3)

dicAttr.Add attr_name, attr_value

End Sub

Private Sub SetMsgSendTypeAttr(ByVal str As String)
' Dim arr
' Dim attr_name, attr_value

'arr = Split(Mid(str, 1, Len(str) - 1), " ")
'attr_name = Mid(arr(2), 3, Len(arr(2)) - 2)
'attr_value = arr(3)
'dicAttr.Add attr_name, attr_value
'TODO: '
attrMsgSendType = Split(Mid(str, InStr(1, str, "ENUM")+5,Len(str)-1), "","")
End Sub

Private Sub sort(ByVal end_row As Integer, ByVal end_col As Integer)

Range("A3:A" + CStr(end_row)).Select
Range(Col_Letter(end_col) + CStr(end_row)).Activate
ActiveSheet.sort.SortFields.clear
ActiveSheet.sort.SortFields.Add Key:=Range("A3:A" + CStr(end_row)), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveSheet.sort.SortFields.Add Key:=Range(Col_Letter(eMultipGrp) & "3:" & Col_Letter(eMultipGrp) & CStr(end_row)), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveSheet.sort.SortFields.Add Key:=Range(Col_Letter(eStartbit) & "3:" & Col_Letter(eStartbit) & CStr(end_row)), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveSheet.sort
    .SetRange Range("A2:" + Col_Letter(end_col) + CStr(end_row))
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

End Sub
Private Sub group(ByVal start_row As Integer, end_row As Integer)
Dim i, j As Integer

'For i = start_row + 1 To end_row
'    For j = 1 To 4
'        ActiveSheet.Cells(i, j) = ActiveSheet.Cells(start_row, j)
'    Next j
'Next i

Rows(CStr(start_row + 1) + ":" + CStr(end_row)).Select
Selection.Rows.group

End Sub

Private Function GetElapsedTime(ByVal starttime As Date, ByVal step As String) As String
Dim text As String
Dim elapsed As Double
Dim endtime As Date

endtime = Now
elapsed = endtime - starttime
text = step + ": " + Format(elapsed * 3600 * 24, "#0") + "s"
GetElapsedTime = text

End Function
Private Sub CheckNode(ByVal str As String)
If Not dicNode.Exists(str) Then
    dicNode.Add str, dicNode.Count + 1
    ActiveSheet.Cells(2, vClmSig + dicNode.Count) = str
End If
End Sub

Private Function SetCycleTime(ByVal str As String) As Double
Dim arr
Dim i As Integer

arr = Split(Mid(str, 1, Len(str) - 1), " ")
i = dicMessage.Item(CStr(arr(3)))
ActiveSheet.Cells(i, 5) = arr(4)

'SetCycleTime = 1000# / arr(4) * (ActiveSheet.Cells(i, 3) * 8 + Fix((ActiveSheet.Cells(i, 3) * 8 + 1 + 32 + 6 + 16) / 5) + 32 + 32)

End Function

Private Function SetSendType(ByVal str As String) As Double
Dim arr
Dim i As Integer

arr = Split(Mid(str, 1, Len(str) - 1), " ")
i = dicMessage.Item(CStr(arr(3)))
'todo: '
ActiveSheet.Cells(i, eTxMethod) = attrMsgSendType(arr(4))

'SetCycleTime = 1000# / arr(4) * (ActiveSheet.Cells(i, 3) * 8 + Fix((ActiveSheet.Cells(i, 3) * 8 + 1 + 32 + 6 + 16) / 5) + 32 + 32)

End Function

Private Function GetComment(ByVal str As String) As SignalComment
Dim arr1, arr2
Dim sm As SignalComment

arr1 = Split(str, """")
arr2 = Split(str, " ")
sm.ID = arr2(2)
sm.Name = arr2(3)
sm.Comment = arr1(1)
sm.OKStart = True
If UBound(arr1) = 2 Then
    sm.OKEnd = True
End If

GetComment = sm

End Function

Private Function GetMessage(ByVal start_row As Integer, ByVal str As String) As Message
Dim m As Message
Dim arr() As String

arr = Split(str, " ")
m.Index = start_row
m.ID = arr(1)
m.Name = Mid(arr(2), 1, Len(arr(2)) - 1)
m.DLC = arr(3)
m.Transmitter = arr(4)

countMessage = countMessage + 1
dicMessage.Add CStr(m.ID), m.Index

GetMessage = m
End Function

Private Function GetSignal(ByVal start_row As Integer, message_id As String, ByVal str As String) As Signal
Dim arr1, arr2, arr3
Dim S As Signal
Dim i1, i2, j As Integer
S.Index = start_row
S.ID = message_id
'unit
arr1 = Split(str, """")
S.Unit = arr1(1)
'i1 = InStr(str, """")
'i2 = InStr(i1 + 1, str, """")
's.Unit = Mid(str, i1 + 1, i2 - i1)
'name
arr2 = Split(arr1(0), " ")
S.Name = arr2(1)
If arr2(3) = ":" Then
    S.Multiplexing_Group = arr2(2)
    j = 1
Else
    S.Multiplexing_Group = "-"
    j = 0
End If
'startbit,length,byte order,sign
i1 = InStr(arr2(3 + j), "|")
i2 = InStr(arr2(3 + j), "@")
S.Startbit = Mid(arr2(3 + j), 1, i1 - 1)
S.Length = Mid(arr2(3 + j), i1 + 1, i2 - i1 - 1)
If Mid(arr2(3 + j), i2 + 1, 1) = "0" Then
    S.ByteOrder = "MSB"
Else
    S.ByteOrder = "LSB"
End If
If Mid(arr2(3 + j), i2 + 2, 1) = "+" Then
    S.ValueType = "Unsigned"
Else
    S.ValueType = "Signed"
End If
'factor,offset
i1 = InStr(arr2(4 + j), ",")
S.Factor = Mid(arr2(4 + j), 2, i1 - 2)
S.Offset = Mid(arr2(4 + j), i1 + 1, Len(arr2(4 + j)) - i1 - 1)

'min,max
i1 = InStr(arr2(5 + j), "|")
S.Minimum = Mid(arr2(5 + j), 2, i1 - 2)
S.Maximum = Mid(arr2(5 + j), i1 + 1, Len(arr2(5 + j)) - i1 - 1)
'receiver
If UBound(arr1) >= 2 Then
    S.Receiver = Split(Trim(arr1(2)), ",")
Else
    vErrLog = vErrLog & "Error with : " & str & vbLf
    S.Receiver = Split(Trim(arr1(1)), " ")
End If
'
S.Range = "[" + CStr(S.Minimum) + "," + CStr(S.Maximum) + "]"
S.Encoding = "E=" + CStr(S.Factor) + "*N+" + CStr(S.Offset)

countSignal = countSignal + 1
dicSignal.Add CStr(S.ID) + "-" + S.Name, S.Index

GetSignal = S
End Function

Private Sub SetInitialValue(str As String)
Dim arr
Dim i As Integer
i = InStr(str, ";")
arr = Split(Mid(str, 1, i - 1), " ")

i = dicSignal.Item(arr(3) + "-" + arr(4))
ActiveSheet.Cells(i, vClmMsg+7) = arr(5) * ActiveSheet.Cells(i, vClmMsg+8).Value + ActiveSheet.Cells(i, vClmMsg+9).Value

End Sub
Private Sub SetValueTable(str As String)
Dim i, j As Integer
Dim arr1, arr2
Dim vt As String
arr1 = Split(str, " ")
arr2 = Split(str, """")

For j = UBound(arr2) / 2 To 2 Step -1
    If Len(arr2(j * 2 - 2)) > 5 Then
        vt = vt + arr2(j * 2 - 2) + "=" + Trim(arr2(j * 2 - 1)) + ";" + vbLf
    Else
        vt = vt + "0x" + ConvertDecHex(CLng(arr2(j * 2 - 2))) + "=" + Trim(arr2(j * 2 - 1)) + ";" + vbLf
    End If
Next j
If Len(arr1(3)) > 5 Then
    vt = vt + arr1(3) + "=" + arr2(1) + ";"
Else
    vt = vt + "0x" + ConvertDecHex(CLng(arr1(3))) + "=" + arr2(1) + ";"
End If


i = dicSignal.Item(arr1(1) + "-" + arr1(2))
ActiveSheet.Cells(i, vClmMsg+13) = vt

End Sub

Private Sub combine(col As String, ByVal start_row As Integer, end_row As Integer)
Range(col + CStr(start_row) + ":" + col + CStr(end_row)).Select
'With Selection
'    .HorizontalAlignment = xlCenter
'    .VerticalAlignment = xlBottom
'    .WrapText = False
'    .Orientation = 0
'    .AddIndent = False
'    .IndentLevel = 0
'    .ShrinkToFit = False
'    .ReadingOrder = xlContext
'    .MergeCells = False
'End With
Selection.Merge
'With Selection
'    .HorizontalAlignment = xlCenter
'    .VerticalAlignment = xlCenter
'    .WrapText = False
'    .Orientation = 0
'    .AddIndent = False
'    .IndentLevel = 0
'    .ShrinkToFit = False
'    .ReadingOrder = xlContext
'    .MergeCells = True
'End With
'With Selection
'    .HorizontalAlignment = xlLeft
'    .VerticalAlignment = xlCenter
'    .WrapText = False
'    .Orientation = 0
'    .AddIndent = False
'    .IndentLevel = 0
'    .ShrinkToFit = False
'    .ReadingOrder = xlContext
'    .MergeCells = True
'End With
End Sub

Private Function Col_Letter(ByVal lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

'Callback for customButton1 onAction
Sub dbc2excel(control As IRibbonControl)
dbc_Click
End Sub

'Private Function ReadUniFile(ByVal sFile As String) As String
'  Dim A As Long
'  A = FileLen(sFile)
'  ReDim Buff(A - 1) As Byte
'  ReDim Buff1(A - 3) As Byte
'  Open sFile For Binary As #1
'    Get #1, , Buff
'  Close #1
'  RtlMoveMemory Buff1(0), Buff(2), A - 2
'  Dim S As String
'  S = StrConv(Buff1, vbNarrow)
'  ReadUniFile = S
'End Function


Private Function ConvertDecHex(Num_Dec As Long) As String

    Dim sTemp As String
   
    If Num_Dec >= 16 Then
        'if greater than 16 then
        'call recursively this function
        sTemp = ConvertDecHex(Num_Dec \ 16) _
            & ConvertDecHex(Num_Dec Mod 16)
           
    ElseIf Num_Dec > 9 Then
        'if within 10 to 15 then assign A...F
        Select Case Num_Dec
            Case 10: sTemp = "A"
            Case 11: sTemp = "B"
            Case 12: sTemp = "C"
            Case 13: sTemp = "D"
            Case 14: sTemp = "E"
            Case 15: sTemp = "F"
        End Select
    Else
        'If within 0 to 9 then no change
        sTemp = Num_Dec
    End If
           
    ConvertDecHex = sTemp
       
End Function


