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
 '    4. dec-->hex, 修改条件判断'
 '    5. 分离dbc读取代码块，改为函数，以备dbc合并
 '    6. 添加一列eConflict，表示dbc信号冲突, 可以手动删除冲突信号，然后再转换为dbc
 '    7. 添加CheckMsgConflict()，CheckSignalName()以查重
 '    8. 修改文件打开代码，以同时打开多个dbc文件
 '    9. add smDict
 '    10. add FileIndex attr --> fileType, Enum, 
 '    11. del CheckMsgConflict() --> super Code
 '    12. arrMessage, 
'''''''''''''''''''''''''''''''''''''''''''''
'Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, source As Any, ByVal Length As Long)

Option Explicit

'clms of Msg and Sig
Const vClmMsg As Integer = 6
Const vClmSig As Integer = 22
Public Enum eCLM
    eMessage = 1
    eID
    eDLC
    eTxMethod
    eCycleTime
    eMsgComment

    'Signal Attr
    eSignal
    eMultipGrp
    eStartbit
    eLength
    eByteOrder
    eValueType
    eInitialValue
    eFactor
    eOffset
    eMinimum
    eMaximum
    eUnit
    eValueTable
    eSigComment
    eConflict
    eFileIndex
End Enum

Dim vErrLog As String

Private Type Message
    Index As Integer
    Name As String
    id As Double
    DLC As Integer
    TxMethod As String
    Transmitter As String
    CycleTime As Integer
    SignalCount As Integer
    MsgComment As String
    Layoutuse(63) As Integer  'bitSt,0 Unuse, 1 use
    Conflict As Integer
End Type

Private Type Signal
    Index As Integer
    id As Double
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
    Conflict As Integer
    FileIndex As Integer
End Type

Private Type MsgComment
    id As Double
    Name As String
    Comment As String
    OKStart As Boolean
    OKEnd As Boolean
End Type

Private Type SignalComment
    id As Double
    Name As String
    Comment As String
    OKStart As Boolean
    OKEnd As Boolean
End Type

Dim dicMessage, dicSignal, dicNode, dicAttr As Scripting.Dictionary
'   (k-id-fl,v-starow), (id-sig-f,strow)
Dim m As Message 
Dim S As Signal  
Dim arrMessage() As Message
Dim arrSignal() As Signal
'   (indx,msg), (indx,sig)  
Dim start_row, emptyMessage As Integer

Dim countMessage, countSignal, countConflictSig As Integer
Dim attrMsgSendType() As String




Private Sub dbc_clear()

Dim i, j As Integer

i = 3
While Len(ActiveSheet.Cells(i, vClmMsg + 1)) > 0
    i = i + 1
Wend

j = vClmSig + 1
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
'cellsdelete'
Cells.Select
Range("B1").Activate
Cells.Delete Shift:=xlShiftUp

End Sub

Private Sub dbc_Click()

Application.DisplayAlerts = False
'On Error Resume Next
Dim MBox As Integer
Dim str, text As String
Dim i, j, Index As Integer
Dim Filename, File, k, v
Dim arr
Dim fso As New FileSystemObject
Dim head As String
Dim temp, temp_high, temp_low
Dim rowHt As Double
Dim starttime, endtime As Date
Dim elapsed As Double
Dim baudrate, totalbit As Double

baudrate = 500000
'MBox = MsgBox("提示：如果要合并多帧报文，请先加载多帧dbc文件", vbYesNoCancel + vbQuestion)

starttime = Now
endtime = starttime

'activesheet.ShowAllData
Call dbc_clear

text = GetElapsedTime(endtime, "Clear")
endtime = Now

Filename = Application.GetOpenFilename("DBC File,*.dbc", 0, "选择需转换、合并（可多选）的文件", Null, True)

head = "Message,ID,DLC [Byte],TxMethod,Cycle Time [Ms],MsgComment,Signal,Multip/Grp,Startbit,Length [Bit],Byte Order,Value Type,Initial Value,Factor,Offset,Minimum,Maximum,Unit,Value Table,Comment,Conflict, File"
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
countConflictSig = 0


start_row = 3

'change from hailing.hu
Columns("A:A").Select
Selection.NumberFormatLocal = "@"

text = text + vbLf + GetElapsedTime(endtime, "Read file")
endtime = Now

'todo: read dbc'

If IsArray(Filename) = True Then
    Application.ScreenUpdating = False
    ' For Each File In Filename
    For i = 1 To UBound(Filename)
        Call dbc_file_read(Filename(i), i)
    Next
    Application.ScreenUpdating = True
Else
    Exit Sub
End If

' Call dbc_file_read(Filename)
' Call dbc_file_read(File2ndname)

text = text + vbLf + GetElapsedTime(endtime, "Fill table")
endtime = Now

k = dicMessage.Keys
v = dicMessage.Items

For i = 0 To dicMessage.Count - 1
    temp = ActiveSheet.Cells(v(i), eID)
    If temp > 65535 Then
        temp_high = Fix(temp / 65536)   '0x10000'
        temp_low = temp - temp_high * 65536
        If temp_high > 32767 Then   '0x8000
            ActiveSheet.Cells(v(i), eID) = "0x" & Right(String(4, "0") & Hex(temp_high - 32768), 4) & Right(String(4, "0") & Hex(temp_low), 4)
        Else
            ActiveSheet.Cells(v(i), eID) = "0x" & Right(String(4, "0") & Hex(temp_high), 4) & Right(String(4, "0") & Hex(temp_low), 4)
        End If
    Else
        ActiveSheet.Cells(v(i), eID) = "0x" + Hex(temp)
    End If
    ' ActiveSheet.Cells(v(i), eID) = "0x" + Hex(temp)
Next i

text = text + vbLf + GetElapsedTime(endtime, "Format message id(dec-->hex)")
endtime = Now

start_row = 3
For i = 4 To countSignal + 3 + emptyMessage
    'same Message Name
    If ActiveSheet.Cells(i, 1) <> ActiveSheet.Cells(i - 1, 1) Then
        If i - start_row > 1 Then
            group start_row, i - 1
            For j = 1 To vClmMsg
                ' combine Col_Letter(j), start_row, i - 1
            Next
		    Range("B" + CStr(start_row) +":"+Col_Letter(vClmMsg)+ CStr(i-1)).Select
			Selection.FillDown  
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

sort countSignal + 2, dicNode.Count + vClmSig

text = text + vbLf + GetElapsedTime(endtime, "Sort")
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
'Auto fit: value_table,Sigcomments, ECUs
Columns(Col_Letter(eID) + ":" + Col_Letter(vClmSig)).Select
Selection.EntireColumn.AutoFit
Columns(Col_Letter(eMsgComment) + ":" + Col_Letter(eMsgComment)).Select
With Selection
    .ColumnWidth = 25
    .WrapText = True
End With
Columns(Col_Letter(eValueTable) + ":" + Col_Letter(eValueTable)).Select
With Selection
    .ColumnWidth = 15
    .WrapText = True
End With
Columns(Col_Letter(eSigComment) + ":" + Col_Letter(eSigComment)).Select
With Selection
    .ColumnWidth = 25
    .WrapText = True
End With

Columns(Col_Letter(vClmSig + 1) + ":" + Col_Letter(vClmSig + dicNode.Count)).Select
Selection.EntireColumn.AutoFit

Range("B3").Select
ActiveWindow.FreezePanes = True

text = text + vbLf + GetElapsedTime(endtime, "Format fillter and frezze")
endtime = Now

str = ""
rowHt = 48
' For Each File In Filename
For i = 1 To UBound(Filename)
        str = str + "DBC File(" + CStr(i) + ") = " + fso.GetFileName(Filename(i)) + vbLf
        rowHt = rowHt + 12
    Next
str = str + "ECU Nodes Count= " + CStr(dicNode.Count) + vbLf
str = str + "Messages Count= " + CStr(dicMessage.Count) + vbLf
str = str + "Signals Count= " + CStr(dicSignal.Count) + vbLf
str = str + "Conflict'Sigs Count= " + CStr(countConflictSig) + vbLf
'str = str + vbLf + "Bus Load= " + Format(totalbit * 100 / baudrate, "0.00") + "%"
ActiveSheet.Cells(1, vClmMsg + 1) = str
    ActiveSheet.Range(Col_Letter(vClmMsg + 1) & "1:" & Col_Letter(vClmSig) & "1").Select
    With Selection
        .Merge Across:=False
        .WrapText = True
        .AutoFit
        .VerticalAlignment = xlVAlignCenter
        .HorizontalAlignment = xlHAlignLeft
        .RowHeight = rowHt
    End With

'Can Type:
ActiveSheet.Cells(1, 2) = "Standard"
    ActiveSheet.Range("B1:" & Col_Letter(vClmMsg) & "1").Select
    With Selection
        .Merge Across:=False
        .HorizontalAlignment = xlHAlignCenter
    End With

Set dicMessage = Nothing
Set dicSignal = Nothing
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

Private Sub GetMsgSendTypeAttr(ByVal str As String)
 Dim vtemp
 Dim attr As String

'arr = Split(Mid(str, 1, Len(str) - 1), " ")
'attr_name = Mid(arr(2), 3, Len(arr(2)) - 2)
'attr_value = arr(3)
'dicAttr.Add attr_name, attr_value
'TODO: '
vtemp = InStr(1, str, "ENUM")
attr = Mid(str, vtemp + 6, Len(str) - vtemp - 6)
attr = Replace(attr, """,""", ",")
attr = Replace(attr, """", "")
attrMsgSendType = Split(attr, ",")
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



'todo: compare the signals'
Private Function CheckSignalName(ByVal Sig As String, ByVal start_row As Integer) As Integer
Dim ret As Integer

If Not dicSignal.Exists(Sig) Then
    dicSignal.Add Sig, start_row ' dicMessage.Count + 1
    ' dicSignal.Count = dicSignal.Count +1
    'ActiveSheet.Cells(2, vClmSig + dicNode.Count) = id
Else
    ActiveSheet.Cells(start_row, eConflict) = "Conflict"
    ActiveSheet.Cells(dicSignal.Item(Sig), eConflict) = "Conflict"
    countConflictSig = countConflictSig + 1 
End If
End Function

Private Function SetCycleTime(ByVal str As String) As Double
Dim arr
Dim i As Integer

arr = Split(Mid(str, 1, Len(str) - 1), " ")
i = dicMessage.Item(CStr(arr(3)))

'msg.CycleTime = arr(4)
ActiveSheet.Cells(i, 5) = arr(4)
' Range(Col_Letter(eCycleTime) + CStr(i) +":"+Col_Letter(eCycleTime)+m.SignalCount).Select
' Selection.FillDown
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

Private Function GetMsgComment(ByVal str As String) As MsgComment
Dim arr1, arr2
Dim Msgm As MsgComment

arr1 = Split(str, """")
arr2 = Split(str, " ")
Msgm.id = arr2(2)
' Msgm.Name = arr2(3)
Msgm.Comment = arr1(1)
Msgm.OKStart = True
If UBound(arr1) = 2 Then
    Msgm.OKEnd = True
End If

GetMsgComment = Msgm

End Function

Private Function GetComment(ByVal str As String) As SignalComment
Dim arr1, arr2
Dim sm As SignalComment

arr1 = Split(str, """")
arr2 = Split(str, " ")
sm.id = arr2(2)
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
Dim ret As Integer
ret = 1

arr = Split(str, " ")
m.Index = start_row
m.id = arr(1)
m.Name = Mid(arr(2), 1, Len(arr(2)) - 1)
m.DLC = arr(3)
m.Transmitter = arr(4)
'todo:'

If Not dicMessage.Exists(CStr(m.id)) Then
    dicMessage.Add CStr(m.id), start_row ' dicMessage.Count + 1
    'ActiveSheet.Cells(2, vClmSig + dicNode.Count) = id
    ret = 0
End If
m.Conflict = ret
countMessage = countMessage + 1
GetMessage = m
End Function

Private Function GetSignal(ByVal start_row As Integer, message_id As String, ByVal str As String) As Signal
Dim arr1, arr2, arr3
Dim S As Signal
Dim i1, i2, j As Integer
Dim vectorXXX(0) As String
vectorXXX(0) = "Vector__XXX"
S.Index = start_row
S.id = message_id
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
    S.Receiver = vectorXXX
End If
'
S.Range = "[" + CStr(S.Minimum) + "," + CStr(S.Maximum) + "]"
S.Encoding = "E=" + CStr(S.Factor) + "*N+" + CStr(S.Offset)

 countSignal = countSignal + 1
CheckSignalName CStr(S.id) + "-" + S.Name, S.Index

GetSignal = S
End Function

Private Sub SetInitialValue(str As String)
Dim arr
Dim i As Integer
i = InStr(str, ";")
arr = Split(Mid(str, 1, i - 1), " ")

i = dicSignal.Item(arr(3) + "-" + arr(4))
ActiveSheet.Cells(i, vClmMsg + 7) = arr(5) * ActiveSheet.Cells(i, vClmMsg + 8).Value + ActiveSheet.Cells(i, vClmMsg + 9).Value

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
ActiveSheet.Cells(i, vClmMsg + 13) = vt

End Sub

Private Sub combine(col As String, ByVal start_row As Integer, end_row As Integer)
Range(col + CStr(start_row) + ":" + col + CStr(end_row)).Select
With Selection
   ' .HorizontalAlignment = xlCenter
   .VerticalAlignment = xlCenter    'xlBottom
'    .WrapText = False
'    .Orientation = 0
'    .AddIndent = False
'    .IndentLevel = 0
'    .ShrinkToFit = False
'    .ReadingOrder = xlContext
'    .MergeCells = False
End With
Selection.Merge

End Sub

Private Function Col_Letter(ByVal lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function


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

'Read dbc file content'
Private Sub dbc_file_read(File, ByVal flidx As Integer)
Dim str, rline, text As String
Dim i, j, Index, ii As Integer
Dim lines() As String
Dim isUnix As Boolean
Dim arr
Dim Msgcmt As MsgComment
Dim sm As SignalComment

Open File For Input As #1
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
            CheckNode (arr(i))
        Next
    ElseIf InStr(1, rline, "BO_ ") = 1 Then
        'move to next row for empty frame
        If m.id > 0 And m.SignalCount = 0 Then
            emptyMessage = emptyMessage + 1
            start_row = start_row + 1
        End If
        m = GetMessage(start_row, rline)
        If m.Conflict = 0 Then
             ActiveSheet.Cells(start_row, eMessage) = m.Name
             ActiveSheet.Cells(start_row, eID) = m.id
             ActiveSheet.Cells(start_row, eDLC) = m.DLC
            CheckNode (m.Transmitter)
        End If
    ElseIf InStr(1, rline, "SG_ ") = 1 Then
        m.SignalCount = m.SignalCount + 1
        S = GetSignal(start_row, CStr(m.id), rline)
        S.FileIndex = flidx
        ActiveSheet.Cells(start_row, eMessage) = m.Name
        ActiveSheet.Cells(start_row, eID) = m.id
        ActiveSheet.Cells(start_row, eDLC) = m.DLC


        ActiveSheet.Cells(start_row, eSignal) = S.Name
        ActiveSheet.Cells(start_row, eMultipGrp) = S.Multiplexing_Group
        ActiveSheet.Cells(start_row, eStartbit) = S.Startbit
        ActiveSheet.Cells(start_row, eLength) = S.Length
        ActiveSheet.Cells(start_row, eByteOrder) = S.ByteOrder
        ActiveSheet.Cells(start_row, eValueType) = S.ValueType
        ActiveSheet.Cells(start_row, eInitialValue) = S.InitialValue
        ActiveSheet.Cells(start_row, eFactor) = S.Factor
        ActiveSheet.Cells(start_row, eOffset) = S.Offset
        ActiveSheet.Cells(start_row, eMinimum) = S.Minimum
        ActiveSheet.Cells(start_row, eMaximum) = S.Maximum
        ActiveSheet.Cells(start_row, eUnit) = S.Unit
        ActiveSheet.Cells(start_row, eFileIndex) = S.FileIndex
        'activesheet.Cells(start_row,  - vClmMsg) = s.Encoding

        'Node and Color'
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
    ElseIf InStr(1, rline, "CM_ BO_ ") = 1 Then
        Msgcmt = GetMsgComment(rline)
    ElseIf Msgcmt.id > 0 And Msgcmt.OKStart And Not Msgcmt.OKEnd Then
        arr = Split(rline, """")
        If UBound(arr) = 1 Then
            Msgcmt.Comment = Msgcmt.Comment + vbLf + arr(0)
            Msgcmt.OKEnd = True
        Else
            Msgcmt.Comment = Msgcmt.Comment + vbLf + rline
        End If
    ElseIf InStr(1, rline, "CM_ SG_ ") = 1 Then
        sm = GetComment(rline)
    ElseIf sm.id > 0 And sm.OKStart And Not sm.OKEnd Then
        arr = Split(rline, """")
        If UBound(arr) = 1 Then
            sm.Comment = sm.Comment + vbLf + arr(0)
            sm.OKEnd = True
        Else
            sm.Comment = sm.Comment + vbLf + rline
        End If
    ElseIf InStr(1, rline, "BA_DEF_ BO_  ") = 1 Then
        If InStr(1, rline, "GenMsgSendType") > 0 Then
            GetMsgSendTypeAttr (rline)
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
    End If
    If Msgcmt.id > 0 And Msgcmt.OKStart And Msgcmt.OKEnd Then
        i = dicMessage.Item(CStr(Msgcmt.id))
        ActiveSheet.Cells(i, eMsgComment) = Msgcmt.Comment      '.AddComment text:
        'm.MsgComment = Msgcmt.Comment
        Msgcmt.id = 0
    End If
    If sm.id > 0 And sm.OKStart And sm.OKEnd Then
        i = dicSignal.Item(CStr(sm.id) + "-" + sm.Name)     ' + File
        'dicSig.add CStr(sm.id) + "-" + sm.Name, sm.Comment
        ActiveSheet.Cells(i, eSigComment) = sm.Comment
        'S.Comment = sm.Comment
        sm.id = 0
    End If
    
Next ii
End Sub

'Callback for customButton1 onAction
Sub dbc2excel(control As IRibbonControl)
dbc_Click
End Sub

