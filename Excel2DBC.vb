''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'If you see this comment, you have already cracked this tool.I hope this code will help you.
'Submit issue if you have problem.
'Author: zhyongquan
'Email: zhyongquan@gmail.com
'GitHub: https://github.com/zhyongquan/Excel2DBC
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
    eSignal     '7
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
    eFileIndex  '22
End Enum


Sub dbc_Click()
Application.DisplayAlerts = False
On Error Resume Next

Dim dicMsgB, dicSigB As Scripting.Dictionary
'   (k-id,v-idx), (id-sig,idx)
Dim i, j, k, node_count, message_count, signal_count As Integer
Dim Filename, arr
Dim nodes As String
Dim message, id, dlc, cycle_time, tx
Dim line, signal, byte_order, value_type, initial_value, value_table, comment, rx As String
Dim initial_value_list, value_table_list, cycle_time_list, comment_list As String
Dim str, text As String
Dim fso As New FileSystemObject
Dim starttime, endtime As Date
Dim elapsed As Double
Dim dbc_type As String
Dim vtmpMulti As String

Filename = Application.GetSaveAsFilename(fileFilter:="DBC Files (*.dbc), *.dbc")

If Filename = False Then
    Exit Sub
End If

Set dicMsgB = New Scripting.Dictionary
Set dicSigB = New Scripting.Dictionary

dbc_type = ActiveSheet.Cells(1, 2)

starttime = Now
endtime = starttime

Open Filename For Output As 1#

If dbc_type <> "J1939" Then
    Print #1, Sheet4.Cells(1, 2)
Else
    Print #1, Sheet5.Cells(1, 2)
End If

Print #1, vbLf

i = vClmSig + 1
nodes = ""
While Len(ActiveSheet.Cells(2, i)) > 0
    node_count = node_count + 1
    nodes = nodes & " " & ActiveSheet.Cells(2, i)
    i = i + 1
Wend
Print #1, "BU_:" & nodes
 
i = 3
message_count = 0
signal_count = 0
While Len(ActiveSheet.Cells(i, vClmMsg + 1)) > 0    'exist Signal
    If Len(ActiveSheet.Cells(i, eMessage)) > 0 Then
        'read Message Attr
        If NOT dicMsgB.Exists(CStr(ActiveSheet.Cells(i, eID))) Then
            dicMsgB.Add CStr(ActiveSheet.Cells(i, eID)), dicMsgB.Count + 1 
            message_count = message_count + 1
            message = ActiveSheet.Cells(i, 1)
            id = Hex2Dec(ActiveSheet.Cells(i, eID))
            If dbc_type <> "Standard" Then
                id = id + 2147483648#
            End If
            dlc = ActiveSheet.Cells(i, eDLC)
            tx = ""
            For j = 1 To node_count
                If ActiveSheet.Cells(i, j + vClmSig) = "T" Then
                    tx = ActiveSheet.Cells(2, j + vClmSig)
                    Exit For
                End If
            Next j
            If tx = "" Then
                tx = "Vector__XXX"
            End If

            'writ Message head 
            Print #1, vbLf
            Print #1, "BO_ " & id & " " & ActiveSheet.Cells(i, eMessage) & ": " & ActiveSheet.Cells(i, eDLC) & " " & tx 
            'CycleTime
            If Len(ActiveSheet.Cells(i, eCycleTime) > 0) Then
                cycle_time = ActiveSheet.Cells(i, eCycleTime) + 0
                'cycle_time_list = cycle_time_list & "BA_ " & """GenMsgILSupport"" BO_ " & id & " " & 1 & ";" & vbLf
                cycle_time_list = cycle_time_list & "BA_ " & """GenMsgSendType"" BO_ " & id & " " & 0 & ";" & vbLf
                cycle_time_list = cycle_time_list & "BA_ " & """GenMsgCycleTime"" BO_ " & id & " " & cycle_time & ";" & vbLf
            End If
        End If
    End If

    'Signal set
    signal_count = signal_count + 1
    signal = ActiveSheet.Cells(i, eSignal)
    If ActiveSheet.Cells(i, eByteOrder) = "MSB" Then
        byte_order = "0"
    Else
        byte_order = "1"
    End If
    If ActiveSheet.Cells(i, eValueType) = "Unsigned" Then
        value_type = "+"
    Else
        value_type = "-"
    End If
    rx = ""
    For j = 1 To node_count
            If ActiveSheet.Cells(i, j + vClmSig) = "R" Then
                rx = rx & ActiveSheet.Cells(2, j + vClmSig) & ","
            End If
    Next j
    If rx = "" Then
        rx = " Vector__XXX"
    Else
        rx = Mid(rx, 1, Len(rx) - 1)
    End If

    'Multi
    If ActiveSheet.Cells(i, eMultipGrp) <> "-" Then
        vtmpMulti = " " & ActiveSheet.Cells(i, eMultipGrp)
    Else
        vtmpMulti = ""
    End If
    Print #1, " SG_ " & ActiveSheet.Cells(i, eSignal) & vtmpMulti & " : " & ActiveSheet.Cells(i, eStartbit) & "|" & ActiveSheet.Cells(i, eLength) & "@" & byte_order & value_type & _
        " (" & Num2Str(ActiveSheet.Cells(i, eFactor)) & "," & Num2Str(ActiveSheet.Cells(i, eOffset)) & ") " & "[" & Num2Str(ActiveSheet.Cells(i, eMinimum)) & "|" & Num2Str(ActiveSheet.Cells(i, eMaximum)) & "] " & _
        """" & ActiveSheet.Cells(i, eUnit) & """ " & rx
    If Len(ActiveSheet.Cells(i, eInitialValue)) > 0 And Len(ActiveSheet.Cells(i, eMessage)) > 0 And ActiveSheet.Cells(i, eFactor) <> 0 Then
        initial_value = (ActiveSheet.Cells(i, eInitialValue) - ActiveSheet.Cells(i, eOffset)) / ActiveSheet.Cells(i, eFactor)
        initial_value_list = initial_value_list & "BA_ ""GenSigStartValue"" SG_ " & id & " " & signal & " " & initial_value & ";" & vbLf
    End If
    If Len(ActiveSheet.Cells(i, eValueTable)) > 0 Then
        arr = Split(ActiveSheet.Cells(i, eValueTable), vbLf)
        value_table = ""
        For j = UBound(arr) To 0 Step -1
            k = InStr(arr(j), "=")
            value_table = value_table & Hex2Dec(Mid(arr(j), 1, k - 1)) & " """ & Mid(arr(j), k + 1, Len(arr(j)) - k - 1) & """ "
        Next j
        value_table_list = value_table_list & "VAL_ " & id & " " & signal & " " & value_table & ";" & vbLf
    End If
    If Len(ActiveSheet.Cells(i, eSigComment)) > 0 Then
        comment = ActiveSheet.Cells(i, eSigComment)
        comment_list = comment_list & "CM_ SG_ " & id & " " & signal & " """ & comment & """;" & vbLf
    End If
    i = i + 1
Wend

Print #1, vbLf
Print #1, comment_list
If dbc_type <> "J1939" Then
    Print #1, Sheet4.Cells(2, 2)
    Print #1, Sheet4.Cells(3, 2)
Else
    Print #1, Sheet5.Cells(2, 2)
    Print #1, Sheet5.Cells(3, 2)
End If
Print #1, "BA_ ""DBName"" """ & fso.GetBaseName(Filename) + """;" + vbLf
Print #1, cycle_time_list
Print #1, initial_value_list
Print #1, value_table_list

Close #1


str = "DBC File= " + fso.GetFileName(Filename) + vbLf
str = str + "ECU Nodes Count= " + CStr(node_count) + vbLf
str = str + "Messages Count= " + CStr(message_count) + vbLf
str = str + "Signals Count= " + CStr(signal_count)
'ActiveSheet.Cells(1, 5) = str

Set fso = Nothing

MsgBox "Finish, " + GetElapsedTime(starttime, "elapsed time")

End Sub

Private Function Num2Str(ByVal num) As String

Dim str As String
str = CStr(num)
If Len(str) > 0 And Mid(str, 1, 1) = "." Then
    str = "0" & str
ElseIf Len(str) > 0 And Mid(str, 1, 2) = "-." Then
     str = "-0." & Mid(str, 3, Len(str) - 2)
End If
Num2Str = str

End Function

Private Function GetElapsedTime(ByVal starttime As Date, ByVal step As String) As String
Dim text As String
Dim elapsed As Double
Dim endtime As Date

endtime = Now
elapsed = endtime - starttime
text = step + ": " + Format(elapsed * 3600 * 24, "#0") + "s"
GetElapsedTime = text

End Function

Function Hex2Dec(h)
    h = Mid(h, 3, Len(h) - 2)
    Dim L As Long: L = Len(h)
    If L < 16 Then               ' CDec results in Overflow error for hex numbers above 16 ^ 8
        Hex2Dec = CDec("&h0" & h)
        If Hex2Dec < 0 Then Hex2Dec = Hex2Dec + 4294967296# ' 2 ^ 32
    ElseIf L < 25 Then
        Hex2Dec = Hex2Dec(Left$(h, L - 9)) * 68719476736# + CDec("&h" & Right$(h, 9)) ' 16 ^ 9 = 68719476736
    End If
End Function

Sub excel2dbc(control As IRibbonControl)
dbc_Click
End Sub
