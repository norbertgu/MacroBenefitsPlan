Attribute VB_Name = "Module2"
Function e(n As String) As Boolean
' check if worksheet exist
  e = False
  For Each ws In Worksheets
    If n = ws.Name Then
      e = True
      Exit Function
    End If
  Next ws
End Function
Sub CheckSheet()

On Error Resume Next
chksht2:
    Set wSheet = Sheets(SheetToChk)
    If wSheet Is Nothing Then
        SheetToChk = InputBox(prompt:="The sheet " & SheetToChk & "does not exist.. CREATE it BEFORE RUNNING THE MACRO ")
'        MsgBox " check the visa fee "
        SheetToChk = "XXXX"
    End If

End Sub
Sub CheckSheetE()

On Error Resume Next
chksht2:
    Set wSheet = Sheets(SheetToChk)
    If wSheet Is Nothing Then GoTo endchk
'    SheetToChk = MsgBox(prompt:="The sheet " & SheetToChk & " already exist.. DELETE it BEFORE RUNNING THE MACRO ")
    SheetToChk = "XXXX"
       
endchk:
End Sub
Function LastRow()

LastRow = Range("A65536").End(xlUp).Row

End Function


Function NLastRow(n As Integer)

Dim sh As Worksheet
Set sh = ThisWorkbook.ActiveSheet
'Dim NLastRow As Long
NLastRow = Cells(Rows.Count, n).End(xlUp).Row

End Function
Function NLastCol(n As Integer)

Dim sh As Worksheet
Set sh = ThisWorkbook.ActiveSheet
'Dim NLastCol As Long
NLastCol = Cells(n, Columns.Count).End(xlToLeft).Column

End Function
Sub WordDec()
                     tlen = Len(tword)
                     tform = tword
                     tform = Replace(tform, ",", "")
                     pos0 = 1
                     pos1 = InStr(pos0, tform, " ")
                     If pos1 > 0 Then
                              tword1 = Mid(tform, pos0, pos1 - pos0)
                              pos2 = InStr(pos1 + 1, tform, " ")
                              If pos2 > 0 Then
                                             tword2 = Mid(tform, pos1 + 1, pos2 - pos1 - 1)
                                             pos3 = InStr(pos2 + 1, tform, " ")
                                             If pos3 > 0 Then
                                                  tword3 = Mid(tform, pos2 + 1, pos3 - pos2 - 1)
                                               Else
                                                  tword3 = Mid(tform, pos2 + 1)
                                             End If
                                         Else
                                          tword2 = Mid(tform, pos1 + 1)
                                         End If
                                Else
                                   tword1 = tword
                     End If
                     
                     If tword3 = "" Then tword3 = tword2
                     

End Sub
Sub TexttoNum()
'format the column employee to num if needed
    Cells(1, 99).Select
    ActiveCell.FormulaR1C1 = "1"
    Selection.Copy
    'to change according to the target range
    empidcolx = 1
    lRowEmpx = 10
    Range(Cells(1, empidcolx), Cells(lRowEmpx, empidcolx)).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
    Cells(1, 99).Value = ""
End Sub

Function ReplaceClean1(sText As String, Optional sSubText As String = " ")
    Dim J As Integer
    Dim vAddText

    vAddText = Array(Chr(129), Chr(141), Chr(143), Chr(144), Chr(157))
    For J = 1 To 31
        sText = Replace(sText, Chr(J), sSubText)
    Next
    For J = 0 To UBound(vAddText)
        sText = Replace(sText, vAddText(J), sSubText)
    Next
    ReplaceClean1 = sText
End Function
