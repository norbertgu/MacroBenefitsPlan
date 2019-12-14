Attribute VB_Name = "Module3"
Sub various()

''''
'''' autofilter
''''
'    Sheets("AccountingCodes").Select
'    Range("A1").Select
'    ActiveSheet.Paste
'    lrowm = LastRow
'    ' filter the rules for the specific project
'    Cells.Select
'    Selection.AutoFilter
'    ActiveSheet.Range(Cells(1, 1), Cells(lrowm, 4)).AutoFilter Field:=3, Criteria1:="=" & proj & " "
'
'    ' delete the rows non selected
''    Dim lRows As Long
'    For lrowS = ActiveSheet.UsedRange.Rows.Count To 1 Step -1
'        If Cells(lrowS, 1).EntireRow.Hidden = True Then Cells(lrowS, 1).EntireRow.Delete
'    Next lrowS



''''
'''' distinct list
''''
                     
'    Range(Cells(1, 7), Cells(trow, 8)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range( _
'        "M1:N1"), Unique:=True
'    Range(Cells(1, 13), Cells(lvrow, 14)).Select
'    ActiveWorkbook.Worksheets("target").Sort.SortFields.Clear
'    ActiveWorkbook.Worksheets("target").Sort.SortFields.Add Key:=Range(Cells(2, 13), Cells(lvrow, 13)) _
'    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    With ActiveWorkbook.Worksheets("target").Sort
'        .SetRange Range(Cells(2, 13), Cells(lvrow, 14))
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
    
''''
''''  save workbook
''''
'savebook:
'    MyNewbookN = Pname & "SP" & runD & runM & runY
' check project name for special characters like in Bright Zhuai
'
'    MyNewbookN = Pname & "-" & runD & runM & runY & "-Import Data"
'    Dim temp2 As String
'    temp2 = MyNewbookN
''    temp1 = Replace(MyNewbookN, ",", " ")
'    temp2 = ReplaceClean1(temp2)
''    temp = PathOnly & MyNewbookN & ".xlsx"
''    k = Len(temp)
'    With MyNewbook
'        .SaveAs Filename:=PathOnly & temp2 & ".xlsx"
'    End With
'    MyNewWbn = MyNewbook.Name

''''
''''  sort
''''
'     ActiveWorkbook.Worksheets("target").Sort.SortFields.Clear
'    ActiveWorkbook.Worksheets("target").Sort.SortFields.Add Key:=Range(Cells(2, 13), Cells(lvrow, 13)) _
'    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    With ActiveWorkbook.Worksheets("target").Sort
''        .SetRange Range("M1:N14")
'        .SetRange Range(Cells(2, 13), Cells(lvrow, 14))
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'    Range(Cells(1, 13), Cells(lRItm, 14)).Select
'      ActiveWorkbook.Worksheets("sourceN").Sort.SortFields.Clear
'     ActiveWorkbook.Worksheets("sourceN").Sort.SortFields.Add2 Key:=Range(Cells(1, 14), Cells(lRItm, 14)) _
'     , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'     ActiveWorkbook.Worksheets("sourceN").Sort.SortFields.Add2 Key:=Range(Cells(1, 13), Cells(lRItm, 13)) _
'     , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
     
'     With ActiveWorkbook.Worksheets("sourceN").Sort
'           .SetRange Range(Cells(1, 13), Cells(lRItm, 14))
'           .Header = xlYes
'           .MatchCase = False
'           .Orientation = xlTopToBottom
'           .SortMethod = xlPinYin
'           .Apply
'      End With


End Sub
