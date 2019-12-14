Attribute VB_Name = "Module1"
'change 141219
Public target As String
Public source As String
Public invoice As String
Public templateN As String
Public MyMainSheet As String
Public MyNewbook As Workbook

Sub A_BenefitPlanReport()
Attribute A_BenefitPlanReport.VB_ProcData.VB_Invoke_Func = "b\n14"
'
' Macro to create Benefit Plan per customer
'

    Dim MyWorkbook As Workbook
    Dim MyWb       As Workbook
    Dim MyPath As Variant
    Dim FilePath, FileOnly, PathOnly, FileNameN As String

    FilePath = ActiveWorkbook.FullName
    FileOnly = ActiveWorkbook.Name
    FileNameN = Left(FileOnly, Len(FileOnly) - 5)
    PathOnly = Left(FilePath, Len(FilePath) - Len(FileOnly))
    
     Set MyWorkbook = ActiveWorkbook
     MyWbName = MyWorkbook.Name
     MyPath = ThisWorkbook.Path
     MyMainSheet = ActiveSheet.Name
     runDate = Date
     
     runD = Application.WorksheetFunction.Text(Day(Date), "00")
     runM = Application.WorksheetFunction.Text(Month(Date), "00")
     runyt = Application.WorksheetFunction.Text(Year(Date), "0000")
     runY = Mid(runyt, 3, 2)
     
     Set MyNewbookT = Workbooks.Add
     MyNewbookT.Activate
     MyNewWbT = MyNewbookT.Name
     ActiveSheet.DisplayRightToLeft = False
     
     MyWorkbook.Activate
     Sheets("Workers").Select
     Selection.AutoFilter
     Cells.Select
     Selection.Copy
     
     MyNewbookT.Activate
     Sheets(1).Select
     Range("A1").Select
     ActiveSheet.Paste
     ActiveSheet.Name = "Workers"
'    Sheets("Workers").Select
    
    '''
    ''' cleaning of the lines and columns not necessary
    ''
    Rows("1:3").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("L:M").Select
    Selection.Delete Shift:=xlToLeft
    
    lrow1 = NLastRow(1)
    
    '''
    ''' list of customer
    '''
    Range(Cells(1, 2), Cells(lrow1, 2)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range( _
        "z1:z1"), Unique:=True
    lrow26 = NLastRow(26)
    
    Range(Cells(1, 26), Cells(lrow26, 26)).Select
    ActiveWorkbook.Worksheets("Workers").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Workers").Sort.SortFields.Add Key:=Range(Cells(2, 26), Cells(lrow26, 26)) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Workers").Sort
        .SetRange Range(Cells(2, 26), Cells(lrow26, 26))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    '''
    ''' autofilter by customer
    '''
    For i = 2 To lrow26
    
          MyNewbookT.Activate
          Sheets("Workers").Select
          cust = Cells(i, 26).Value
          Cells.Select
          Selection.AutoFilter
'          ActiveSheet.Range("$A$1:$Q$685").AutoFilter Field:=1, Criteria1:= _
'              "Bright Machines"
          ActiveSheet.Range(Cells(1, 1), Cells(lrow1, 11)).AutoFilter Field:=2, Criteria1:="=" & cust & " "
          lrowf1 = NLastRow(1)
'          Range("A1:K474").Select
          Range(Cells(1, 1), Cells(lrowf1, 11)).Select
          Selection.Copy
          Sheets.Add After:=ActiveSheet
          Range("A1").Select
          ActiveSheet.Paste
'          pos1 = InStr(1, cust, ",")
          pos1 = 0
          name1 = Mid(cust, pos1 + 1, 30)
          ActiveSheet.Name = name1
          Columns("A:A").EntireColumn.AutoFit
          Columns("B:B").EntireColumn.AutoFit
          Columns("D:D").EntireColumn.AutoFit
          Columns("E:E").EntireColumn.AutoFit
          
          Cells.Select
          ActiveWorkbook.Worksheets(name1).Sort.SortFields.Clear
           ActiveWorkbook.Worksheets(name1).Sort.SortFields.Add2 Key:=Range( _
              Cells(2, 5), Cells(lrowf1, 5)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
              xlSortNormal
          ActiveWorkbook.Worksheets(name1).Sort.SortFields.Add2 Key:=Range( _
              Cells(2, 3), Cells(lrowf1, 3)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
              xlSortNormal
          With ActiveWorkbook.Worksheets(name1).Sort
'              .SetRange Range("A1:K474")
              .SetRange Range(Cells(1, 1), Cells(lrowf1, 11))
              .Header = xlYes
              .MatchCase = False
              .Orientation = xlTopToBottom
              .SortMethod = xlPinYin
              .Apply
          End With
          
          Set MyNewbook = Workbooks.Add
          MyNewbook.Activate
          MyNewWb = MyNewbook.Name
          ActiveSheet.DisplayRightToLeft = False
          
          MyNewbookT.Activate
          Sheets(name1).Select
'          Sheets(cust).Copy
          Cells.Select
          Selection.Copy
          
          MyNewbook.Activate
          Sheets(1).Select
          Range("A1").Select
          ActiveSheet.Paste
          ActiveSheet.Name = name1
              
         MyNewbookN = "Benefit_plans_" & name1 & "-" & runD & runM & runY
         Dim temp2 As String
         temp2 = MyNewbookN
         With MyNewbook
             .SaveAs Filename:=PathOnly & temp2 & ".xlsx"
         End With
         
'         MyNewWbn = MyNewbook.Name
         MyNewbook.Close SaveChanges:=False
     Next i
    
    '''
    '''

    
End Sub

