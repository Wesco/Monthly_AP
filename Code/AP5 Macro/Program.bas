Attribute VB_Name = "Program"
Option Explicit
Public Const VersionNumber As String = "1.0.2"
Public Const RepositoryName As String = "Monthly_AP"

'---------------------------------------------------------------------------------------
' Procedure : Macro1
' Author    : TReische
' Date      : 1/4/2013
' Purpose   : Remove blanks, headers, and duplicates, then add a basic data summary
'---------------------------------------------------------------------------------------
'
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim Wkbk As Workbook    'Workbook containing finished report
    Dim s As Worksheet      'Copy of
    Dim iRows As Long       'Number of used rows
    Dim i As Long           'Loop Counter
    Dim x As Long           'Formula variable

    Application.ScreenUpdating = False

    Sheets("DropIn").Select
    iRows = ActiveSheet.UsedRange.Rows.Count

    'Separates into columns
    Range(Cells(1, 1), Cells(iRows, 1)).TextToColumns _
            Destination:=Range("A1"), _
            DataType:=xlFixedWidth, _
            FieldInfo:= _
            Array( _
            Array(0, 1), _
            Array(2, 1), _
            Array(6, 1), _
            Array(13, 1), _
            Array(21, 1), _
            Array(35, 1), _
            Array(45, 1), _
            Array(51, 1), _
            Array(58, 1), _
            Array(65, 1), _
            Array(76, 1)), _
            TrailingMinusNumbers:=True

    Range(Cells(1, 1), Cells(iRows, 1)).Delete Shift:=xlToLeft
    Rows("1:2").Delete

    FilterSheet "507-01", 2, True

    'Concatenate by removing headers and blanks
    Range("A2").Select
    i = 2
    Do While i <= ActiveSheet.UsedRange.Rows.Count
        Select Case Cells(i, 1).Value
            Case "AP10"
                Rows(i).Delete
            Case "BR."
                Rows(i).Delete
            Case ""
                Rows(i).Delete
            Case Else
                i = i + 1
        End Select
    Loop
    iRows = ActiveSheet.UsedRange.Rows.Count

    'Convert column data from MDY to m/d/yyyy
    Range(Cells(2, 7), Cells(iRows, 7)).TextToColumns Destination:=Range("G2"), _
                                                      DataType:=xlDelimited, _
                                                      TextQualifier:=xlDoubleQuote, _
                                                      ConsecutiveDelimiter:=False, _
                                                      Tab:=True, _
                                                      Semicolon:=False, _
                                                      Comma:=False, _
                                                      Space:=False, _
                                                      Other:=False, _
                                                      FieldInfo:=Array(1, 3), _
                                                      TrailingMinusNumbers:=True
    Range(Cells(2, 6), Cells(iRows, 6)).NumberFormat = "m/d/yyyy"
    Columns("G:G").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit

    'Sort and remove duplicates
    ActiveWorkbook.Worksheets("DropIn").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("DropIn").Sort.SortFields.Add Key:=Range("E1"), _
                                                            SortOn:=xlSortOnValues, _
                                                            Order:=xlAscending, _
                                                            DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DropIn").Sort
        .SetRange Range(Cells(2, 1), Cells(iRows, 10))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.Range(Cells(2, 1), Cells(iRows, 10)).RemoveDuplicates Columns:=5, Header:=xlYes
    iRows = Rows(Rows.Count).End(xlUp).Row

    'Fix the year if needed
    For i = 2 To iRows
        If CDate(Format(Cells(i, 6).Text, "yyyy-mm-dd")) > CDate(Format(Date, "yyyy-mm-dd")) Then
            Cells(i, 6).Value = Format(Cells(i, 6).Text, "m/d") & "/" & Year(Date) - 1
        End If
    Next
    iRows = ActiveSheet.UsedRange.Rows.Count

    Range("K2").Formula = "=F2-G2"
    Range("K2").AutoFill Destination:=Range(Cells(2, 11), Cells(iRows, 11))
    Range(Cells(2, 11), Cells(iRows, 11)).Value = Range(Cells(2, 11), Cells(iRows, 11)).Value

    ActiveWorkbook.Worksheets("DropIn").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("DropIn").Sort.SortFields.Add Key:=Range("K1"), _
                                                            SortOn:=xlSortOnValues, _
                                                            Order:=xlDescending, _
                                                            DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DropIn").Sort
        .SetRange Range(Cells(2, 1), Cells(iRows, 11))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Highlight and count cells greater than 20
    Range("K1").Value = "Days"
    iRows = Rows(Rows.Count).End(xlUp).Row
    For i = 2 To iRows
        If Cells(i, 11).Value >= 15 Then
            x = x + 1
            Cells(i, 11).Interior.Color = RGB(255, 255, 0)
        End If
    Next
    iRows = ActiveSheet.UsedRange.Rows.Count

    'Add data summary
    Cells(iRows + 1, 1).Value = "# over 15:"
    Cells(iRows + 1, 1).Interior.Color = RGB(255, 255, 0)
    Cells(iRows + 1, 2).Value = x
    Cells(iRows + 1, 2).Interior.Color = RGB(255, 255, 0)

    Cells(iRows + 2, 1).Value = "% of total:"
    Cells(iRows + 2, 1).Interior.Color = RGB(255, 255, 0)
    Cells(iRows + 2, 2).Value = Round(((x / iRows) * 100), 2)
    Cells(iRows + 2, 2).Interior.Color = RGB(255, 255, 0)

    'Save results in a new workbook
    Sheets("DropIn").Copy
    Set Wkbk = ActiveWorkbook
    Set s = ActiveSheet
    s.Name = "Sheet1"
    Application.Dialogs(xlDialogSaveAs).Show
    Cells(iRows, 1).Select

    'Clean this workbook
    ThisWorkbook.Activate
    Application.DisplayAlerts = False
    Sheets("DropIn").Cells.Delete
    Application.DisplayAlerts = True
    Range("A1").Select
    Sheets("Macro").Select
    Range("C7").Select
    Wkbk.Activate

    Application.ScreenUpdating = True
End Sub
