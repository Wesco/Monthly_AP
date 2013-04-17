Attribute VB_Name = "All_Helper_Functions"
Option Explicit
'Pauses for x# of milliseconds
'Used for email function to prevent
'all emails from being sent at once
'Example: "Sleep 1500" will pause for 1.5 seconds
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Used when importing 117 to determine the type of report to pull
Enum ReportType
    DS
    Bo
End Enum

'---------------------------------------------------------------------------------------
' Proc  : Function FileExists
' Date  : 10/10/2012
' Type  : Boolean
' Desc  : Checks if a file exists
' Ex    : FileExists "C:\autoexec.bat"
'---------------------------------------------------------------------------------------
Function FileExists(ByVal sPath As String) As Boolean
    'Remove trailing backslash
    If InStr(Len(sPath), sPath, "\") > 0 Then sPath = Left(sPath, Len(sPath) - 1)
    'Check to see if the directory exists and return true/false
    If Dir(sPath, vbDirectory) <> "" Then FileExists = True
End Function

'---------------------------------------------------------------------------------------
' Proc  : Function FolderExists
' Date  : 10/10/2012
' Type  : Boolean
' Desc  : Checks if a folder exists
' Ex    : FolderExists "C:\Program Files\"
'---------------------------------------------------------------------------------------
Function FolderExists(ByVal sPath As String) As Boolean
    'Add trailing backslash
    If InStr(Len(sPath), sPath, "\") = 0 Then sPath = sPath & "\"
    'If the folder exists return true
    If Dir(sPath, vbDirectory) <> "" Then FolderExists = True
End Function

'---------------------------------------------------------------------------------------
' Proc  : Sub RecMkDir
' Date  : 10/10/2012
' Desc  : Creates an entire directory tree
' Ex    : RecMkDir "C:\Dir1\Dir2\Dir3\"
'---------------------------------------------------------------------------------------
Sub RecMkDir(ByVal sPath As String)
    Dim sDirArray() As String   'Folder names
    Dim sDrive As String        'Base drive
    Dim sNewPath As String      'Path builder
    Dim i As Long               'Counter

    'Add trailing slash
    If Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If

    'Split at each \
    sDirArray = Split(sPath, "\")
    sDrive = sDirArray(0) & "\"

    'Loop through each directory
    For i = 1 To UBound(sDirArray) - 1
        If Len(sNewPath) = 0 Then
            sNewPath = sDrive & sNewPath & sDirArray(i) & "\"
        Else
            sNewPath = sNewPath & sDirArray(i) & "\"
        End If

        If Not FolderExists(sNewPath) Then
            MkDir sNewPath
        End If
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc  : Sub Email
' Date  : 10/11/2012
' Desc  : Sends an email
' Ex    : Email SendTo:=email@example.com, Subject:="example email", Body:="Email Body"
'---------------------------------------------------------------------------------------
Sub Email(SendTo As String, Optional CC As String, Optional BCC As String, Optional Subject As String, Optional Body As String, Optional Attachment As Variant)
    Dim s As Variant              'Attachment string if array is passed
    Dim Mail_Object As Variant    'Outlook application object
    Dim Mail_Single As Variant    'Email object

    Set Mail_Object = CreateObject("Outlook.Application")
    Set Mail_Single = Mail_Object.CreateItem(0)

    With Mail_Single
        'Add attachments
        Select Case TypeName(Attachment)
            Case "Variant()"
                For Each s In Attachment
                    If s <> Empty Then
                        If FileExists(s) = True Then
                            Mail_Single.attachments.Add s
                        End If
                    End If
                Next
            Case "String"
                If Attachment <> Empty Then
                    If FileExists(Attachment) = True Then
                        Mail_Single.attachments.Add Attachment
                    End If
                End If
        End Select

        'Setup email
        .Subject = Subject
        .To = SendTo
        .CC = CC
        .BCC = BCC
        .HTMLbody = Body
        On Error GoTo SEND_FAILED
        .Send
        On Error GoTo 0
    End With

    'Give the email time to send
    Sleep 1500
    Exit Sub

SEND_FAILED:
    With Mail_Single
        MsgBox "Mail to '" & .To & "' could not be sent."
        .Delete
    End With
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Proc  : Sub ImportGaps
' Date  : 12/12/2012
' Desc  : Imports gaps to the workbook containing this macro.
' Ex    : ImportGaps
'---------------------------------------------------------------------------------------
Sub ImportGaps()
    Dim sPath As String     'Gaps file path
    Dim sName As String     'Gaps Sheet Name
    Dim iCounter As Long    'Counter to decrement the date
    Dim iRows As Long       'Total number of rows
    Dim dt As Date          'Date for gaps file name and path
    Dim Result As VbMsgBoxResult    'Yes/No to proceed with old gaps file if current one isn't found
    Dim Gaps As Worksheet           'The sheet named gaps if it exists, else this = nothing
    Dim StartTime As Double         'The time this function was started
    Dim FileFound As Boolean        'Indicates whether or not gaps was found

    StartTime = Timer
    dt = Date - iCounter
    sPath = "\\br3615gaps\gaps\3615 Gaps Download\" & Format(dt, "yyyy") & "\"
    sName = "3615 " & Format(dt, "m-dd-yy") & ".xlsx"
    FileFound = False

    'This error is bypassed so you can determine whether or not the sheet exists
    On Error GoTo CREATE_GAPS
    Set Gaps = ThisWorkbook.Sheets("Gaps")
    On Error GoTo 0

    Application.DisplayAlerts = False

    'Find gaps
    For iCounter = 0 To 15
        dt = Date - iCounter
        sPath = "\\br3615gaps\gaps\3615 Gaps Download\" & Format(dt, "yyyy") & "\"
        sName = "3615 " & Format(dt, "m-dd-yy") & ".xlsx"
        If FileExists(sPath & sName) Then
            FileFound = True
            Exit For
        End If
    Next

    'Make sure Gaps file was found
    If FileFound = True Then
        If dt <> Date Then
            Result = MsgBox( _
                     Prompt:="Gaps from " & Format(dt, "mmm dd, yyyy") & " was found." & vbCrLf & "Would you like to continue?", _
                     Buttons:=vbYesNo, _
                     Title:="Gaps not up to date")
        End If

        If Result <> vbNo Then
            If ThisWorkbook.Sheets("Gaps").Range("A1").Value <> "" Then
                Gaps.Cells.Delete
            End If

            Workbooks.Open sPath & sName
            ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Gaps").Range("A1")
            ActiveWorkbook.Close

            Sheets("Gaps").Select
            iRows = ActiveSheet.UsedRange.Rows.Count
            Columns(1).EntireColumn.Insert
            Range("A1").Value = "SIM"
            Range("A2").Formula = "=C2&D2"
            Range("A2").AutoFill Destination:=Range(Cells(2, 1), Cells(iRows, 1))
            Range(Cells(2, 1), Cells(iRows, 1)).Value = Range(Cells(2, 1), Cells(iRows, 1)).Value

            FillInfo FunctionName:="Gaps", _
                     FileDate:=Format(dt, "mm/dd/yy"), _
                     Parameters:="", _
                     ExecutionTime:=Timer - StartTime, _
                     Result:="Complete"
        Else
            FillInfo FunctionName:="Gaps", _
                     FileDate:=Format(dt, "mm/dd/yy"), _
                     Parameters:="", _
                     ExecutionTime:=Timer - StartTime, _
                     Result:="Failed - User Aborted"
            ERR.Raise 18
        End If
    Else
        MsgBox Prompt:="Gaps could not be found.", Title:="Gaps not found"
        FillInfo FunctionName:="Gaps", _
                 FileDate:=Format(dt, "mm/dd/yy"), _
                 Parameters:="", _
                 ExecutionTime:=Timer - StartTime, _
                 Result:="Failed - Gaps not found"
        ERR.Raise 53
    End If

    Application.DisplayAlerts = True
    Exit Sub

CREATE_GAPS:
    ThisWorkbook.Sheets.Add After:=Sheets(ThisWorkbook.Sheets.Count)
    ActiveSheet.Name = "Gaps"
    Resume

End Sub

'---------------------------------------------------------------------------------------
' Proc : FilterSheet
' Date : 1/29/2013
' Desc : Remove all rows that do not match a specified string
'---------------------------------------------------------------------------------------
Sub FilterSheet(sFilter As String, ColNum As Integer, Match As Boolean)
    Dim Rng As Range
    Dim aRng() As Variant
    Dim aHeaders As Variant
    Dim StartTime As Double
    Dim iCounter As Long
    Dim i As Long
    Dim y As Long

    StartTime = Timer
    Set Rng = ActiveSheet.UsedRange
    aHeaders = Range(Cells(1, 1), Cells(1, ActiveSheet.UsedRange.Columns.Count))
    iCounter = 1

    Do While iCounter <= Rng.Rows.Count
        If Match = True Then
            If Rng(iCounter, ColNum).Value = sFilter Then
                i = i + 1
            End If
        Else
            If Rng(iCounter, ColNum).Value <> sFilter Then
                i = i + 1
            End If
        End If
        iCounter = iCounter + 1
    Loop

    ReDim aRng(1 To i, 1 To Rng.Columns.Count) As Variant

    iCounter = 1
    i = 0
    Do While iCounter <= Rng.Rows.Count
        If Match = True Then
            If Rng(iCounter, ColNum).Value = sFilter Then
                i = i + 1
                For y = 1 To Rng.Columns.Count
                    aRng(i, y) = Rng(iCounter, y)
                Next
            End If
        Else
            If Rng(iCounter, ColNum).Value <> sFilter Then
                i = i + 1
                For y = 1 To Rng.Columns.Count
                    aRng(i, y) = Rng(iCounter, y)
                Next
            End If
        End If
        iCounter = iCounter + 1
    Loop

    ActiveSheet.Cells.Delete
    Range(Cells(1, 1), Cells(UBound(aRng, 1), UBound(aRng, 2))) = aRng
    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, UBound(aHeaders, 2))) = aHeaders
    FillInfo "FilterSheet", _
             "", _
             "Filter: " & sFilter & vbCrLf & "Col: " & Columns(ColNum).Address(False, False) & vbCrLf & "Match: " & Match, _
             Timer - StartTime, _
             "Complete"
End Sub


'---------------------------------------------------------------------------------------
' Proc : UserImportFile
' Date : 1/29/2013
' Desc : Prompts the user to select a file for import
'---------------------------------------------------------------------------------------
Sub UserImportFile(DestRange As Range, DelFile As Boolean)
    Dim StartTime As Double         'The time this function was started
    Dim File As String              'Full path to user selected file
    Dim FileDate As String          'Date the file was last modified
    Dim OldDispAlert As Boolean     'Original state of Application.DisplayAlerts

    OldDispAlert = Application.DisplayAlerts
    StartTime = Timer
    File = Application.GetOpenFilename()

    Application.DisplayAlerts = False
    If File <> "False" Then
        FileDate = Format(FileDateTime(File), "mm/dd/yy")
        Workbooks.Open File

        ActiveSheet.UsedRange.Copy Destination:=DestRange
        ActiveWorkbook.Close
        ThisWorkbook.Activate

        If DelFile = True Then
            DeleteFile File
        End If

        FillInfo FunctionName:="UserImportFile", _
                 Parameters:="FileName: " & File, _
                 FileDate:=FileDate, _
                 ExecutionTime:=Timer - StartTime, _
                 Result:="Complete"

        FillInfo FunctionName:="", _
                 Parameters:="DestRange: " & DestRange.Address(False, False), _
                 Result:="Complete"
    Else
        FillInfo FunctionName:="UserImportFile", _
                 Parameters:="DestRange: " & DestRange.Address(False, False), _
                 ExecutionTime:=Timer - StartTime, _
                 Result:="Failed - User Aborted"
        Sheets("Info").Select
        ERR.Raise 18
    End If

End Sub

'---------------------------------------------------------------------------------------
' Proc : FillInfo
' Date : 1/29/2013
' Desc : Used to add a line to the Info sheet
'---------------------------------------------------------------------------------------
Sub FillInfo(Optional FunctionName As String = "", Optional Result As String = "", Optional ExecutionTime As String = "", Optional Parameters As String = "", Optional FileDate As String = "")
    Dim Info As Worksheet           'Info worksheet if it exists, else this = nothing
    Dim LastSheet As Worksheet      'The previously selected worksheet
    Dim LastWorkbook As Workbook    'The previously activated workbook
    Set LastSheet = ActiveSheet
    Set LastWorkbook = ActiveWorkbook
    Dim Row As Long

    'This error is bypassed so you can determine whether or not the sheet exists
    On Error GoTo CREATE_INFO
    Set Info = ThisWorkbook.Sheets("Info")
    On Error GoTo 0

    ThisWorkbook.Activate
    Sheets("Info").Select
    Range("A1").Value = "Function"
    Range("B1").Value = "Created"
    Range("C1").Value = "Params"
    Range("D1").Value = "Exec Time"
    Range("E1").Value = "Result"

    Row = ActiveSheet.UsedRange.Rows.Count + 1
    Cells(Row, 1).Value = FunctionName
    Cells(Row, 2).Value = FileDate
    Cells(Row, 3).Value = Parameters
    Cells(Row, 4).Value = ExecutionTime
    Cells(Row, 5).Value = Result

    ActiveSheet.UsedRange.Columns.EntireColumn.AutoFit

    LastWorkbook.Activate
    LastSheet.Select
    Exit Sub

CREATE_INFO:
    Sheets.Add After:=Sheets(ThisWorkbook.Sheets.Count)
    ActiveSheet.Name = "Info"
    Resume

End Sub

'---------------------------------------------------------------------------------------
' Proc : ExportCode
' Date : 3/19/2013
' Desc : Exports all modules
'---------------------------------------------------------------------------------------
Sub ExportCode()
    Dim comp As Variant
    Dim codeFolder As String
    Dim FileName As String
    Dim File As String

    'References Microsoft Visual Basic for Applications Extensibility 5.3
    AddReference "{0002E157-0000-0000-C000-000000000046}", 5, 3
    codeFolder = GetWorkbookPath & "Code\" & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & "\"

    On Error Resume Next
    RecMkDir codeFolder
    On Error GoTo 0

    'Remove all previously exported modules
    File = Dir(codeFolder)
    Do While File <> ""
        DeleteFile codeFolder & File
        File = Dir
    Loop

    'Export modules in current project
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1
                FileName = codeFolder & comp.Name & ".bas"
                comp.Export FileName
            Case 2
                FileName = codeFolder & comp.Name & ".cls"
                comp.Export FileName
            Case 3
                FileName = codeFolder & comp.Name & ".frm"
                comp.Export FileName
        End Select
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : ImportModule
' Date : 4/4/2013
' Desc : Imports a code module into VBProject
'---------------------------------------------------------------------------------------
Sub ImportModule()
    Dim comp As Variant
    Dim codeFolder As String
    Dim FileName As String
    Dim WkbkPath As String

    'Adds a reference to Microsoft Visual Basic for Applications Extensibility 5.3
    AddReference "{0002E157-0000-0000-C000-000000000046}", 5, 3

    'Gets the path to this workbook
    WkbkPath = Left$(ThisWorkbook.fullName, InStr(1, ThisWorkbook.fullName, ThisWorkbook.Name, vbTextCompare) - 1)

    'Gets the path to this workbooks code
    codeFolder = WkbkPath & "Code\" & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & "\"

    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Name <> "All_Helper_Functions" Then
            Select Case comp.Type
                Case 1
                    FileName = codeFolder & comp.Name & ".bas"
                    ThisWorkbook.VBProject.VBComponents.Remove comp
                    ThisWorkbook.VBProject.VBComponents.Import FileName
                Case 2
                    FileName = codeFolder & comp.Name & ".cls"
                    ThisWorkbook.VBProject.VBComponents.Remove comp
                    ThisWorkbook.VBProject.VBComponents.Import FileName
                Case 3
                    FileName = codeFolder & comp.Name & ".frm"
                    ThisWorkbook.VBProject.VBComponents.Remove comp
                    ThisWorkbook.VBProject.VBComponents.Import FileName
            End Select
        End If
    Next

End Sub

'---------------------------------------------------------------------------------------
' Proc : DeleteFile
' Date : 3/19/2013
' Desc : Deletes a file
'---------------------------------------------------------------------------------------
Sub DeleteFile(FileName As String, Optional LogEntry As Boolean = False)
    On Error GoTo File_Error
    Kill FileName

    If LogEntry = True Then
        FillInfo FunctionName:="DeleteFile", _
                 Parameters:=FileName, _
                 Result:="Complete"
    End If
    Exit Sub

File_Error:
    If LogEntry = True Then
        FillInfo FunctionName:="DeleteFile", _
                 Result:="Err #: " & ERR.Number
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : GetWorkbookPath
' Date : 3/19/2013
' Desc : Gets the full path of ThisWorkbook
'---------------------------------------------------------------------------------------
Function GetWorkbookPath() As String
    Dim fullName As String
    Dim wrkbookName As String
    Dim pos As Long

    wrkbookName = ThisWorkbook.Name
    fullName = ThisWorkbook.fullName

    pos = InStr(1, fullName, wrkbookName, vbTextCompare)

    GetWorkbookPath = Left$(fullName, pos - 1)
End Function

'---------------------------------------------------------------------------------------
' Proc : EndsWith
' Date : 3/19/2013
' Desc : Checks if a string ends in a specified character
'---------------------------------------------------------------------------------------
Function EndsWith(ByVal InString As String, ByVal TestString As String) As Boolean
    EndsWith = (Right$(InString, Len(TestString)) = TestString)
End Function

'---------------------------------------------------------------------------------------
' Proc : AddReferences
' Date : 3/19/2013
' Desc : Adds a reference to VBProject
'---------------------------------------------------------------------------------------
Sub AddReference(GUID As String, Major As Integer, Minor As Integer)
    Dim ID As Variant
    Dim Ref As Variant
    Dim Result As Boolean


    For Each Ref In ThisWorkbook.VBProject.References
        If Ref.GUID = GUID And Ref.Major = Major And Ref.Minor = Minor Then
            Result = True
        End If
    Next

    'References Microsoft Visual Basic for Applications Extensibility 5.3
    If Result = False Then
        ThisWorkbook.VBProject.References.AddFromGuid GUID, Major, Minor
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : RemoveReferences
' Date : 3/19/2013
' Desc : Removes a reference from VBProject
'---------------------------------------------------------------------------------------
Sub RemoveReference(GUID As String, Major As Integer, Minor As Integer)
    Dim Ref As Variant

    For Each Ref In ThisWorkbook.VBProject.References
        If Ref.GUID = GUID And Ref.Major = Major And Ref.Minor = Minor Then
            Application.VBE.ActiveVBProject.References.Remove Ref
        End If
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : ShowReferences
' Date : 4/4/2013
' Desc : Lists all VBProject references
'---------------------------------------------------------------------------------------
Sub ShowReferences()
    Dim i As Variant
    Dim n As Integer

    ThisWorkbook.Activate
    On Error GoTo SHEET_EXISTS
    Sheets("VBA References").Select
    ActiveSheet.Cells.Delete
    On Error GoTo 0

    [A1].Value = "Name"
    [B1].Value = "Description"
    [C1].Value = "GUID"
    [D1].Value = "Major"
    [E1].Value = "Minor"

    For i = 1 To ThisWorkbook.VBProject.References.Count
        n = i + 1
        With ThisWorkbook.VBProject.References(i)
            Cells(n, 1).Value = .Name
            Cells(n, 2).Value = .Description
            Cells(n, 3).Value = .GUID
            Cells(n, 4).Value = .Major
            Cells(n, 5).Value = .Minor
        End With
    Next
    Columns.AutoFit

    Exit Sub

SHEET_EXISTS:
    ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count), Count:=1
    ActiveSheet.Name = "VBA References"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : Import117byISN
' Date : 4/10/2013
' Desc : Imports the most recent 117 report for the specified sales number
'---------------------------------------------------------------------------------------
Sub Import117byISN(RepType As ReportType, Destination As Range, Optional ByVal ISN As String = "", Optional Cancel As Boolean = False)
    Dim sPath As String
    Dim FileName As String

    If ISN = "" And Cancel = False Then
        ISN = InputBox("Inside Sales Number:", "Please enter the ISN#")
    Else
        If ISN = "" Then
            FillInfo "Import117byISN", "Failed - User Aborted", Parameters:="ReportType: " & ReportTypeText(RepType)
            ERR.Raise 53
        End If
    End If

    If ISN <> "" Then
        Select Case RepType
            Case ReportType.DS:
                FileName = "3615 " & Format(Date, "m-dd-yy") & " DSORDERS.xlsx"

            Case ReportType.Bo:
                FileName = "3615 " & Format(Date, "m-dd-yy") & " BACKORDERS.xlsx"
        End Select

        sPath = "\\br3615gaps\gaps\3615 117 Report\ByInsideSalesNumber\" & ISN & "\" & FileName

        If FileExists(sPath) Then
            Workbooks.Open sPath
            ActiveSheet.UsedRange.Copy Destination:=Destination
            Application.DisplayAlerts = False
            ActiveWorkbook.Close
            Application.DisplayAlerts = True

            FillInfo FunctionName:="Import117byISN", _
                     Parameters:="Sales #: " & ISN, _
                     Result:="Complete"
            FillInfo Parameters:="Report Type: " & ReportTypeText(RepType)
            FillInfo Parameters:="Destination: " & Destination.Address(False, False)
        Else
            FillInfo FunctionName:="Import117byISN", _
                     Parameters:="Sales #: " & ISN, _
                     Result:="Failed - File not found"
            FillInfo Parameters:="Report Type: " & ReportTypeText(RepType)
            FillInfo Parameters:="Destination: " & Destination.Address(False, False)
            MsgBox Prompt:=ReportTypeText(RepType) & " report not found.", Title:="Error 53"
        End If
    Else
        FillInfo "Import117byISN", "Failed - Missing ISN", Parameters:="ReportType: " & ReportTypeText(RepType)
        ERR.Raise 18
    End If

End Sub

'---------------------------------------------------------------------------------------
' Proc : Import473
' Date : 4/11/2013
' Desc : Imports a 473 report from the current day
'---------------------------------------------------------------------------------------
Sub Import473(Destination As Range)
    Dim sPath As String
    Dim FileName As String
    Dim AlertStatus As Boolean

    FileName = "473 " & Format(Date, "m-dd-yy") & ".xlsx"
    sPath = "\\br3615gaps\gaps\473 Download\" & FileName
    AlertStatus = Application.DisplayAlerts

    If FileExists(sPath) Then
        Workbooks.Open sPath
        ActiveSheet.UsedRange.Copy Destination:=Destination

        Application.DisplayAlerts = False
        ActiveWorkbook.Close
        Application.DisplayAlerts = AlertStatus

        FillInfo FunctionName:="Import473", _
                 Parameters:=Destination.Address(False, False), _
                 Result:="Complete"
    Else
        FillInfo FunctionName:="Import473", _
                 Parameters:="Destination: " & Destination.Address(False, False), _
                 Result:="Failed - File not found"
        MsgBox Prompt:="473 report not found."
        ERR.Raise 18
    End If

End Sub

'---------------------------------------------------------------------------------------
' Proc : ReportTypeText
' Date : 4/10/2013
' Desc : Returns the report type as a string
'---------------------------------------------------------------------------------------
Function ReportTypeText(RepType As ReportType) As String
    Select Case RepType
        Case ReportType.Bo:
            ReportTypeText = "BO"
        Case ReportType.DS:
            ReportTypeText = "DS"
    End Select
End Function

'---------------------------------------------------------------------------------------
' Proc : DeleteColumn
' Date : 4/11/2013
' Desc : Removes a column based on text in the column header
'---------------------------------------------------------------------------------------
Sub DeleteColumn(HeaderText As String)
    Dim i As Integer

    For i = ActiveSheet.UsedRange.Columns.Count To 1 Step -1
        If Trim(Cells(1, i).Value) = HeaderText Then
            Columns(i).Delete
            Exit For
        End If
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : FindColumn
' Date : 4/11/2013
' Desc : Returns the column number if a match is found
'---------------------------------------------------------------------------------------
Function FindColumn(HeaderText As String, Optional SearchArea As Range) As Integer
    Dim i As Integer: i = 0
    
    If TypeName(SearchArea) = Empty Then
        SearchArea = ActiveSheet.UsedRange
    End If
    
    For i = 1 To SearchArea.Columns.Count
        If Trim(SearchArea.Cells(1, i).Value) = HeaderText Then
            FindColumn = i
            Exit For
        End If
    Next
End Function



