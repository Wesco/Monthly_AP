Attribute VB_Name = "All_Helper_Functions"
Option Explicit

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
    Dim WkbkPath As String


    'References Microsoft Visual Basic for Applications Extensibility 5.3
    AddReference "{0002E157-0000-0000-C000-000000000046}", 5, 3
    WkbkPath = Left$(ThisWorkbook.fullName, InStr(1, ThisWorkbook.fullName, ThisWorkbook.Name, vbTextCompare) - 1)
    codeFolder = WkbkPath & "Code\" & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & "\"

    On Error Resume Next
    If Dir(codeFolder) = "" Then
        RecMkDir codeFolder
    End If
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
            Case 100
                If comp.Name = "ThisWorkbook" Then
                    FileName = codeFolder & comp.Name & ".bas"
                    comp.Export FileName
                End If
        End Select
    Next
End Sub

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
