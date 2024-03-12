Attribute VB_Name = "Sheet_Loop"
Option Explicit

Sub Sheet__Loop()
Dim rng As Range
Dim objOL As Object
Dim objMI As Object
Dim Recip As Object

If Range("P1") <> "NONE" Then

    Worksheets("AUDIT").Range("A6:G5000").AutoFilter Field:=1, Criteria1:=2
    Worksheets("AUDIT").Range("A6:G5000").AutoFilter Field:=6, Criteria1:=""
    Range("A6", Range("A6").End(xlToRight).End(xlToRight).End(xlDown)).Select
    Selection.Copy
    Sheets.Add.Name = "PlaceHolder"
    Worksheets("PlaceHolder").Select
    Range("A1").Select
    ActiveSheet.Paste
    Worksheets("PlaceHolder").Range("A1:G1").Columns.AutoFit
    Set rng = Range("A1:G40")
    
    ' Outlook application object
    Set objOL = CreateObject("Outlook.Application")
    ' Create e-mail message
    Set objMI = objOL.CreateItem(0) ' 0 = olMailItem
    ' Set some properties and send the message
        With objMI
            .To = "destination@email.com"
            .Cc = "appropriate@email.com"
            .Subject = "Mock Subject"
            .HTMLBody = "These account(s) are still missing from the database. Thanks! " & Chr(13) & RangetoHTML(rng)
            .Display
        End With
    
    Sheets("PlaceHolder").Delete
    Worksheets("AUDIT").AutoFilter.ShowAllData


End If

End Sub

'This Function will generate the appropriate information for the body of the email.
Function RangetoHTML(rng As Range)

    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
    
End Function


