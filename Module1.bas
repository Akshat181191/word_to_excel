Attribute VB_Name = "Module1"
Option Explicit
Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)

' This section turns on high overhead operations
Sub AppTrue()
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

' This section increases performance by turning off high overhead operations
Sub AppFalse()
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

End Sub


Sub TransferTables()

Call AppFalse

Dim UserInterface As Worksheet
Set UserInterface = ThisWorkbook.Sheets("User Interface")
Dim Word As Word.Application
Dim TargetWB As Workbook
Dim TargetSheet As Worksheet
Dim SourceDoc As Document
Dim Table As Table
Dim Excel As Excel.Application
Set TargetWB = ThisWorkbook
Set Excel = TargetWB.Application
Set Word = New Word.Application

'clean up from last run
On Error Resume Next
Set SourceDoc = GetObject(UserInterface.OLEObjects("TextBox21").Object.Text)
Set Word = SourceDoc.Application
    'set saved = true so no pop up
SourceDoc.Saved = True
SourceDoc.Close SaveChanges:=False
Word.Application.Quit
On Error GoTo 0

'set word objects
Set Word = New Word.Application
Set SourceDoc = Word.Documents.Open(UserInterface.OLEObjects("TextBox21").Object.Text)
Word.Visible = True

'set excel objects
Set TargetWB = ThisWorkbook
Set Excel = TargetWB.Application
Set TargetSheet = TargetWB.Sheets.Add(, Sheets("User Interface"))
On Error Resume Next
TargetWB.Sheets(SourceDoc.Name).Delete
On Error GoTo 0
TargetSheet.Name = SourceDoc.Name

Dim i As Integer
i = 1

'set ID values to Word Tables
For Each Table In SourceDoc.Tables
    Table.ID = i
    i = i + 1
Next Table

i = 1
Dim TrueTableID As Integer
Dim j As Integer
Dim DetailedBOENumber As String
Dim k As Long
k = 1


For Each Table In SourceDoc.Tables
    'find tables with T6, capture ID value
    If Table.Rows.Count > 5 And Table.Columns.Count > 2 Then
        If Left(Table.Cell(4, 2), 1) = "T" Then
            TrueTableID = Table.ID
        End If
    End If
    
    If TrueTableID <> 0 Then
        'paste in the table before the table containing T6
        SourceDoc.Tables(TrueTableID - 1).Select
        Word.Selection.Copy
        TargetSheet.Cells(i, 2).Select
        'sleep added to provide time for data to move to clipboard, was throwing error
        Sleep 50
        TargetSheet.Paste
        DetailedBOENumber = Trim(Right(Cells(i, 2), 5))
        While Not IsNumeric(DetailedBOENumber) And Len(DetailedBOENumber) > 3
            DetailedBOENumber = Trim(Right(DetailedBOENumber, Len(DetailedBOENumber) - 1))
        Wend
        
        i = i + SourceDoc.Tables(TrueTableID - 1).Rows.Count
        
        'paste in the table after the table containing T6
        SourceDoc.Tables(TrueTableID + 2).Select
        Word.Selection.Copy
        TargetSheet.Cells(i, 2).Select
        Sleep 50
        TargetSheet.Paste
        
        'add DetailedBOENumbers if the data is there
        If IsNumeric(DetailedBOENumber) Then
            For j = i + 1 To i + SourceDoc.Tables(TrueTableID + 2).Rows.Count - 1
                TargetSheet.Cells(j, 1) = DetailedBOENumber & NumberToLetter(k)
                k = k + 1
            Next
        Else
            For j = i + 1 To i + SourceDoc.Tables(TrueTableID + 2).Rows.Count - 1
                TargetSheet.Cells(j, 1) = "No BOE Number provided"
                k = k + 1
            Next
            
            MsgBox "One of the tables did not have an associated BOE Number."
        End If
        
        k = 1
        i = i + SourceDoc.Tables(TrueTableID + 2).Rows.Count
        TrueTableID = 0
        
    End If
    
    Call AppTrue
    
Next Table


End Sub

Function NumberToLetter(Number As Long) As String
     On Error Resume Next
     NumberToLetter = Application.Substitute(Application.ConvertFormula("R1C" & Number, xlR1C1, xlA1, 4), "1", "")
 End Function
