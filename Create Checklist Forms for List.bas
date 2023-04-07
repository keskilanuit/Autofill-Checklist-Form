Attribute VB_Name = "Module1"
Option Explicit
Sub NewSheets()
Dim i As Integer
Dim ws As Worksheet
Dim sh As Worksheet
Set ws = Sheets("Template")
Set sh = Sheets("Sheets Insert")
Application.ScreenUpdating = 0

    For i = 2 To Range("A" & Rows.Count).End(xlUp).Row
            Sheets("Template").Copy After:=sh
            ActiveSheet.Name = sh.Range("A" & i).Value
    Next i
End Sub
