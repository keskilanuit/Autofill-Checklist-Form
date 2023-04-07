
Sub create_checklist_forms_based_on_your_list()
    
Dim i As Integer
Dim ws As Worksheet
Dim sh As Worksheet
Set ws = Sheets("Template")
        'loop through listed items from your checklist'
        Set sh = Sheets("Your List")
Application.ScreenUpdating = 0

    For i = 2 To Range("A" & Rows.Count).End(xlUp).Row
            Sheets("Template").Copy After:=sh
            ActiveSheet.Name = sh.Range("A" & i).Value
    Next i

End Sub
