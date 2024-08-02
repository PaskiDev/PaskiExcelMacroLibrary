Attribute VB_Name = "Module1"
Option Explicit

Sub AddSheetsFromList()
    ' variables
    Dim col As Range
    Dim tSheet As String
    Dim cell As Range
    Dim wsName As String
    
    ' set range for column
    Set col = Application.InputBox(Prompt:="Select the column for the sheet names.", Type:=8)
    tSheet = Application.InputBox(Prompt:="Set the name of the sheet you want to use as template.", Type:=2)
    
    ' iterator to check values from list
    For Each cell In col
        wsName = Trim(cell.Value)
        tSheet = Trim(tSheet)
        
        ' check if exist
        If wsName <> "" Then
            Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = wsName
            Worksheets(tSheet).Range("A1:Z90").Copy _
                Destination:=Worksheets(wsName).Range("A1:Z90")
            
        End If
        
    Next cell
End Sub
