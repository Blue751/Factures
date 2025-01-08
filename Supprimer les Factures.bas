Attribute VB_Name = "Module2"
Sub RŽcommencer()
    Dim ws As Worksheet
    Dim wsName As String
    Dim deleteList As Collection
    Dim item As Variant
    
    'Create a collection to hold the names of the worksheets to delete
    Set deleteList = New Collection
    
    'Loop though each sheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        wsName = ws.Name
    
        'Check if the sheen name contains a comma, which suggests lastName, firstName.
        If InStr(wsName, ", ") > 0 Then
            deleteList.Add wsName
        End If
    Next ws
    
    'Delete the worksheets listed in deleteList
    For Each item In deleteList
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(item).Delete
        Application.DisplayAlerts = True
    Next item
    
    MsgBox "Complet"
End Sub

