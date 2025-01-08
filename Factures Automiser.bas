Attribute VB_Name = "Module1"
Sub GenerateInvoices()
    'Declaring all the variables this function will use.
    Dim wsMain As Worksheet
    Dim wsAttendance As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long
    Dim childLastName As String, childFirstName As String
    Dim parent1 As String, parent2 As String
    Dim drdParent1 As String, drdParent2 As String
    Dim percentageParent1 As Double, percentageParent2 As Double
    Dim hasOneBill As Boolean
        
    'Set the primary and dates sheets
    Set wsMain = ThisWorkbook.Sheets("Program Principal")
    Set wsAttendance = ThisWorkbook.Sheets("List de Dates")
    
    lastRow = wsMain.Cells(wsMain.Rows.Count, 1).End(xlUp).Row
    lastCol = wsMain.Cells(1, wsMain.Columns.Count).End(xlToLeft).Column
    
    
    'Loop through each student in the current worksheet.
    For i = 2 To lastRow
        childLastName = wsMain.Cells(i, 1).Value
        childFirstName = wsMain.Cells(i, 2).Value
        
        'Parent 1's information
        parent1 = wsMain.Cells(i, 3).Value
        drdParent1 = wsMain.Cells(i, 4).Value
        percentageParent1 = wsMain.Cells(i, 5).Value
        
        'Parent 2's information
        parent2 = wsMain.Cells(i, 6).Value
        drdParent2 = wsMain.Cells(i, 7).Value
        percentageParent2 = wsMain.Cells(i, 8).Value
        
        hasOneBill = wsMain.Cells(i, 9).Value
        
        'Create two sheets if there are two bills, otherwise create 1
        If hasOneBill = False Then
            CreateSheet wsMain, wsAttendance, childLastName & ", " & childFirstName & " -1", i, parent1, drdParent1, percentageParent1, lastRow, lastCol
            CreateSheet wsMain, wsAttendance, childLastName & ", " & childFirstName & " -2", i, parent2, drdParent2, percentageParent2, lastRow, lastCol
        End If
            
        If hasOneBill = True Then
            If parent2 = "" Then
                CreateSheet wsMain, wsAttendance, childLastName & ", " & childFirstName, i, parent1, drdParent1, percentageParent1, lastRow, lastCol
            Else
                CreateSheet wsMain, wsAttendance, childLastName & ", " & childFirstName, i, parent1 & " & " & parent2, drdParent1, percentageParent1, lastRow, lastCol
            End If
        End If

    
    'Move on to the next student
    Next i
    
End Sub
    
Sub CreateSheet(wsMain As Worksheet, wsAttendance As Worksheet, sheetName As String, i As Long, parentName As String, parentDRD As String, parentPercentage As Double, lastRow As Long, lastCol As Long)
    'Declaring all the variables this function will use.
    Dim wsChild As Worksheet
    Dim j As Long, k As Long
    Dim cost As Double
    Dim dayType As String
    Dim attendanceDate As Date
    Dim totalCharge As Double
    Dim totalCost As Double
    Dim percentageString As Double
    
    
    'Add a new sheet
    Set wsChild = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsChild.Name = sheetName
    
    'Add logo, date, student and parent information at the top of the sheet
    'Assuming logo is saved as "logo.png" in the same directory.
    Dim logoPath As String
    logoPath = "D:\logo.png"
    
    On Error Resume Next
    wsChild.Pictures.Insert(logoPath).Select
    With Selection
        .Top = wsChild.Cells(1, 1).Top
        .Left = wsChild.Cells(1, 1).Left
        .Lock AspectRation - msoTrue
        .Height = 80
    End With
    On Error GoTo 0
    
     If parentDRD <> "" Then
        wsChild.Cells(1, 7).Value = "Facture #" & parentDRD
        wsChild.Cells(1, 7).Font.Name = "Avenir Next LT Pro Light"
        wsChild.Cells(1, 7).HorizontalAlignment = xlRight
    End If
    
    wsChild.Cells(7, 1).Value = "Date :"
    wsChild.Cells(7, 1).Font.Name = "Avenir Next LT Pro Light"
    wsChild.Cells(7, 1).HorizontalAlignment = xlRight
    
    wsChild.Cells(7, 2).Value = "=Today()"
    wsChild.Cells(7, 2).Font.Name = "Avenir Next LT Pro Light"
    wsChild.Cells(7, 2).HorizontalAlignment = xlLeft
    
    wsChild.Cells(8, 1).Value = "Nom du parent :"
    wsChild.Cells(8, 1).Font.Name = "Avenir Next LT Pro Light"
    wsChild.Cells(8, 1).HorizontalAlignment = xlRight
    
    wsChild.Range("B8:G8").MergeCells = True
    wsChild.Cells(8, 2).Value = parentName
    wsChild.Cells(8, 2).Font.Name = "Avenir Next LT Pro Light"
    wsChild.Cells(8, 2).HorizontalAlignment = xlLeft
    
    wsChild.Range("A10:G10").MergeCells = True
    wsChild.Cells(10, 1).Value = "Facture paiement pŽdagogique et tempte"
    wsChild.Cells(10, 1).HorizontalAlignment = xlCenter
    wsChild.Cells(10, 1).Font.Name = "Avenir Next LT Pro Light"
    wsChild.Cells(10, 1).Font.Bold = True
    
    wsChild.Cells(11, 1).Value = "Nom de l'enfant :"
    wsChild.Cells(11, 1).Font.Name = "Avenir Next LT Pro Light"
    wsChild.Cells(11, 1).HorizontalAlignment = xlRight
    
    wsChild.Cells(11, 2).Value = childFirstName & " " & childLastName
    wsChild.Cells(11, 2).Font.Name = "Avenir Next LT Pro Light"
    wsChild.Cells(11, 2).HorizontalAlignment = xlLeft
    
    Columns("A").ColumnWidth = 17
    
    'Initialize Cost
    totalCost = 0
        
    j = 13
    For k = 11 To lastCol
        If wsMain.Cells(i, k).Value = True Then
            
            dayType = wsAttendance.Cells(j, 2).Value
            attendanceDate = wsMain.Cells(1, k).Value
            
            If wsMain.Cells(i, 10).Value = True Then
                cost = 20
            Else
                cost = 0
            End If
            
            If dayType = "Jour pŽdagogique" Then
                        cost = cost + 15
            End If
            
            If dayType = "Jour tempte" Thenn
                    cost = cost + 6
            End If
                    
            If dayType = "Demi-jour pŽdagogique" Then
                    cost = cost + 12
            Else
                cost = cost + 0
            End If
            
            If cost > 0 Then
                wsChild.Cells(j, 1).Value = Format(attendanceDate, "yyyy/mm/dd")
                wsChild.Cells(j, 1).Font.Name = "Avenir Next LT Pro Light"
                wsChild.Cells(j, 1).HorizontalAlignment = xlRight
                
                wsChild.Cells(j, 4).Value = dayType
                wsChild.Cells(j, 4).Font.Name = "Avenir Next LT Pro Light"
                wsChild.Cells(j, 4).HorizontalAlignment = xlCenter
                
                wsChild.Cells(j, 7).Value = FormatCurrency(cost, 0)
                wsChild.Cells(j, 7).Font.Name = "Avenir Next LT Pro Light"
                wsChild.Cells(j, 7).HorizontalAlignment = xlLeft
                j = j + 1
            End If
            
            totalCost = totalCost + cost
            
        End If
    Next k
    
    wsChild.Cells(j, 6).Value = "Total "
    wsChild.Cells(j, 6).Font.Name = "Avenir Next LT Pro Light"
    
    wsChild.Cells(j, 7).Value = cost
    wsChild.Cells(j, 7).Font.Name = "Avenir Next LT Pro Light"

    
    'If there are two bills, Calculate the parents share.
    If parentPercentage < 1 Then
        wsChild.Cells(j, 6).Value = "Total : "
        wsChild.Cells(j, 6).Font.Name = "Avenir Next LT Pro Light"
        wsChild.Cells(j, 6).HorizontalAlignment = xlRight
        
        wsChild.Cells(j, 7).Value = FormatCurrency(totalCost, 0)
        wsChild.Cells(j, 7).Font.Name = "Avenir Next LT Pro Light"
        wsChild.Cells(j, 7).HorizontalAlignment = xlLeft
        
        j = j + 1
        percentageString = parentPercentage * 100
        wsChild.Cells(j, 6).Value = "Total qui sera chargŽ le " & wsAttendance.Cells(2, 8).Value & "(" & percentageString & "%) : "
        wsChild.Cells(j, 6).Font.Name = "Avenir Next LT Pro Light"
        wsChild.Cells(j, 6).HorizontalAlignment = xlRight
        wsChild.Cells(j, 6).Font.Bold = True
        
        totalCharge = totalCost * parentPercentage
        wsChild.Cells(j, 7).Value = FormatCurrency(totalCharge, 2)
        wsChild.Cells(j, 7).Font.Name = "Avenir Next LT Pro Light"
        wsChild.Cells(j, 7).HorizontalAlignment = xlLeft
        wsChild.Cells(j, 7).Font.Bold = True
        
    Else
        'Add total cost to the new sheet
        wsChild.Cells(j, 6).Value = "Total qui sera chargŽ le " & wsAttendance.Cells(2, 8).Value & " : "
        wsChild.Cells(j, 6).Font.Name = "Avenir Next LT Pro Light"
        wsChild.Cells(j, 6).HorizontalAlignment = xlRight
        wsChild.Cells(j, 6).Font.Bold = True
        
        wsChild.Cells(j, 7).Value = FormatCurrency(totalCost, 2)
        wsChild.Cells(j, 7).Font.Name = "Avenir Next LT Pro Light"
        wsChild.Cells(j, 7).HorizontalAlignment = xlLeft
        wsChild.Cells(j, 7).Font.Bold = True
    End If
    
    j = j + 2
    wsChild.Cells(j, 4).Value = "*Ce montant ne compte pas le camp de No‘l, ni le camp de mars"
    wsChild.Cells(j, 4).Font.Name = "Avenir Next LT Pro Light"
    wsChild.Cells(j, 4).HorizontalAlignment = xlCenter
    wsChild.Cells(j, 4).Font.Size = 9
    
    j = j + 1
    wsChild.Cells(j, 4).Value = "____________________________________________________________________________________________"
    wsChild.Cells(j, 4).Font.Name = "Avenir Next LT Pro Light"
    wsChild.Cells(j, 4).HorizontalAlignment = xlCenter
    wsChild.Cells(j, 4).Font.Size = 9
    
    j = j + 1
    wsChild.Cells(j, 1).Value = "160-A, promenade Eco Terra"
    wsChild.Cells(j, 1).Font.Name = "Avenir Next LT Pro Light"
    wsChild.Cells(j, 1).HorizontalAlignment = xlLeft
    wsChild.Cells(j, 1).Font.Size = 9
    
    wsChild.Cells(j, 7).Value = "TŽlŽphone : 506 453-7700"
    wsChild.Cells(j, 7).Font.Name = "Avenir Next LT Pro Light"
    wsChild.Cells(j, 7).HorizontalAlignment = xlRight
    wsChild.Cells(j, 7).Font.Size = 9
    
    j = j + 1
    wsChild.Cells(j, 1).Value = "Fredericton (Nouveau-Brunswick)  E3A 9M1"
    wsChild.Cells(j, 1).Font.Name = "Avenir Next LT Pro Light"
    wsChild.Cells(j, 1).HorizontalAlignment = xlLeft
    wsChild.Cells(j, 1).Font.Size = 9
    
    wsChild.Cells(j, 7).Value = "Cellulaire : 506 478-2198"
    wsChild.Cells(j, 7).Font.Name = "Avenir Next LT Pro Light"
    wsChild.Cells(j, 7).HorizontalAlignment = xlRight
    wsChild.Cells(j, 7).Font.Size = 9
    
    j = j + 1
    wsChild.Cells(j, 4).Value = "direction@garderielenvolee.ca"
    wsChild.Cells(j, 4).Font.Name = "Avenir Next LT Pro Light"
    wsChild.Cells(j, 4).HorizontalAlignment = xlCenter
    wsChild.Cells(j, 4).Font.Size = 9
    
    j = j + 1
    wsChild.Cells(j, 4).Value = "renee@garderielenvolee.ca"
    wsChild.Cells(j, 4).Font.Name = "Avenir Next LT Pro Light"
    wsChild.Cells(j, 4).HorizontalAlignment = xlCenter
    wsChild.Cells(j, 4).Font.Size = 9
End Sub


