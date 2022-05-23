
Function in_array(my_array, my_value, max)
    
    'https://www.excel-pratique.com/en/vba_tricks/search-in-array-function
    
    in_array = False
    
    For i = 5 To max Step 1
        If my_array(i) = my_value Then 'If value found
            in_array = True
            Exit For
        End If
    Next
    
End Function

Private Sub CreateTableOfContents()
    Dim wsSheet     As Worksheet
    Dim ws          As Worksheet
    Dim Counter     As Long
    Dim arrKey(100000000) As Variant

    On Error Resume Next
    Set wsSheet = Sheets("Hiep123")

    

    'Kiem tra su ton tai cua Sheet
    On Error GoTo 0
    If wsSheet Is Nothing Then
        'Neu chua co thi them vao vi tri dau tien cua Workbook
        Set wsSheet = ActiveWorkbook.Sheets.Add(Before:=Worksheets(1))
        wsSheet.Name = "Hiep123"
    End If

'   msgbox("Value stored in Array index 0 : ")

    Counter = 1
    For Each ws In Worksheets
        If ws.Name <> wsSheet.Name Then
            wsSheet.Cells(Counter + 4, 1).Value = Counter
            ' wsSheet.Cells(Counter + 4, 3).Value =  ws.Range("C5:C100").value
            wsSheet.Cells(Counter + 4, 2).Value = ws.Name
            wsSheet.Cells(Counter + 4, 2).Font.Bold = True
            wsSheet.Cells(Counter + 4, 2).Font.Size = 12
                                        
           
            Dim indexRange
            indexRange = 5
            
            Dim valueFist
            valueFist = ws.Cells(5, 3)
            
                'Set ColumnWidth
            With Columns("C:D")
                .ColumnWidth = 30
            End With
            
            For Each Item In ws.Range("C5:C64")
            
                Dim isExist
                   ' MsgBox ("Value stored in Array index 0 : " & Item)

                
                isExist = in_array(arrKey, Item, indexRange + 1)

                If isExist = False Then
                    arrKey(indexRange) = Item
                     wsSheet.Cells(Counter + 4, 3).Value = Item
                     wsSheet.Cells(Counter + 4, 4).Value = valueFist
                     indexRange = indexRange + 1
                End If

                
                If Len(Item) > 0 Then
                    Counter = Counter + 1
                End If
            Next Item
            Counter = Counter + 1
        End If
    Next ws
    
    
    Set xlSheet = Nothing
End Sub


