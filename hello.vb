
Function in_array(my_array, my_value)
    
    'https://www.excel-pratique.com/en/vba_tricks/search-in-array-function
	
    in_array = False
    
    For i = 0 To UBound(my_array)
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
    Dim arrKey()

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
           
            Dim indexRange
            indexRange = 5

            For Each item In ws.Range("C5:C100")
                wsSheet.Cells(Counter + 4, 3).Value = item

                ' Range(ws.Cells(indexRange, 3),  ws.Cells(indexRange, 24)).Select
                ' Selection.Copy

                ' ws.Range("C5:Y5").Select
                ' Selection.Copy

                ' Range(wsSheet.Cells(Counter + 4)).Select
                ' wsSheet.Range("C"&Counter + 4).Select
                ' ActiveSheet.Paste
                Dim isExist
                isExist = in_array(arrKey, item)

                If isExist = False Then
                    arrKey(UBound(arrKey) + 1) = item
                End If

                indexRange = indexRange + 1
                if Len(item) > 0 Then
                    Counter = Counter + 1
                End If
            Next item
            Counter = Counter + 1
        End If
    Next ws
    
    
    Set xlSheet = Nothing
End Sub
