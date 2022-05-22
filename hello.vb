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

            ' 'Gan gia tri cot thu tu
            ' wsSheet.Cells(Counter + 4, 1).Value = Counter
            ' wsSheet.Cells(Counter + 4, 3).Value =  ws.cells(1,1).value
            ' 'Tao lien ket
            ' wsSheet.Hyperlinks.Add Anchor:=wsSheet.Cells(Counter + 4, 2), _
            ' Address:="", _
            ' SubAddress:=ws.Name & "!A1", _
            ' ScreenTip:=ws.Name, _
            ' TextToDisplay:=ws.Name
            ' 'Them nut Quay ve Sheet Muc luc tai moi Sheet
            ' With ws
            '     .Hyperlinks.Add Anchor:=.Range("H1"), Address:="", SubAddress:="Index", TextToDisplay:="Quay ve"
            ' End With

            wsSheet.Cells(Counter + 4, 1).Value = Counter
            ' wsSheet.Cells(Counter + 4, 3).Value =  ws.cells(Counter + 5,3).value
            wsSheet.Cells(Counter + 4, 3).Value =  ws.Range("C5:C100").value
            wsSheet.Cells(Counter + 4, 2).Value = ws.Name
           
            For Each item In ws.Range("C5:C100")
                wsSheet.Cells(Counter + 4, 3).Value = item
                Counter = Counter + 1
            Next item

            Counter = Counter + 1


        End If
    Next ws
    
    
    Set xlSheet = Nothing
End Sub

Function in_array(my_array, my_value)
    
    'https://www.excel-pratique.com/en/vba_tricks/search-in-array-function
	
    in_array = False
    
    For i = LBound(my_array) To UBound(my_array)
        If my_array(i) = my_value Then 'If value found
            in_array = True
            Exit For
        End If
    Next
    
End Function