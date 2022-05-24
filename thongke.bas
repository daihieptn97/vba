
Function in_array(my_array, my_value, max)
    'https://www.excel-pratique.com/en/vba_tricks/search-in-array-function
    in_array = False
    For i = 1 To max Step 1
        If my_array(i) = my_value Then 'If value found
            in_array = True
            Exit For
        End If
    Next
End Function

Function log_cel(my_array, my_value, max)
     For Each Item In ws.Range("C6:C64")
        Dim isExist
        isExist = in_array(arrAccessories, Item, countAccessories)
        If isExist = False Then
            arrAccessories(countAccessories) = Item
            ' wsSheet.Cells(Counter + 4, 3).Value = Item
            ' wsSheet.Cells(Counter + 4, 4).Value = valueFist
            countAccessories = countAccessories + 1
        End If
    Next Item
    
End Function

Private Sub thongkelinhkien()

Dim wsSheet     As Worksheet
Dim ws          As Worksheet
Dim Counter     As Long
Dim arrAccessories(1000000) As Variant
Dim countAccessories As Long


Dim arrProduct(1000000) As Variant
Dim countProduct As Long


On Error Resume Next
Set wsSheet = Sheets("ReportSheet")

' init
countAccessories = 1
countProduct = 1


On Error GoTo 0
If wsSheet Is Nothing Then
    Set wsSheet = ActiveWorkbook.Sheets.Add(Before:=Worksheets(1))
    wsSheet.Name = "ReportSheet"
End If

Columns("B:B").ColumnWidth = 30
Columns("C:C").ColumnWidth = 30
Columns("D:D").ColumnWidth = 30
Columns("G:G").ColumnWidth = 30
Columns("H:H").ColumnWidth = 30
Columns("I:I").ColumnWidth = 30

Range("A3") = "STT"
Range("A3").Font.Bold = TRUE

Range("B3") = "Tên sheet"
Range("B3").Font.Bold = TRUE

Range("C3") = "ten linh kien"
Range("C3").Font.Bold = TRUE

Range("D3") = "ten thanh pham"
Range("D3").Font.Bold = TRUE



For Each ws In Worksheets
    If ws.Name <> wsSheet.Name And ws.Name <> "Hiep123" Then
        ' arrAccessories(countAccessories) =
        ' arrProduct(countProduct) = ws.Name
        ' countProduct = countProduct + 1

        For Each Item In ws.Range("C6:C64")
            Dim isExist
            isExist = in_array(arrAccessories, Item, countAccessories)
            If isExist = False And Item <> "" Then
                arrAccessories(countAccessories) = Item
                ' wsSheet.Cells(Counter + 4, 3).Value = Item
                ' wsSheet.Cells(Counter + 4, 4).Value = valueFist
                countAccessories = countAccessories + 1
            End If
        Next Item
        ' end for

    End If
Next ws



For i = 1 To (countAccessories - 1)
    wsSheet.Cells(Counter + 4, 1).Value = Counter
    ' Begin for Worksheets
    For Each ws In Worksheets
        If ws.Name <> wsSheet.Name And ws.Name <> "Hiep123" Then

            Dim valueFist
            valueFist = ws.Cells(5, 3)
            
            Dim indexRange
            indexRange = 6
            For Each Item In ws.Range("C6:C64")
                If Item = arrAccessories(i) Then
                    wsSheet.Cells(Counter + 4, 2).Value = ws.Name
                    wsSheet.Cells(Counter + 4, 3).Value = Item
                    wsSheet.Cells(Counter + 4, 4).Value = valueFist
                    
                    ' wsSheet.Cells(Counter + 4, 4).ColumnWidth = 30
                    ' wsSheet.Cells(Counter + 4, 3).ColumnWidth = 30
                    
                   ' ws.Range("1:1").Copy wsSheet.Range("5:5")
                   
                   ' ws.Range("1:1").Copy wsSheet.Range("5:5")
                    
                   '  ws.Range("A5:Y5").Copy Destination:=wsSheet.Cells(Counter + 4, 5)
                      
                    '  ws.Range("A5:Y5").Copy Destination:=wsSheet.Cells(Counter + 4, 5)
                    
                   '  wsSheet.Cells(Counter + 4, 5).Value = Item.Rows
                   
                    ws.Range("A" & indexRange & ":Y" & indexRange).Copy Destination:=wsSheet.Cells(Counter + 4, 5)

                   ' ws.Rows("6:6").Activate
                   ' Selection.Copy
                    
                   ' wsSheet.Cells(Counter + 4, 4).Activate
                   ' ActiveSheet.Paste
                    
                    Counter = Counter + 1
                    Exit For
                End If
                indexRange = indexRange + 1
            Next Item
            ' end for
        End If
    Next ws
    ' end for Worksheets
    Counter = Counter + 1
Next i


' For i = 1 To (countAccessories - 1)
' wsSheet.Cells(Counter + 4, 3).Value = Counter
'      wsSheet.Cells(Counter + 4, 4).Value = arrAccessories(i)
'         Counter = Counter + 1
' Next i


End Sub







