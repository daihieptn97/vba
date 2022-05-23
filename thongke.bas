
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

Private Sub thongkelinhkien()

Dim wsSheet     As Worksheet
Dim ws          As Worksheet
Dim Counter     As Long
Dim arrAccessories As Variant
Dim countAccessories As Long


Dim arrProduct As Variant
Dim countProduct As Long


On Error Resume Next
Set wsSheet = Sheets("ReportSheet")

' init 
countAccessories = 0
countProduct = 0


On Error GoTo 0
If wsSheet Is Nothing Then
    Set wsSheet = ActiveWorkbook.Sheets.Add(Before:=Worksheets(1))
    wsSheet.Name = "ReportSheet"
End If

For Each ws In Worksheets
    If ws.Name <> wsSheet.Name Then
        ' arrAccessories(countAccessories) = 
        arrProduct(countProduct) = ws.Name
        countProduct = countProduct + 1

        For Each Item In ws.Range("C6:C64")
            Dim isExist
            isExist = in_array(arrAccessories, Item, countAccessories)
            If isExist = False Then
                arrAccessories(countAccessories) = Item
                wsSheet.Cells(Counter + 4, 3).Value = Item
                wsSheet.Cells(Counter + 4, 4).Value = valueFist
                countAccessories = countAccessories + 1
            End If
        Next Item
        ' end for

    End If
Next ws




End Sub


