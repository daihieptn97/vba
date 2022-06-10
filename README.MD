# Code Report Accessories

```vb
Private Sub CreateTableOfContents()
    Dim wsSheet     As Worksheet
    Dim ws          As Worksheet
    Dim Counter     As Long
    On Error Resume Next
    Set wsSheet = Sheets("Mucluc")
    'Kiem tra su ton tai cua Sheet
    On Error GoTo 0
    If wsSheet Is Nothing Then
        'Neu chua co thi them vao vi tri dau tien cua Workbook
        Set wsSheet = ActiveWorkbook.Sheets.Add(Before:=Worksheets(1))
        wsSheet.Name = "Mucluc"
    End If
    
    With wsSheet
        .Cells(2, 1) = "DANH SACH CAC SHEET"
        .Cells(2, 1).Name = "Index"
        .Cells(4, 1).Value = "STT"
        .Cells(4, 2).Value = "Ten Sheet"
    End With
    
    'Merge Cell
    With Range("A2:B2")
        .Merge
        .HorizontalAlignment = xlCenter
        .Font.Bold = TRUE
    End With
    
    'Set ColumnWidth
    With Columns("A:A")
        .ColumnWidth = 8
        .HorizontalAlignment = xlCenter
    End With
    
    With Range("A4")
        .HorizontalAlignment = xlCenter
        .Font.Bold = TRUE
    End With
    
    Columns("B:B").ColumnWidth = 30
    With Range("B4")
        .HorizontalAlignment = xlCenter
        .Font.Bold = TRUE
    End With
    
    Counter = 1
    For Each ws In Worksheets
        If ws.Name <> wsSheet.Name Then
            'Gan gia tri cot thu tu
            wsSheet.Cells(Counter + 4, 1).Value = Counter
            'Tao lien ket
            wsSheet.Hyperlinks.Add Anchor:=wsSheet.Cells(Counter + 4, 2), _
            Address:="", _
            SubAddress:=ws.Name & "!A1", _
            ScreenTip:=ws.Name, _
            TextToDisplay:=ws.Name
            'Them nut Quay ve Sheet Muc luc tai moi Sheet
            ' With ws
            '     .Hyperlinks.Add Anchor:=.Range("H1"), Address:="", SubAddress:="Index", TextToDisplay:="Quay ve"
            ' End With
            Counter = Counter + 1
        End If
    Next ws
    Set xlSheet = Nothing
End Sub

```