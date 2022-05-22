
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