Attribute VB_Name = "dl_functions"
Option Explicit




'
' function to get bond history given an isin, an isin list and a block of data
' history of same number of rows as the isin list
'
Function dl_Get_Hist(isin As String, isin_list As Variant, history_block As Range) As Variant
    
    Dim hb_values As Variant
    Dim x As Variant ' dummy variable for testing
    Dim isin_list_values As Variant
    Dim number_if_isins As Integer
    Dim i As Integer
    Dim width_of_dates As Integer
    Dim the_spreads_list As Variant
    Dim counter As Integer
    Dim lookup_value As Integer
    
    
    isin_list_values = isin_list.Value
    number_if_isins = UBound(isin_list_values)
    
    
    ' hb_values we get here
    Debug.Print "hello world" + Str(number_if_isins)
    
    hb_values = history_block.Value
    Debug.Print hb_values(1, 1)
    
    
    width_of_dates = UBound(hb_values, 2)
    
    Debug.Print "the width is" + Str(width_of_dates)
    
    
    ' lookup the isin from the isin list
    lookup_value = Application.WorksheetFunction.XMatch(isin, isin_list_values)
    Debug.Print "the loopup value is: " + Str(lookup_value)
    
    
    
    the_spreads_list = Application.WorksheetFunction.Index(hb_values, lookup_value, 0)
    
'    Debug.Print the_spreads
    
'    Debug.Print "list of values"
'    For i = 1 To UBound(hb_values)
'        Debug.Print hb_values(i, 1)
'    Next i


    counter = 0
    For i = 1 To UBound(the_spreads_list)
        If the_spreads_list(i) = "" Then
            counter = counter + 1
            the_spreads_list(i) = "-"
        End If

        
    Next i
    
    Debug.Print "the blanks: " + Str(counter)

    dl_Get_Hist = "hello world " + Str(number_if_isins) + ", " + Str(width_of_dates)
    dl_Get_Hist = the_spreads_list
    
    
    
End Function




'
' function to get the spread between two bonds
'

Function dl_spread_between_two_bonds(first_isin As String, second_isin As String, isin_list As Variant, history_block As Range) As Variant

    Dim first_bond_spreads As Variant
    Dim second_bond_spreads As Variant
    Dim spread_array As Variant
    Dim i As Integer
    
    
    
    first_bond_spreads = dl_Get_Hist(first_isin, isin_list, history_block)
    second_bond_spreads = dl_Get_Hist(second_isin, isin_list, history_block)
    Debug.Print "got to point 2...."
    
    ReDim spread_array(1 To UBound(first_bond_spreads))
    
    Debug.Print "length of spread array: " + Str(UBound(first_bond_spreads))
    
    For i = 1 To UBound(first_bond_spreads)
        If (first_bond_spreads(i) = "-") Or (second_bond_spreads(i) = "-") Then
            spread_array(i) = "-"
        Else
            spread_array(i) = first_bond_spreads(i) - second_bond_spreads(i)
        End If
        
        
    Next i
    
    dl_spread_between_two_bonds = spread_array
    

End Function
