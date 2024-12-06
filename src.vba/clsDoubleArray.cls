' I want to implement jagged array.

Option Explicit


Private data_() As Double


Public Sub redim_array(from_index As Integer, to_index As Integer)

    ReDim Preserve data_(from_index To to_index) As Double

End Sub

Public Function get_data(inx As Integer) As Double
    
    get_data = data_(inx)

End Function

Public Function get_array() As Double()
    
    get_array = data_

End Function


Public Sub set_data(inx As Integer, the_value As Double)
    
    data_(inx) = the_value

End Sub


Public Sub set_array(the_array() As Double)
    
    data_ = the_array

End Sub