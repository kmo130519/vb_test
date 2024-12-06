Option Explicit




Public number As Long
Public source As String
Public description As String

Public Sub setError(ByVal number_in As Long, ByVal source_in As String, ByVal description_in As String)

    number = number_in
    source = source_in
    description = description_in

End Sub

Public Sub clear()

    number = 0
    source = ""
    description = ""
    
End Sub