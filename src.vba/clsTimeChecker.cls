Option Explicit

'Private Declare Function GetTickCount Lib "kernel32.dll" () As Long


Private tick_counter_() As Long
Private stop_point_name_() As String

Public Sub initailize()
    
    Erase tick_counter_
    Erase stop_point_name_

End Sub




Public Function get_array_count() As Long
    
    Dim rtn_value As Long
    
On Error Resume Next

    rtn_value = UBound(tick_counter_)
    
    If Err.number = 9 Then
    
        rtn_value = 0
        
    End If
    
    
    get_array_count = rtn_value


End Function

Public Sub add_tick_counter(ByVal stop_point_name As String)

    Dim inx As Integer
    Dim current_size As Integer
    
On Error Resume Next
    current_size = UBound(tick_counter_)
    
    If Err.number = 9 Then
        
        current_size = 0
    
    End If
    
On Error GoTo ErrorHandler

    ReDim Preserve tick_counter_(1 To current_size + 1) As Long
    ReDim Preserve stop_point_name_(1 To current_size + 1) As String
    
    tick_counter_(UBound(tick_counter_)) = GetTickCount
    stop_point_name_(UBound(stop_point_name_)) = stop_point_name

    Exit Sub
    
ErrorHandler:

    raise_err "clsTimeChecker.add_tick_counter"
    

End Sub



Public Function tick_counter() As Long()

    tick_counter = tick_counter_

End Function

Public Function stop_point_name() As String()

    stop_point_name = stop_point_name_

End Function