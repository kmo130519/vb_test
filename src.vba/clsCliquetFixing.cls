Option Explicit


Public fixing_date As Date
Public local_floor As Double
Public local_cap As Double

Public fixing_value As Double


Public Function copy_obj() As clsCliquetFixing
    
    Dim rtn_obj As clsCliquetFixing
    
    Set rtn_obj = New clsCliquetFixing
    
    rtn_obj.fixing_date = fixing_date
    rtn_obj.local_cap = local_cap
    rtn_obj.local_floor = local_floor
    rtn_obj.fixing_value = fixing_value
    
    Set copy_obj = rtn_obj


End Function