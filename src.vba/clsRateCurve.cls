Option Explicit

'----------------------------------------------------
' This class uses qunatlib xl addin
'----------------------------------------------------
Private rate_dates_() As Long
Private dcfs_() As Double


'----------------------------------------------------
' This class uses qunatlib xl addin
'----------------------------------------------------
'Public obj_id As String            'quantlib object id

Public interpolator As clsLinearInterpolator



Public Sub initialize(dates As Range, dc As Range)

On Error GoTo ErrorHandler

    Dim double_dates() As Double
    Dim log_dcfs() As Double
    Dim inx As Integer
    
'    obj_id = Application.run("qlInterpolatedYieldCurve", "_DC_097_", dates, dc, "NullCalendar", "Actual/365", , , "Discount", , False, , True)
    
    rate_dates_ = range_to_array_long(dates)
    dcfs_ = range_to_array(dc)
    
    '-----------------------
    ' Appended on 2013/01/22
    '-----------------------
    ReDim double_dates(LBound(rate_dates_) To UBound(rate_dates_)) As Double
    ReDim log_dcfs(LBound(dcfs_) To UBound(dcfs_)) As Double
    
    For inx = LBound(rate_dates_) To UBound(rate_dates_)
    
        double_dates(inx) = CDbl(rate_dates_(inx))
        log_dcfs(inx) = Log(dcfs_(inx))
    
    Next inx
    
    Set interpolator = New clsLinearInterpolator
    interpolator.initialize double_dates, log_dcfs

    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsRateCurve.initialize :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Sub
    
End Sub

Public Sub initialize_by_array(dates() As Date, dc() As Double)

On Error GoTo ErrorHandler

    Dim double_dates() As Double
    Dim log_dcfs() As Double
    Dim inx As Integer
    
'    obj_id = Application.run("qlInterpolatedYieldCurve", "_DC_097_", dates, dc, "NullCalendar", "Actual/365", , , "Discount", , False, , True)
    
    rate_dates_ = date_to_array_long(dates)
    dcfs_ = double_to_array(dc)
    
    '-----------------------
    ' Appended on 2013/01/22
    '-----------------------
    ReDim double_dates(LBound(rate_dates_) To UBound(rate_dates_)) As Double
    ReDim log_dcfs(LBound(dcfs_) To UBound(dcfs_)) As Double
    
    For inx = LBound(rate_dates_) To UBound(rate_dates_)
    
        double_dates(inx) = CDbl(rate_dates_(inx))
        log_dcfs(inx) = Log(dcfs_(inx))
    
    Next inx
    
    Set interpolator = New clsLinearInterpolator
    interpolator.initialize double_dates, log_dcfs

    Exit Sub
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsRateCurve.initialize_by_array :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Sub
    
End Sub


Public Sub set_rate_dates(rate_dates_in() As Long)

    rate_dates_ = rate_dates_in

End Sub

Public Sub set_dcfs(dcfs_in() As Double)

    dcfs_ = dcfs_in

End Sub



Public Function copy_obj(Optional day_shift As Long = 0, Optional parallel_shift As Double = 0) As clsRateCurve

    Dim obj As clsRateCurve
    Dim new_dates() As Long
    Dim new_dcfs() As Double
    Dim inx As Long
    
On Error GoTo ErrorHandler

    ReDim new_dates(0 To UBound(rate_dates_)) As Long
    ReDim new_dcfs(0 To UBound(rate_dates_)) As Double
    

    For inx = 0 To UBound(rate_dates_)
    
        new_dates(inx) = rate_dates_(inx) + day_shift
        'new_dcfs(inx) = 1 / Exp(get_fwd_rate(CDate(new_dates(0)), CDate(new_dates(inx))) * (new_dates(inx) - new_dates(0)) / 365)
        
        If parallel_shift <> 0 And inx > 0 Then
        
            new_dcfs(inx) = Exp(-(-Log(dcfs_(inx)) / (new_dates(inx) - new_dates(0)) * 365 + parallel_shift) * (new_dates(inx) - new_dates(0)) / 365)
        Else
            new_dcfs(inx) = dcfs_(inx)
        End If
        
    
    Next inx
    
    Set obj = New clsRateCurve
    
    obj.set_rate_dates new_dates
    obj.set_dcfs new_dcfs
    
 '   obj.obj_id = Application.run("qlInterpolatedYieldCurve", "_DC_CP_" & obj_name, new_dates, new_dcfs, "NullCalendar", "Actual/365", , , "Discount", , False, , False)
 
    
    
    Dim double_dates() As Double
    Dim log_dcfs() As Double
    
    ReDim double_dates(LBound(rate_dates_) To UBound(rate_dates_)) As Double
    ReDim log_dcfs(LBound(dcfs_) To UBound(dcfs_)) As Double
    
    For inx = LBound(rate_dates_) To UBound(rate_dates_)
    
        double_dates(inx) = CDbl(rate_dates_(inx))
        log_dcfs(inx) = Log(dcfs_(inx))
    
    Next inx
    
    Set obj.interpolator = New clsLinearInterpolator
    
    obj.interpolator.initialize double_dates, log_dcfs
    
    Set copy_obj = obj
    
    
    Exit Function
    
ErrorHandler:

    raise_err "clsRateCurve.copy_obj", Err.description
    


End Function

Public Function rate_dates_shifted(Optional day_shift As Long = 0) As Long()
       
    Dim new_dates() As Long
    Dim inx As Integer
    
    ReDim new_dates(0 To UBound(rate_dates_)) As Long
    
    For inx = 0 To UBound(rate_dates)
    
        new_dates(inx) = rate_dates_(inx) + day_shift
    
    Next inx
    
    rate_dates_shifted = new_dates

End Function
Public Function rate_dates() As Long()
       
    
    rate_dates = rate_dates_ 'new_dates

End Function

Public Function spread_dcf(rate_spread As Double) As Double()
    
    Dim rtn_array() As Double
    Dim inx As Integer
    
On Error GoTo ErrorHandler

    ReDim rtn_array(LBound(dcfs_) To UBound(dcfs_)) As Double
    
    For inx = LBound(dcfs_) To UBound(dcfs_)
    
        rtn_array(inx) = dcfs_(inx) * Exp(-rate_spread * (rate_dates_(inx) - rate_dates_(0)) / 365)
    
    Next inx
    
    spread_dcf = rtn_array
    
    Exit Function
    
ErrorHandler:

    raise_err "clsRateCurve.spread_dcf"
    


End Function

Public Function dcfs() As Double()
    
    dcfs = dcfs_

End Function

Public Function get_fwd_rate(d1 As Date, d2 As Date) As Double

    Dim d1_long As Long
    Dim d2_long As Long

On Error GoTo ErrorHandler

    d1_long = d1
    
    If d2 > d1 + 1 Then
        d2_long = d2
    Else
        d2_long = d2 + 1
    End If
    
    If d1 < rate_dates_(0) Then
        
        d1_long = rate_dates_(0)
        d2_long = rate_dates_(0) + d2 - d1

    End If

    If d1_long = d2_long Then
        get_fwd_rate = 0
    Else
        get_fwd_rate = -(interpolator.interpolated_value(CDbl(d2_long)) - interpolator.interpolated_value(CDbl(d1_long))) / ((d2_long - d1_long) / 365)
    End If
    
 '   get_fwd_rate = Application.run("qlYieldTSForwardRate", obj_id, d1_long, d2_long, "Actual/365", , , True)

    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsRateCurve.get_fwd_rate :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Function
    

End Function

Public Function get_discount_factor(d1 As Date, d2 As Date) As Double


On Error GoTo ErrorHandler

    Dim d1_long As Long
    Dim d2_long As Long
    
    d1_long = d1
    d2_long = d2

    If d1 < rate_dates_(0) Then
        
        d1_long = rate_dates_(0)
        d2_long = rate_dates_(0) + d2 - d1
        

    End If


    get_discount_factor = Exp(interpolator.interpolated_value(CDbl(d2_long)) - interpolator.interpolated_value(CDbl(d1_long)))
'    get_discount_factor = Application.run("qlYieldTSDiscount", obj_id, d2_long, True) / Application.run("qlYieldTSDiscount", obj_id, d1_long, True)

    
    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 1002, "clsRateCurve.get_discount_factor :" & Chr(13) & Err.source, Err.description '1002 <-- initialization error
    Exit Function
    

End Function
'
'Public Sub parallel_shift(spread As Double)
'
'    Dim inx As Integer
'
'On Error GoTo ErrorHandler
'
'
'
'    Exit Sub
'
'ErrorHandler:
'
'    raise_err "parallel_shift"
'
'End Sub
'

'Public Sub move_date(date_offset As Integer)
'
'    Dim inx As Integer
'
'    For inx = LBound(rate_dates_) To UBound(rate_dates_)
'
'        rate_dates_(inx) = rate_dates_(inx) + date_offset
'
'    Next inx
'
'End Sub