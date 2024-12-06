Option Explicit

'###################################################################################
' Author: YK Jeon
'###################################################################################

Public Function ToBinary(ByVal n As Double) As String

'This function return the binary expansion of N where N is natural number - ykjeon

Dim temp As String

On Error GoTo ErrorHandler

    If Int(n) <> n Or n < 0 Then
        MsgBox "INPUT MUST BE A NONNEGATIVE INTEGER"
        Exit Function
        
    End If
    
    temp = ""
    
    Do
    
        temp = CStr(n Mod 2) & temp

        n = n \ 2

    Loop While n > 0
    
    ToBinary = temp
    
Exit Function

ErrorHandler:
    
    MsgBox "FAIL IN BINARY CONVERSION"
    Exit Function

End Function

Public Function DigitConversion( _
ByVal n As Double, _
ByVal BaseNum As Long) As String

Dim temp As String

On Error GoTo ErrorHandler

    If Int(n) <> n Or n < 0 Then
    
        MsgBox "INPUT MUST BE A NONNEGATIVE INTEGER"
        Exit Function
        
    Else
    
        If BaseNum = 1 Then
        
            DigitConversion = "-1"
            
        Else
    
            temp = ""
            
            Do
            
                temp = CStr(n Mod BaseNum) & temp
        
                n = n \ BaseNum
        
            Loop While n > 0
            
            DigitConversion = temp
        
        End If
    
    End If
    
Exit Function

ErrorHandler:
    
    MsgBox "FAIL IN DIGIT CONVERSION"
    Exit Function

End Function

Public Sub ReturnPermutator( _
ByVal Dimension As Long, _
ByVal BaseNum As Long, _
ByRef PermutArray() As Long)

Dim i As Long
Dim j As Long

Dim temp As Long
Dim Strtemp As String

    temp = BaseNum ^ (Dimension) - 1

    ReDim PermutArray(0 To temp, 1 To Dimension) As Long
    
    For i = 0 To temp
    
        Strtemp = StrReverse(DigitConversion(i, BaseNum))
        
        For j = 1 To Len(Strtemp)
        
            PermutArray(i, j) = CLng(Mid(Strtemp, j, 1))
        
        Next j
        
        For j = Len(Strtemp) + 1 To Dimension
        
            PermutArray(i, j) = 0
        
        Next j

    Next i

End Sub

Public Function Summation(ByRef values() As Double) As Double

Dim IndexMax As Long
Dim IndexMin As Long
Dim i As Integer
Dim temp As Double
    
    IndexMin = LBound(values)
    IndexMax = UBound(values)
    temp = 0
    
    For i = IndexMin To IndexMax
    
        temp = temp + values(i)
    
    Next i
    
    Summation = temp

End Function

Public Function max(ByVal x As Double, ByVal y As Double) As Double

    max = (x + y + Abs(x - y)) * 0.5

End Function

Public Function min(ByVal x As Double, ByVal y As Double) As Double

    'Min = x + y - Max(x, y)
    min = (x + y - Abs(x - y)) * 0.5

End Function

Public Function max_date(ByVal x As Date, ByVal y As Date) As Date

    max_date = (x + y + Abs(x - y)) * 0.5

End Function
Public Function min_date(ByVal x As Date, ByVal y As Date) As Date

    min_date = (x + y - Abs(x - y)) * 0.5

End Function

Public Function ReturnMax(ByRef values() As Double) As Double

Dim IndexMax As Long
Dim IndexMin As Long
Dim i As Integer
Dim temp As Double
    
    IndexMin = LBound(values)
    IndexMax = UBound(values)
    temp = values(IndexMin)
    
    If IndexMax = IndexMin Then
    
        ReturnMax = values(IndexMin)
        
    Else
    
        For i = IndexMin + 1 To IndexMax
        
            temp = max(temp, values(i))
        
        Next i
        
        ReturnMax = temp
        
    End If

End Function

Public Function ReturnMin(ByRef values() As Double) As Double

Dim IndexMax As Long
Dim IndexMin As Long
Dim i As Integer
Dim temp As Double
    
    IndexMin = LBound(values)
    IndexMax = UBound(values)
    temp = values(IndexMin)
    
    If IndexMax = IndexMin Then
    
        ReturnMin = values(IndexMin)
        
    Else
    
        For i = IndexMin + 1 To IndexMax
        
            temp = min(temp, values(i))
        
        Next i
        
        ReturnMin = temp
        
    End If

End Function

Public Function ReturnMinDate(ByRef values() As Date) As Date

Dim IndexMax As Long
Dim IndexMin As Long
Dim i As Integer
Dim temp As Date
    
    IndexMin = LBound(values)
    IndexMax = UBound(values)
    temp = values(IndexMin)
    
    If IndexMax = IndexMin Then
    
        ReturnMinDate = values(IndexMin)
        
    Else
    
        For i = IndexMin + 1 To IndexMax
        
            temp = min(temp, values(i))
        
        Next i
        
        ReturnMinDate = temp
        
    End If

End Function

Public Function lReturnMax(ByRef values() As Long) As Long

Dim IndexMax As Long
Dim IndexMin As Long
Dim i As Integer
Dim temp As Double
    
    IndexMin = LBound(values)
    IndexMax = UBound(values)
    temp = values(IndexMin)
    
    If IndexMax = IndexMin Then
    
        lReturnMax = values(IndexMin)
        
    Else
    
        For i = IndexMin + 1 To IndexMax
        
            temp = max(temp, values(i))
        
        Next i
        
        lReturnMax = temp
        
    End If

End Function

Public Function lReturnMin(ByRef values() As Long) As Long

Dim IndexMax As Long
Dim IndexMin As Long
Dim i As Integer
Dim temp As Double
    
    IndexMin = LBound(values)
    IndexMax = UBound(values)
    temp = values(IndexMin)
    
    If IndexMax = IndexMin Then
    
        lReturnMin = values(IndexMin)
        
    Else
    
        For i = IndexMin + 1 To IndexMax
        
            temp = min(temp, values(i))
        
        Next i
        
        lReturnMin = temp
        
    End If

End Function

Public Function Combin(ByVal n As Long, ByVal m As Long) As Double

Dim i As Long
Dim k As Long
Dim temp As Double
    
    temp = 1

    If Int(n) <> n Or Int(m) <> m Or n < m Then
    
        MsgBox "Input Values for Combination function are not valid"
        Exit Function
        
    Else
    
        If m = 0 Or m = n Then
        
            Combin = 1
        
        ElseIf m = 1 Or n - m = 1 Then
        
            Combin = n
            
        Else
        
            k = max(n - m, m)
            
            For i = 0 To n - k - 1
                
                temp = temp * (n - i) / (n - k - i)
            
            Next i
            
            Combin = temp
                        
        End If
        
    End If
        
End Function

Public Sub ReturnTangentVector( _
ByRef p1() As Double, _
ByRef p2() As Double, _
ByRef TangentVector() As Double)

Dim i As Long
Dim j As Long
Dim Dimension As Long

    Dimension = UBound(p1(), 1) - LBound(p1(), 1) + 1
    
    ReDim TangentVector(1 To Dimension) As Double
    
    For i = LBound(p1(), 1) To UBound(p1(), 1)
    
        TangentVector(i - LBound(p1(), 1) + 1) = _
        (p2(i) - p1(i))
    
    Next i

End Sub

Public Function ReturnVectorNorm( _
ByRef p1() As Double, _
ByRef p2() As Double, _
Optional ByVal NormIndex As Long = 2) As Double

Dim i As Long
Dim v() As Double

Dim temp As Double

Dim Dimension As Long

    Dimension = UBound(p1(), 1) - LBound(p1(), 1) + 1
    
    ReDim v(1 To Dimension) As Double
    
    For i = LBound(p1(), 1) To UBound(p1(), 1)
    
        v(i - LBound(p1(), 1) + 1) = Abs(p2(i) - p1(i))
    
    Next i

    If NormIndex = 0 Then
    
        temp = ReturnMax(v())
        
        ReturnVectorNorm = temp
        
    Else
    
        For i = 1 To Dimension
        
            temp = temp + (v(i)) ^ (NormIndex)
        
        Next i
        
        ReturnVectorNorm = (temp) ^ (1 / NormIndex)
        
    End If

End Function

Public Function Sorting_dbl( _
ByRef values() As Double, _
ByRef OrderedValues() As Double, _
Optional ByVal DataOdering As String = "ASC", Optional min_inx As Integer = -1, Optional max_inx As Integer = -1) As Double

Dim i As Long
Dim j As Long

Dim IndexMax As Long
Dim IndexMin As Long
Dim SwapArray() As Double
Dim temp As Double

On Error GoTo ErrorHandler

    IndexMax = UBound(values)
    IndexMin = LBound(values)
    
    ReDim OrderedValues(IndexMin To IndexMax) As Double
    ReDim SwapArray(IndexMin To IndexMax) As Double
    
    For i = IndexMin To IndexMax
    
        SwapArray(i) = values(i)
    
    Next i
    
    For i = IndexMin + 1 To IndexMax
    
        temp = values(i)
        
        For j = i - 1 To IndexMin Step -1
    
            If SwapArray(j) > values(i) Then
            
                SwapArray(j + 1) = SwapArray(j)
                SwapArray(j) = temp
                
            End If
    
        Next j
        
    Next i
    
    If min_inx > 0 Then
    
        IndexMin = min_inx
        
    End If
    
    If max_inx > 0 Then
    
        IndexMax = max_inx
        
    End If
    
    ReDim OrderedValues(LBound(values) To LBound(values) + IndexMax - IndexMin) As Double
    
    If LCase(DataOdering) = LCase("ASC") Then
        
        For i = IndexMin To IndexMax
        
            OrderedValues(LBound(values) + i - IndexMin) = SwapArray(i)
        
        Next i
        
    Else
    
        For i = IndexMin To IndexMax
        
            OrderedValues(LBound(values) + i - IndexMin) = SwapArray(IndexMax - i + IndexMin)
        
        Next i
    
    End If
    
    Sorting_dbl = 1
    
    Exit Function
    
ErrorHandler:

    Sorting_dbl = -1
    
    Exit Function

End Function

Public Function Sorting( _
ByRef values() As Variant, _
ByRef OrderedValues() As Variant, _
Optional ByVal DataOdering As String = "ASC") As Double

Dim i As Long
Dim j As Long

Dim IndexMax As Long
Dim IndexMin As Long
Dim SwapArray() As Variant
Dim temp As Variant

On Error GoTo ErrorHandler

    IndexMax = UBound(values)
    IndexMin = LBound(values)
    
    ReDim OrderedValues(IndexMin To IndexMax) As Variant
    ReDim SwapArray(IndexMin To IndexMax) As Variant
    
    For i = IndexMin To IndexMax
    
        SwapArray(i) = values(i)
    
    Next i
    
    For i = IndexMin + 1 To IndexMax
    
        temp = values(i)
        
        For j = i - 1 To IndexMin Step -1
    
            If SwapArray(j) > values(i) Then
            
                SwapArray(j + 1) = SwapArray(j)
                SwapArray(j) = temp
                
            End If
    
        Next j
        
    Next i
    
    If LCase(DataOdering) = LCase("ASC") Then
        
        For i = IndexMin To IndexMax
        
            OrderedValues(i) = SwapArray(i)
        
        Next i
        
    Else
    
        For i = IndexMin To IndexMax
        
            OrderedValues(i) = SwapArray(IndexMax - i + IndexMin)
        
        Next i
    
    End If
    
    Sorting = 1
    
    Exit Function
    
ErrorHandler:

    Sorting = -1
    
    Exit Function

End Function

Public Function Moment( _
ByRef values() As Double, _
ByVal MomentOrder As Long) As Double

Dim IndexMax As Long
Dim IndexMin As Long
Dim i As Long
Dim temp As Double
    
    temp = 0
    IndexMax = UBound(values)
    IndexMin = LBound(values)
    
    If IndexMax >= IndexMin Then
    
        For i = IndexMin To IndexMax
    
            temp = temp + values(i) ^ MomentOrder / (IndexMax - IndexMin + 1)
    
        Next i
        
        Moment = temp
    
    Else
    
        Moment = -999
        
    End If

End Function

Public Function CMoment( _
ByRef values() As Double, _
ByVal MomentOrder As Long) As Double

Dim i As Long
Dim IndexMax As Long
Dim IndexMin As Long
Dim temp As Double
Dim Average As Double
    
    temp = 0
    IndexMax = UBound(values)
    IndexMin = LBound(values)
    
    Average = Moment(values, 1)
    
    If MomentOrder = 1 Then
    
        CMoment = Average
    
    Else
    
        If IndexMax > IndexMin Then
    
            For i = IndexMin To IndexMax
        
                temp = temp + (values(i) - Average) ^ MomentOrder / (IndexMax - IndexMin)
        
            Next i
            
            CMoment = temp
            
        ElseIf IndexMax = IndexMin Then
        
            CMoment = temp
            
        Else
        
            CMoment = -999
            
        End If
        
    End If

End Function

Public Function ReturnAVG( _
ByRef values() As Double) As Double
    
    ReturnAVG = Moment(values(), 1)

End Function

Public Function ReturnSTDEV( _
ByRef values() As Double) As Double

Dim temp As Double

    temp = Sqr(CMoment(values(), 2))

    If temp > 0 Then

        ReturnSTDEV = Sqr(CMoment(values(), 2))
        
    Else
    
        ReturnSTDEV = -1
        
    End If

End Function

Public Function ReturnNCSTDEV( _
ByRef values() As Double) As Double

Dim IndexMax As Long
Dim IndexMin As Long
Dim i As Long
Dim temp As Double

    temp = 0
    
    IndexMax = UBound(values)
    IndexMin = LBound(values)
    
    If IndexMax - IndexMin >= 0 Then
        
        For i = IndexMin To IndexMax
    
            temp = temp + values(i) * values(i) / (IndexMax - IndexMin + 1)
    
        Next i
    
        ReturnNCSTDEV = Sqr(temp)

    Else
    
        ReturnNCSTDEV = -1
        
    End If

End Function

Public Function ReturnABSSTDEV( _
ByRef values() As Double) As Double

Dim i As Long
Dim IndexMax As Long
Dim IndexMin As Long
Dim Average As Double
Dim temp As Double

    temp = 0
    
    IndexMax = UBound(values)
    IndexMin = LBound(values)

    If IndexMax > IndexMin Then
    
        Average = ReturnAVG(values())
    
        For i = IndexMin To IndexMax
    
            temp = temp + Abs(values(i) - Average) / (IndexMax - IndexMin)
    
        Next i
        
        ReturnABSSTDEV = temp
    
    ElseIf IndexMax = IndexMin Then
    
        ReturnABSSTDEV = 0
        
    Else
    
        ReturnABSSTDEV = -999
        
    End If

End Function


Function LogSummation( _
ByRef ReturnArray() As Double, _
ByVal MomentOrder As Long) As Double

Dim i As Long
Dim IndexMax As Long
Dim IndexMin As Long
Dim LogReturn() As Double

    IndexMax = UBound(ReturnArray)
    IndexMin = LBound(ReturnArray)
        
    If IndexMax <= IndexMin + 1 Then
        
        LogSummation = -999
        
        Exit Function
        
    Else
    
        ReDim LogReturn(IndexMin To IndexMax) As Double
        
    
        
        For i = IndexMin To IndexMax
            
            LogReturn(i) = Log(1 + ReturnArray(i))
        
        Next i
        
            
        LogSummation = Summation(LogReturn())
        
    End If

End Function

Function LogMoment( _
ByRef PriceArray() As Double, _
ByVal MomentOrder As Long, _
Optional ByVal DateOrdering As String = "ASC") As Double

Dim i As Long
Dim IndexMax As Long
Dim IndexMin As Long
Dim LogReturn() As Double

    IndexMax = UBound(PriceArray)
    IndexMin = LBound(PriceArray)
        
    If IndexMax <= IndexMin + 1 Then
        
        LogMoment = -999
        
        Exit Function
        
    Else
    
        ReDim LogReturn(IndexMin To IndexMax - 1) As Double
        
        If LCase(DateOrdering) = LCase("ASC") Then
        
            For i = IndexMin To IndexMax - 1
                
                LogReturn(i) = Log(PriceArray(i + 1) / PriceArray(i))
            
            Next i

        Else
        
            For i = IndexMin To IndexMax - 1
            
                LogReturn(i) = Log(PriceArray(i) / PriceArray(i + 1))
            
            Next i
        
        End If
            
        LogMoment = Moment(LogReturn(), MomentOrder)
        
    End If

End Function

Function CLogMoment( _
ByRef PriceArray() As Double, _
ByVal MomentOrder As Long, _
Optional ByVal DateOrdering As String = "ASC") As Double

Dim i As Long
Dim IndexMax As Long
Dim IndexMin As Long
Dim LogReturn() As Double

    IndexMax = UBound(PriceArray)
    IndexMin = LBound(PriceArray)
        
    If IndexMax <= IndexMin + 1 Then
        
        CLogMoment = -999
        
        Exit Function
        
    Else
    
        ReDim LogReturn(IndexMin To IndexMax - 1) As Double
        
        If LCase(DateOrdering) = LCase("ASC") Then
        
            For i = IndexMin To IndexMax - 1
                
                LogReturn(i) = Log(PriceArray(i + 1) / PriceArray(i))
            
            Next i

        Else
        
            For i = IndexMin To IndexMax - 1
            
                LogReturn(i) = Log(PriceArray(i) / PriceArray(i + 1))
            
            Next i
        
        End If
            
        CLogMoment = CMoment(LogReturn(), MomentOrder)
        
    End If

End Function

Function GetDrift( _
ByRef PriceArray() As Double, _
Optional ByVal AnnualizeFactor As Long = 252, _
Optional ByVal DateOrdering As String = "ASC") As Double

    GetDrift = LogMoment(PriceArray(), 1, DateOrdering) * AnnualizeFactor

End Function

Function GetVol( _
ByRef PriceArray() As Double, _
Optional ByVal AnnualizeFactor As Long = 252, _
Optional ByVal DateOrdering As String = "ASC") As Double

    GetVol = Sqr(LogMoment(PriceArray(), 2, DateOrdering) * AnnualizeFactor)

End Function

Function GetCVol( _
ByRef PriceArray() As Double, _
Optional ByVal AnnualizeFactor As Long = 252, _
Optional ByVal DateOrdering As String = "ASC") As Double

    GetCVol = Sqr(CLogMoment(PriceArray(), 2, DateOrdering) * AnnualizeFactor)

End Function

Function GetSkew( _
ByRef PriceArray() As Double, _
Optional ByVal DateOrdering As String = "ASC") As Double

    GetSkew = LogMoment(PriceArray(), 3, DateOrdering) / (GetVol(PriceArray(), 1, DateOrdering) ^ 3)

End Function

Function GetCSkew( _
ByRef PriceArray() As Double, _
Optional ByVal DateOrdering As String = "ASC") As Double

    GetCSkew = CLogMoment(PriceArray(), 3, DateOrdering) / (GetCVol(PriceArray(), 1, DateOrdering) ^ 3)

End Function

Function GetKurtosis( _
ByRef PriceArray() As Double, _
Optional ByVal DateOrdering As String = "ASC") As Double

    GetKurtosis = LogMoment(PriceArray(), 4, DateOrdering) / (GetVol(PriceArray(), 1, DateOrdering) ^ 4)

End Function

Function GetCKurtosis( _
ByRef PriceArray() As Double, _
Optional ByVal DateOrdering As String = "ASC") As Double

    GetCKurtosis = CLogMoment(PriceArray(), 4, DateOrdering) / (GetCVol(PriceArray(), 1, DateOrdering) ^ 4)

End Function

Function GetCCORR( _
ByRef PriceArrayX() As Double, _
ByRef PriceArrayY() As Double, _
Optional ByVal DateOrdering As String = "ASC" _
) As Double

Dim i As Long
Dim temp As Double
Dim IndexMax As Long
Dim IndexMin As Long
Dim AverageX As Double
Dim AverageY As Double

    temp = 0
    IndexMax = UBound(PriceArrayX)
    IndexMin = LBound(PriceArrayX)
    
    If (IndexMax <= IndexMin) Or ((IndexMax - IndexMin) <> (UBound(PriceArrayY(), 1) - LBound(PriceArrayY(), 1))) Then
    
        GetCCORR = -999
            
        Exit Function
        
    Else
            
        AverageX = GetDrift(PriceArrayX(), 1, DateOrdering)
        AverageY = GetDrift(PriceArrayY(), 1, DateOrdering)
        
        If LCase(DateOrdering) = LCase("ASC") Then
        
            For i = IndexMin To IndexMax - 1
            
                temp = temp + _
                (Log(PriceArrayX(i + 1) / PriceArrayX(i)) - AverageX) * _
                (Log(PriceArrayY(i + 1) / PriceArrayY(i)) - AverageY) _
                / (IndexMax - IndexMin - 1)
            
            Next i
        
        Else
        
            For i = IndexMin To IndexMax - 1
            
                temp = temp + _
                (Log(PriceArrayX(i) / PriceArrayX(i + 1)) - AverageX) * _
                (Log(PriceArrayY(i) / PriceArrayY(i + 1)) - AverageY) _
                / (IndexMax - IndexMin - 1)
            
            Next i
        
        End If
        
        GetCCORR = temp / (GetCVol(PriceArrayX(), 1, DateOrdering) * GetCVol(PriceArrayY(), 1, DateOrdering))
        
    End If

End Function

Function GetCORR( _
ByRef PriceArrayX() As Double, _
ByRef PriceArrayY() As Double _
) As Double

Dim i As Long
Dim temp As Double
Dim IndexMax As Long
Dim IndexMin As Long

    temp = 0
    IndexMax = UBound(PriceArrayX)
    IndexMin = LBound(PriceArrayX)
    
    If (IndexMax <= IndexMin) Or ((IndexMax - IndexMin) <> (UBound(PriceArrayY(), 1) - LBound(PriceArrayY(), 1))) Then
    
        GetCORR = -999
            
        Exit Function
        
    Else
        
        For i = IndexMin To IndexMax - 1
        
            temp = temp + _
            (Log(PriceArrayX(i + 1) / PriceArrayX(i)) * Log(PriceArrayY(i + 1) / PriceArrayY(i))) _
            / (IndexMax - IndexMin)
        
        Next i
    
        GetCORR = temp / (GetVol(PriceArrayX(), 1, "ASC") * GetVol(PriceArrayY(), 1, "ASC"))
        
    End If

End Function


Public Function indicator(condition As Boolean) As Double

    If condition Then
    
        indicator = 1
        
    Else
        
        indicator = 0
        
    End If

End Function