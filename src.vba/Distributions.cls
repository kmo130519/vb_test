Option Explicit

'Reference Papers - ykjeon
'1) Better approximations to cumulative normal functions - Graeme west
'2) A short note on the numerical approximatioin of standard normal cumulative distribution and its inverse - Dridi
'3) Numerical computation of rectangular bivariate and trivariate normal and t probabilities  - Alan Genz

Const pi As Double = 3.14159265358979
Const tiny As Double = 0.0000000001
Const one As Double = 0.9999999999

Public Function NormInv(ByVal p As Double) As Double

'  Adapted for Microsoft Visual Basic from Peter Acklam's
'  This Algorithm originally written by John Herrero(2003-01-03)
'  for more details visit the site, [http://home.online.no/~pjacklam/notes/invnorm/]

'Define coefficients in rational approximations

'This function returns a inverse value of standard normal cumulative density function
'i.e. returns x such that Prob[Z <= x] = p where Z is a standard normal random variable is - ykjeon

Const A1 = -39.6968302866538
Const A2 = 220.946098424521
Const A3 = -275.928510446969
Const a4 = 138.357751867269
Const a5 = -30.6647980661472
Const a6 = 2.50662827745924

Const b1 = -54.4760987982241
Const b2 = 161.585836858041
Const b3 = -155.698979859887
Const b4 = 66.8013118877197
Const b5 = -13.2806815528857

Const c1 = -7.78489400243029E-03
Const C2 = -0.322396458041136
Const c3 = -2.40075827716184
Const c4 = -2.54973253934373
Const c5 = 4.37466414146497
Const c6 = 2.93816398269878

Const d1 = 7.78469570904146E-03
Const d2 = 0.32246712907004
Const d3 = 2.445134137143
Const d4 = 3.75440866190742

'Define break-points
Const p_low = 0.02425
Const p_high = 1 - p_low

'Define work variables
Dim q As Double, r As Double

'If argument out of bounds, raise error
If p <= 0 Or p >= 1 Then

    MsgBox "INPUT IS OUT OF RANGE"
    Exit Function
    
End If

If p < p_low Then
  'Rational approximation for lower region
  q = Sqr(-2 * Log(p))
  NormInv = (((((c1 * q + C2) * q + c3) * q + c4) * q + c5) * q + c6) / _
    ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
ElseIf p <= p_high Then
  'Rational approximation for lower region
  q = p - 0.5
  r = q * q
  NormInv = (((((A1 * r + A2) * r + A3) * r + a4) * r + a5) * r + a6) * q / _
    (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1)
ElseIf p < 1 Then
  'Rational approximation for upper region
  q = Sqr(-2 * Log(1 - p))
  NormInv = -(((((c1 * q + C2) * q + c3) * q + c4) * q + c5) * q + c6) / _
    ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
End If

End Function

Public Function NormCDF(ByVal x As Double) As Double

'This function returns Prob[Z <= x] where Z is a standard normal random variable - ykjeon

Dim xabs As Double
Dim temp As Double
Dim Exponential As Double

xabs = Abs(x)

If xabs > 37 Then

    NormCDF = 0

Else

    Exponential = Exp(-0.5 * xabs * xabs)
    
    If xabs < 7.07106781186547 Then
    
        temp = 3.52624965998911E-02 * xabs + 0.700383064443688
        temp = temp * xabs + 6.37396220353165
        temp = temp * xabs + 33.912866078383
        temp = temp * xabs + 112.079291497871
        temp = temp * xabs + 221.213596169931
        temp = temp * xabs + 220.206867912376
        
        NormCDF = Exponential * temp
        
        temp = 8.83883476483184E-02 * xabs + 1.75566716318264
        temp = temp * xabs + 16.064177579207
        temp = temp * xabs + 86.7807322029461
        temp = temp * xabs + 296.564248779674
        temp = temp * xabs + 637.333633378831
        temp = temp * xabs + 793.826512519948
        temp = temp * xabs + 440.413735824752
        
        NormCDF = NormCDF / temp
    
    Else
    temp = xabs + 0.65
    temp = xabs + 4 / temp
    temp = xabs + 3 / temp
    temp = xabs + 2 / temp
    temp = xabs + 1 / temp
    NormCDF = (Exponential / temp) / 2.506628274631
    
    End If

End If

If x > 0 Then

    NormCDF = 1 - NormCDF
    
End If

End Function

Public Function NormD( _
ByVal x As Double, _
Optional ByVal Mean As Double = 0, _
Optional ByVal sigma As Double = 1) As Double

'This function returns a value of standard normal density function at x - ykjeon

    NormD = 1 / (sigma * Sqr(2 * pi)) * Exp(-0.5 * (x - Mean) * (x - Mean) / (sigma * sigma))

End Function

Public Function BiNormCDFDW2( _
ByVal a As Double, _
ByVal b As Double, _
ByVal r As Double) _
As Double

'This approximation was developed by Drezner & Wesolowsky
'as their second verison(thus, DW2) in 1989
'DW2 method is single precision

Dim i As Integer
Dim x As Variant, w As Variant
Dim h1 As Double, h2 As Double
Dim LH As Double, h12 As Double
Dim h3 As Double, h5 As Double, h6 As Double, h7 As Double
Dim r1 As Double, r2 As Double, r3 As Double, rr As Double
Dim AA As Double, ab As Double

On Error GoTo ErrorHandler

x = Array(0.04691008, 0.23076534, 0.5, 0.76923466, 0.95308992)
w = Array(0.018854042, 0.038088059, 0.0452707394, 0.038088059, 0.018854042)
h1 = a
h2 = b

h12 = (h1 * h1 + h2 * h2) / 2

    If Abs(r) >= 0.7 Then
        r2 = 1 - r * r
        r3 = Sqr(r2)
        
        If r < 0 Then h2 = -h2
        
        h3 = h1 * h2
        h7 = Exp(-h3 / 2)
        
        If Abs(r) < 1 Then
        
            h6 = Abs(h1 - h2)
            h5 = h6 * h6 / 2
            h6 = h6 / r3
            AA = 0.5 - h3 / 8
            ab = 3 - 2 * AA * h5
            LH = 0.13298076 * h6 * ab * (1 - NormCDF(h6)) _
            - Exp(-h5 / r2) * (ab + AA * r2) * 0.053051647
            
            For i = 0 To 4
            
                r1 = r3 * x(i)
                rr = r1 * r1
                r2 = Sqr(1 - rr)
                LH = LH - w(i) * Exp(-h5 / rr) * (Exp(-h3 / (1 + r2)) / r2 / h7 - 1 - AA * rr)
            
            Next i
            
        End If
            
        BiNormCDFDW2 = LH * r3 * h7 + NormCDF((h1 + h2 - Abs(h1 - h2)) / 2)
            
        If r < 0 Then
        
            BiNormCDFDW2 = NormCDF(h1) - BiNormCDFDW2
            
        End If
        
    Else
        h3 = h1 * h2
        
        If r <> 0 Then
        
            For i = 0 To 4
            
                r1 = r * x(i)
                r2 = 1 - r1 * r1
                LH = LH + w(i) * Exp((r1 * h3 - h12) / r2) / Sqr(r2)
            
            Next i
            
        End If
        
        BiNormCDFDW2 = NormCDF(h1) * NormCDF(h2) + r * LH
        
    End If
    
Exit Function

ErrorHandler:
    
    MsgBox "ERROR IN BINORMAL APPROX."
    Exit Function
    
End Function

'General Brownian motion Z(t) = a*t + b*dB(t) (where a and b are constans, B(t) is standard Brownian motion)
'Geometric Brownian Motion G(t) = exp(Z(t)) - ykjeon

Public Function GeneralBMCDF( _
ByVal sigma As Double, _
ByVal Mu As Double, _
ByVal t As Double, _
ByVal x As Double)

'This function returns Prob[Z(t) <= x] where a = mu, b = sigma - ykjeon

    GeneralBMCDF = NormCDF((x - Mu * t) / (sigma * Sqr(t)))

End Function

Public Function GeneralMinima( _
ByVal sigma As Double, _
ByVal Mu As Double, _
ByVal t As Double, _
ByVal x As Double) As Double

'This function returns Prob[m(t) <= x] where m(t) = min{Z(s): 0<= s <= t} - ykjeon

    GeneralMinima = Exp(2 * Mu * x / (sigma * sigma)) * NormCDF((x + Mu * t) / (sigma * Sqr(t))) + _
    GeneralBMCDF(sigma, Mu, t, x)

End Function

Public Function GeneralMaxima( _
ByVal sigma As Double, _
ByVal Mu As Double, _
ByVal t As Double, _
ByVal x As Double) As Double

'This function returns Prob[M(t) >= x] where M(t) = max{Z(s): 0<= s <= t} - ykjeon

    GeneralMaxima = Exp(2 * Mu * x / (sigma * sigma)) * NormCDF((-x - Mu * t) / (sigma * Sqr(t))) + _
    (1 - GeneralBMCDF(sigma, Mu, t, x))

End Function

Public Function GeometricBMCDF( _
ByVal sigma As Double, _
ByVal Mu As Double, _
ByVal t As Double, _
ByVal x As Double)

'This function returns Prob[G(t) <= x] (recall G(t) denotes a geometric Brownian motion) - ykjeon

    GeometricBMCDF = GeneralBMCDF(sigma, Mu, t, Log(x))

End Function

Public Function GeometricMinima( _
ByVal sigma As Double, _
ByVal Mu As Double, _
ByVal t As Double, _
ByVal x As Double) As Double

'This function return Prob[m(t) <= x] where m(t) = min{(G(s): 0 <= s <= t} - ykjeon

    GeometricMinima = GeneralMinima(sigma, Mu, t, Log(x))
    
End Function

Public Function GeometricMaxima( _
ByVal sigma As Double, _
ByVal Mu As Double, _
ByVal t As Double, _
ByVal x As Double) As Double

'This function returns Prob[M(t) >= x] where M(t) = max{G(s): 0<= s <= t} - ykjeon

    GeometricMaxima = GeneralMaxima(sigma, Mu, t, Log(x))

End Function