Option Explicit

Private greek_value_ As Double

Public date_location As Integer
Public strike_location As Integer
Public greek_type As String

Private strike_ As Double
Private maturity_ As Date
Private strike_abs_() As Double
Private date_abs_() As Date

Public greek_1_1 As Double
Public greek_1_2 As Double
Public greek_2_1 As Double
Public greek_2_2 As Double


Public Sub set_input_greek(greek_type As String, greek_value As Double, strike_abs() As Double, date_abs() As Date, strike As Double, Maturity As Date)

    strike_abs_ = strike_abs
    date_abs_ = date_abs
    greek_value_ = greek_value
    
    strike_ = strike
    maturity_ = Maturity
    Me.greek_type = greek_type
    

End Sub

Public Sub split_greek()
                     

    '---------------------------------------
    ' Handle the boundaries
    '---------------------------------------
    If maturity_ < date_abs_(LBound(date_abs_)) Then
        maturity_ = date_abs_(LBound(date_abs_))
    ElseIf maturity_ > date_abs_(UBound(date_abs_)) Then
        maturity_ = date_abs_(UBound(date_abs_))
    End If
                     
    If strike_ < strike_abs_(LBound(strike_abs_)) Then
        strike_ = strike_abs_(LBound(strike_abs_))
    ElseIf strike_ > strike_abs_(UBound(strike_abs_)) Then
        strike_ = strike_abs_(UBound(strike_abs_))
    End If

    date_location = find_location_date(date_abs_, maturity_)
    strike_location = find_location(strike_abs_, strike_)
    
    greek_1_1 = (strike_abs_(strike_location + 1) - strike_) * (date_abs_(date_location + 1) - maturity_) _
             / (strike_abs_(strike_location + 1) - strike_abs_(strike_location)) / (date_abs_(date_location + 1) - date_abs_(date_location)) _
             * greek_value_
    greek_1_2 = (strike_ - strike_abs_(strike_location)) * (date_abs_(date_location + 1) - maturity_) _
              / (strike_abs_(strike_location + 1) - strike_abs_(strike_location)) / (date_abs_(date_location + 1) - date_abs_(date_location)) _
              * greek_value_
    greek_2_1 = (strike_abs_(strike_location + 1) - strike_) * (maturity_ - date_abs_(date_location)) _
             / (strike_abs_(strike_location + 1) - strike_abs_(strike_location)) / (date_abs_(date_location + 1) - date_abs_(date_location)) _
             * greek_value_
    greek_2_2 = (strike_ - strike_abs_(strike_location)) * (maturity_ - date_abs_(date_location)) _
             / (strike_abs_(strike_location + 1) - strike_abs_(strike_location)) / (date_abs_(date_location + 1) - date_abs_(date_location)) _
             * greek_value_


End Sub