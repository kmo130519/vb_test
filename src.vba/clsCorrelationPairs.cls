Option Explicit


Private asset1_() As String
Private asset2_() As String
Private corr_() As Double
Private corr_origin_() As Double
'Private min_corr_() As Double 'local correlation 추가 2019.3.27

Public Function copy_obj() As clsCorrelationPairs

    Dim inx As Integer
    Dim no_of_assets As Integer
    Dim rtn_obj As clsCorrelationPairs

On Error Resume Next

    no_of_assets = UBound(asset1_)
    
    If Err.number = 8 Then
        no_of_assets = 0
    End If

On Error GoTo ErrorHandler

    Set rtn_obj = New clsCorrelationPairs
    
    For inx = 1 To no_of_assets
    
        rtn_obj.set_corr asset1_(inx), asset2_(inx), corr_(inx)
        'rtn_obj.set_min_corr asset1_(inx), asset2_(inx), min_corr_(inx) 'local correlation 추가 2019.3.27
        
    Next inx
    
    Set copy_obj = rtn_obj
    
    Exit Function
    
ErrorHandler:

    raise_err "clsCorrelationPairs.get_corr", Err.description

End Function


Public Function rewind(asset_name_1 As String, asset_name_2 As String)

    Dim inx As Integer
    
    
On Error GoTo ErrorHandler

    inx = find_pair(asset_name_1, asset_name_2)
    
    If inx > 0 Then
    
        set_corr asset_name_1, asset_name_2, corr_origin_(inx)
        
    End If

    Exit Function
    
ErrorHandler:

    raise_err "bump_corr", Err.description

End Function


Public Function bump_corr(asset_name_1 As String, asset_name_2 As String, bump_value As Double)

    Dim inx As Integer
    
    
On Error GoTo ErrorHandler

    inx = find_pair(asset_name_1, asset_name_2)
    
    If inx > 0 Then
    
        set_corr asset_name_1, asset_name_2, corr_(inx) + bump_value
        
    End If

    Exit Function
    
ErrorHandler:

    raise_err "bump_corr", Err.description

End Function

Public Function get_corr(asset_name_1 As String, asset_name_2 As String) As Double

    Dim inx As Integer
    Dim rtn_value As Double
    
On Error GoTo ErrorHandler
            
    inx = find_pair(asset_name_1, asset_name_2)
    
    If inx > 0 Then
    
        rtn_value = corr_(inx)
    
    Else
    
        rtn_value = 0
        
    End If
    
    get_corr = rtn_value
    
    Exit Function
    
ErrorHandler:

    raise_err "clsCorrelationPairs.get_corr", Err.description
    

End Function

'local correlation 추가 2019. 3. 27
'Public Function get_min_corr(asset_name_1 As String, asset_name_2 As String) As Double
'
'    Dim inx As Integer
'    Dim rtn_value As Double
'
'On Error GoTo ErrorHandler
'
'    inx = find_pair(asset_name_1, asset_name_2)
'
'    If inx > 0 Then
'
'        rtn_value = min_corr_(inx)
'
'    Else
'
'        rtn_value = 0
'
'    End If
'
'    get_min_corr = rtn_value
'
'    Exit Function
'
'ErrorHandler:
'
'    raise_err "clsCorrelationPairs.get_min_corr", Err.description
'
'
'End Function

Public Sub set_corr(asset_name_1 As String, asset_name_2 As String, corr As Double)
    
    Dim inx As Integer
    
On Error GoTo ErrorHandler
    
    inx = find_pair(asset_name_1, asset_name_2)
    
    If inx <= 0 Then
    
        push_back_string asset1_, asset_name_1
        push_back_string asset2_, asset_name_2
        push_back_double corr_, corr
        push_back_double corr_origin_, corr
        
    Else
        
        corr_(inx) = corr
        'corr_origin_(inx) = corr
    
    End If
    

    Exit Sub
    
ErrorHandler:

    raise_err "clsCorrelationPairs.set_corr", Err.description
    

End Sub

'local correlation 추가 2019. 3. 27
'Public Sub set_min_corr(asset_name_1 As String, asset_name_2 As String, min_corr As Double)
'
'    Dim inx As Integer
'
'On Error GoTo ErrorHandler
'
'    inx = find_pair(asset_name_1, asset_name_2)
'
'    If inx <= 0 Then
'
'        push_back_string asset1_, asset_name_1
'        push_back_string asset2_, asset_name_2
'        push_back_double min_corr_, min_corr
'
'    Else
'
'        min_corr_(inx) = min_corr
'        'corr_origin_(inx) = corr
'
'    End If
'
'
'    Exit Sub
'
'ErrorHandler:
'
'    raise_err "clsCorrelationPairs.set_min_corr", Err.description
'
'
'End Sub


Private Function find_pair(asset_name_1 As String, asset_name_2 As String) As Integer

    Dim inx As Integer
    Dim rtn_value As Integer
    Dim no_of_pairs As Integer
    
On Error Resume Next

    no_of_pairs = UBound(asset1_)
    
    
    If Err.number = 9 Then
    
        rtn_value = 0
        
    Else
    
On Error GoTo ErrorHandler
            
        For inx = 1 To no_of_pairs
        
            If (asset_name_1 = asset1_(inx) And asset_name_2 = asset2_(inx)) _
            Or (asset_name_1 = asset2_(inx) And asset_name_2 = asset1_(inx)) Then
            
                rtn_value = inx
                Exit For
            
            End If
        
        Next inx
        
    End If
    
    find_pair = rtn_value
    
    Exit Function
    
ErrorHandler:

    raise_err "clsCorrelationPairs.find_pair", Err.description
    

End Function