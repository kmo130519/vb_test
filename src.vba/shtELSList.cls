Option Explicit
Option Base 0

Private Enum HEDG_TR_TYPE_CODE
    SELF = 1
    BTB = 2
    all = 3
End Enum

Public Sub btnReadELSList_Click()

    Dim target_date As Date
    target_date = Range("target_date").value
    
    Dim tdate_str As String
    tdate_str = date2str(target_date)
   
    Dim selected_hedge_type_code As String
    Select Case Range("hedge_type").value
    Case "자체": selected_hedge_type_code = HEDG_TR_TYPE_CODE.SELF
    Case "BTB": selected_hedge_type_code = HEDG_TR_TYPE_CODE.BTB
    Case "전체": selected_hedge_type_code = HEDG_TR_TYPE_CODE.all
    End Select
    
    Dim rng As Range
    Set rng = Range("TBL_ELS_LIST")
    rng.ClearContents
    
    Dim els_list As Variant
    els_list = rng.value

    Dim adoCon As New adoDB.Connection
    Dim oCmd As New adoDB.Command
    Dim oRS As New adoDB.Recordset
    
    Call connectDB(adoCon, TNS_SERVICE_NAME, USER_ID, PASSWORD)
    
    Dim bind_variable() As String
    Dim bind_value() As Variant
    ReDim bind_variable(1) As String
    ReDim bind_value(1) As Variant
    bind_variable(1) = ":tdate"
    bind_value(1) = tdate_str
    
    Dim sql As String
    sql = getSQL(SQL_PATH_ELS, bind_variable, bind_value)
    
    If Range("including_ids_yn").value = True Then
        sql = getSQL(SQL_PATH_IDS, bind_variable, bind_value) + " union all " + sql
    End If
    
    With oCmd

        .ActiveConnection = adoCon
        .CommandType = adCmdText
        .CommandText = sql

        oRS.Open .Execute

    End With

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    Dim i As Integer
    Dim j As Integer
    i = 1

    Do Until oRS.EOF
        
        If selected_hedge_type_code = HEDG_TR_TYPE_CODE.all Or selected_hedge_type_code = oRS(0) Then
    
            els_list(i, 1) = i
            
            For j = 0 To oRS.Fields.count - 1
                els_list(i, j + 2) = oRS(j)
            Next j
            
            i = i + 1
        End If

        oRS.MoveNext

    Loop
    
    With rng.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    For i = 1 To rng.Rows.count
        If i Mod 2 = 0 Then
            With rng.Rows(i).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent3
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With
        End If
    Next i
    
    rng.value = els_list
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    oRS.Close
    
    Call disconnectDB(adoCon)
    
End Sub