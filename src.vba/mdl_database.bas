Option Explicit

Public Function connectDB(ByRef objADODB As adoDB.Connection, service_name As String, user_name As String, PASSWORD As String) As Boolean

    Dim strConnection As String
    strConnection = "Provider=" & PROVIDER & ";Data Source=" & service_name & ";User ID=" & user_name & ";Password=" & PASSWORD & ";Persist Security Info=True"
    
    objADODB.Open (strConnection)
    
    If objADODB.State <> adStateOpen Then
        connectDB = False
        Exit Function
    End If
    
    connectDB = True

End Function

Public Function disconnectDB(ByRef objADODB As adoDB.Connection) As Boolean

    objADODB.Close
    
    If objADODB.State <> adStateClosed Then
        disconnectDB = False
        Exit Function
    End If
    
    disconnectDB = True

End Function

Public Function getSQL(sql_filepath As String, bind_variable() As String, bind_value() As Variant) As String
'sql_filepath에 있는 sql 파일을 읽어, bind_variable을 bind_value binding 한 sql string 생성

    Dim sql As String
    Dim fileline As String
    Dim comment_start As Variant
    Dim quotation_start As Variant
    Dim quotation_end As Variant
    Dim quotation As String
    Dim i As Integer
        
    Open sql_filepath For Input As #1
    
    Do While Not EOF(1)
    
        Line Input #1, fileline
        
        '주석 제거
        comment_start = InStr(1, fileline, "--", vbTextCompare)
        If comment_start > 0 Then
            fileline = Left(fileline, comment_start - 1)
        End If
        
        '따옴표 제거
        quotation_start = InStr(1, fileline, """", vbTextCompare)
        quotation_end = InStr(quotation_start + 1, fileline, """", vbTextCompare)
        
        If quotation_start > 0 Then
            quotation = Mid(fileline, quotation_start, quotation_end - quotation_start + 1)
            fileline = Replace(fileline, quotation, "", , , vbTextCompare)
        Else
            quotation = ""
        End If
        
        'bind variable 처리
        For i = 1 To UBound(bind_variable)
            Select Case VarType(bind_value(i))
            Case vbString
                fileline = Replace(fileline, bind_variable(i), "'" & CStr(bind_value(i)) & "'", , , vbTextCompare)
            Case vbDouble
                fileline = Replace(fileline, bind_variable(i), CStr(bind_value(i)), , , vbTextCompare)
            Case vbNull, vbEmpty
                fileline = Replace(fileline, bind_variable(i), "Null", , , vbTextCompare)
            End Select
        Next i
        
        sql = sql & Chr(32) & fileline
    
    Loop
    
    Close #1
    
    getSQL = sql

End Function