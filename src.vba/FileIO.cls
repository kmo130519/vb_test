Option Explicit

Public Function WrtFile( _
ByVal Dimension As Long, _
ByVal SavePathName As String, _
ByVal DerivID As String, _
ByRef NodeValue() As Double, _
ByRef px() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long _
) As Integer

Dim i As Long
Dim j As Long
Dim filesystem As Variant
Dim txtfile As Variant
Dim LineStr As String

Dim Series() As Long
Dim tmpNode() As Long
Dim tmpSeq() As String

On Error GoTo ErrorHandler

    Dim nodecls As New NodeControl

    Set filesystem = CreateObject("Scripting.FileSystemObject")
    Set txtfile = filesystem.CreateTextFile(SavePathName & DerivID, True)
    
    Call nodecls.ToSeries(Dimension, Series(), nmin(), nmax())
    
    txtfile.writeline (Dimension)
    
    For j = 1 To Dimension
    
        ReDim tmpSeq(nmin(j) To nmax(j)) As String
    
        For i = nmin(j) To nmax(j)
        
           tmpSeq(i) = CStr(px(j, i))
                    
        Next i
        
        LineStr = Join(tmpSeq(), "*")
        
        txtfile.writeline (LineStr)
    
    Next j
    
    ReDim tmpSeq(0 To Series(Dimension) - 1) As String
    
    For i = 0 To Series(Dimension) - 1
    
        Call nodecls.SeqToNode(Dimension, i, tmpNode(), Series(), nmin(), nmax())
        
        tmpSeq(i) = CStr(nodecls.ReturnNodeValue(tmpNode(), NodeValue()))
        
    Next i

    LineStr = Join(tmpSeq(), "*")
    
    txtfile.writeline (LineStr)
        
    txtfile.Close
    WrtFile = 1
    Exit Function
    
ErrorHandler:
    
    If txtfile <> Null Then
        txtfile.Close
    End If
    WrtFile = -1
    
    Err.Raise Err.number, "FileIO.wrtFile", "ERROR IN WRITING TEXT FILE " & Err.description
    'MsgBox Err.Number
    Exit Function

End Function

Public Function ReadFile( _
ByVal SaveFileName As String, _
ByRef NodeValue() As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long _
) As Integer

Dim i As Long
Dim j As Long
Dim Dimension As Long

Dim LineStrArray() As String
Dim LineStr As String

Dim Series() As Long
Dim tmpNode() As Long

On Error GoTo ErrorHandler

    Dim nodecls As New NodeControl
    Dim funcls As New Functions
    
    Open SaveFileName For Input As #1
    
    i = 0
    
    Do While Not EOF(1)
    
        Line Input #1, LineStr

        If i = 0 Then

            Dimension = CLng(LineStr)
            
            If Dimension = 1 Then
            
                ReDim NodeValue(nmin(1) To nmax(1)) As Double
            
            ElseIf Dimension = 2 Then
            
                ReDim NodeValue(nmin(1) To nmax(1), nmin(2) To nmax(2)) As Double
            
            ElseIf Dimension = 3 Then
            
                ReDim NodeValue(nmin(1) To nmax(1), nmin(2) To nmax(2), nmin(3) To nmax(3)) As Double
            
            End If
            
            ReDim px(1 To Dimension, funcls.lReturnMin(nmin()) To funcls.lReturnMax(nmax())) As Double
            ReDim dx(1 To Dimension, funcls.lReturnMin(nmin()) + 1 To funcls.lReturnMax(nmax())) As Double
            
            Call nodecls.ToSeries(Dimension, Series(), nmin(), nmax())
        
        ElseIf i >= 1 And i <= Dimension Then

            LineStrArray = Split(CStr(LineStr), "*")
            
            For j = LBound(LineStrArray()) To UBound(LineStrArray())
            
                px(i, j + nmin(i) - LBound(LineStrArray())) = CDbl(LineStrArray(j))
                
                If j > LBound(LineStrArray()) Then
                
                    dx(i, j) = px(i, j + nmin(i) - LBound(LineStrArray())) - px(i, j - 1 + nmin(i) - LBound(LineStrArray()))
                    
                End If
            
            Next j

        ElseIf i = Dimension + 1 Then
        
            LineStrArray = Split(CStr(LineStr), "*")
            
            For j = LBound(LineStrArray()) To UBound(LineStrArray())
            
                Call nodecls.SeqToNode(Dimension, j - LBound(LineStrArray()), tmpNode(), Series(), nmin(), nmax())
                
                Call nodecls.GetNodeValue(Dimension, tmpNode(), NodeValue(), CDbl(LineStrArray(j)))
            
            Next j
        
        End If
        
        i = i + 1

    Loop
    
    Close #1
    ReadFile = 1
    Exit Function
    
ErrorHandler:

    Close #1
    MsgBox Err.description
    ReadFile = -1
    
    Exit Function

End Function