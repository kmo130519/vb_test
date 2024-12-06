Option Explicit

Public Function ToBinary(ByVal n As Long) As String

Dim temp As String

On Error GoTo ErrorHandler

    If Int(n) <> n Or n < 0 Then
    
        MsgBox "INPUT MUST BE NONNEGATIVE INTEGER"
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
    
    MsgBox "ERROR IN BINARY FUNCTION " & Err.description
    Exit Function

End Function

Public Sub ToSeries( _
ByVal Dimension As Long, _
ByRef Series() As Long, _
ByRef nmin() As Long, _
ByRef nmax() As Long)

Dim i As Long

    ReDim Series(0 To Dimension) As Long
    
    Series(0) = 1
    
    For i = 1 To Dimension
    
        Series(i) = Series(i - 1) * (nmax(i) - nmin(i) + 1)
    
    Next i

End Sub

Public Sub ToSeries2( _
ByVal Dimension As Long, _
ByRef Series() As Long, _
ByRef SeriesNode() As Long, _
ByRef nmin() As Long, _
ByRef nmax() As Long)

Dim i As Long
Dim j As Long
Dim Node() As Long

    ReDim Series(0 To Dimension) As Long
    
    Series(0) = 1
    
    For i = 1 To Dimension
    
        Series(i) = Series(i - 1) * (nmax(i) - nmin(i) + 1)
    
    Next i
    
    ReDim SeriesNode(0 To Series(Dimension) - 1, 1 To Dimension) As Long
    
    For i = 1 To Series(Dimension) - 1
    
        Call SeqToNode(Dimension, i, Node(), Series(), nmin(), nmax())
        
        For j = 1 To Dimension
        
            SeriesNode(i, j) = Node(j)
        
        Next j
    
    Next i

End Sub

Public Sub SeqToNode( _
ByVal Dimension As Long, _
ByVal seq As Long, _
ByRef Node() As Long, _
ByRef Series() As Long, _
ByRef nmin() As Long, _
ByRef nmax() As Long)

Dim i As Long
Dim q As Long

    ReDim Node(1 To Dimension) As Long
    
    q = seq
    
    For i = Dimension To 1 Step -1
    
        Node(i) = q \ Series(i - 1) + nmin(i)
        
        q = q Mod Series(i - 1)
    
    Next i

End Sub

Public Function NodeToSeq( _
ByVal Dimension As Long, _
ByRef Node() As Long, _
ByRef Series() As Long, _
ByRef nmin() As Long, _
ByRef nmax() As Long) As Long

Dim i As Long
Dim Lngtmp As Long
    
    For i = 1 To Dimension
    
        Lngtmp = Lngtmp + Series(i - 1) * (Node(i) - nmin(i))
    
    Next i
    
    NodeToSeq = Lngtmp

End Function
'--------------------------------------------------
' Input : Point, px, dx, nmin, nmax
' Output : Node, Ratio
' DESC: This function looks up the nearest node to the point which does not exceed the point.
'       The ratio is the ratio between (point - node ) / dx, in case the point does not exist
'        on the node, but between nodes.
'--------------------------------------------------
Public Function ReturnNode( _
ByRef Point() As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long, _
ByRef Node() As Long, _
ByRef Ratio() As Double) _
As Long

Dim i As Long
Dim j As Long
Dim k As Long
    
'On Error GoTo ErrorHandler

    ReDim Node(LBound(Point(), 1) To UBound(Point(), 1)) As Long
    ReDim Ratio(LBound(Point(), 1) To UBound(Point(), 1)) As Double
    
    For i = LBound(Point(), 1) To UBound(Point(), 1)
    
        If Point(i) >= px(i, nmax(i)) Then
                    
            Ratio(i) = 1
            Node(i) = nmax(i) - 1
        
        ElseIf Point(i) <= px(i, nmin(i)) Then
        
            Ratio(i) = 0
            Node(i) = nmin(i)
        
        Else
        
        'LBarrier가 Px=0에서 거리가 얼마나 되는지 계산
            k = Int(CDbl((Point(i) - px(i, nmin(i))) / dx(i, nmin(i) + 1))) + nmin(i)
            
            If k >= nmax(i) Then
                k = nmax(i)
            End If
            
            If Point(i) >= px(i, k) Then
            
                For j = k To nmax(i) - 1
                
                    If Point(i) >= px(i, j) And Point(i) < px(i, j + 1) Then
                    
                        Ratio(i) = (Point(i) - px(i, j)) / dx(i, j + 1)
                        Node(i) = j
                        
                        Exit For
                        
                    End If
                
                Next j
                
            Else
            
                For j = k - 1 To nmin(i) Step -1
                
                    If Point(i) >= px(i, j) And Point(i) < px(i, j + 1) Then
                    
                        Ratio(i) = (Point(i) - px(i, j)) / dx(i, j + 1)
                        Node(i) = j
                        
                        Exit For
                        
                    End If
                
                Next j
            
            End If
                        
        End If
    
    Next i
    
    ReturnNode = 1
    
    Exit Function
    
'ErrorHandler:
'
'    MsgBox "ERROR IN NODE RETURN FUNCTION " & Err.Description
'    ReturnNode = -1
'
'    Exit Function

End Function

'Public Function ReturnNode( _
'ByRef Point() As Double, _
'ByRef Px() As Double, _
'ByRef dx() As Double, _
'ByRef Nmin() As Long, _
'ByRef Nmax() As Long, _
'ByRef Node() As Long, _
'ByRef Ratio() As Double) _
'As Long
'
'Dim i As Long
'Dim j As Long
'
'On Error GoTo ErrorHandler
'
'    ReDim Node(LBound(Point(), 1) To UBound(Point(), 1)) As Long
'    ReDim Ratio(LBound(Point(), 1) To UBound(Point(), 1)) As Double
'
'    For i = LBound(Point(), 1) To UBound(Point(), 1)
'
'        If Point(i) >= Px(i, Nmax(i)) Then
'
'            Ratio(i) = 1
'            Node(i) = Nmax(i) - 1
'
'        ElseIf Point(i) <= Px(i, Nmin(i)) Then
'
'            Ratio(i) = 0
'            Node(i) = Nmin(i)
'
'        Else
'
'            For j = Nmin(i) To Nmax(i) - 1
'
'                If Point(i) >= Px(i, j) And Point(i) < Px(i, j + 1) Then
'
'                    Ratio(i) = (Point(i) - Px(i, j)) / dx(i, j + 1)
'                    Node(i) = j
'
'                    Exit For
'
'                End If
'
'            Next j
'
'        End If
'
'    Next i
'
'    ReturnNode = 1
'
'    Exit Function
'
'ErrorHandler:
'
'    MsgBox "ERROR IN NODE RETURN FUNCTION " & Err.Description
'    ReturnNode = -1
'
'    Exit Function
'
'End Function

Public Function ReturnWeight( _
ByVal Dimension As Long, _
ByRef Ratio() As Double, _
ByRef WeightFactor() As Double _
) As Long

Dim i As Long
Dim j As Long

Dim BinaryPoint() As Long

Dim Strtemp As String

On Error GoTo ErrorHandler
    
    ReDim BinaryPoint(0 To 2 ^ (Dimension) - 1, 1 To Dimension) As Long
    ReDim WeightFactor(0 To 2 ^ (Dimension) - 1) As Double
    
    For i = 0 To 2 ^ (Dimension) - 1
    
        Strtemp = ToBinary(i)
        WeightFactor(i) = 1
        
        For j = 1 To Len(Strtemp)
        
            BinaryPoint(i, j) = CLng(Mid(Strtemp, Len(Strtemp) + 1 - j, 1))
            
        Next j
        
        For j = Len(Strtemp) + 1 To Dimension
        
            BinaryPoint(i, j) = 0
    
        Next j
    
        For j = 1 To Dimension
            
            WeightFactor(i) = _
            WeightFactor(i) * ((1 - BinaryPoint(i, j)) * (1 - Ratio(j)) + BinaryPoint(i, j) * Ratio(j))
        
        Next j

    Next i
    
    ReturnWeight = 1
    
    Exit Function
    
ErrorHandler:

    MsgBox "ERROR IN BinaryPoint FUNCTION " & Err.description
    
    ReturnWeight = -1

End Function

Public Sub GetNodeValue( _
ByVal Dimension As Long, _
ByRef Node() As Long, _
ByRef ResultVector() As Double, _
ByVal InPutVal As Double _
)

On Error GoTo ErrorHandler
    
    Select Case Dimension
    
        Case 1
        
            ResultVector(Node(1)) = InPutVal
        
        Case 2
        
            ResultVector(Node(1), Node(2)) = InPutVal
        
        Case 3
        
            ResultVector(Node(1), Node(2), Node(3)) = InPutVal
        
        Case Else
        
            MsgBox "OVER DIMENSION LIMIT"
            
            Exit Sub
        
    End Select
    
    Exit Sub
    
ErrorHandler:

    MsgBox "ERROR IN VALUE INSERT AT NODE " & Err.description
    
    Exit Sub

End Sub

Public Function ReturnNodeValue( _
ByRef Node() As Long, _
ByRef ResultVector() As Double _
) As Double
        
Dim Dimension As Long

On Error GoTo ErrorHandler

    Dimension = UBound(Node(), 1) - LBound(Node(), 1) + 1
    
    Select Case Dimension
    
        Case 1
        
            ReturnNodeValue = ResultVector(Node(1))
        
        Case 2
        
            ReturnNodeValue = ResultVector(Node(1), Node(2))
        
        Case 3
        
            ReturnNodeValue = ResultVector(Node(1), Node(2), Node(3))
        
        Case Else
        
            MsgBox "OVER DIMENSION LIMIT"
            
            ReturnNodeValue = -1
            
            Exit Function
        
    End Select
    
    Exit Function
    
ErrorHandler:

    MsgBox "ERROR IN VALUE RETURN AT NODE " & Err.description
    
    ReturnNodeValue = -1
    
    Exit Function

End Function

Public Function ReturnValue( _
ByRef Point() As Double, _
ByRef ResultVector() As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long _
) As Double

Dim i As Long
Dim j As Long

Dim Node() As Long
Dim Nodetemp() As Long
Dim Ratio() As Double

Dim Dimension As Long
Dim baseIndex As Long

Dim BinaryPoint() As Long
Dim WeightFactor() As Double

Dim Strtemp As String
Dim Dbltemp As Double
        
'On Error GoTo ErrorHandler

    Dimension = UBound(Point(), 1) - LBound(Point(), 1) + 1
    baseIndex = LBound(Point(), 1)
    
    i = ReturnNode(Point(), px(), dx(), nmin(), nmax(), Node(), Ratio())
        
    ReDim Nodetemp(LBound(Node(), 1) To UBound(Node(), 1)) As Long
            
    Dbltemp = 0

    ReDim BinaryPoint(0 To 2 ^ (Dimension) - 1, 1 To Dimension) As Long
    ReDim WeightFactor(0 To 2 ^ (Dimension) - 1) As Double

    For i = 0 To 2 ^ (Dimension) - 1
    
        Strtemp = ToBinary(i)
        WeightFactor(i) = 1

        For j = 1 To Len(Strtemp)

            BinaryPoint(i, j) = CLng(Strings.Mid(Strtemp, Len(Strtemp) + 1 - j, 1))

        Next j

        For j = Len(Strtemp) + 1 To Dimension

            BinaryPoint(i, j) = 0

        Next j

        For j = 1 To Dimension

            Nodetemp(baseIndex + j - 1) = Node(baseIndex + j - 1) + BinaryPoint(i, j)

            WeightFactor(i) = _
            WeightFactor(i) * ((1 - BinaryPoint(i, j)) * (1 - Ratio(j)) + BinaryPoint(i, j) * Ratio(j))

        Next j
        
        Dbltemp = Dbltemp + WeightFactor(i) * ReturnNodeValue(Nodetemp(), ResultVector())
    
    Next i
    
    ReturnValue = Dbltemp
    
    Exit Function
    
'ErrorHandler:
'
'    MsgBox "ERROR IN VALUE RETURN FUNCTION " & Err.Description
'
'    ReturnValue = -1
'
'    Exit Function

End Function

Public Function ReturnNodeTangent( _
ByRef Node() As Long, _
ByRef ResultVector() As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long, _
ByVal Axis As Long, _
Optional ByVal Direction As Long = 0, _
Optional ByVal NodeStep As Long = 1 _
) As Double

Dim i As Long
Dim j As Long

Dim UNode() As Long
Dim DNode() As Long

Dim Dimension As Long
Dim baseIndex As Long

Dim ULength As Double
Dim DLength As Double

Dim UValue As Double
Dim DValue As Double

Dim temp As Double
        
On Error GoTo ErrorHandler

    Dimension = UBound(Node(), 1) - LBound(Node(), 1) + 1
    baseIndex = LBound(Node(), 1)
        
    ReDim UNode(LBound(Node(), 1) To UBound(Node(), 1)) As Long
    ReDim DNode(LBound(Node(), 1) To UBound(Node(), 1)) As Long
    
    For i = 1 To Dimension
    
        UNode(baseIndex + i - 1) = Node(baseIndex + i - 1)
        DNode(baseIndex + i - 1) = Node(baseIndex + i - 1)
    
    Next i
    
    If Node(baseIndex + Axis - 1) >= nmax(Axis) - NodeStep + 1 Then
    
        DNode(baseIndex + Axis - 1) = nmax(Axis) - NodeStep
        
        temp = ReturnNodeTangent(DNode(), ResultVector, px(), dx(), nmin(), nmax(), Axis, Direction, NodeStep)
        
        ReturnNodeTangent = temp
        
    ElseIf Node(baseIndex + Axis - 1) <= nmin(Axis) + NodeStep - 1 Then
    
        UNode(baseIndex + Axis - 1) = nmin(Axis) + NodeStep
        
        temp = ReturnNodeTangent(UNode(), ResultVector, px(), dx(), nmin(), nmax(), Axis, Direction, NodeStep)
        
        ReturnNodeTangent = temp
        
    Else

        ULength = px(Axis, Node(baseIndex + Axis - 1) + NodeStep) - px(Axis, Node(baseIndex + Axis - 1))
        DLength = px(Axis, Node(baseIndex + Axis - 1)) - px(Axis, Node(baseIndex + Axis - 1) - NodeStep)
    
        Select Case Direction
        
            Case -1
            
                DNode(baseIndex + Axis - 1) = Node(baseIndex + Axis - 1) - NodeStep
                
                UValue = ReturnNodeValue(UNode(), ResultVector())
                DValue = ReturnNodeValue(DNode(), ResultVector())
            
                ReturnNodeTangent = (UValue - DValue) / DLength
            
            Case 0
            
                DNode(baseIndex + Axis - 1) = Node(baseIndex + Axis - 1) - NodeStep
                UNode(baseIndex + Axis - 1) = Node(baseIndex + Axis - 1) + NodeStep
                
                UValue = ReturnNodeValue(UNode(), ResultVector())
                DValue = ReturnNodeValue(DNode(), ResultVector())
            
                ReturnNodeTangent = _
                DLength * UValue / (ULength * (ULength + DLength)) + _
                (ULength - DLength) * ReturnNodeValue(Node(), ResultVector()) / (ULength * DLength) - _
                ULength * DValue / (DLength * (ULength + DLength))
            
            Case 1
            
                UNode(baseIndex + Axis - 1) = Node(baseIndex + Axis - 1) + NodeStep
                
                UValue = ReturnNodeValue(UNode(), ResultVector())
                DValue = ReturnNodeValue(DNode(), ResultVector())
            
                ReturnNodeTangent = (UValue - DValue) / ULength
            
            Case Else
            
                MsgBox "CHECK DIFFERENTIAL DIRECTION"
                
                ReturnNodeTangent = -99
        
        End Select
        
    End If
    
    Exit Function
    
ErrorHandler:

    MsgBox "ERROR IN RETURN TANGENT AT NODE " & Err.description
    
    ReturnNodeTangent = -99
    
    Exit Function

End Function

Public Function ReturnTangent( _
ByRef Point() As Double, _
ByRef ResultVector() As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long, _
ByVal Axis As Long, _
Optional ByVal Direction As Long = 0, _
Optional ByVal NodeStep As Long = 1 _
) As Double

Dim i As Long
Dim j As Long

Dim Node() As Long
Dim Nodetemp() As Long
Dim Ratio() As Double

Dim Dimension As Long
Dim baseIndex As Long

Dim BinaryPoint() As Long
Dim WeightFactor() As Double

Dim Strtemp As String
Dim Dbltemp As Double
        
On Error GoTo ErrorHandler

    Dimension = UBound(Point(), 1) - LBound(Point(), 1) + 1
    baseIndex = LBound(Point(), 1)
    
    i = ReturnNode(Point(), px(), dx(), nmin(), nmax(), Node(), Ratio())
        
    ReDim Nodetemp(LBound(Node()) To UBound(Node())) As Long
    ReDim BinaryPoint(0 To 2 ^ (Dimension) - 1, 1 To Dimension) As Long
    ReDim WeightFactor(0 To 2 ^ (Dimension) - 1) As Double
    
    Dbltemp = 0
    
    For i = 0 To 2 ^ (Dimension) - 1
    
        Strtemp = ToBinary(i)
        WeightFactor(i) = 1
        
        For j = 1 To Len(Strtemp)
        
            BinaryPoint(i, j) = CLng(Strings.Mid(Strtemp, Len(Strtemp) + 1 - j, 1))
            
        Next j
        
        For j = Len(Strtemp) + 1 To Dimension
        
            BinaryPoint(i, j) = 0
    
        Next j
    
        For j = 1 To Dimension
        
            Nodetemp(baseIndex + j - 1) = Node(baseIndex + j - 1) + BinaryPoint(i, j)
            
            WeightFactor(i) = _
            WeightFactor(i) * ((1 - BinaryPoint(i, j)) * (1 - Ratio(j)) + BinaryPoint(i, j) * Ratio(j))
        
        Next j
        
        Dbltemp = Dbltemp + WeightFactor(i) * _
        ReturnNodeTangent(Nodetemp(), ResultVector(), px(), dx(), nmin(), nmax(), Axis, Direction, NodeStep)
    
    Next i
    
    ReturnTangent = Dbltemp
    
    Exit Function
    
ErrorHandler:

    MsgBox "ERROR IN TANGENT RETURN FUNCTION " & Err.description
    
    ReturnTangent = -99
    
    Exit Function

End Function

Public Function ReturnNodeConvexity( _
ByRef Node() As Long, _
ByRef ResultVector() As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long, _
ByVal Axis As Long, _
Optional ByVal NodeStep As Long = 1 _
) As Double

Dim i As Long
Dim j As Long

Dim UNode() As Long
Dim DNode() As Long

Dim Dimension As Long
Dim baseIndex As Long

Dim ULength As Double
Dim DLength As Double

Dim UValue As Double
Dim DValue As Double

Dim temp As Double
        
On Error GoTo ErrorHandler

    Dimension = UBound(Node(), 1) - LBound(Node(), 1) + 1
    baseIndex = LBound(Node(), 1)
        
    ReDim UNode(LBound(Node()) To UBound(Node())) As Long
    ReDim DNode(LBound(Node()) To UBound(Node())) As Long
    
    For i = 1 To Dimension
    
        UNode(baseIndex + i - 1) = Node(baseIndex + i - 1)
        DNode(baseIndex + i - 1) = Node(baseIndex + i - 1)
    
    Next i
    
    If Node(baseIndex + Axis - 1) >= nmax(Axis) - NodeStep + 1 Then
    
        DNode(baseIndex + Axis - 1) = nmax(Axis) - NodeStep
        
        temp = ReturnNodeConvexity(DNode(), ResultVector, px(), dx(), nmin(), nmax(), Axis, NodeStep)
        
        ReturnNodeConvexity = temp
        
    ElseIf Node(baseIndex + Axis - 1) <= nmin(Axis) + NodeStep - 1 Then
    
        UNode(baseIndex + Axis - 1) = nmin(Axis) + NodeStep
        
        temp = ReturnNodeConvexity(UNode(), ResultVector, px(), dx(), nmin(), nmax(), Axis, NodeStep)
        
        ReturnNodeConvexity = temp
        
    Else
    
        ULength = px(Axis, Node(baseIndex + Axis - 1) + NodeStep) - px(Axis, Node(baseIndex + Axis - 1))
        DLength = px(Axis, Node(baseIndex + Axis - 1)) - px(Axis, Node(baseIndex + Axis - 1) - NodeStep)
            
        DNode(baseIndex + Axis - 1) = Node(baseIndex + Axis - 1) - NodeStep
        UNode(baseIndex + Axis - 1) = Node(baseIndex + Axis - 1) + NodeStep
        
        UValue = ReturnNodeValue(UNode(), ResultVector())
        DValue = ReturnNodeValue(DNode(), ResultVector())
        
        ReturnNodeConvexity = _
        2 * UValue / (ULength * (ULength + DLength)) - _
        2 * ReturnNodeValue(Node(), ResultVector()) / (ULength * DLength) + _
        2 * DValue / (DLength * (ULength + DLength))
        
    End If
    
    Exit Function
    
ErrorHandler:

    MsgBox "ERROR IN RETURN CONVEXITY AT NODE " & Err.description
    
    ReturnNodeConvexity = -99
    
    Exit Function

End Function

Public Function ReturnConvexity( _
ByRef Point() As Double, _
ByRef ResultVector() As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long, _
ByVal Axis As Long, _
Optional ByVal NodeStep As Long = 1 _
) As Double

Dim i As Long
Dim j As Long

Dim Node() As Long
Dim Nodetemp() As Long
Dim Ratio() As Double

Dim Dimension As Long
Dim baseIndex As Long

Dim BinaryPoint() As Long
Dim WeightFactor() As Double

Dim Strtemp As String
Dim Dbltemp As Double
        
On Error GoTo ErrorHandler

    Dimension = UBound(Point(), 1) - LBound(Point(), 1) + 1
    baseIndex = LBound(Point(), 1)
    
    i = ReturnNode(Point(), px(), dx(), nmin(), nmax(), Node(), Ratio())
        
    ReDim Nodetemp(LBound(Node()) To UBound(Node())) As Long
    ReDim BinaryPoint(0 To 2 ^ (Dimension) - 1, 1 To Dimension) As Long
    ReDim WeightFactor(0 To 2 ^ (Dimension) - 1) As Double
    
    Dbltemp = 0
    
    For i = 0 To 2 ^ (Dimension) - 1
    
        Strtemp = ToBinary(i)
        WeightFactor(i) = 1
        
        For j = 1 To Len(Strtemp)
        
            BinaryPoint(i, j) = CLng(Strings.Mid(Strtemp, Len(Strtemp) + 1 - j, 1))
            
        Next j
        
        For j = Len(Strtemp) + 1 To Dimension
        
            BinaryPoint(i, j) = 0
    
        Next j
    
        For j = 1 To Dimension
        
            Nodetemp(baseIndex + j - 1) = Node(baseIndex + j - 1) + BinaryPoint(i, j)
            
            WeightFactor(i) = _
            WeightFactor(i) * ((1 - BinaryPoint(i, j)) * (1 - Ratio(j)) + BinaryPoint(i, j) * Ratio(j))
        
        Next j
        
        Dbltemp = Dbltemp + WeightFactor(i) * _
        ReturnNodeConvexity(Nodetemp(), ResultVector(), px(), dx(), nmin(), nmax(), Axis, NodeStep)
    
    Next i
    
    ReturnConvexity = Dbltemp
    
    Exit Function
    
ErrorHandler:

    MsgBox "ERROR IN CONVEXITY RETURN FUNCTION " & Err.description
    
    ReturnConvexity = -99
    
    Exit Function

End Function

Public Function ReturnNodeMixedDeriv( _
ByRef Node() As Long, _
ByRef ResultVector() As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long, _
ByVal Axis1 As Long, _
ByVal Axis2 As Long, _
Optional ByVal NodeStep As Long = 1 _
) As Double

Dim i As Long
Dim j As Long

Dim UNode() As Long
Dim DNode() As Long

Dim Dimension As Long
Dim baseIndex As Long

Dim ULength As Double
Dim DLength As Double

Dim UValue As Double
Dim DValue As Double

Dim temp As Double
        
On Error GoTo ErrorHandler

    Dimension = UBound(Node(), 1) - LBound(Node(), 1) + 1
    baseIndex = LBound(Node(), 1)
        
    ReDim UNode(LBound(Node()) To UBound(Node())) As Long
    ReDim DNode(LBound(Node()) To UBound(Node())) As Long
    
    For i = 1 To Dimension
    
        UNode(baseIndex + i - 1) = Node(baseIndex + i - 1)
        DNode(baseIndex + i - 1) = Node(baseIndex + i - 1)
    
    Next i
    
    If Node(baseIndex + Axis1 - 1) >= nmax(Axis1) - NodeStep + 1 Then

        DNode(baseIndex + Axis1 - 1) = nmax(Axis1) - NodeStep

        temp = ReturnNodeMixedDeriv(DNode(), ResultVector, px(), dx(), nmin(), nmax(), Axis1, Axis2, NodeStep)

        ReturnNodeMixedDeriv = temp
        
    ElseIf Node(baseIndex + Axis2 - 1) >= nmax(Axis2) - NodeStep + 1 Then
    
         DNode(baseIndex + Axis2 - 1) = nmax(Axis2) - NodeStep

        temp = ReturnNodeMixedDeriv(DNode(), ResultVector, px(), dx(), nmin(), nmax(), Axis1, Axis2, NodeStep)

        ReturnNodeMixedDeriv = temp

    ElseIf Node(baseIndex + Axis1 - 1) <= nmin(Axis1) + NodeStep - 1 Then

        UNode(baseIndex + Axis1 - 1) = nmin(Axis1) + NodeStep

        temp = ReturnNodeMixedDeriv(UNode(), ResultVector, px(), dx(), nmin(), nmax(), Axis1, Axis2, NodeStep)

        ReturnNodeMixedDeriv = temp

    ElseIf Node(baseIndex + Axis2 - 1) <= nmin(Axis2) + NodeStep - 1 Then

        UNode(baseIndex + Axis2 - 1) = nmin(Axis2) + NodeStep

        temp = ReturnNodeMixedDeriv(UNode(), ResultVector, px(), dx(), nmin(), nmax(), Axis1, Axis2, NodeStep)

        ReturnNodeMixedDeriv = temp

    Else
    
    
        ULength = px(Axis1, Node(baseIndex + Axis1 - 1) + 1) - px(Axis1, Node(baseIndex + Axis1 - 1))
        DLength = px(Axis1, Node(baseIndex + Axis1 - 1)) - px(Axis1, Node(baseIndex + Axis1 - 1) - 1)
        
        UNode(baseIndex + Axis1 - 1) = Node(baseIndex + Axis1 - 1) + 1
        DNode(baseIndex + Axis1 - 1) = Node(baseIndex + Axis1 - 1) - 1
        
        UValue = ReturnNodeTangent(UNode(), ResultVector(), px(), dx(), nmin(), nmax(), Axis2, 0, 1)
        DValue = ReturnNodeTangent(DNode(), ResultVector(), px(), dx(), nmin(), nmax(), Axis2, 0, 1)
    
        ReturnNodeMixedDeriv = _
        DLength * UValue / (ULength * (ULength + DLength)) + _
        (ULength - DLength) * ReturnNodeTangent(Node(), ResultVector(), px(), dx(), nmin(), nmax(), Axis2, 0, 1) / (ULength * DLength) - _
        ULength * DValue / (DLength * (ULength + DLength))
        
    End If
    
    Exit Function
    
ErrorHandler:

    MsgBox "ERROR IN RETURN CONVEXITY AT NODE " & Err.description
    
    ReturnNodeMixedDeriv = -99
    
    Exit Function

End Function

Public Function ReturnMixedDeriv( _
ByRef Point() As Double, _
ByRef ResultVector() As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long, _
ByVal Axis1 As Long, _
ByVal Axis2 As Long _
) As Double

Dim i As Long
Dim j As Long

Dim Node() As Long
Dim Nodetemp() As Long
Dim Ratio() As Double

Dim NodeTangent() As Double

Dim Dimension As Long
Dim baseIndex As Long

Dim BinaryPoint() As Long
Dim WeightFactor() As Double

Dim Strtemp As String
Dim Dbltemp As Double
        
On Error GoTo ErrorHandler

    Dimension = UBound(Point(), 1) - LBound(Point(), 1) + 1
    baseIndex = LBound(Point(), 1)
    
    i = ReturnNode(Point(), px(), dx(), nmin(), nmax(), Node(), Ratio())
        
    ReDim Nodetemp(LBound(Node()) To UBound(Node())) As Long
    ReDim BinaryPoint(0 To 2 ^ (Dimension) - 1, 1 To Dimension) As Long
    ReDim WeightFactor(0 To 2 ^ (Dimension) - 1) As Double
    
    Dbltemp = 0
    
    For i = 0 To 2 ^ (Dimension) - 1
    
        Strtemp = ToBinary(i)
        WeightFactor(i) = 1
        
        For j = 1 To Len(Strtemp)
        
            BinaryPoint(i, j) = CLng(Strings.Mid(Strtemp, Len(Strtemp) + 1 - j, 1))
            
        Next j
        
        For j = Len(Strtemp) + 1 To Dimension
        
            BinaryPoint(i, j) = 0
    
        Next j
    
        For j = 1 To Dimension
        
            Nodetemp(baseIndex + j - 1) = Node(baseIndex + j - 1) + BinaryPoint(i, j)
            
            WeightFactor(i) = _
            WeightFactor(i) * ((1 - BinaryPoint(i, j)) * (1 - Ratio(j)) + BinaryPoint(i, j) * Ratio(j))
        
        Next j
        
        Dbltemp = Dbltemp + WeightFactor(i) * _
        ReturnNodeMixedDeriv(Nodetemp(), ResultVector(), px(), dx(), nmin(), nmax(), Axis1, Axis2)
    
    Next i
    
    ReturnMixedDeriv = Dbltemp
    
    Exit Function
    
ErrorHandler:

    MsgBox "ERROR IN CONVEXITY RETURN FUNCTION " & Err.description
    
    ReturnMixedDeriv = -99
    
    Exit Function

End Function

Public Function ReturnNodeThirdMoment( _
ByRef Node() As Long, _
ByRef ResultVector() As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long, _
ByVal Axis As Long _
) As Double

Dim i As Long
Dim j As Long

Dim UNode() As Long
Dim DNode() As Long

Dim Dimension As Long
Dim baseIndex As Long

Dim ULength As Double
Dim DLength As Double

Dim UValue As Double
Dim DValue As Double

Dim temp As Double
        
On Error GoTo ErrorHandler

    Dimension = UBound(Node(), 1) - LBound(Node(), 1) + 1
    baseIndex = LBound(Node(), 1)
        
    ReDim UNode(LBound(Node()) To UBound(Node())) As Long
    ReDim DNode(LBound(Node()) To UBound(Node())) As Long
    
    For i = 1 To Dimension
    
        UNode(baseIndex + i - 1) = Node(baseIndex + i - 1)
        DNode(baseIndex + i - 1) = Node(baseIndex + i - 1)
    
    Next i
    
    If Node(baseIndex + Axis - 1) >= nmax(Axis) - 1 Then
    
        DNode(baseIndex + Axis - 1) = nmax(Axis) - 2
        
        temp = ReturnNodeThirdMoment(DNode(), ResultVector, px(), dx(), nmin(), nmax(), Axis)
        
        ReturnNodeThirdMoment = temp
        
    ElseIf Node(baseIndex + Axis - 1) <= nmin(Axis) + 1 Then
    
        UNode(baseIndex + Axis - 1) = nmin(Axis) + 2
        
        temp = ReturnNodeThirdMoment(UNode(), ResultVector, px(), dx(), nmin(), nmax(), Axis)
        
        ReturnNodeThirdMoment = temp
        
    Else
    
        ULength = px(Axis, Node(baseIndex + Axis - 1) + 1) - px(Axis, Node(baseIndex + Axis - 1))
        DLength = px(Axis, Node(baseIndex + Axis - 1)) - px(Axis, Node(baseIndex + Axis - 1) - 1)
            
        DNode(baseIndex + Axis - 1) = Node(baseIndex + Axis - 1) - 1
        UNode(baseIndex + Axis - 1) = Node(baseIndex + Axis - 1) + 1
        
        UValue = ReturnNodeConvexity(UNode(), ResultVector(), px(), dx(), nmin(), nmax(), Axis, 1)
        DValue = ReturnNodeConvexity(DNode(), ResultVector(), px(), dx(), nmin(), nmax(), Axis, 1)
        
        ReturnNodeThirdMoment = (UValue - DValue) / (ULength + DLength)
        
    End If
    
    Exit Function
    
ErrorHandler:

    MsgBox "ERROR IN RETURN THIRD MOMENT AT NODE " & Err.description
    
    ReturnNodeThirdMoment = -99
    
    Exit Function

End Function

Public Function ReturnThirdMoment( _
ByRef Point() As Double, _
ByRef ResultVector() As Double, _
ByRef px() As Double, _
ByRef dx() As Double, _
ByRef nmin() As Long, _
ByRef nmax() As Long, _
ByVal Axis As Long _
) As Double

Dim i As Long
Dim j As Long

Dim Node() As Long
Dim Nodetemp() As Long
Dim Ratio() As Double

Dim Dimension As Long
Dim baseIndex As Long

Dim BinaryPoint() As Long
Dim WeightFactor() As Double

Dim Strtemp As String
Dim Dbltemp As Double
        
On Error GoTo ErrorHandler

    Dimension = UBound(Point(), 1) - LBound(Point(), 1) + 1
    baseIndex = LBound(Point(), 1)
    
    i = ReturnNode(Point(), px(), dx(), nmin(), nmax(), Node(), Ratio())
        
    ReDim Nodetemp(LBound(Node()) To UBound(Node())) As Long
    ReDim BinaryPoint(0 To 2 ^ (Dimension) - 1, 1 To Dimension) As Long
    ReDim WeightFactor(0 To 2 ^ (Dimension) - 1) As Double
    
    Dbltemp = 0
    
    For i = 0 To 2 ^ (Dimension) - 1
    
        Strtemp = ToBinary(i)
        WeightFactor(i) = 1
        
        For j = 1 To Len(Strtemp)
        
            BinaryPoint(i, j) = CLng(Mid(Strtemp, Len(Strtemp) + 1 - j, 1))
            
        Next j
        
        For j = Len(Strtemp) + 1 To Dimension
        
            BinaryPoint(i, j) = 0
    
        Next j
    
        For j = 1 To Dimension
        
            Nodetemp(baseIndex + j - 1) = Node(baseIndex + j - 1) + BinaryPoint(i, j)
            
            WeightFactor(i) = _
            WeightFactor(i) * ((1 - BinaryPoint(i, j)) * (1 - Ratio(j)) + BinaryPoint(i, j) * Ratio(j))
        
        Next j
        
        Dbltemp = Dbltemp + WeightFactor(i) * _
        ReturnNodeThirdMoment(Nodetemp(), ResultVector(), px(), dx(), nmin(), nmax(), Axis)
    
    Next i
    
    ReturnThirdMoment = Dbltemp
    
    Exit Function
    
ErrorHandler:

    MsgBox "ERROR IN CONVEXITY RETURN FUNCTION " & Err.description
    
    ReturnThirdMoment = -99
    
    Exit Function

End Function