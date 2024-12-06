Option Explicit

Private distribution_helper_ As New Distributions 'Helper class. Written by Jeon, YK.

Private Const tiny As Double = 0.0000000001
Private Const one As Double = 0.9999999999
'----------------------------------------------
' This class utilizes qunatlib xl addin
'----------------------------------------------


'----------------------------------------------
' This function utilizes qunatlib xl addin
' Function : sobol_numbers_2d
' DESC: generates sobol sequence with dimension of 2 * steps, size of no_of_paths
'       The array returns have dimension ( 1 to no_of_paths, 1 to steps, 1 to 2 )
'----------------------------------------------

Public Function sobol_numbers_2d(no_of_paths As Integer, steps As Long) As Double()

    Dim obj_id As String
    Dim numbers_variant As Variant '() As Variant
    Dim numbers() As Double
    Dim inx As Long
    Dim jnx As Integer
    
On Error GoTo ErrorHandler

    ReDim numbers(1 To no_of_paths, 1 To steps, 1 To 2) As Double
    
    '-----------------------------------------------------------------
    ' Dimension should be 2 times the number of time steps.
    '------------------------------------------------------------------
    
    obj_id = Application.run("qlSobolRsg", "_sobol_", steps * 2, 11, False, , True)
    
    numbers_variant = Application.run("qlVariates", obj_id, no_of_paths)
    
    For inx = 1 To no_of_paths
        
        For jnx = 1 To steps
        
            numbers(inx, jnx, 1) = numbers_variant(inx, 1 + (jnx - 1) * 2)
            numbers(inx, jnx, 2) = numbers_variant(inx, jnx * 2)
            
        Next jnx
        
    Next inx
    
    sobol_numbers_2d = numbers
    
    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsRandomGenerator.sobol_numbers_2d :" & Chr(13) & Err.source, Err.description '1003 <-- initialization error
    Exit Function

End Function


'----------------------------------------------
' This function utilizes qunatlib xl addin
' Function : sobol_numbers_1d
' DESC: generates sobol sequence with dimension of  steps, size of no_of_paths
'       The array returns have dimension ( 1 to no_of_paths, 1 to steps )
'----------------------------------------------
Public Function sobol_numbers_1d(no_of_paths As Integer, steps As Long) As Double()

    Dim obj_id As String
    Dim numbers_variant As Variant '() As Variant
    Dim numbers() As Double
    Dim inx As Long
    Dim jnx As Integer
    
On Error GoTo ErrorHandler

    ReDim numbers(1 To no_of_paths, 1 To steps) As Double
    
    '-----------------------------------------------------------------
    ' Dimension should be 2 times the number of time steps.
    '------------------------------------------------------------------
    
    obj_id = Application.run("qlSobolRsg", "_sobol_", steps, 1, False, , True)
    
    numbers_variant = Application.run("qlVariates", obj_id, no_of_paths)
    
    For inx = 1 To no_of_paths
        
        For jnx = 1 To steps
        
            numbers(inx, jnx) = numbers_variant(inx, jnx)
            
        Next jnx
        
    Next inx
    
    sobol_numbers_1d = numbers
    
    
    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsRandomGenerator.sobol_numbers_1d :" & Chr(13) & Err.source, Err.description '1003 <-- initialization error
    Exit Function

End Function

'----------------------------------------------
' This function utilizes qunatlib xl addin
' Function : norm_sobol_2d
' DESC: Generates quasi random numbers following normal distribution and correlation of corr.
'       Calls sobol_numbers_2d
'       The array returns have dimension ( 1 to no_of_paths, 1 to steps, 1 to 2 )
'----------------------------------------------

Public Function norm_sobol_2d(no_of_paths As Integer, steps As Long, Optional corr As Double = 0) As Double()

    Dim numbers() As Double
    Dim norm_numbers() As Double
    
    Dim temp_number1 As Double
    Dim temp_number2 As Double
    
    Dim inx As Integer
    Dim jnx As Integer
    
On Error GoTo ErrorHandler
    
    numbers = sobol_numbers_2d(no_of_paths, steps)
    
    ReDim norm_numbers(LBound(numbers, 1) To UBound(numbers, 2), LBound(numbers, 2) To UBound(numbers, 2), LBound(numbers, 3) To UBound(numbers, 3)) As Double
    
    For inx = LBound(numbers, 1) To UBound(numbers, 1)
    
        For jnx = LBound(numbers, 2) To UBound(numbers, 2)
    
            norm_numbers(inx, jnx, 1) = distribution_helper_.NormInv(numbers(inx, jnx, 1))
            norm_numbers(inx, jnx, 2) = corr * norm_numbers(inx, jnx, 1) + Sqr(1 - corr ^ 2) * distribution_helper_.NormInv(numbers(inx, jnx, 2))
                                    
        Next jnx
    
    Next inx
    
    norm_sobol_2d = norm_numbers

    Exit Function
    
ErrorHandler:

    Err.Raise vbObjectError + 1003, "clsRandomGenerator.norm_sobol_2d :" & Chr(13) & Err.source, Err.description '1003 <-- initialization error
    Exit Function


End Function


' OBSOLETE
Public Function uniform_numbers_2d(no_of_paths As Long, steps As Integer) As Double()

    Dim random_numbers() As Double
    Dim inx As Long
    Dim jnx As Integer

    
    ReDim random_numbers(1 To no_of_paths, 1 To steps, 1 To 2) As Double
    
    For inx = 1 To no_of_paths
    
        For jnx = 1 To steps
        
            
            random_numbers(inx, jnx, 1) = rnd
            random_numbers(inx, jnx, 2) = rnd
            
            If random_numbers(inx, jnx, 1) = 0 Then
            
                random_numbers(inx, jnx, 1) = tiny
                
            End If
            
            If random_numbers(inx, jnx, 2) = 0 Then
            
                random_numbers(inx, jnx, 2) = tiny
                
            End If
            
            If random_numbers(inx, jnx, 1) = 1 Then
            
                random_numbers(inx, jnx, 1) = one
                
            End If
            
            If random_numbers(inx, jnx, 2) = 1 Then
            
                random_numbers(inx, jnx, 2) = one
                
            End If
                
      
        Next jnx
        
    Next inx
    
    uniform_numbers_2d = random_numbers
    
End Function