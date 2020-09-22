VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSimplestFraction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function FindFraction(ByVal strNum As String, ByVal strDenum As String, ByRef dblNewNum As Double, ByRef dblNewDenum As Double, ByRef strDesc As String) As Boolean
    'Returns:
    '   'False' if the function has failed for some reason.
    '       ** reason is placed into referenced variable 'strDesc'
    '   'True' if the function has succeeded.
    '       ** succeeded does not necessarly mean it has simplified.
    '
    'Parameters:
    '   strNum      -   Any numerator
    '   strDenum    -   Any denumerator
    '   dblNewNum   -   The new simplified numerator
    '   dblNewDenum -   The new simplified denumerator
    '   strDesc     -   Description of what has occured.
    '
    'Example:
    '   (1) Input:
    '           strNum          = 3
    '           strDenum        = 9
    '       Output:
    '           dblNewNum       = 1
    '           dblNewDenum     = 3
    '           strDesc         = [depends, success]
    '
    '   (2) Input:
    '           strNum          = -5
    '           strDenum        = 30
    '       Output:
    '           dblNewNum       = -1
    '           dblNewDenum     = 6
    '           strDesc         = [depends, success]
    '
    '   (3) Input:
    '           strNum          = [ decimal , !numeric ]
    '           strDenum        = [ decimal , !numeric ]
    '       Output:
    '           dblNewNum       = 0
    '           dblNewDenum     = 0
    '           strDesc         = [depends, error]
    '
    'The previous examples pretty much explain how to use the funtion.
    '
    
    Dim dblTempNum          As Double       'Var to be manipulated.
    Dim dblTempNum2         As Double       'Original Numerator.
    Dim dblTempDenum        As Double       'Var to be manipulated.
    Dim dblTempDenum2       As Double       'Original Denumerator.
    Dim bMultiply           As Boolean      'This is used to know if an
                                            'operation occured.
    Dim bPositive           As Boolean      'Will our final value be positive?
    
    'Quickly check for errors in values that were passed..
    If ValidFraction(strNum, strDenum) = False Then
        strDesc = "Invalid fraction."
        Exit Function
    End If
    
    'Lets set our original variables..
    dblTempNum = Int(CDbl(strNum))
    dblTempNum2 = dblTempNum
    dblTempDenum = Int(CDbl(strDenum))
    dblTempDenum2 = dblTempDenum
    
    'Lets check our signs
    bPositive = IIf((dblTempNum / dblTempDenum) > 0, True, False)
    dblTempNum = Abs(dblTempNum)
    dblTempNum2 = Abs(dblTempNum2)
    dblTempDenum = Abs(dblTempDenum)
    dblTempDenum2 = Abs(dblTempDenum2)
    
    'Lets quickly make sure no decimals were passed..
    If (dblTempNum <> Abs(strNum)) Or (dblTempDenum <> Abs(strDenum)) Then
        strDesc = "Please do not use decimals."
        Exit Function
    End If
    
    'Try to see if the fraction is a multiple of the numerator (unless numerator is 1)..
    bMultiply = False
    
    Do Until (Int(dblTempNum) <> dblTempNum) Or (Int(dblTempDenum) <> dblTempDenum) Or (dblTempNum2 = 1)
        bMultiply = True
        dblTempNum = dblTempNum / dblTempNum2
        dblTempDenum = dblTempDenum / dblTempNum2
    Loop
    
    If bMultiply = True Then
    
        dblTempNum = dblTempNum * dblTempNum2
        dblTempDenum = dblTempDenum * dblTempNum2
        
    End If
    
    'Try to see if the fraction is a multiple of the denumerator (unless denumerator is 1)..
    bMultiply = False
    
    Do Until (Int(dblTempNum) <> dblTempNum) Or (Int(dblTempDenum) <> dblTempDenum) Or (dblTempDenum2 = 1)
        bMultiply = True
        dblTempNum = dblTempNum / dblTempDenum2
        dblTempDenum = dblTempDenum / dblTempDenum2
    Loop
    
    If bMultiply = True Then
    
        dblTempNum = dblTempNum * dblTempDenum2
        dblTempDenum = dblTempDenum * dblTempDenum2
        
    End If
    
    'Lets try dividing by 3, since it is the next odd number bigger then 1.
    bMultiply = False
    Do Until (Int(dblTempNum) <> dblTempNum) Or (Int(dblTempDenum) <> dblTempDenum)
        bMultiply = True
        dblTempNum = dblTempNum / 3
        dblTempDenum = dblTempDenum / 3
    Loop
    
    If bMultiply = True Then
        
        dblTempNum = dblTempNum * 3
        dblTempDenum = dblTempDenum * 3
        
    End If

    'Lets try dividing by 2, since it is the next pair number bigger then 0.
    bMultiply = False
    Do Until (Int(dblTempNum) <> dblTempNum) Or (Int(dblTempDenum) <> dblTempDenum)
        bMultiply = True
        dblTempNum = dblTempNum / 2
        dblTempDenum = dblTempDenum / 2
    Loop
    
    If bMultiply = True Then
        
        dblTempNum = dblTempNum * 2
        dblTempDenum = dblTempDenum * 2
        
    End If
    
    'This will quickly verify (mathematicly) if the values we found are correct.
    If (dblTempNum / dblTempDenum) = (dblTempNum2 / dblTempDenum2) Then
        dblNewNum = IIf(bPositive, dblTempNum, -dblTempNum)
        dblNewDenum = dblTempDenum
        strDesc = "Found " & dblNewNum & " / " & dblNewDenum
        FindFraction = True
    Else
        strDesc = "Failed check."
    End If
End Function

Private Function ValidFraction(ByVal strNum As String, ByVal strDenum As String) As Boolean
    'This will return 'True' if strNum and strDenum are valid.
    'This will remove any invalid errors, like 0 for denum.
    On Error GoTo err_handler:
    
    Dim dblTemp As Double
    
    dblTemp = CDbl(strNum) / CDbl(strDenum)
    
    If dblTemp = 0 Then
        GoTo exit_func:
    End If
    
    ValidFraction = True
    
exit_func:
    Exit Function
err_handler:
    Resume exit_func:
End Function

