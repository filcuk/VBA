Attribute VB_Name = "modUniqueUDF"
Option Explicit

'---/ UNIQUE \---------------------------------------------
' UDF to return unique list of values from range.
'----------------------------------------------------------
' 2019-10-08    Initial
'----------------------------------------------------------
Function UNIQUE_DEPR( _
        ByRef Target As Range, _
        Optional ByVal Filler As String = "NOT_SET", _
        Optional ByVal Ender As String, _
        Optional ByVal Order As XlSortOrder _
    )

Const SHORT_MSG As String = "<more values>"

Dim pref() As Variant
Dim cell As Range
Dim i As Long
Dim fnRows As Long
Dim tmp As Variant
Dim j As Long
    
    On Error GoTo err_UNIQUE
    
    ' Create list of unique values
    ReDim pref(0)
    
    For Each cell In Target
        ' Check if value exists in array
        For i = LBound(pref) To UBound(pref)
            If pref(i) = cell Then
                i = i + 1
                Exit For
            End If
        Next
        
        If pref(i - 1) <> cell And cell <> "" Then
            pref(UBound(pref)) = cell
            ReDim Preserve pref(UBound(pref) + 1)
        End If
    Next
    
    ' Sort list, if requested
    If Order <> 0 Then
        For i = LBound(pref) To UBound(pref) - 1
            tmp = pref(i)
            For j = LBound(pref) To UBound(pref)
                If (Order = xlAscending And pref(j) > tmp) _
                Or (Order = xlDescending And pref(j) < tmp) Then
                    pref(i) = pref(j)
                    pref(j) = tmp
                    tmp = pref(i)
                End If
            Next
        Next
    End If
    
    ' Add ending value, if defined
    If Ender = "" Then
        ReDim Preserve pref(UBound(pref) - 1)
    Else
        pref(UBound(pref)) = Ender
    End If
    
    ' Compare input length to output length
    fnRows = Range(Application.Caller.Address).Rows.CountLarge
    
    If fnRows < UBound(pref) + 1 Then
        pref(fnRows - 1) = "<more values>" 'SHORT_MSG
    Else
        ReDim Preserve pref(fnRows)
        
        For i = LBound(pref) To UBound(pref)
            If pref(i) = "" Then
                pref(i) = IIf(Filler = "NOT_SET", CVErr(xlErrNA), Filler)
            End If
        Next
    End If
    
    ' Output
    UNIQUE = Application.Transpose(pref)
    
    Exit Function
    
err_UNIQUE:
    
    UNIQUE = CVErr(xlErrNA)
    
End Function

' Add fn description to the 'Insert function' dialog
Private Sub RegisterUDF()
    Application.MacroOptions _
        Macro:="UNIQUE", _
        Description:="Returns list of unique values from range." & vbLf _
                        & "UNIQUE(<range>, <filler>, <ender>, <order>)", _
        category:="Custom", _
        ArgumentDescriptions:=Array( _
            "Selection of values.", _
            "String to fill into empty spaces.", _
            "String to insert after last entry.", _
            "Bubble sort of output list.")
End Sub

Private Sub UnregisterUDF()
    Application.MacroOptions _
        Macro:="UNIQUE", _
        Description:=Empty, _
        category:=Empty
End Sub
