Attribute VB_Name = "modAutoFit"
Option Explicit

' Autofit cell height to contents
Sub AdjustCellHeight( _
    ByRef pRng As Range, _
    ByVal pMinH As Double, _
    ByVal pMaxH As Double _
)

Dim Ar As Range

    On Error GoTo AdjustCellHeight_Err
    Application.ScreenUpdating = False

    For Each Ar In pRng.Areas
        Ar.EntireRow.AutoFit
        With Ar.EntireRow
            If .RowHeight < pMinH Then .RowHeight = pMinH
            If .RowHeight > pMaxH Then .RowHeight = pMaxH
        End With
    Next

AdjustCellHeight_Exit:
    Application.ScreenUpdating = True
    Exit Sub

AdjustCellHeight_Err:
    Resume AdjustCellHeight_Exit
End Sub
