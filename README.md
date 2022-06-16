# VBA

![GitHub commit activity](https://img.shields.io/github/commit-activity/m/filcuk/VBA?label=commits)
![GitHub last commit](https://img.shields.io/github/last-commit/filcuk/VBA)

Repository of useful modules.

Procedure header example:
```vba
'==========================================================
' Purpose   Creates Excel instance for saving files
'           to SharePoint the easy way.
' Input     pSource()   local file path array
'           pTarget()   remote file URL array
' Output    False if array lengths differ or local paths
'           aren't valid.
'           Does not verify whether upload was successful.
' Author    Filip Kraus, github.com/filcuk
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Log       2022-06-16  Initial
'----------------------------------------------------------
Function SaveToSharepoint( _
    ByRef pSource() As Variant, _
    ByRef pTarget() As Variant _
    ) As Boolean
...
End Function
```