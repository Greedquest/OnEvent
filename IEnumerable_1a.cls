VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEnumerable_1a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "CBuse - Array Iterators.Method1a"
Option Explicit

Public Function ItemByIndex(ByVal iIndex As Long _
                            , ByRef outVar As Variant) As Boolean
    Err.Raise 5, TypeName(Me), "This class should be used as an interface only"
End Function


