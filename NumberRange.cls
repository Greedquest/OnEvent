VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NumberRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Implementation.ValueProvider"
Implements IValueProvider

Private Type TRange
    isGenerator As Boolean
    currentCount As Long
    maxCount As Long
End Type

Private this As TRange

Private Function IValueProvider_GetNext() As Variant
    IValueProvider_GetNext = this.currentCount 'Chr$(65 + this.currentCount)
    this.currentCount = this.currentCount + 1
End Function

Private Function IValueProvider_HasMore() As Boolean
    IValueProvider_HasMore = this.isGenerator And (this.currentCount <= this.maxCount)
End Function

' === Public interface ============================================
Public Sub generatorTo(ByVal Count As Long)
    this.isGenerator = True
    this.currentCount = 0
    this.maxCount = Count - 1
End Sub

'@Ignore ProcedureNotUsed
'@Enumerator
Public Property Get NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = NewValueProviderIterator(Me)
End Property
