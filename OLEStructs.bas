Attribute VB_Name = "OLEStructs"
'@IgnoreModule IntegerDataType
'@Folder "Libs.COMTools.OLE"
Option Explicit

Public Type GUIDt
    Data1 As Long
    '@Ignore IntegerDataType
    Data2 As Integer
    '@Ignore IntegerDataType
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Type EXCEPINFOt
    wCode As Integer
    wReserved As Integer
    bstrSource As String
    bstrDescription As String
    bstrHelpFile As String
    dwHelpContext As Long
    pvReserved As LongPtr
    pfnDeferredFillIn As LongPtr
    scode As Long
End Type

Public Type DISPPARAMSt
    rgvarg As LongPtr                            '  VARIANTARG *rgvarg;
    rgdispidNamedArgs As LongPtr                 '  DISPID     *rgdispidNamedArgs;
    cArgs As Long                                '  UINT       cArgs;
    cNamedArgs As Long                           '  UINT       cNamedArgs;
End Type

Public Enum DISPID
    DISPID_COLLECT = -8                          'The Collect property. You use this property if the method you are calling through Invoke is an accessor function.
    DISPID_CONSTRUCTOR = -6                      'The C++ constructor function for the object.
    DISPID_DESTRUCTOR = -7                       'The C++ destructor function for the object.
    DISPID_EVALUATE = -5                         'The Evaluate method. This method is implicitly invoked when the ActiveX client encloses the arguments in square brackets. For example, the following two lines are equivalent: x.[A1:C1].value = 10: x.Evaluate("A1:C1").value = 10
    DISPID_NEWENUM = -4                          'The _NewEnum property. This special, restricted property is required for collection objects. It returns an enumerator object that supports IEnumVARIANT, and should have the restricted attribute specified.
    DISPID_PROPERTYPUT = -3                      'The parameter that receives the value of an assignment in a PROPERTYPUT.
    DISPID_UNKNOWN = -1                          'The value returned by IDispatch::GetIDsOfNames to indicate that a member or parameter name was not found.
    DISPID_VALUE = 0                             'The default member for the object. This property or method is invoked when an ActiveX client specifies the object name without a property or method.
End Enum


Public Enum API
    E_NOTIMPL = &H80004001
    E_NOINTERFACE = &H80004002
    E_POINTER = &H80004003
End Enum

