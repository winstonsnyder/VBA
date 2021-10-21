Attribute VB_Name = "M_Essbase"
'Declare Function EssVCalculate Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ByVal calcScript As Variant, ByVal synchronous As Variant) As Long
'Declare Function EssVCancelCalc Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant) As Long
'Declare Function EssVCascade Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ByVal range As Variant, ByVal selection As Variant, ByVal path As Variant, ByVal prefix As Variant, ByVal suffix As Variant, ByVal level As Variant, ByVal openFile As Variant, ByVal copyFormats As Variant, ByVal overwrite As Variant, ByVal listFile As Variant) As Long
'Declare Function EssVCell Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ParamArray memberlist() As Variant) As Variant
'Declare Function EssVConnect Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ByVal username As Variant, ByVal password As Variant, ByVal server As Variant, ByVal application As Variant, ByVal database As Variant) As Long
'Declare Function EssVDisconnect Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant) As Long
'Declare Function EssVFlashBack Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant) As Long
'Declare Function EssVGetCurrency Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant) As Variant
'Declare Function EssVGetDataPoint Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ByVal cell As Variant, ByVal range As Variant, ByVal aliases As Variant) As Variant
'Declare Function EssVFreeDataPoint Lib "ESSEXCLN.XLL" (ByVal Info As Variant) As Long
'Declare Function EssVGetGlobalOption Lib "ESSEXCLN.XLL" (ByVal item As Long) As Variant
'Declare Function EssVGetHctxFromSheet Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant) As Long
'Declare Function EsbExport Lib "ESBAPIN.DLL" (ByVal hCtx As Long, ByVal AppName As String, ByVal DbName As String, ByVal FilePath As String, ByVal level As Integer, ByVal isColumns As Integer) As Long
'Declare Function EssVGetMemberInfo Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ByVal mbrName As Variant, ByVal action As Variant, ByVal aliases As Variant) As Variant
'Declare Function EssVFreeMemberInfo Lib "ESSEXCLN.XLL" (ByRef memInfo As Variant) As Long
'Declare Function EssVGetSheetOption Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ByVal item As Variant) As Variant
'Declare Function EssVGetStyle Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ByVal styleType As Variant, ByVal dimName As Variant, ByVal item As Long) As Variant
'Declare Function EssVKeepOnly Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ByVal range As Variant, ByVal selection As Variant) As Long
'Declare Function EssVLoginSetPassword Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ByVal newPassword As Variant, ByVal oldPassword As Variant, ByVal server As Variant, ByVal username As Variant) As Long
'Declare Function EssVPivot Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ByVal range As Variant, ByVal startPoint As Variant, ByVal endPoint As Variant) As Long
'Declare Function EssVRemoveOnly Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ByVal range As Variant, ByVal selection As Variant) As Long
'Declare Function EssVRetrieve Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ByVal range As Variant, ByVal lockFlag As Variant) As Long
'Declare Function EssVSendData Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ByVal range As Variant) As Long
'Declare Function EssVUnlock Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant) As Long
'Declare Function EssVSetCurrency Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ByVal currencyIdentifier As Variant) As Long
'Declare Function EssVSetGlobalOption Lib "ESSEXCLN.XLL" (ByVal item As Long, ByVal globalOption As Variant) As Long
'Declare Function EssVSetMenu Lib "ESSEXCLN.XLL" (ByVal setMenu As Boolean) As Long
'Declare Function EssVSetSheetOption Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ByVal item As Variant, ByVal sheetOption As Variant) As Long
'Declare Function EssVSetStyle Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ByVal styleType As Variant, ByVal dimName As Variant, ByVal item As Long, ByVal newValue As Variant) As Long
'Declare Function EssVZoomIn Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ByVal range As Variant, ByVal selection As Variant, ByVal level As Variant, ByVal across As Variant) As Long
'Declare Function EssVZoomOut Lib "ESSEXCLN.XLL" (ByVal sheetName As Variant, ByVal range As Variant, ByVal selection As Variant) As Long
'
