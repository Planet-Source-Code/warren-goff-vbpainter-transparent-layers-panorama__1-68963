Attribute VB_Name = "ezzTwain"
Declare Function TWAIN_AcquireToFilename Lib "EZTW32.DLL" (ByVal hwndApp%, ByVal bmpFileName$) As Integer
Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hwndApp&) As Long
Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hwndApp As Long, ByVal wPixTypes As Long) As Long
Declare Function TWAIN_IsAvailable Lib "EZTW32.DLL" () As Long
Declare Function TWAIN_EasyVersion Lib "EZTW32.DLL" () As Long


