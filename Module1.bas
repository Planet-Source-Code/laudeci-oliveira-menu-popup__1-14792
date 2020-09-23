Attribute VB_Name = "basAPI"
'Declaration section
Public Declare Function CreatePopupMenu Lib "User32.dll" () As Long
Public Declare Function DestroyMenu Lib "User32.dll" (ByVal hMenu As Long) As Long
Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_SUBMENU = &H4
Public Const MIIM_CHECKMARKS = &H8
Public Const MIIM_DATA = &H20
Public Const MIIM_TYPE = &H10

Public Declare Function InsertMenuItem Lib "User32.dll" Alias "InsertMenuItemA" _
        (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Public Declare Function TrackPopupMenu Lib "User32.dll" _
        (ByVal hMenu As Long, ByVal uFlags As Long, ByVal x As Long, ByVal y As Long, _
        ByVal nReserved As Long, ByVal hwnd As Long, ByVal prcRect As Long) As Long
Public Const TPM_RIGHTALIGN = &H8&
Public Const TPM_CENTERALIGN = &H4&
Public Const TPM_LEFTALIGN = &H0
Public Const TPM_TOPALIGN = &H0
Public Const TPM_NONOTIFY = &H80
Public Const TPM_RETURNCMD = &H100
Public Const TPM_LEFTBUTTON = &H0
Public Const TPM_RIGHTBUTTON = &H2&
Public Type POINT_TYPE
    x As Long
    y As Long
End Type
Public Declare Function GetCursorPos Lib "User32.dll" (lpPoint As POINT_TYPE) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

