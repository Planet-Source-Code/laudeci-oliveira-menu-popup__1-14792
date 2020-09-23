VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVBSubMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Private hPopupMenu As Long 'handle to the popup menu to display
Private MenuInfo As MENUITEMINFO   'describes menu items to add
Private MenuCount As Long
Dim retval As Long        'generic return value


Friend Function SubMenuPopUp() As Long
    hPopupMenu = CreatePopupMenu()
    With MenuInfo
        .cbSize = Len(MenuInfo) 'The size of this structure.
        .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU 'Which elements of the structure to use.
    End With
    SubMenuPopUp = hPopupMenu
End Function

Public Sub Add(ByVal MenuCaption As String, MenuType As MenuTypeEnum, MenuState As MenuStateEnum, MenuID As Long)
    Dim mPic As StdPicture
    With MenuInfo '(SubMenu/teste)
        .fType = MenuType
        .fState = MenuState
        .wID = MenuID 'Assign this item an item identifier.
        .dwTypeData = MenuCaption
        .cch = Len(MenuCaption)
        .hSubMenu = 0
        
    End With
    retval = InsertMenuItem(hPopupMenu, MenuCount, 1, MenuInfo)
    MenuCount = MenuCount + 1
End Sub

