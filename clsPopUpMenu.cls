VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVBPopUpMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsSubMenu"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarclsSubMenu As cVBSubMenu
Private hPopupMenu As Long 'handle to the popup menu to display
Private hPopupSubMenu As Long 'handle to the popup menu to display
Private MenuInfo As MENUITEMINFO   'describes menu items to add
Private SubMenuInfo As MENUITEMINFO   'describes menu items to add
Private MenuCount As Long
Private SubMenuCreated As Boolean
Private SubMenuCount As Long

Dim menusel As Long       'ID of what the user selected in the popup menu
Dim retval As Long        'generic return value
Dim curpos As POINT_TYPE

Public Enum MenuTypeEnum
    MFT_BITMAP = &H4
    MFT_MENUBARBREAK = &H20
    MFT_MENUBREAK = &H40
    MFT_OWNERDRAW = &H100
    MFT_RADIOCHECK = &H200
    MFT_RIGHTJUSTIFY = &H4000
    MFT_RIGHTORDER = &H2000
    MFT_SEPARATOR = &H800
    MFT_STRING = &H0
End Enum
Public Enum MenuStateEnum
    MFS_CHECKED = &H8
    MFS_DEFAULT = &H1000
    MFS_DISABLED = &H2
    MFS_ENABLED = &H0
    MFS_GRAYED = &H1
    MFS_HILITE = &H80
    MFS_UNCHECKED = &H0
    MFS_UNHILITE = &H0
End Enum

Public Sub CreateMenuPopUp()
    hPopupMenu = CreatePopupMenu()
    With MenuInfo
        .cbSize = Len(MenuInfo) 'The size of this structure.
        .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU  'Which elements of the structure to use.
    End With
End Sub

Public Sub AddMenu(ByVal MenuCaption As String, MenuType As MenuTypeEnum, MenuState As MenuStateEnum, MenuID As Long)
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

Public Sub CreateSubMenu(ByVal MenuCaption As String, MenuType As MenuTypeEnum, MenuState As MenuStateEnum, MenuID As Long)
    Set mvarclsSubMenu = New cVBSubMenu
    With MenuInfo '(SubMenu/teste)
        .fType = MenuType
        .fState = MenuState
        .wID = MenuID 'Assign this item an item identifier.
        .dwTypeData = MenuCaption
        .cch = Len(MenuCaption)
        .hSubMenu = mvarclsSubMenu.SubMenuPopUp()
    End With
    retval = InsertMenuItem(hPopupMenu, MenuCount, 1, MenuInfo)
    MenuCount = MenuCount + 1
End Sub

Public Property Get SubMenu() As cVBSubMenu
    Set SubMenu = mvarclsSubMenu
End Property

Public Function ShowMenu(mHWND As Long) As Long
    retval = GetCursorPos(curpos)
    ShowMenu = TrackPopupMenu(hPopupMenu, _
            TPM_TOPALIGN Or TPM_NONOTIFY _
            Or TPM_RETURNCMD Or TPM_LEFTALIGN _
            Or TPM_LEFTBUTTON Or TPM_RIGHTBUTTON, _
            curpos.x, curpos.y, 0, mHWND, 0)

End Function

Public Sub MenuUnload()
    retval = DestroyMenu(hPopupMenu)
End Sub



