VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "With Columns"
      Height          =   285
      Index           =   1
      Left            =   375
      TabIndex        =   1
      Top             =   2790
      Width           =   1740
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Default"
      Height          =   285
      Index           =   0
      Left            =   390
      TabIndex        =   0
      Top             =   2460
      Value           =   -1  'True
      Width           =   1740
   End
   Begin VB.Label Label2 
      Caption         =   "laudeci@quality-al.com.br"
      Height          =   165
      Left            =   1500
      TabIndex        =   3
      Top             =   3465
      Width           =   3045
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Right Click  anywhere in this form to get your Dynamic Menu"
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   4245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mMenu As cVBPopUpMenu

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
If Button = vbRightButton Then
        Dim mResult As Long
        Set mMenu = New cVBPopUpMenu
        mMenu.CreateMenuPopUp
        With mMenu
            If Option1(0).Value Then
                .AddMenu "Teste1", MFT_STRING, MFS_ENABLED, 100
                .AddMenu "Teste2", MFT_STRING, MFS_ENABLED + MFS_DEFAULT, 200
                .AddMenu "dgjsdjg ", MFT_STRING, MFS_ENABLED, 350
                .AddMenu "Teste3", MFT_STRING, MFS_ENABLED, 400
                .CreateSubMenu "Sub Test1", MFT_STRING, MFS_ENABLED, 401
                .SubMenu.Add "Item 1", MFT_STRING, MFS_ENABLED + MFS_HILITE, 402
                .SubMenu.Add "/Separator/", MFT_SEPARATOR, MFS_ENABLED, 403
                .SubMenu.Add "Item 2", MFT_STRING, MFS_ENABLED + MFS_CHECKED, 404
                .CreateSubMenu "Sub Test2", MFT_STRING, MFS_ENABLED, 405
                .SubMenu.Add "Item 1", MFT_STRING, MFS_ENABLED + MFS_ENABLED, 406
                .SubMenu.Add "/Separator/", MFT_SEPARATOR, MFS_ENABLED, 407
                .SubMenu.Add "Item 2", MFT_STRING, MFS_ENABLED + MFS_DEFAULT, 408
            
            Else
                .AddMenu "Teste1", MFT_STRING, MFS_ENABLED, 100
                .AddMenu "Teste2", MFT_STRING, MFS_ENABLED + MFS_DEFAULT, 200
                .AddMenu "dgjsdjg ", MFT_STRING + MFT_MENUBARBREAK, MFS_ENABLED, 350
                .AddMenu "Teste3", MFT_STRING, MFS_ENABLED, 400
                .CreateSubMenu "Sub Test1", MFT_STRING + MFT_MENUBARBREAK, MFS_ENABLED, 401
                .SubMenu.Add "Item 1", MFT_STRING, MFS_ENABLED + MFS_HILITE, 402
                .SubMenu.Add "/Separator/", MFT_SEPARATOR, MFS_ENABLED, 403
                .SubMenu.Add "Item 2", MFT_STRING, MFS_ENABLED + MFS_CHECKED, 404
                .CreateSubMenu "Sub Test2", MFT_STRING, MFS_ENABLED, 405
                .SubMenu.Add "Item 1", MFT_STRING, MFS_ENABLED + MFS_ENABLED, 406
                .SubMenu.Add "/Separator/", MFT_SEPARATOR, MFS_ENABLED, 407
                .SubMenu.Add "Item 2", MFT_STRING, MFS_ENABLED + MFS_DEFAULT, 408
            End If
        End With

        mResult = mMenu.ShowMenu(Me.hWnd)
        If mResult <> 0 Then MsgBox "Menu id selected = " & mResult
        mMenu.MenuUnload
        Set mMenu = Nothing
    End If
End Sub

