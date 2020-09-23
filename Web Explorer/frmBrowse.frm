VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBrowse 
   Caption         =   "Web Explorer"
   ClientHeight    =   7395
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   13035
   Icon            =   "frmBrowse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   13035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Appearance      =   0  'Flat
      Default         =   -1  'True
      Height          =   255
      Left            =   8880
      Picture         =   "frmBrowse.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Go"
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   6600
      TabIndex        =   11
      ToolTipText     =   "Search For..."
      Top             =   90
      Visible         =   0   'False
      Width           =   2295
   End
   Begin SHDocVwCtl.WebBrowser WB1 
      Height          =   6015
      Left            =   60
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   2895
      ExtentX         =   5106
      ExtentY         =   10610
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer tmrsearch 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   2280
      Top             =   6600
   End
   Begin VB.CommandButton cmdGo 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   6120
      Picture         =   "frmBrowse.frx":07FC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Go"
      Top             =   480
      Width           =   255
   End
   Begin VB.ComboBox cmbUrl 
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   450
      Width           =   5295
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6255
      Left            =   3360
      TabIndex        =   1
      Top             =   840
      Width           =   6735
      ExtentX         =   11880
      ExtentY         =   11033
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   240
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":0CEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":1008
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":1322
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":163C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":1956
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":1C70
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":1F8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":22A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":25BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":2A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":2D2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":3044
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":335E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":3678
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":3992
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":3CAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":3FC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":42E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   1200
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":4732
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":4A4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":4D66
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":5080
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":51DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":54F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":580E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":5B28
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":5E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":615C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   1320
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Tb1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   741
      ButtonWidth     =   714
      ButtonHeight    =   688
      Appearance      =   1
      Style           =   1
      ImageList       =   "tbrImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Description     =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
            Style           =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Description     =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   2
            Style           =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Description     =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Description     =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Home"
            Description     =   "Home"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            Description     =   "Search"
            Object.ToolTipText     =   "Search"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Favorites"
            Description     =   "Favorites"
            Object.ToolTipText     =   "Favorites"
            ImageIndex      =   7
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Add"
                  Text            =   "Add To Favorites"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "View"
                  Text            =   "View Favorites"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "IE"
                  Text            =   "View IE Favorites"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "History"
            Description     =   "History"
            Object.ToolTipText     =   "History"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Mail"
            Description     =   "Mail"
            Object.ToolTipText     =   "Mail"
            ImageIndex      =   9
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MC Mail"
                  Text            =   "Check Mail"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MS Mail"
                  Text            =   "Send Mail"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Media"
            Description     =   "Media"
            Object.ToolTipText     =   "Media"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Description     =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   11
         EndProperty
      EndProperty
      MousePointer    =   1
   End
   Begin MSComctlLib.ImageList tbrImageList 
      Left            =   360
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":6476
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":69D2
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":6F2E
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":748A
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":79E6
            Key             =   "Home"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":7F42
            Key             =   "History"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":840A
            Key             =   "Favorites"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":885A
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":8DB6
            Key             =   "Mail"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":92BA
            Key             =   "Media"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowse.frx":93E2
            Key             =   "Print"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMedia 
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
      ButtonWidth     =   1773
      ButtonHeight    =   582
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Media..."
            Key             =   "tbrMedia"
            Description     =   "Media"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "X"
            Key             =   "Exit"
            Object.ToolTipText     =   "Close Window"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrHistory 
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
      ButtonWidth     =   1931
      ButtonHeight    =   582
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "History..."
            Key             =   "tbrHistory"
            Description     =   "History"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "X"
            Key             =   "Exit"
            Object.ToolTipText     =   "Close Window"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrFavorites 
      Height          =   330
      Left            =   0
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
      ButtonWidth     =   2223
      ButtonHeight    =   582
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Favorites..."
            Key             =   "tbrFavorites"
            Description     =   "Favorites"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "X"
            Key             =   "Exit"
            Object.ToolTipText     =   "Close Window"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPopup 
      Caption         =   "Blocking Pop-Ups"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   11520
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   0
      Top             =   480
      Width           =   4935
   End
   Begin VB.Menu nmuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New Window"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuFilePagesetup 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Preview"
      End
      Begin VB.Menu sepa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileEditMode 
         Caption         =   "Edit Mode ON"
      End
      Begin VB.Menu mnuFileEditModeOFF 
         Caption         =   "Edit Mode OFF"
      End
      Begin VB.Menu sepaa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuWorkOffline 
         Caption         =   "Work Offline"
      End
      Begin VB.Menu sepaaa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu seppa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu seppaa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find (on this page)..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuGoTo 
         Caption         =   "Go To"
         Begin VB.Menu mnuBack 
            Caption         =   "Back"
         End
         Begin VB.Menu mnuForward 
            Caption         =   "Forward"
         End
         Begin VB.Menu sepaaaaa 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHome 
            Caption         =   "Home Page"
         End
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu sepp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "&Zoom"
         Begin VB.Menu mnuZoom200 
            Caption         =   "Large"
         End
         Begin VB.Menu mnuZoom300 
            Caption         =   "Larger"
         End
         Begin VB.Menu mnuZoomNormal 
            Caption         =   "Normal"
         End
         Begin VB.Menu mnuZoomSmall 
            Caption         =   "Small"
         End
         Begin VB.Menu mnuZoomSmaller 
            Caption         =   "Smaller"
         End
      End
      Begin VB.Menu mnuEditViewSource 
         Caption         =   "View Source"
      End
   End
   Begin VB.Menu mnuFavorites 
      Caption         =   "F&avorites"
      Begin VB.Menu mnuFavoritesAdd 
         Caption         =   "Add to Favorites"
      End
      Begin VB.Menu mnuFavoritesShow 
         Caption         =   "Show Favorites"
      End
      Begin VB.Menu mnuIEFav 
         Caption         =   "IE Favorites"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Tools"
      Begin VB.Menu mnuMailNews 
         Caption         =   "Mail and News"
      End
      Begin VB.Menu mnuAddress 
         Caption         =   "Address Book"
      End
      Begin VB.Menu mnuWinupdate 
         Caption         =   "Windows Update"
      End
      Begin VB.Menu seppp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsHome 
         Caption         =   "Set Home Page"
      End
      Begin VB.Menu mnuOptionsAllow 
         Caption         =   "Allow Pop-Up windows"
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "History"
      End
      Begin VB.Menu mnuLookup 
         Caption         =   "Domain LookUp"
      End
      Begin VB.Menu sepaaaa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInetOpts 
         Caption         =   "Internet Options"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About Power Browser"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'AllowPopup coded by Dustin Davis
'View Source coded by Ali Hussain
'if I have left any credits out - please let me know

Public Hist As Boolean
Dim a
Public AllowPopup As Boolean 'This is for Pop-up windows
Option Explicit

Public homepage As String
Dim TheUrl As String

Private Sub cmbUrl_KeyPress(KeyAscii As Integer)
'******************************************************
    On Error Resume Next
    
    'Navigate to the Url typed when the return key _
    is pressed
    If KeyAscii = vbKeyReturn Then
        WebBrowser1.Navigate (cmbUrl.Text)
        cmbUrl.AddItem (cmbUrl.Text)
    End If
'******************************************************
End Sub



Private Sub cmdGo_Click()
'Go to web page

        WebBrowser1.Navigate (cmbUrl.Text)
        cmbUrl.AddItem (cmbUrl.Text)
End Sub



Private Sub cmdSearch_Click()
   frmMulti.wb1.Navigate ("http://google.yahoo.com/bin/query?p=" & txtSearch.Text & "&hc=0&hs=0")

   frmMulti.wb2.Navigate ("http://www.google.com/search?q=" & txtSearch.Text)

   frmMulti.wb3.Navigate ("http://search.dmoz.org/cgi-bin/search?search=" & txtSearch.Text)

   frmMulti.wb4.Navigate ("http://search.excite.com/search.gw?c=web&search=" & txtSearch.Text)
   
   frmMulti.wb5.Navigate ("http://hotbot.lycos.com/?MT=" & txtSearch.Text & "&SQ=1&AM1=MC")
   
   frmMulti.wb6.Navigate ("http://dpxml.webcrawler.com/_1_2VE6UK7034NF7HR__info.wbcrwl/dog/results?otmpl=dog/webresults.htm&qkw=" & txtSearch.Text & "recipes&qcat=web&qk=20&top=1&start=&ver=16670")
frmMulti.Show
txtSearch.Visible = False
cmdSearch.Visible = False
Tb1.Buttons(8).Value = tbrUnpressed
End Sub

Private Sub Form_Load()
homepage = GetSetting(App.Path, "HP", "HP")
WebBrowser1.Navigate homepage
'Pre-Load the Media window
wb1.Navigate "http://www.windowsmedia.com/mg/Radio.asp?rf=1#radTop"
           
'Resize and place objects
With WebBrowser1
    .Width = frmBrowse.Width - 200
    .Left = 50
    .Height = frmBrowse.Height - 1670
End With

wb1.Height = frmBrowse.Height - 2170
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu mnuEdit
End If
End Sub

Private Sub Form_Resize()
'Resizes everything to fit to the form
On Error Resume Next

With WebBrowser1
    .Width = frmBrowse.Width - 200
    .Left = 50
    .Height = frmBrowse.Height - 1670
End With

wb1.Height = frmBrowse.Height - 2170

End Sub

Private Sub mnuAddress_Click()
On Error Resume Next
Dim x
x = Shell("C:\Program Files\Outlook Express\wab.exe", 1)

End Sub

Private Sub mnuBack_Click()
On Error Resume Next
Hist = False
 For a = 1 To Tb1.Buttons("Forward").ButtonMenus.Count
            If Tb1.Buttons("Forward").ButtonMenus.Item(a).Text = WebBrowser1.LocationURL Then
                Hist = True
            End If
        Next a
        If Hist = False Then Tb1.Buttons("Forward").ButtonMenus.add Text:=WebBrowser1.LocationURL
        WebBrowser1.GoBack
End Sub

Private Sub mnuCut_Click()
    WebBrowser1.SetFocus
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_CUT, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub mnuEditCopy_Click()
'******************************************************
    On Error Resume Next
    'Copy selected text/picture etc...
    WebBrowser1.SetFocus
    WebBrowser1.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
'******************************************************
End Sub


Private Sub mnuEditSelAll_Click()
'*******************************************************
    On Error Resume Next
    'Select all webpage
    WebBrowser1.SetFocus
    WebBrowser1.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
'*******************************************************
End Sub

Private Sub mnuEditViewSource_Click()
Dim cap As String
On Error GoTo esource
frmSource.Text1.Text = WebBrowser1.Document.documentElement.innerHTML
frmSource.Caption = cap
frmSource.Show
esource:
End Sub

Private Sub mnuexit_Click()
'Exit program
Unload Me
End Sub

Private Sub mnuFavoritesAdd_Click()
    'Add current webpage to favorites list
    Dim add
    add = cmbUrl.Text
        If add = "" Then
        add = InputBox("Enter website you wish to add to favorites", "Add", "www.")
            If add = "" Then
                Exit Sub
            End If
        Else
                frmFav.List1.AddItem (add)
            
        End If
End Sub

Private Sub mnuFavoritesShow_Click()
frmFav.Show
End Sub

Private Sub mnuFileEditMode_Click()
WebBrowser1.Document.designMode = "On"
End Sub

Private Sub mnuFileEditModeOFF_Click()
WebBrowser1.Document.designMode = "Off"
End Sub

Private Sub mnuFileNew_Click()
    'Display a new window
    Dim f As Form
        Set f = New frmBrowse
        f.Show
End Sub

Private Sub mnuFilePagesetup_Click()
'*******************************************************
    On Error Resume Next
    'call page setup function
    WebBrowser1.SetFocus
    WebBrowser1.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT
'*******************************************************
End Sub

Private Sub mnuFilePrint_Click()
'*******************************************************
    On Error Resume Next
    'Print current page
    WebBrowser1.SetFocus
    WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
'*******************************************************
End Sub

Private Sub mnuFilePrintPreview_Click()
'*******************************************************
    On Error Resume Next
    'Call the print preview function
    WebBrowser1.SetFocus
    WebBrowser1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT
'*******************************************************
End Sub

Private Sub mnuFileSaveAs_Click()

'*******************************************************
    On Error Resume Next
    'Save current page as
    WebBrowser1.SetFocus
    WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
'*******************************************************
End Sub

Private Sub mnuFind_Click()
WebBrowser1.SetFocus
    SendKeys "^f"
End Sub

Private Sub mnuForward_Click()
On Error Resume Next
        For a = 1 To Tb1.Buttons("Back").ButtonMenus.Count
        If Tb1.Buttons("Back").ButtonMenus.Item(a).Text = WebBrowser1.LocationURL Then
            Hist = True
        End If
    Next a
    If Hist = False Then Tb1.Buttons("Back").ButtonMenus.add Text:=WebBrowser1.LocationURL
       WebBrowser1.GoForward
End Sub

Private Sub mnuHistory_Click()
Dim nFolder As SpecialShellFolderIDs
  Dim pidl As Long
  Dim cbpidl As Integer
  Dim abpidl() As Byte
  Dim avpidl As Variant
  Dim sPath As Long
  nFolder = CSIDL_HISTORY
  If SHGetSpecialFolderLocation(hWnd, nFolder, pidl) = NOERROR Then
    If pidl Then
       cbpidl = GetPIDLSize(pidl)
      If cbpidl Then
        ReDim abpidl(cbpidl - 1)
        MoveMemory abpidl(0), ByVal pidl, cbpidl
         avpidl = abpidl
        wb1.Navigate2 avpidl
        wb1.Visible = True
      End If
      Call CoTaskMemFree(pidl)
    End If
  End If
End Sub

Private Sub mnuHome_Click()
WebBrowser1.GoHome
End Sub

Private Sub mnuIEFav_Click()
Dim nFolder As SpecialShellFolderIDs
  Dim pidl As Long
  Dim cbpidl As Integer
  Dim abpidl() As Byte
  Dim avpidl As Variant
  Dim sPath As Long
  nFolder = CSIDL_FAVORITES
  If SHGetSpecialFolderLocation(hWnd, nFolder, pidl) = NOERROR Then
    If pidl Then
       cbpidl = GetPIDLSize(pidl)
      If cbpidl Then
        ReDim abpidl(cbpidl - 1)
        MoveMemory abpidl(0), ByVal pidl, cbpidl
         avpidl = abpidl
        wb1.Navigate2 avpidl
        wb1.Visible = True
      End If
      Call CoTaskMemFree(pidl)
    End If
  End If
End Sub

Private Sub mnuInetOpts_Click()
Dim RetVal
    RetVal = Shell("rundll32.exe shell32.dll,Control_RunDLL Inetcpl.cpl", vbNormalFocus)
End Sub

Private Sub mnuLookup_Click()
frmLookup.Show
End Sub

Private Sub mnuMailNews_Click()
    Shell "C:\Program Files\Outlook Express\MSIMN.EXE", vbNormalFocus
End Sub

Private Sub mnuOptionsAllow_Click()
'Turn on/off pop-up windows
If AllowPopup = True Then
    AllowPopup = False
    mnuOptionsAllow.Checked = False
    lblPopup.Caption = "Blocking Pop-Ups"
    lblPopup.ForeColor = vbBlue
ElseIf AllowPopup = False Then
    AllowPopup = True
    mnuOptionsAllow.Checked = True
    lblPopup.Caption = "Allowing Pop-Ups"
    lblPopup.ForeColor = vbRed
End If
End Sub

Private Sub mnuOptionsHome_Click()
homepage = InputBox("Please enter the website of the homepage.")
SaveSetting App.Path, "HP", "HP", homepage
End Sub



Private Sub mnuPaste_Click()
    WebBrowser1.SetFocus
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuProperties_Click()
    WebBrowser1.SetFocus
    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuRefresh_Click()
WebBrowser1.Refresh
End Sub

Private Sub mnuStop_Click()
WebBrowser1.Stop
End Sub

Private Sub mnuWinupdate_Click()
WebBrowser1.Navigate "http://windowsupdate.microsoft.com/"
End Sub

Private Sub mnuWorkOffline_Click()
mnuWorkOffline.Checked = Not mnuWorkOffline.Checked
If mnuWorkOffline.Checked = True Then
   mnuWorkOffline.Checked = True
    WebBrowser1.Offline = True
        
    ElseIf mnuWorkOffline.Checked = False Then
        
        mnuWorkOffline.Checked = False
        WebBrowser1.Offline = False
        
    
    End If
End Sub

Private Sub mnuZoom200_Click()
WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(3), vbNull

End Sub

Private Sub mnuZoom300_Click()
WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(4), vbNull

End Sub

Private Sub mnuZoomNormal_Click()
WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2), vbNull

End Sub

Private Sub mnuZoomSmall_Click()
WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(1), vbNull

End Sub

Private Sub mnuZoomSmaller_Click()
WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(0), vbNull

End Sub





Private Sub tbrFavorites_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

Select Case Button.Key
    Case "Exit"
            tbrFavorites.Visible = False
            wb1.Visible = False
WebBrowser1.Width = frmBrowse.Width - 200
WebBrowser1.Left = 50
        
End Select
End Sub

Private Sub tbrHistory_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

Select Case Button.Key
    Case "Exit"
         Tb1.Buttons(10).Value = tbrUnpressed
            tbrHistory.Visible = False
            wb1.Visible = False
WebBrowser1.Width = frmBrowse.Width - 200
WebBrowser1.Left = 50
        
End Select
End Sub

Private Sub tbrMedia_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

Select Case Button.Key
    Case "Exit"
         Tb1.Buttons(14).Value = tbrUnpressed
            tbrMedia.Visible = False
            wb1.Visible = False
WebBrowser1.Width = frmBrowse.Width - 200
WebBrowser1.Left = 50
        
End Select
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'shows done in the status bar
lblStatus.Caption = "Done"
cmbUrl.Text = WebBrowser1.LocationURL
frmBrowse.Caption = WebBrowser1.LocationName
End Sub

Private Sub WebBrowser1_DownloadBegin()
'Starting download
lblStatus.Caption = "Starting Download"
End Sub

Private Sub WebBrowser1_DownloadComplete()
'Done downloading
lblStatus.Caption = "Download Done!"
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'Loaded page
lblStatus.Caption = "Done Loading!"
frmBrowse.Caption = WebBrowser1.LocationName  'Shows webpage in title bar
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)

'Set ppDisp = New frmBrowse
'ppDisp.Show
'ppDisp.WebBrowser1.Navigate TheUrl
Dim frm As frmBrowse
Set frm = New frmBrowse
Set ppDisp = frm.WebBrowser1.object
frm.Show
'This will allow a pop-up window to load or to be blocked!
'If AllowPopup = True Then
'    Cancel = False
'    DoEvents
'ElseIf AllowPopup = False Then
'    Cancel = True
'End If
End Sub

Private Sub WebBrowser1_OnToolBar(ByVal toolbar As Boolean)
    On Error Resume Next


    If toolbar = False And mnuOptionsAllow.Checked = False Then
        Unload Me
    End If
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
'Shows progress in status bar
lblStatus.Caption = "Reading " & Progress & "  of  " & ProgressMax
End Sub

Private Sub webBrowser1_StatusTextChange(ByVal Text As String)
'shows new text in status bar
lblStatus.Caption = Text
End Sub



Private Sub WebBrowser1_TitleChange(ByVal Text As String)
cmbUrl.Text = WebBrowser1.LocationURL
Me.Caption = Text
End Sub

Private Sub Tb1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim starting
On Error Resume Next

Select Case Button.Key
    Case "Back"
        mnuBack_Click
    Case "Forward"
        mnuForward_Click
    Case "Stop"
        WebBrowser1.Stop
    Case "Refresh"
        WebBrowser1.Refresh
    Case "Home"
        WebBrowser1.GoHome
    Case "Search"
    
        'frmMulti.Show
        If txtSearch.Visible = True Then
        Tb1.Buttons(8).Value = tbrUnpressed
       txtSearch.Visible = False
       cmdSearch.Visible = False
       Else
               Tb1.Buttons(8).Value = tbrPressed
       txtSearch.Visible = True
       cmdSearch.Visible = True
       End If
    Case "History"
        If tbrHistory.Visible = True Then
        Tb1.Buttons(10).Value = tbrUnpressed
            tbrHistory.Visible = False
            wb1.Visible = False
WebBrowser1.Width = frmBrowse.Width - 200
WebBrowser1.Left = 50
        Else
        Tb1.Buttons(10).Value = tbrPressed
            tbrHistory.Visible = True
            wb1.Visible = True
WebBrowser1.Width = frmBrowse.Width - tbrMedia.Width - 200
WebBrowser1.Left = 3050
    mnuHistory_Click
        End If
     
    Case "Media"
        If tbrMedia.Visible = True Then
        Tb1.Buttons(14).Value = tbrUnpressed
            tbrMedia.Visible = False
            wb1.Visible = False
WebBrowser1.Width = frmBrowse.Width - 200
WebBrowser1.Left = 50
        Else
        Tb1.Buttons(14).Value = tbrPressed
            tbrMedia.Visible = True
    wb1.Navigate "http://www.windowsmedia.com/mg/Radio.asp?rf=1#radTop"
 
            wb1.Visible = True
WebBrowser1.Width = frmBrowse.Width - tbrMedia.Width - 200
WebBrowser1.Left = 3050
        End If
    
    Case "Print"
        WebBrowser1.SetFocus
        On Error Resume Next
        WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
        
End Select

End Sub

Private Sub Tb1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

On Error Resume Next
Select Case ButtonMenu.Key

'Check Mail
Case "MC Mail"
    Shell "C:\Program Files\Outlook Express\MSIMN.EXE", vbNormalFocus
   
'Send Mail
Case "MS Mail"
    Dim subject, person
        person = InputBox("Enter email address", "email")
        subject = InputBox("Enter subject for email", "subject")
        WebBrowser1.Navigate ("mailto:" & person & "?subject=" & subject)
    
'Favorites
Case "Add"
    mnuFavoritesAdd_Click
Case "View"
    frmFav.Show
Case "IE"
    'mnuIEFav_Click
        If tbrFavorites.Visible = True Then
                'Tb1.Buttons(10).Value = tbrUnpressed
            tbrFavorites.Visible = False
            wb1.Visible = False
WebBrowser1.Width = frmBrowse.Width - 200
WebBrowser1.Left = 50
        Else
                'Tb1.Buttons(10).Value = tbrPressed
            tbrFavorites.Visible = True
            wb1.Visible = True
WebBrowser1.Width = frmBrowse.Width - tbrMedia.Width - 200
WebBrowser1.Left = 3050
    mnuIEFav_Click
        End If

End Select

End Sub


