VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMulti 
   BackColor       =   &H8000000D&
   Caption         =   "Web Explorer Search"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8910
   Icon            =   "frmMulti.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5040
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   255
      Left            =   7320
      TabIndex        =   1
      Top             =   150
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8310
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   14658
      _Version        =   393216
      Tabs            =   6
      Tab             =   2
      TabsPerRow      =   6
      TabHeight       =   420
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Yahoo"
      TabPicture(0)   =   "frmMulti.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "wb1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Google"
      TabPicture(1)   =   "frmMulti.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "wb2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "DMOZ"
      TabPicture(2)   =   "frmMulti.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "wb3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Excite"
      TabPicture(3)   =   "frmMulti.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "wb4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Hot Bot"
      TabPicture(4)   =   "frmMulti.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "wb5"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Web Crawler"
      TabPicture(5)   =   "frmMulti.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "wb6"
      Tab(5).ControlCount=   1
      Begin SHDocVwCtl.WebBrowser wb4 
         Height          =   7815
         Left            =   -75000
         TabIndex        =   3
         Top             =   360
         Width           =   8775
         ExtentX         =   15478
         ExtentY         =   13785
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
      Begin SHDocVwCtl.WebBrowser wb5 
         Height          =   7815
         Left            =   -75000
         TabIndex        =   4
         Top             =   360
         Width           =   8775
         ExtentX         =   15478
         ExtentY         =   13785
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
      Begin SHDocVwCtl.WebBrowser wb6 
         Height          =   7815
         Left            =   -75000
         TabIndex        =   5
         Top             =   600
         Width           =   8775
         ExtentX         =   15478
         ExtentY         =   13785
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
      Begin SHDocVwCtl.WebBrowser wb1 
         Height          =   7695
         Left            =   -75000
         TabIndex        =   6
         Top             =   360
         Width           =   8775
         ExtentX         =   15478
         ExtentY         =   13573
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
      Begin SHDocVwCtl.WebBrowser wb2 
         Height          =   7815
         Left            =   -75000
         TabIndex        =   7
         Top             =   360
         Width           =   8775
         ExtentX         =   15478
         ExtentY         =   13785
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
      Begin SHDocVwCtl.WebBrowser wb3 
         Height          =   7815
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   8775
         ExtentX         =   15478
         ExtentY         =   13785
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
   End
End
Attribute VB_Name = "frmMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   wb1.Navigate ("http://google.yahoo.com/bin/query?p=" & Text1.Text & "&hc=0&hs=0")

   wb2.Navigate ("http://www.google.com/search?q=" & Text1.Text)

   wb3.Navigate ("http://search.dmoz.org/cgi-bin/search?search=" & Text1.Text)

   wb4.Navigate ("http://search.excite.com/search.gw?c=web&search=" & Text1.Text)
   
   wb5.Navigate ("http://hotbot.lycos.com/?MT=" & Text1.Text & "&SQ=1&AM1=MC")
   
   wb6.Navigate ("http://dpxml.webcrawler.com/_1_2VE6UK7034NF7HR__info.wbcrwl/dog/results?otmpl=dog/webresults.htm&qkw=" & Text1.Text & "qcat=web&qk=20&top=1&start=&ver=16670")

End Sub

Private Sub Form_Resize()
SSTab1.Width = Me.Width
wb1.Width = SSTab1.Width - 80
wb2.Width = SSTab1.Width - 80
wb3.Width = SSTab1.Width - 80
wb4.Width = SSTab1.Width - 80
wb5.Width = SSTab1.Width - 80
wb6.Width = SSTab1.Width - 80
SSTab1.Height = frmMulti.ScaleHeight
SSTab1.Width = frmMulti.ScaleWidth
wb1.Height = SSTab1.Height
wb2.Height = SSTab1.Height
wb3.Height = SSTab1.Height
wb4.Height = SSTab1.Height
wb5.Height = SSTab1.Height
wb6.Height = SSTab1.Height
'Command1.Top = SSTab1.Height + 170
'Text1.Top = SSTab1.Height + 150
End Sub

