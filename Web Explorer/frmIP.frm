VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmIP 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   1508
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
      Location        =   ""
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 


Private Sub Command1_Click()
    WebBrowser1.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER
    WebBrowser1.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DONTPROMPTUSER
    Text1.Text = Clipboard.GetText
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate "http://pchelplive.com/ip.php$ip_address = $REMOTE_ADDR; echo $ip_address;"

End Sub
