VERSION 5.00
Begin VB.Form frmLookup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Domian Look Up "
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5040
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Look&Up"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Domian Lookup:"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter The Domian Name You Want To Look Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GlobalSearch As String


Private Sub Command1_Click()
GlobalSearch = "http://www.internic.net/cgi-bin/whois?whois_nic=" & Text1.Text & "&type=domain"
 If Text1.Text <> "" Then
 frmBrowse.WebBrowser1.Navigate GlobalSearch
 Unload Me
 End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
End Sub

