VERSION 5.00
Begin VB.Form frmSource 
   Caption         =   "Source Code"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   Icon            =   "frmSource.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6570
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   6615
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
Text1.Height = frmSource.ScaleHeight
Text1.Width = frmSource.ScaleWidth
End Sub

