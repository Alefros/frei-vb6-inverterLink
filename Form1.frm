VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inverter link"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Limpar"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
            Text1 = Empty
            Text2 = Empty
            Text1.SetFocus
End Sub

Private Sub Text1_GotFocus()
            Text2.Enabled = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                Call Text1_LostFocus
            End If
End Sub

Private Sub Text1_LostFocus()
            Text2.Enabled = True
            Command1.SetFocus
            Text2.Text = StrReverse(Text1.Text)
End Sub

Private Sub Text2_Change()
            If Text2 = Empty Then
                Text1.SetFocus
            End If
End Sub
