VERSION 5.00
Begin VB.Form Frmabout 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frmabout.frx":0000
   ScaleHeight     =   3285
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   420
      Picture         =   "Frmabout.frx":3F4C4
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   945
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "eProperty - All rights reserved 2009-2010"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1905
      TabIndex        =   4
      Top             =   1455
      Width           =   3585
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   2220
      TabIndex        =   3
      Top             =   2985
      Width           =   2790
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   2220
      TabIndex        =   2
      Top             =   2700
      Width           =   3150
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This Product is Liencesed to :"
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   45
      TabIndex        =   1
      Top             =   2700
      Width           =   2385
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   330
      Left            =   1860
      Top             =   1410
      Width           =   3555
   End
End
Attribute VB_Name = "Frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Label2.Caption = RegisterName
    Label3.Caption = RegisterAddress
End Sub
