VERSION 5.00
Begin VB.Form FrmAgentPhone 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Outgoing Phone ( Agent Phone )"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   Icon            =   "FrmAgentPhone.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4245
      TabIndex        =   14
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Canecl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5820
      TabIndex        =   13
      Top             =   3120
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Rseidential Client"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1605
      TabIndex        =   4
      Top             =   120
      Width           =   2010
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Commercial Client"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1605
      TabIndex        =   3
      Top             =   420
      Width           =   2010
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Vendor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4605
      TabIndex        =   2
      Top             =   120
      Width           =   2010
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Comercial Vender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4455
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.OptionButton Option9 
      Caption         =   "Others"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4620
      TabIndex        =   0
      Top             =   435
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1785
      Left            =   90
      TabIndex        =   15
      Top             =   1125
      Width           =   7275
      Begin VB.CommandButton Command3 
         Caption         =   "&Find Refrence"
         Height          =   345
         Left            =   5250
         TabIndex        =   20
         Top             =   135
         Width           =   1500
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1740
         TabIndex        =   19
         Top             =   150
         Width           =   3060
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   1740
         TabIndex        =   18
         Top             =   525
         Width           =   5385
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   1740
         MaxLength       =   255
         TabIndex        =   17
         Top             =   900
         Width           =   5385
      End
      Begin VB.TextBox Text6 
         Height          =   465
         Left            =   1740
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   1275
         Width           =   5385
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Reference No. :"
         Height          =   195
         Left            =   255
         TabIndex        =   24
         Top             =   150
         Width           =   1140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Name :"
         Height          =   195
         Left            =   255
         TabIndex        =   23
         Top             =   510
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Telephone No. :"
         Height          =   195
         Left            =   255
         TabIndex        =   22
         Top             =   870
         Width           =   1155
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Note :"
         Height          =   195
         Left            =   255
         TabIndex        =   21
         Top             =   1230
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1785
      Left            =   90
      TabIndex        =   6
      Top             =   1125
      Width           =   7275
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   1710
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1170
         Width           =   5160
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   1710
         MaxLength       =   255
         TabIndex        =   10
         Top             =   705
         Width           =   5160
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   1710
         TabIndex        =   8
         Top             =   255
         Width           =   5160
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Note :"
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   1185
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Telephone No :"
         Height          =   195
         Left            =   270
         TabIndex        =   9
         Top             =   727
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name :"
         Height          =   195
         Left            =   270
         TabIndex        =   7
         Top             =   270
         Width           =   510
      End
   End
   Begin VB.Line Line4 
      X1              =   90
      X2              =   7350
      Y1              =   2955
      Y2              =   2955
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   90
      X2              =   7320
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Line Line2 
      X1              =   90
      X2              =   7350
      Y1              =   975
      Y2              =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   90
      X2              =   7320
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Phone To :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   5
      Top             =   135
      Width           =   960
   End
End
Attribute VB_Name = "FrmAgentPhone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim RsP As New ADODB.Recordset
Dim RsC As New ADODB.Recordset
Dim RsV As New ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim LastNo As Integer
Dim trec As Integer
Dim i As Integer

Private Sub Combo1_Click()
'display record
        If Rs1.State = 1 Then Rs1.Close
        If Option1.Value = True Or Option2.Value = True Then
            Rs1.Open "select * from clientmaster where clientrefno =" & Val(Combo1.Text), db, adOpenStatic, adLockReadOnly
            
            Text4.Text = Rs1("title") & " " & Rs1("fname") & " " & Rs1("sname") & " " & Rs1("surname")
            Text5.Text = Rs1("telhome") & "/ " & Rs1("telhome1") & "; Office : " & Rs1("teloffice") & " Ext : " & Rs1("telofficeext") & "; Mobile : " & Rs1("mobileno")
            
        End If
        If Option3.Value = True Or Option4.Value = True Then
            Rs1.Open "select * from vendermaster where venderrefno =" & Val(Combo1.Text), db, adOpenStatic, adLockReadOnly
            
            Text4.Text = Rs1("title") & " " & Rs1("fname") & " " & Rs1("sname") & " " & Rs1("surname")
            Text5.Text = Rs1("telehome") & "/ " & Rs1("telehome1") & "; Office : " & Rs1("telework") & " Ext : " & Rs1("teleworkext") & "; Mobile : " & Rs1("mobile")
            
        End If
        

End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
If Option9.Value = True Then
    If Len(Trim(Text1.Text)) = 0 Then
        MsgBox "Type the Name in the Text Box", vbExclamation, "Type Name"
        Exit Sub
    End If
End If
If Frame2.Visible = True Then
    If Combo1.Text = "" Or Len(Trim(Text4.Text)) = 0 Then
        MsgBox "Enter Name and Refrence No", vbInformation, "Can't Register"
        Exit Sub
    End If
End If

If RsP.State = 1 Then RsP.Close
RsP.Open "select * from outgoingphone order by phonesrno", db, adOpenDynamic, adLockOptimistic


If RsP.RecordCount > 0 Then
    RsP.MoveLast
    LastNo = RsP.RecordCount
End If
LastNo = LastNo + 1
If Frame1.Visible = True Then
    RsP.AddNew
    RsP("phonesrno") = LastNo
    RsP("phoneto") = "Other"
    RsP("clientvendertype") = "Other"
    RsP("refno") = 0
    RsP("name") = IIf(Len(Text1.Text) = 0, Space(1), Text1.Text)
    RsP("telephone") = IIf(Len(Text2.Text) = 0, Space(1), Text2.Text)
    RsP("dateandtime") = Now
    RsP("note") = Text3.Text
    RsP.Update
End If

If Frame2.Visible = True Then
    RsP.AddNew
    RsP("phonesrno") = LastNo
    RsP("phoneto") = "Old"
    If Option1.Value = True Then RsP("clientvendertype") = "RC"
    If Option2.Value = True Then RsP("clientvendertype") = "CC"
    If Option3.Value = True Then RsP("clientvendertype") = "RV"
    If Option4.Value = True Then RsP("clientvendertype") = "CV"
    RsP("refno") = Val(Combo1.Text)
    RsP("name") = IIf(Len(Text4.Text) = 0, Space(1), Text4.Text)
    RsP("telephone") = IIf(Len(Text5.Text) = 0, Space(1), Text5.Text)
    RsP("dateandtime") = Now
    RsP("note") = Text6.Text
    RsP.Update
End If

Dim R
R = MsgBox("Do you want to continue with the phone call entry, and make new Entry", vbYesNoCancel + vbInformation, "Outgoing Call Entered Successfully")
If R = vbNo Then
    Unload Me
End If
If R = vbYes Then
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    
    Combo1.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    
End If
    

End Sub

Private Sub Command3_Click()
On Error Resume Next
If Option3.Value = True Then
    Set Com = Combo1
    Load FrmFindVender
    FrmFindVender.Show 1
End If
If Option1.Value = True Or Option2.Value = True Then
    Set Com = Combo1
    Load FrmFindClient
    FrmFindClient.Show 1
End If
Combo1_Click
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    If RsP.State = 1 Then RsP.Close
    RsP.Open "select * from outgoingphone", db, adOpenDynamic, adLockOptimistic
    
    Option9.Value = True
    Frame1.Visible = True
    Frame2.Visible = False
End Sub

Private Sub CheckFrame()
    Combo1.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    If Option9.Value = True Then
        Frame1.Visible = True
        Frame2.Visible = False
    Else
        Frame2.Visible = True
        Frame1.Visible = False
        Combo1.Clear
        If Option1.Value = True Then
            If Rs1.State = 1 Then Rs1.Close
            Rs1.Open "select * from clientmaster where clienttype = 'Residential' order by clientrefno", db, adOpenDynamic, adLockOptimistic
                If Rs1.RecordCount > 0 Then
                    Rs1.MoveLast
                    trec = Rs1.RecordCount
                    For i = 0 To trec - 1
                        Rs1.MoveFirst
                        Rs1.Move i
                        Combo1.AddItem Rs1("clientrefno")
                    Next
                End If
        End If
        If Option2.Value = True Then
            If Rs1.State = 1 Then Rs1.Close
            Rs1.Open "select * from clientmaster where clienttype = 'Commercial' order by clientrefno", db, adOpenDynamic, adLockOptimistic
            
                If Rs1.RecordCount > 0 Then
                    Rs1.MoveLast
                    trec = Rs1.RecordCount
                    For i = 0 To trec - 1
                        Rs1.MoveFirst
                        Rs1.Move i
                        Combo1.AddItem Rs1("clientrefno")
                    Next
                End If
        End If

        If Option3.Value = True Or Option4.Value = True Then
            If Rs1.State = 1 Then Rs1.Close
            Rs1.Open "select * from vendermaster order by venderrefno", db, adOpenDynamic, adLockOptimistic
           
                If Rs1.RecordCount > 0 Then
                    Rs1.MoveLast
                    trec = Rs1.RecordCount
                    For i = 0 To trec - 1
                        Rs1.MoveFirst
                        Rs1.Move i
                        Combo1.AddItem Rs1("venderrefno")
                    Next
                End If
        End If
        
        
    End If
End Sub

Private Sub Option1_Click()
    Call CheckFrame
End Sub

Private Sub Option2_Click()
    Call CheckFrame
End Sub

Private Sub Option3_Click()
    Call CheckFrame
End Sub

Private Sub Option4_Click()
        Call CheckFrame

End Sub

Private Sub Option9_Click()
    Call CheckFrame
End Sub


'*****

