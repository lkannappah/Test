VERSION 5.00
Begin VB.Form FrmEnquiry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incoming Phone and Enquiry"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   Icon            =   "FrmEnquery.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
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
      Height          =   390
      Left            =   5250
      TabIndex        =   13
      Top             =   3840
      Width           =   1365
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6780
      TabIndex        =   12
      Top             =   3840
      Width           =   1365
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
      Left            =   6645
      TabIndex        =   10
      Top             =   150
      Width           =   1170
   End
   Begin VB.OptionButton Option8 
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
      Left            =   3000
      TabIndex        =   9
      Top             =   1515
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.OptionButton Option7 
      Caption         =   "Vender"
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
      Left            =   2985
      TabIndex        =   8
      Top             =   1185
      Width           =   2010
   End
   Begin VB.OptionButton Option6 
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
      Left            =   3000
      TabIndex        =   7
      Top             =   720
      Width           =   2010
   End
   Begin VB.OptionButton Option5 
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
      Left            =   3000
      TabIndex        =   6
      Top             =   165
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
      Left            =   435
      TabIndex        =   4
      Top             =   1515
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Vender"
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
      Left            =   420
      TabIndex        =   3
      Top             =   1185
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
      Left            =   435
      TabIndex        =   2
      Top             =   720
      Width           =   2010
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
      Left            =   435
      TabIndex        =   1
      Top             =   165
      Width           =   2010
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1770
      Left            =   90
      TabIndex        =   14
      Top             =   1875
      Width           =   8235
      Begin VB.TextBox Text3 
         Height          =   660
         Left            =   1545
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   945
         Width           =   6585
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   6030
         TabIndex        =   21
         Top             =   525
         Width           =   2100
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1530
         TabIndex        =   19
         Top             =   525
         Width           =   3465
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Find Ref."
         Height          =   345
         Left            =   5100
         TabIndex        =   17
         Top             =   75
         Width           =   1500
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1530
         TabIndex        =   16
         Top             =   75
         Width           =   3495
      End
      Begin VB.Label Label7 
         Caption         =   "Note :"
         Height          =   270
         Left            =   150
         TabIndex        =   22
         Top             =   945
         Width           =   1020
      End
      Begin VB.Label Label6 
         Caption         =   "Date of Reg."
         Height          =   270
         Left            =   5100
         TabIndex        =   20
         Top             =   585
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Name :"
         Height          =   255
         Left            =   135
         TabIndex        =   18
         Top             =   540
         Width           =   1305
      End
      Begin VB.Label Label4 
         Caption         =   "Reference No :"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1770
      Left            =   90
      TabIndex        =   30
      Top             =   1860
      Width           =   8235
      Begin VB.TextBox Text7 
         Height          =   660
         Left            =   1515
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   660
         Width           =   6525
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   1515
         TabIndex        =   31
         Top             =   210
         Width           =   3465
      End
      Begin VB.Label Label11 
         Caption         =   "Note :"
         Height          =   270
         Left            =   135
         TabIndex        =   34
         Top             =   615
         Width           =   1020
      End
      Begin VB.Label Label9 
         Caption         =   "Name :"
         Height          =   255
         Left            =   135
         TabIndex        =   33
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1770
      Left            =   90
      TabIndex        =   24
      Top             =   1875
      Width           =   8235
      Begin VB.CheckBox Check1 
         Caption         =   "After Saving the Enquiry Register New Client/ Vender"
         Height          =   255
         Left            =   1545
         TabIndex        =   29
         Top             =   1395
         Width           =   4470
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   1515
         TabIndex        =   26
         Top             =   210
         Width           =   3465
      End
      Begin VB.TextBox Text4 
         Height          =   660
         Left            =   1515
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   660
         Width           =   6525
      End
      Begin VB.Label Label10 
         Caption         =   "Name :"
         Height          =   255
         Left            =   135
         TabIndex        =   28
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label8 
         Caption         =   "Note :"
         Height          =   270
         Left            =   135
         TabIndex        =   27
         Top             =   615
         Width           =   1020
      End
   End
   Begin VB.Line Line4 
      X1              =   90
      X2              =   8325
      Y1              =   3690
      Y2              =   3690
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   8295
      Y1              =   3705
      Y2              =   3705
   End
   Begin VB.Line Line2 
      X1              =   105
      X2              =   8340
      Y1              =   1785
      Y2              =   1785
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   135
      X2              =   8310
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label3 
      Caption         =   "(Nither Old or New Clients and Venders)"
      Height          =   495
      Left            =   6645
      TabIndex        =   11
      Top             =   465
      Width           =   1470
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "N  E  W"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1020
      Left            =   2655
      TabIndex        =   5
      Top             =   345
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "O  L  D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1020
      Left            =   105
      TabIndex        =   0
      Top             =   330
      Width           =   150
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1665
      Left            =   180
      Top             =   45
      Width           =   2325
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      Height          =   1665
      Left            =   2745
      Top             =   45
      Width           =   3660
   End
End
Attribute VB_Name = "FrmEnquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim RsP As New ADODB.Recordset
Dim LastNo As Integer
Dim Rs1 As New ADODB.Recordset
Dim trec As Integer
Dim i As Integer

Private Sub Combo1_Click()
If Rs1.State = 1 Then Rs1.Close
If Frame1.Visible = True Then
        If Option1.Value = True Or Option2.Value = True Then
            Rs1.Open "select * from clientmaster where clientrefno =" & Val(Combo1.Text), db, adOpenStatic, adLockReadOnly
            
            Text1.Text = Rs1("title") & " " & Rs1("fname") & " " & Rs1("sname") & " " & Rs1("surname")
            Text2.Text = Rs1("dor")
            
        End If
        If Option3.Value = True Or Option4.Value = True Then
            Rs1.Open "select * from vendermaster where venderrefno =" & Val(Combo1.Text), db, adOpenStatic, adLockReadOnly
            
            Text1.Text = Rs1("title") & " " & Rs1("fname") & " " & Rs1("sname") & " " & Rs1("surname")
            Text2.Text = Rs1("dor")
         End If

End If
End Sub



Private Sub Command2_Click()
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

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
If RsP.RecordCount > 0 Then
    RsP.MoveLast
    LastNo = RsP("phonesrno")
End If
LastNo = LastNo + 1

If Frame1.Visible = True Then
    RsP.AddNew
    RsP("phonesrno") = LastNo
    RsP("enquiryby") = "Old"
    If Option1.Value = True Then RsP("clientvendertype") = "RC"
    If Option2.Value = True Then RsP("clientvendertype") = "CC"
    If Option3.Value = True Then RsP("clientvendertype") = "RV"
    If Option4.Value = True Then RsP("clientvendertype") = "CV"
    RsP("refno") = Val(Combo1.Text)
    RsP("name") = IIf(Len(Text1.Text) = 0, Space(1), Text1.Text)
    RsP("dateandtime") = Now
    RsP("note") = Text3.Text
    RsP.Update
End If
If Frame2.Visible = True Then
    RsP.AddNew
    RsP("phonesrno") = LastNo
    RsP("enquiryby") = "New"
    If Option5.Value = True Then RsP("clientvendertype") = "RC"
    If Option6.Value = True Then RsP("clientvendertype") = "CC"
    If Option7.Value = True Then RsP("clientvendertype") = "RV"
    If Option8.Value = True Then RsP("clientvendertype") = "CV"
    RsP("refno") = 0
    RsP("name") = IIf(Len(Text6.Text) = 0, Space(1), Text6.Text)
    RsP("dateandtime") = Now
    RsP("note") = Text4.Text
    RsP.Update
    If Check1.Value = 1 Then
        Dim Str As String
        If Option5.Value = True Then Str = Option5.Caption
        If Option6.Value = True Then Str = Option6.Caption
        If Option7.Value = True Then Str = Option7.Caption
        If Option8.Value = True Then Str = Option8.Caption
        If MsgBox("Do You want to Register the new " & Str & " Now.", vbYesNo + vbExclamation, "Register New " & Str) = vbYes Then
            'open the master
        
        
        Unload Me
        Exit Sub
        End If
    End If
                 
End If
If Frame3.Visible = True Then
    RsP.AddNew
    RsP("phonesrno") = LastNo
    RsP("enquiryby") = "Other"
    RsP("refno") = 0
    RsP("name") = IIf(Len(Text5.Text) = 0, Space(1), Text5.Text)
    RsP("dateandtime") = Now
    RsP("note") = Text7.Text
    RsP.Update
End If
If MsgBox("Succesfully Saved the Record. Do You to continue with the Incoming Enqiry Phone", vbYesNo + vbInformation, "Continue with Phone Enquiry") = vbYes Then
    Unload Me
    Load FrmEnquiry
    FrmEnquiry.Show 1
Else
    Unload Me
End If



End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Option9.Value = True
Frame3.Visible = True
Frame2.Visible = False
Frame1.Visible = False

If RsP.State = 1 Then RsP.Close
RsP.Open "select * from incomingphone", db, adOpenDynamic, adLockOptimistic

Check1.Value = 1
End Sub

Private Sub Option1_Click()
Call CheckButton
End Sub
Private Sub Option2_Click()
Call CheckButton
End Sub
Private Sub Option3_Click()
Call CheckButton
End Sub
Private Sub Option4_Click()
Call CheckButton
End Sub
Private Sub Option5_Click()
Call CheckButton
End Sub
Private Sub Option6_Click()
Call CheckButton
End Sub
Private Sub Option7_Click()
Call CheckButton
End Sub
Private Sub Option8_Click()
Call CheckButton
End Sub
Private Sub Option9_Click()
Call CheckButton
End Sub

Private Sub CheckButton()
    Text1.Text = ""
    Text3.Text = ""
    Text2.Text = ""
    Combo1.Text = ""
    Text5.Text = ""
    Text7.Text = ""
    If Option5.Value = True Or Option6.Value = True Or Option7.Value = True Or Option8.Value = True Then
        
        Frame2.Visible = True
    Else
        
        Frame2.Visible = False
    End If
    If Option9.Value = True Then
        Frame3.Visible = True
    Else
        Frame3.Visible = False
    End If
    If Rs1.State = 1 Then Rs1.Close
    If Option1.Value = True Or Option2.Value = True Or Option3.Value = True Or Option4.Value = True Then
        Frame1.Visible = True
        'call combo fill
        Combo1.Clear
        If Option1.Value = True Then
            Rs1.Open "select * from clientmaster where clienttype = 'Residential' order by clientrefno", db, adOpenStatic, adLockReadOnly
            
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
            Rs1.Open "select * from clientmaster where clienttype = 'Commercial' order by clientrefno", db, adOpenStatic, adLockReadOnly
           
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
            Rs1.Open "select * from vendermaster order by venderrefno", db, adOpenStatic, adLockReadOnly
            
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
    Else
        Frame1.Visible = False
    End If
    
    
End Sub
