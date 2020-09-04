VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmAppoinment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Appoinments"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8715
   Icon            =   "FrmAppoinment.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option4 
      Caption         =   "View with Agent"
      Height          =   225
      Left            =   1965
      TabIndex        =   45
      Top             =   4350
      Width           =   1605
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Client Direct"
      Height          =   225
      Left            =   300
      TabIndex        =   44
      Top             =   4350
      Width           =   1605
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   2295
      TabIndex        =   35
      Top             =   3870
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Format          =   69074946
      CurrentDate     =   37673
   End
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
      Height          =   390
      Left            =   5310
      TabIndex        =   34
      Top             =   4335
      Width           =   1350
   End
   Begin VB.CommandButton Command1 
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
      Left            =   7125
      TabIndex        =   33
      Top             =   4335
      Width           =   1350
   End
   Begin VB.Frame Frame3 
      Caption         =   "Appointment :"
      ForeColor       =   &H00FF0000&
      Height          =   750
      Left            =   4785
      TabIndex        =   30
      Top             =   3405
      Width           =   3780
      Begin VB.OptionButton Option2 
         Caption         =   "Not Confirmed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1905
         TabIndex        =   32
         Top             =   330
         Width           =   1800
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Confirmed"
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
         Left            =   300
         TabIndex        =   31
         Top             =   330
         Width           =   1515
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   2295
      TabIndex        =   28
      Top             =   3465
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   529
      _Version        =   393216
      Format          =   69074945
      CurrentDate     =   37673
   End
   Begin VB.Frame Frame2 
      Caption         =   "Property :"
      ForeColor       =   &H00FF0000&
      Height          =   1980
      Left            =   120
      TabIndex        =   14
      Top             =   1380
      Width           =   8535
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   1410
         TabIndex        =   43
         Top             =   1575
         Width           =   2850
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   5490
         TabIndex        =   42
         Top             =   1575
         Width           =   2880
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   5445
         TabIndex        =   37
         Top             =   1110
         Width           =   1380
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   7455
         TabIndex        =   36
         Top             =   1110
         Width           =   900
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   3430
         TabIndex        =   26
         Top             =   1110
         Width           =   1380
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1410
         TabIndex        =   24
         Top             =   1110
         Width           =   1380
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   5460
         TabIndex        =   22
         Top             =   780
         Width           =   2895
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Height          =   285
         Left            =   1410
         TabIndex        =   20
         Top             =   765
         Width           =   1380
      End
      Begin VB.TextBox Text8 
         Height          =   525
         Left            =   5460
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   195
         Width           =   2910
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Height          =   285
         Left            =   1410
         TabIndex        =   16
         Top             =   195
         Width           =   1380
      End
      Begin VB.Label Label19 
         Caption         =   "Telephone1 :"
         Height          =   225
         Left            =   4410
         TabIndex        =   41
         Top             =   1590
         Width           =   960
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Tenet Name :"
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   1590
         Width           =   975
      End
      Begin VB.Line Line3 
         X1              =   225
         X2              =   8325
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   225
         X2              =   8340
         Y1              =   1455
         Y2              =   1455
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Ext."
         Height          =   195
         Left            =   7005
         TabIndex        =   39
         Top             =   1170
         Width           =   270
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Office :"
         Height          =   195
         Left            =   4890
         TabIndex        =   38
         Top             =   1140
         Width           =   510
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "(2)"
         Height          =   195
         Left            =   3015
         TabIndex        =   25
         Top             =   1125
         Width           =   180
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Telephone1 :"
         Height          =   195
         Left            =   165
         TabIndex        =   23
         Top             =   1140
         Width           =   945
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Name of Vendor :"
         Height          =   195
         Left            =   4140
         TabIndex        =   21
         Top             =   795
         Width           =   1245
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Vendor Ref. No. :"
         Height          =   195
         Left            =   165
         TabIndex        =   19
         Top             =   750
         Width           =   1245
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Address :"
         Height          =   195
         Left            =   4725
         TabIndex        =   17
         Top             =   210
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Property Ref. No. :"
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   285
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Client :"
      ForeColor       =   &H00FF0000&
      Height          =   960
      Left            =   105
      TabIndex        =   1
      Top             =   390
      Width           =   8490
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   7440
         TabIndex        =   13
         Top             =   600
         Width           =   900
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   5435
         TabIndex        =   11
         Top             =   600
         Width           =   1380
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3430
         TabIndex        =   9
         Top             =   600
         Width           =   1380
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1425
         TabIndex        =   7
         Top             =   600
         Width           =   1380
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5430
         TabIndex        =   5
         Top             =   225
         Width           =   2940
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Height          =   285
         Left            =   1425
         TabIndex        =   3
         Top             =   225
         Width           =   1380
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ext."
         Height          =   195
         Left            =   7035
         TabIndex        =   12
         Top             =   615
         Width           =   270
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Office :"
         Height          =   195
         Left            =   4875
         TabIndex        =   10
         Top             =   615
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(H2)"
         Height          =   195
         Left            =   2985
         TabIndex        =   8
         Top             =   615
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Telephone (H1) :"
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   615
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Client Name :"
         Height          =   195
         Left            =   4380
         TabIndex        =   4
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Client Ref. No :"
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Appointment No : 1"
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
      Height          =   195
      Left            =   465
      TabIndex        =   46
      Top             =   60
      Width           =   1650
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Request Appointment Time :"
      Height          =   195
      Left            =   135
      TabIndex        =   29
      Top             =   3870
      Width           =   2010
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Request Appointment Date  :"
      Height          =   195
      Left            =   135
      TabIndex        =   27
      Top             =   3465
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   6300
      X2              =   465
      Y1              =   300
      Y2              =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Appointments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   6300
      TabIndex        =   0
      Top             =   75
      Width           =   1695
   End
End
Attribute VB_Name = "FrmAppoinment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim RSA As New ADODB.Recordset
Dim RSPM As New ADODB.Recordset
Dim RsCM As New ADODB.Recordset
Dim RsVM As New ADODB.Recordset
Dim LastNo As Integer

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    If Option2.Value = True Then
        If MsgBox("Appointment is Not Confirmed with " & Text15.Text & "." & vbCrLf & vbCrLf & "Do you want to Save the Appointment for Future Confirmation.", vbInformation + vbYesNo, "Appointment is Not Confirmed") = vbNo Then
            Exit Sub
        End If
    End If
    RSA.AddNew
    RSA("appointmentno") = LastNo
    RSA("appointmenttype") = "Vender"
    RSA("appointmentwithname") = IIf(Len(Trim(Text15.Text)) = 0, Text10.Text, Text15.Text)
    RSA("telephone") = IIf(Len(Trim(Text15.Text)) = 0, Text11.Text & "/ " & Text12.Text, Text16.Text)
    RSA("address") = Text8.Text
    
    RSA("clientrefno") = Val(Text1.Text)
    RSA("proprefno") = Val(Text7.Text)
    RSA("venderrefno") = Val(Text9.Text)
    If Option1.Value = True Then
        RSA("appointmentstatus") = "Confirmed"
    Else
        RSA("appointmentstatus") = "Not Confirmed"
    End If
    RSA("date") = DTPicker1.Value
    RSA("time") = CDate(DTPicker1.Value & " " & Format(DTPicker2.Value, "hh:mm:ss AMPM"))
    If Option3.Value = True Then
        RSA("view") = "Direct Client"
    Else
        RSA("view") = "View with Agent"
    End If
    RSA("seen") = "No"
    RSA("like") = "No"
    RSA("note") = "Nil"
    
    
    RSA.Update
    If Option1.Value = True Then MsgBox "Appointment is Confirmed with " & Text15.Text, vbInformation, "Appointment is Confirmed"
    Unload Me
End Sub

Private Sub Form_Load()
    If RSA.State = 1 Then RSA.Close
    RSA.Open "select * from appointment order by appointmentno", db, adOpenDynamic, adLockOptimistic
    If RSPM.State = 1 Then RSPM.Close
    RSPM.Open "select * from propertymaster where proprefno = " & PropertyRefNo, db, adOpenStatic, adLockReadOnly
    If RsCM.State = 1 Then RsCM.Close
    RsCM.Open "select * from clientmaster where clientrefno = " & ClientRefNo, db, adOpenStatic, adLockReadOnly
     
    
    If RSPM.RecordCount > 0 And RsCM.RecordCount > 0 Then
        Text1.Text = ClientRefNo
        Text2.Text = RsCM("title") & " " & RsCM("fname") & " " & RsCM("sname") & " " & RsCM("surname")
        Text3.Text = RsCM("telhome")
        Text4.Text = RsCM("telhome1")
        Text5.Text = RsCM("teloffice")
        Text6.Text = RsCM("telofficeext")
        
        Text7.Text = PropertyRefNo
        Text8.Text = RSPM("buildingno") & ", " & RSPM("streetname") & ", " & RSPM("city") & ", " & RSPM("county") & " - " & RSPM("postalcode") & ", " & RSPM("country")
        Text9.Text = RSPM("venderrefno")
        If RsVM.State = 1 Then RsVM.Close
        RsVM.Open "select * from vendermaster where venderrefno = " & RSPM("venderrefno"), db, adOpenStatic, adLockReadOnly
        Text10.Text = RsVM("title") & " " & RsVM("fname") & " " & RsVM("sname") & " " & RsVM("surname")
        Text11.Text = RsVM("telehome")
        Text12.Text = RsVM("telehome1")
        Text14.Text = RsVM("telework")
        Text13.Text = RsVM("teleworkext")
        Text15.Text = RsVM("personname")
        Text16.Text = RsVM("tenenttel1") & " / " & RsVM("tenenttel2")
        
        If RSA.RecordCount > 0 Then
            RSA.MoveLast
            LastNo = RSA("appointmentno")
        End If
        LastNo = LastNo + 1
        Label20.Caption = "Appointment No : " & LastNo
    Else
        MsgBox "Sorry Related Record Not Found For Appointment"
    End If
    
    Option2.Value = True
    Option4.Value = True
    DTPicker1.Value = Date
    DTPicker2.Value = Now
End Sub

Private Sub NewAppointment()
    DTPicker1.Value = Date
    DTPicker2.Value = Now
End Sub
