VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmFindVender 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Vender"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   Icon            =   "FrmFindVender.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   195
      Index           =   3
      Left            =   7665
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   18
      Top             =   570
      Width           =   195
   End
   Begin VB.PictureBox Picture1 
      Height          =   195
      Index           =   2
      Left            =   7665
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   17
      Top             =   120
      Width           =   195
   End
   Begin VB.PictureBox Picture1 
      Height          =   195
      Index           =   1
      Left            =   3660
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   16
      Top             =   570
      Width           =   195
   End
   Begin VB.PictureBox Picture1 
      Height          =   195
      Index           =   0
      Left            =   3660
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   15
      Top             =   120
      Width           =   195
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Result (Double Click the Required Vender Information to go Back)"
      Height          =   2310
      Left            =   75
      TabIndex        =   13
      Top             =   1620
      Width           =   7935
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Mf1 
         Height          =   1980
         Left            =   105
         TabIndex        =   14
         Top             =   225
         Width           =   7740
         _ExtentX        =   13653
         _ExtentY        =   3493
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
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
      Height          =   345
      Left            =   6930
      TabIndex        =   12
      Top             =   1095
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5775
      TabIndex        =   11
      Top             =   1095
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Search Now"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4185
      TabIndex        =   10
      Top             =   1095
      Width           =   1470
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Search For Any value"
      Height          =   195
      Left            =   2010
      TabIndex        =   9
      Top             =   1140
      Width           =   1860
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Search For All value"
      Height          =   195
      Left            =   105
      TabIndex        =   8
      Top             =   1140
      Width           =   1830
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   5415
      TabIndex        =   7
      Text            =   "Combo4"
      Top             =   570
      Width           =   2220
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1410
      TabIndex        =   5
      Text            =   "Combo3"
      Top             =   120
      Width           =   2220
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1410
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   570
      Width           =   2220
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5415
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   2220
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   60
      X2              =   7995
      Y1              =   1515
      Y2              =   1515
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   60
      X2              =   8010
      Y1              =   1035
      Y2              =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Sur Name :"
      Height          =   195
      Left            =   4530
      TabIndex        =   6
      Top             =   615
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Vender Ref. No :"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   165
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Second Name :"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   615
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "First Name :"
      Height          =   195
      Left            =   4500
      TabIndex        =   0
      Top             =   150
      Width           =   840
   End
End
Attribute VB_Name = "FrmFindVender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
     
    Dim RsTemp As New ADODB.Recordset

Private Sub Combo1_Change()
    If Len(Combo1.Text) > 0 Then
        Picture1(2).BackColor = vbGreen
    Else
        Picture1(2).BackColor = vbRed
    End If
End Sub

Private Sub Combo1_Click()
    If Len(Combo1.Text) > 0 Then
        Picture1(2).BackColor = vbGreen
    Else
        Picture1(2).BackColor = vbRed
    End If
End Sub

Private Sub Combo3_Change()
    If Len(Combo3.Text) > 0 Then
        Picture1(0).BackColor = vbGreen
    Else
        Picture1(0).BackColor = vbRed
    End If
End Sub

Private Sub Combo3_Click()
    If Len(Combo3.Text) > 0 Then
        Picture1(0).BackColor = vbGreen
    Else
        Picture1(0).BackColor = vbRed
    End If
End Sub
Private Sub Combo2_Change()
    If Len(Combo2.Text) > 0 Then
        Picture1(1).BackColor = vbGreen
    Else
        Picture1(1).BackColor = vbRed
    End If
End Sub

Private Sub Combo2_Click()
    If Len(Combo2.Text) > 0 Then
        Picture1(1).BackColor = vbGreen
    Else
        Picture1(1).BackColor = vbRed
    End If
End Sub

Private Sub Combo4_Change()
    If Len(Combo4.Text) > 0 Then
        Picture1(3).BackColor = vbGreen
    Else
        Picture1(3).BackColor = vbRed
    End If
End Sub

Private Sub Combo4_Click()
    If Len(Combo4.Text) > 0 Then
        Picture1(3).BackColor = vbGreen
    Else
        Picture1(3).BackColor = vbRed
    End If
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Dim SqlStr1 As String
    Dim i As Integer
    Dim op1 As String
    Dim Ch As Boolean
    For i = 0 To Picture1.Count - 1
        If Picture1(i).BackColor = vbGreen Then
            Ch = True
            Exit For
        End If
    Next
    If Ch = False Then
        MsgBox "Type or Select Search condition", vbInformation, "Can't Search"
        Exit Sub
    End If
        
    
    
    If Option1.Value = True Then
        op1 = " And "
    Else
        op1 = " Or "
    End If
    SqlStr1 = ""
    'concat conditions
        If Picture1(0).BackColor = vbGreen Then SqlStr1 = SqlStr1 & "venderRefNo = " & Val(Combo3.Text)
        If Picture1(1).BackColor = vbGreen Then
            If SqlStr1 <> "" Then SqlStr1 = SqlStr1 & op1
            SqlStr1 = SqlStr1 & "sname = " & "'" & Combo2.Text & "'"
        End If
        If Picture1(2).BackColor = vbGreen Then
            If SqlStr1 <> "" Then SqlStr1 = SqlStr1 & op1
            SqlStr1 = SqlStr1 & "fname = " & "'" & Combo1.Text & "'"
        End If
        If Picture1(3).BackColor = vbGreen Then
            If SqlStr1 <> "" Then SqlStr1 = SqlStr1 & op1
            SqlStr1 = SqlStr1 & "surname = " & "'" & Combo4.Text & "'"
        End If
SqlStr1 = "Select * from vendermaster where " & SqlStr1
Dim TEmpRsVM As New ADODB.Recordset
If TEmpRsVM.State = 1 Then TEmpRsVM.Close
TEmpRsVM.Open SqlStr1, db, adOpenStatic, adLockReadOnly
If TEmpRsVM.RecordCount > 0 Then
    TEmpRsVM.MoveLast
    If TEmpRsVM.RecordCount = 1 Then
        'RsVM.FindFirst "venderrefno = " & TempRsVM(0)
        'Call ShowRecord
        'FrmPropMaster.Combo12.Text = TEmpRsVM("venderrefno")
        Com.Text = TEmpRsVM("venderrefno")
        Unload Me
    Else
        'more than one records
'            Label6.Caption = TEmpRsVM.RecordCount & " Records Found"
'            Frame3.Visible = False
'            Frame4.Visible = True
'            Frame6.Visible = True
'            Frame5.Visible = False
            Dim trec As Integer
            Dim TFld As Integer
            'Dim i As Integer
            Dim j As Integer
            TEmpRsVM.MoveLast
            trec = TEmpRsVM.RecordCount
            Mf1.Rows = trec + 1
            TFld = TEmpRsVM.Fields.Count
            Mf1.Cols = TFld
            
            Mf1.Row = 0
            For j = 0 To TFld - 1
                Mf1.Col = j
                Mf1.Text = TEmpRsVM.Fields(j).Name
            Next
            
            For i = 1 To trec
                TEmpRsVM.MoveFirst
                TEmpRsVM.Move i - 1
                Mf1.Row = i
                For j = 0 To TFld - 1
                    Mf1.Col = j
                    If j = 15 Or j = 16 Or j = 17 Or j = 22 Or j = 23 Or j = 24 Or j = 26 Or j = 28 Then
                        If TEmpRsVM(j) = 1 Then
                            Mf1.Text = "Yes"
                        Else
                            Mf1.Text = "No"
                        End If
                    Else
                        Mf1.Text = TEmpRsVM(j)
                    End If
                Next
            Next
    End If
'    For i = 0 To Picture1.Count - 1
'        Picture1(i).Visible = False
'    Next
Else 'for no match
    If MsgBox("Sorry No record Found, Do you want to try again ? ", vbExclamation + vbYesNo + vbDefaultButton1, "No Match") = vbNo Then
        Unload Me
    End If
End If
End Sub

Private Sub Command2_Click()
Dim i As Integer
    Combo1.Text = ""
    Combo2.Text = ""
    Combo3.Text = ""
    For i = 0 To Picture1.Count - 1
        Picture1(i).BackColor = vbRed
        'Picture1(i).Visible = False
        
    Next

End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim trec As Integer
    Dim i As Integer
    Option1.Value = True
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Combo1.Clear
    Combo2.Clear
    Combo3.Clear
    Combo4.Clear
    For i = 0 To Picture1.Count - 1
        Picture1(i).BackColor = vbRed
        'Picture1(i).Visible = False
    Next
    
    If RsTemp.State = 1 Then RsTemp.Close
    RsTemp.Open "select venderrefno from vendermaster group by venderrefno", db, adOpenStatic, adLockReadOnly
    If RsTemp.RecordCount > 0 Then
        RsTemp.MoveLast
        trec = RsTemp.RecordCount
        For i = 0 To trec - 1
            RsTemp.MoveFirst
            RsTemp.Move i
            Combo3.AddItem RsTemp(0)
        Next
    End If
    If RsTemp.State = 1 Then RsTemp.Close
    RsTemp.Open "select sname from vendermaster group by sname", db, adOpenStatic, adLockReadOnly
    If RsTemp.RecordCount > 0 Then
        RsTemp.MoveLast
        trec = RsTemp.RecordCount
        For i = 0 To trec - 1
            RsTemp.MoveFirst
            RsTemp.Move i
            Combo2.AddItem RsTemp(0)
        Next
    End If
    If RsTemp.State = 1 Then RsTemp.Close
    RsTemp.Open "select fname from vendermaster group by fname", db, adOpenStatic, adLockReadOnly
    
    If RsTemp.RecordCount > 0 Then
        RsTemp.MoveLast
        trec = RsTemp.RecordCount
        For i = 0 To trec - 1
            RsTemp.MoveFirst
            RsTemp.Move i
            Combo1.AddItem RsTemp(0)
        Next
    End If
    If RsTemp.State = 1 Then RsTemp.Close
    RsTemp.Open "select surname from vendermaster group by surname", db, adOpenStatic, adLockReadOnly
    If RsTemp.RecordCount > 0 Then
        RsTemp.MoveLast
        trec = RsTemp.RecordCount
        For i = 0 To trec - 1
            RsTemp.MoveFirst
            RsTemp.Move i
            If IsNull(RsTemp(0)) = False Then
            Combo4.AddItem RsTemp(0)
            End If
        Next
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmFindVender = Nothing
End Sub

Private Sub Mf1_dblClick()
    If Mf1.Row > 0 Then
        'FrmPropMaster.Combo12.Text = Mf1.TextMatrix(Mf1.Row, 0)
        Com.Text = Mf1.TextMatrix(Mf1.Row, 0)
        Unload Me
    End If

End Sub
