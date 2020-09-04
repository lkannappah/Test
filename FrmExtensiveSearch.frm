VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmExtensiveSearch 
   Caption         =   "Extensive Search "
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11010
   Icon            =   "FrmExtensiveSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   11010
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   2700
      Left            =   2430
      TabIndex        =   16
      Top             =   2760
      Width           =   6900
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   1995
         Left            =   255
         TabIndex        =   17
         Top             =   465
         Width           =   6120
         Begin VB.CommandButton Command6 
            Caption         =   "&Reset Search"
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
            Left            =   1185
            TabIndex        =   19
            Top             =   1605
            Width           =   1515
         End
         Begin VB.CommandButton Command5 
            Caption         =   "&Close Search"
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
            Left            =   3495
            TabIndex        =   18
            Top             =   1605
            Width           =   1515
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   105
            TabIndex        =   20
            Top             =   375
            Width           =   5910
         End
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4620
      Left            =   180
      TabIndex        =   9
      Top             =   1980
      Width           =   9885
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   375
         Left            =   1740
         TabIndex        =   12
         Top             =   4200
         Width           =   4800
         Begin VB.CommandButton Command2 
            Caption         =   "&Open Record"
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
            Left            =   0
            TabIndex        =   15
            Top             =   15
            Width           =   1470
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Reset Search"
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
            Left            =   1635
            TabIndex        =   14
            Top             =   0
            Width           =   1470
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Close Search"
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
            Left            =   3240
            TabIndex        =   13
            Top             =   0
            Width           =   1470
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MF1 
         Height          =   1980
         Left            =   60
         TabIndex        =   11
         Top             =   210
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   3493
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1890
      Left            =   165
      TabIndex        =   0
      Top             =   0
      Width           =   9870
      Begin VB.PictureBox PicProgress 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
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
         Left            =   135
         ScaleHeight     =   255
         ScaleWidth      =   9585
         TabIndex        =   10
         Top             =   1545
         Width           =   9585
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   2010
         TabIndex        =   5
         Top             =   165
         Width           =   5985
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Let's Search"
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
         Left            =   8100
         TabIndex        =   4
         Top             =   165
         Width           =   1560
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Search for All Words in Sequence, Match whole Word Only"
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
         Left            =   2010
         TabIndex        =   3
         Top             =   585
         Width           =   6045
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Search for Any Words Not in Sequence, Match whole Word Only"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2010
         TabIndex        =   2
         Top             =   915
         Width           =   6045
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Search for Part of the Records (Wild Search)"
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
         Left            =   2025
         TabIndex        =   1
         Top             =   1245
         Width           =   6045
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "It is Strongly Recomanded that you update keyword table regularly"
         ForeColor       =   &H80000008&
         Height          =   810
         Left            =   8130
         TabIndex        =   8
         Top             =   585
         Width           =   1560
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Your Keyword Table was last updated on 12/12/02 and having 22220000 Keywords"
         ForeColor       =   &H80000008&
         Height          =   840
         Left            =   105
         TabIndex        =   7
         Top             =   570
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Enter Search String :"
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
         Left            =   150
         TabIndex        =   6
         Top             =   210
         Width           =   1800
      End
   End
End
Attribute VB_Name = "FrmExtensiveSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Fs As New FileSystemObject

Dim RsK As New ADODB.Recordset
Dim TempRs As New ADODB.Recordset


Private Sub Command1_Click()
Frame2.Visible = False


Dim sqlstr As String
    Label4.ForeColor = vbBlack
    Label4.Caption = "Wait Search is on ..."
    UpdateStatus FrmExtensiveSearch.PicProgress, 1 / 100, False
    'first condition
    If Option1.Value = True Then
        If TempRs.State = 1 Then TempRs.Close
        TempRs.Open "select * from keywords where word = " & "'" & Text1.Text & "'", db, adOpenStatic, adLockReadOnly
        If TempRs.RecordCount > 0 Then
            TempRs.MoveLast
            UpdateStatus FrmExtensiveSearch.PicProgress, 10 / 100, False
            Frame2.Visible = True
            Frame4.Visible = False
            Call DisRecord
        Else
            UpdateStatus FrmExtensiveSearch.PicProgress, 1, True
            Frame2.Visible = False
            Frame4.Visible = True
            Label4.ForeColor = vbRed
            Label4.Caption = "No Match Found"
            PicProgress.Cls
            PicProgress.BackColor = Me.BackColor

             
        End If
    End If
    
    
    
    
    'second condition
    If Option2.Value = True Then
            Dim Txt As String
        Txt = Text1.Text
        If Len(Txt) > 500 Then
            MsgBox "Please Enter Small String for Search, Maximum character is 250"
            Exit Sub
        End If
            Dim X3 As String
            Dim i1 As Integer
            X3 = Trim(Txt)
            If X3 = "" Or X3 = " " Then 'check search string blank space or empty
                    MsgBox "        Please enter a Search string      ", vbExclamation, "Enter a String"
                    PicProgress.Cls
                    PicProgress.BackColor = Me.BackColor
                    Label4.ForeColor = vbBlue
                    Label4.Caption = "Search is Ready"
                
'                PB1.Value = 0
'                PB1.Visible = False
                Exit Sub
            End If
                For i1 = 50 To 2 Step -1
                    X3 = Replace(X3, Space(i1), " ")
                Next
            Txt = X3 & " "
            
            'count the nos. of word by counting the space
            'seperate each word
            'put each word in the pocket of array Xx
            'we use maximum nos. of word is 20
            'you can increase it by changing the suffix, also change the nos of varible declearation later on..
            Dim k As Integer
            Dim XX(20) As String
            Dim NewPos As Integer
            Dim StartCount As Integer
            Dim C As Integer
            Dim C1 As Integer
            Dim L As Integer
            Dim X As Integer
            Dim S As Integer
            Dim j As Integer
                C = 0
                k = 0
                NewPos = 0
                StartCount = 0
                C1 = 0
                S = 0
                X = 0
                L = Len(Txt)
                For k = 1 To L
                    If Mid(Txt, k, 1) = " " Then
                    C = C + 1
                    End If
                Next
                If C > 20 Then ' checking the nos. of maximum words
                    MsgBox "Please Enter a small search string, Maximum capacity is 20 words", vbExclamation, "Enter a Small String"
                    PicProgress.Cls
                    PicProgress.BackColor = Me.BackColor
                    Label4.ForeColor = vbBlue
                    Label4.Caption = "Search is Ready"
                    
                    
                    Exit Sub
                End If
                StartCount = 1
                S = 1
top2:
                For k = StartCount To L
                    If Mid(Txt, k, 1) = " " Then StartCount = k + 1: S = k + 1: C1 = 0: Exit For
                    C1 = C1 + 1
                    XX(X) = Mid(Txt, S, C1)
                    
                Next
                X = X + 1
                If X >= C Then
                GoTo bottom2
                Else
                GoTo top2
                End If
bottom2:

        'declare the variable for each pocket of array
        'you can use array for it
        'if you use array then code will become smaller
        Dim S1 As String, S2 As String, S3 As String, S4 As String, S5 As String
        Dim S6 As String, S7 As String, S8 As String, S9 As String, S10 As String
        Dim S11 As String, S12 As String, S13 As String, S14 As String, S15 As String
        Dim S16 As String, S17 As String, S18 As String, S19 As String, S20 As String
        Dim SqlStr1 As String
        
        
            S1 = XX(0): If Len(S1) > 0 Then SqlStr1 = SqlStr1 & " word = '" & S1 & "'"
            S2 = XX(1): If Len(S2) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S2 & "'"
            S3 = XX(2): If Len(S3) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S3 & "'"
            S4 = XX(3): If Len(S4) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S4 & "'"
            S5 = XX(4): If Len(S5) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S5 & "'"
            S6 = XX(5): If Len(S6) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S6 & "'"
            S7 = XX(6): If Len(S7) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S7 & "'"
            S8 = XX(7): If Len(S8) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S8 & "'"
            S9 = XX(8): If Len(S9) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S9 & "'"
            S10 = XX(9): If Len(S10) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S10 & "'"
            S11 = XX(10): If Len(S11) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S11 & "'"
            S12 = XX(11): If Len(S12) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S12 & "'"
            S13 = XX(12): If Len(S13) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S13 & "'"
            S14 = XX(13): If Len(S14) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S14 & "'"
            S15 = XX(14): If Len(S15) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S15 & "'"
            S16 = XX(15): If Len(S16) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S16 & "'"
            S17 = XX(16): If Len(S17) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S17 & "'"
            S18 = XX(17): If Len(S18) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S18 & "'"
            S19 = XX(18): If Len(S19) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S19 & "'"
            S20 = XX(19): If Len(S20) > 0 Then SqlStr1 = SqlStr1 & " or word = '" & S20 & "'"
            sqlstr = "Select * from keywords where " & SqlStr1 & " order by word"
            
            If TempRs.State = 1 Then TempRs.Close
            TempRs.Open sqlstr, db, adOpenStatic, adLockReadOnly
            
            
            
            
            If TempRs.RecordCount > 0 Then
                TempRs.MoveLast
                UpdateStatus FrmExtensiveSearch.PicProgress, 10 / 100, False
                Frame2.Visible = True
                Frame4.Visible = False
                Call DisRecord
            Else
                UpdateStatus FrmExtensiveSearch.PicProgress, 1, True
                Frame2.Visible = False
                Frame4.Visible = True
                Label4.ForeColor = vbRed
                Label4.Caption = "No Match Found"
                PicProgress.Cls
                PicProgress.BackColor = Me.BackColor
    
                'MsgBox "NOOOOOO"
            End If
End If
 
'for third condition
If Option3.Value = True Then
        If Text1.Text = "" Then
                MsgBox "        Please enter a Search string      ", vbExclamation, "Enter a String"
                PicProgress.Cls
                PicProgress.BackColor = Me.BackColor
                Label4.ForeColor = vbBlue
                Label4.Caption = "Search is Ready"
                Exit Sub
        End If
        
        If TempRs.State = 1 Then TempRs.Close
        TempRs.Open "select * from keywords where word like " & "'%" & Text1.Text & "%'", db, adOpenStatic, adLockReadOnly
                    
        If TempRs.RecordCount > 0 Then
            TempRs.MoveLast
            UpdateStatus FrmExtensiveSearch.PicProgress, 10 / 100, False
            Frame2.Visible = True
            Frame4.Visible = False
            Call DisRecord
        Else
            UpdateStatus FrmExtensiveSearch.PicProgress, 1, True
            Frame2.Visible = False
            Frame4.Visible = True
            Label4.ForeColor = vbRed
            Label4.Caption = "No Match Found"
            PicProgress.Cls
            PicProgress.BackColor = Me.BackColor

            'MsgBox "NOOOOOO"
        End If
End If
 
End Sub

Private Sub Command3_Click()
    Frame2.Visible = False
    Frame4.Visible = True
    Label4.ForeColor = vbBlue
    Label4.Caption = "Search is Ready"
    
    Text1.Text = ""
        'UpdateStatus FrmExtensiveSearch.PicProgress, 0 / 100, False
    PicProgress.Cls
    PicProgress.BackColor = Me.BackColor
End Sub

Private Sub Command4_Click()
Unload Me

End Sub

Private Sub Command5_Click()
Command4_Click
End Sub

Private Sub Command6_Click()
Command3_Click
End Sub

Private Sub Form_Load()
    Label4.ForeColor = vbBlue
    Label4.Caption = "Search is Ready"

    Frame2.Visible = False
    If RsK.State = 1 Then RsK.Close
    RsK.Open "select count(*) from Keywords", db, adOpenStatic, adLockReadOnly
    If RsK.RecordCount > 0 Then
        RsK.MoveLast
    Else
        MsgBox "No keyword are available for Searching, Update keyword first", vbCritical, "Can't Search : No Keyword Found"
        Unload Me
    End If
    
    Option1.Value = True
    Label2.Caption = "Your Keyword Table having " & RsK(0) & " Keywords"
End Sub

Private Sub Form_Resize()
On Error Resume Next
Frame1.Left = (Me.Width - Frame1.Width) / 2
    Frame2.Left = 50
    Frame2.Width = Me.Width - 200
    Frame2.Top = Frame1.Top + Frame1.Height + 50
    Frame2.Height = Me.Height - Frame1.Height - 500
    Mf1.Width = Frame2.Width - 150
    Mf1.Height = Frame2.Height - 750
    Frame3.Left = (Frame2.Width - Frame3.Width) / 2
    Frame3.Top = Mf1.Top + Mf1.Height + 100
    
    Frame4.Top = Frame2.Top
    Frame4.Left = Frame2.Left
    Frame4.Width = Frame2.Width
    Frame4.Height = Frame2.Height
    
    Frame5.Left = (Frame4.Width - Frame5.Width) / 2
    Frame5.Top = (Frame4.Height - Frame5.Height) / 2
    
    
    Dim X As Long
    Dim R As Long
    Dim T As Long
    X = Mf1.Width - 500
    R = X \ 5
    Dim i As Integer
    For i = 1 To 4
        Mf1.ColWidth(i) = R
        T = T + R
    Next
    Mf1.ColWidth(0) = X - T
End Sub

Private Sub DisRecord()
On Error Resume Next
            Dim trec As Integer
            Dim TFld As Integer
            Dim i As Integer
            Dim j As Integer
            'TempRS.MoveLast
            trec = TempRs.RecordCount
            Mf1.Rows = trec + 1
            TFld = TempRs.Fields.Count
            Mf1.Cols = TFld + 1
            
            Mf1.Row = 0
            Mf1.Col = 0
            Mf1.Text = " Sr. "
            
            Mf1.Col = 1
            Mf1.Text = " Search String "

            Mf1.Col = 2
            Mf1.Text = " Field Name "
            
            Mf1.Col = 3
            Mf1.Text = " Table Name "

            Mf1.Col = 4
            Mf1.Text = " Record No. "
            
            For i = 1 To trec
            UpdateStatus FrmExtensiveSearch.PicProgress, ((90 / (trec - 1)) * i) / 100, False

                TempRs.MoveFirst
                TempRs.Move i - 1
                Mf1.Row = i
                Mf1.Col = 0
                Mf1.Text = i + 1
                For j = 0 To TFld - 1
                    Mf1.Col = j + 1
                    If j = 15 Or j = 16 Or j = 17 Or j = 22 Or j = 23 Or j = 24 Or j = 26 Or j = 28 Then
                        If TempRs(j) = 1 Then
                            Mf1.Text = "Yes"
                        Else
                            Mf1.Text = "No"
                        End If
                    Else
                        Mf1.Text = TempRs(j)
                    End If
                Next
                
            Next
            UpdateStatus FrmExtensiveSearch.PicProgress, 1, True
            

End Sub

