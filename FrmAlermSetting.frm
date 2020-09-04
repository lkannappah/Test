VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmAlermSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Appointment Reminder"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "FrmAlermSetting.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4665
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   5280
      Picture         =   "FrmAlermSetting.frx":1272
      ScaleHeight     =   525
      ScaleWidth      =   645
      TabIndex        =   16
      Top             =   120
      Width           =   645
   End
   Begin VB.Frame Frame2 
      Caption         =   "Alarm Sound :"
      Height          =   2520
      Left            =   75
      TabIndex        =   7
      Top             =   690
      Width           =   6105
      Begin VB.CommandButton Command5 
         Caption         =   "&OK"
         Height          =   360
         Left            =   4770
         TabIndex        =   14
         Top             =   2070
         Width           =   1200
      End
      Begin VB.ListBox List1 
         Height          =   840
         Left            =   105
         TabIndex        =   8
         Top             =   240
         Width           =   5880
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Give Warning Until Stop The Alarm."
         Height          =   270
         Left            =   3150
         TabIndex        =   20
         Top             =   1125
         Width           =   2835
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Give Alarm Once for 10 Seconds."
         Height          =   195
         Left            =   165
         TabIndex        =   19
         Top             =   1155
         Width           =   2745
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   5220
         Picture         =   "FrmAlermSetting.frx":16B4
         ScaleHeight     =   555
         ScaleWidth      =   525
         TabIndex        =   15
         Top             =   1440
         Width           =   525
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Remove"
         Height          =   360
         Left            =   3513
         TabIndex        =   13
         Top             =   2070
         Width           =   1245
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Import"
         Height          =   360
         Left            =   2392
         TabIndex        =   12
         Top             =   2070
         Width           =   1110
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Stop"
         Height          =   360
         Left            =   1271
         TabIndex        =   11
         Top             =   2070
         Width           =   1110
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Play"
         Height          =   360
         Left            =   150
         TabIndex        =   10
         Top             =   2070
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   165
         TabIndex        =   17
         Top             =   1740
         Width           =   45
      End
      Begin VB.Label Label3 
         Caption         =   "Available Wave (*.WAV) File :"
         Height          =   225
         Left            =   165
         TabIndex        =   9
         Top             =   1425
         Width           =   2475
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Alerm Setting :"
      Height          =   1020
      Left            =   5910
      TabIndex        =   6
      Top             =   780
      Visible         =   0   'False
      Width           =   390
      Begin VB.ListBox List2 
         Height          =   450
         Left            =   5130
         TabIndex        =   18
         Top             =   540
         Width           =   540
      End
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   255
      Left            =   3780
      TabIndex        =   4
      Top             =   150
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   450
      _Version        =   393216
      Value           =   1
      Max             =   60
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   3150
      TabIndex        =   3
      Text            =   "1"
      Top             =   120
      Width           =   900
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Don't Remind Appointment."
      Height          =   195
      Left            =   105
      TabIndex        =   1
      Top             =   450
      Width           =   2310
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Remind Appointment."
      Height          =   270
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   1890
   End
   Begin VB.Label Label2 
      Caption         =   "Minutes."
      Height          =   255
      Left            =   4140
      TabIndex        =   5
      Top             =   165
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Before :"
      Height          =   210
      Left            =   2550
      TabIndex        =   2
      Top             =   135
      Width           =   630
   End
End
Attribute VB_Name = "FrmAlermSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Fs As New FileSystemObject
Dim f1 As File

Private Sub Command1_Click()
If List1.ListIndex >= 0 Then
     If CanPlayWaves Then PlayWave (List1.List(List1.ListIndex))
Else
    MsgBox "Select a Wave File from the List", vbInformation, "Select File"
End If
    
End Sub

Private Sub Command2_Click()
    StopPlayingWave
End Sub

Private Sub Command3_Click()
    CD1.DialogTitle = "Select Wave file (*.WAV)"
    CD1.Filter = "Wave|*.WAV"
    CD1.ShowOpen
    If CD1.FileName <> "" Then
        List1.AddItem CD1.FileName
    End If
End Sub

Private Sub Command4_Click()
    If List1.ListCount > 0 Then
        List1.RemoveItem List1.ListIndex
        Label4.Caption = ""
    End If
End Sub

Private Sub Command5_Click()



Dim XX(100) As String
Dim ind As Integer
If List1.ListCount > 0 Then
    Dim Fol As Folder
    Dim AllFile As Files
    Dim ThisFile As File
    Dim Str As String
        Set Fol = Fs.GetFolder(App.Path & "\data\sound\")
        Set AllFile = Fol.Files
        If Fol.Files.Count > 0 Then
            For Each ThisFile In AllFile
                Str = UCase(Right(ThisFile.Name, 3))
                If Str = "WAV" Then
                    XX(ind) = ThisFile.Path
                    ind = ind + 1
                End If
            Next
        End If
        Dim i As Integer
        Dim j As Integer
        Dim Check As Boolean
        For i = 0 To List1.ListCount - 1
            Check = False
            For j = 0 To 99
                If UCase(XX(j)) = UCase(List1.List(i)) Then
                    Check = True
                    Exit For
                End If
            Next
            If Check = False Then
                Fs.CopyFile List1.List(i), App.Path & "\data\sound\"
            End If
                
            
        Next

Else
    MsgBox "At least One Wav File Required To Continue", vbExclamation
    Exit Sub
End If
If List1.ListIndex >= 0 Then
    SaveSetting "MyAppPP", "Alarm", "Alarmsound", List1.List(List1.ListIndex)
Else
    MsgBox "Select a wave file name for Alarm"
    Exit Sub
End If
If Option1.Value = True Then
    SaveSetting "MyAppPP", "Alarm", "AlarmONOFF", "1"
Else
    SaveSetting "MyAppPP", "Alarm", "AlarmONOFF", "0"
End If
    SaveSetting "MyAppPP", "Alarm", "AlarmBeforeTime", Val(Text1.Text)
If Option3.Value = True Then
    SaveSetting "MyAppPP", "Alarm", "Continue", "1"
Else
    SaveSetting "MyAppPP", "Alarm", "Continue", "0"
End If
    

Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Option1.Value = True
    Option3.Value = True
    Text1.Text = 1
    
    List1.Clear
    List2.Clear
    Dim Fol As Folder
    Dim AllFile As Files
    Dim ThisFile As File
    Dim Str As String
        Set Fol = Fs.GetFolder(App.Path & "\data\sound\")
        Set AllFile = Fol.Files
        If Fol.Files.Count > 0 Then
            For Each ThisFile In AllFile
                Str = UCase(Right(ThisFile.Name, 3))
                If Str = "WAV" Then
                    List1.AddItem ThisFile.Path
                    List2.AddItem ThisFile.Path
                End If
            Next
        End If
    If List1.ListCount > 0 Then
        If Fs.FileExists(List1.List(0)) = True Then
            Set f1 = Fs.GetFile(List1.List(0))
            Label4.Caption = "Current Alerm Sound : " & f1.Name
            List1.ListIndex = 0
        End If
    End If
    
    Dim i As Integer
    If List1.ListCount > 0 Then
        For i = 0 To List1.ListCount - 1
            If UCase(GetSetting("MyAppPP", "Alarm", "Alarmsound")) = UCase(List1.List(List1.ListIndex)) Then
                List1.ListIndex = i
                Exit For
            End If
        Next
    End If
    If GetSetting("MyAppPP", "Alarm", "AlarmONOFF") = "1" Then
        Option1.Value = True
    Else
        Option2.Value = True
    End If
    Text1.Text = GetSetting("MyAppPP", "Alarm", "AlarmBeforeTime", 1)
    UpDown1.Value = Val(Text1.Text)
    If GetSetting("MyAppPP", "Alarm", "Continue") = "1" Then
        Option3.Value = True
    Else
        Option4.Value = True
    End If
    
End Sub

Private Sub List1_Click()
        Set f1 = Fs.GetFile(List1.List(List1.ListIndex))
        Label4.Caption = "Current Alerm Sound : " & f1.Name
End Sub

Private Sub UpDown1_Change()
Text1.Text = UpDown1.Value
End Sub
