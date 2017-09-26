VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FormPopUp 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6510
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormPopUp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FramePopUp 
      BackColor       =   &H00000000&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton CmdOk 
         Caption         =   "OK"
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.Timer TimerPopUp 
         Left            =   4560
         Top             =   600
      End
      Begin MSComctlLib.ProgressBar ProgressBarPopUp 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label LabelJudulNama 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Selamat Datang Kembali"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label LabelNama 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label LabelJudulJumlah 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Terdapat"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.Label LabelJumlah 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label LabelHitungMundur 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Hitung Mundur"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
   End
End
Attribute VB_Name = "FormPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''********************************************************************''
''                                                                    ''
'' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\' ''
'' \\ Project Name       :   Panda Pustaka                        \\' ''
'' \\ Project Version    :   1.0                                  \\' ''
'' \\ Project Author     :   Raka Suryaardi Widjaja               \\' ''
'' \\ Project Home Page  :   www.Gurisa.Com                       \\' ''
'' \\ Project License    :   All Right Reserved Gurisa © 2015     \\' ''
'' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\' ''
''                                                                    ''
''********************************************************************''

Private Sub CmdOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
TimerPopUp.Enabled = True
TimerPopUp.Interval = 75

Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height / 1 - Me.Height / 1
AlwaysOnTopForm hwnd
End Sub

Private Sub TimerPopUp_Timer()
ProgressBarPopUp.Value = Val(ProgressBarPopUp.Value + 1)
If ProgressBarPopUp.Value = 10 Then
    LabelHitungMundur.Caption = "Otomatis Tertutup Dalam 10"
ElseIf ProgressBarPopUp.Value = 20 Then
    LabelHitungMundur.Caption = "Otomatis Tertutup Dalam 9"
ElseIf ProgressBarPopUp.Value = 30 Then
    LabelHitungMundur.Caption = "Otomatis Tertutup Dalam 8"
ElseIf ProgressBarPopUp.Value = 40 Then
    LabelHitungMundur.Caption = "Otomatis Tertutup Dalam 7"
ElseIf ProgressBarPopUp.Value = 50 Then
    LabelHitungMundur.Caption = "Otomatis Tertutup Dalam 6"
ElseIf ProgressBarPopUp.Value = 60 Then
    LabelHitungMundur.Caption = "Otomatis Tertutup Dalam 5"
ElseIf ProgressBarPopUp.Value = 70 Then
    LabelHitungMundur.Caption = "Otomatis Tertutup Dalam 4"
ElseIf ProgressBarPopUp.Value = 80 Then
    LabelHitungMundur.Caption = "Otomatis Tertutup Dalam 3"
ElseIf ProgressBarPopUp.Value = 85 Then
    LabelHitungMundur.Caption = "Otomatis Tertutup Dalam 2"
ElseIf ProgressBarPopUp.Value = 90 Then
    LabelHitungMundur.Caption = "Otomatis Tertutup Dalam 1"
ElseIf ProgressBarPopUp.Value = 95 Then
    LabelHitungMundur.Caption = "Menutup Pemberitahuan"
ElseIf ProgressBarPopUp.Value = 100 Then
    Unload Me
End If
End Sub
