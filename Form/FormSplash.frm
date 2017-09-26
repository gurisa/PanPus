VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FormSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Splashing Perpustakaan"
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBarSplash 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer TimerSplash 
      Left            =   6240
      Top             =   1320
   End
   Begin VB.Label LabelPersen 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label LabelPeriksa 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Sedang Memeriksa Komponen File"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   5175
   End
   Begin VB.Image ImageLogo 
      Height          =   1395
      Left            =   120
      Picture         =   "FormSplash.frx":0CCA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label LabelPerpustakaan 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "PANDA PUSTAKA"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
End
Attribute VB_Name = "FormSplash"
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

Option Explicit
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crColor As Long, ByVal nAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private Sub Form_Load()
If App.PrevInstance = True Then
    MsgBox "Jika Panda Pustaka Tidak Berjalan" & vbNewLine & "Kill App Dengan Name Proccess Perpustakaan.exe Di Task Manager Secara Manual" & vbNewLine & "Lalu Jalankan Ulang", vbExclamation + vbOKOnly, "Panda Pustaka Sedang Di Jalankan"
    End
End If

TimerSplash.Interval = 50
TimerSplash.Enabled = True

Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
Call SetLayeredWindowAttributes(Me.hwnd, RGB(255, 255, 255), 128, LWA_ALPHA Or LWA_COLORKEY)
End Sub

Private Sub TimerSplash_Timer()
ProgressBarSplash.Value = Val(ProgressBarSplash.Value + 1)
LabelPersen.Caption = Val(ProgressBarSplash.Value) & " %"

If ProgressBarSplash.Value = 10 Then
    LabelPeriksa.Caption = "Sedang Memeriksa Komponen File"
ElseIf ProgressBarSplash.Value = 20 Then
    If Dir$(App.Path & "\Perpustakaan.exe") <> "" Then
        LabelPeriksa.Caption = "Program Inti Siap Di Akses"
    Else
        LabelPeriksa.Caption = "Program Inti Tidak Siap Di Akses"
        LabelPeriksa.ForeColor = &HFF&
        MsgBox "Silahkan Jalankan Aplikasi Dalam Kondisi Utuh", vbInformation, "Jalankan Aplikasi Secara Utuh"
    End If
ElseIf ProgressBarSplash.Value = 30 Then
    If Dir$(App.Path & "\Crystl32.ocx") <> "" Then
        LabelPeriksa.Caption = "File Crystl32.ocx Tersedia"
    Else
        LabelPeriksa.Caption = "File Crystl32.ocx Tidak Tersedia"
        LabelPeriksa.ForeColor = &HFF&
        MsgBox "Silahkan Periksa Struktur Kelengkapan Aplikasi", vbInformation, "Hubungi Pengembang Aplikasi"
    End If
ElseIf ProgressBarSplash.Value = 40 Then
    If Dir$(App.Path & "\MSCOMCT2.ocx") <> "" Then
        LabelPeriksa.Caption = "File MSCOMCT2.ocx Tersedia"
    Else
        LabelPeriksa.Caption = "File MSCOMCT2.ocx Tidak Tersedia"
        LabelPeriksa.ForeColor = &HFF&
        MsgBox "Silahkan Periksa Struktur Kelengkapan Aplikasi", vbInformation, "Hubungi Pengembang Aplikasi"
    End If
ElseIf ProgressBarSplash.Value = 50 Then
    If Dir$(App.Path & "\MSCOMCTL.ocx") <> "" Then
        LabelPeriksa.Caption = "File MSCOMCTL.ocx Tersedia"
    Else
        LabelPeriksa.Caption = "File MSCOMCTL.ocx Tidak Tersedia"
        LabelPeriksa.ForeColor = &HFF&
        MsgBox "Silahkan Periksa Struktur Kelengkapan Aplikasi", vbInformation, "Hubungi Pengembang Aplikasi"
    End If
ElseIf ProgressBarSplash.Value = 60 Then
    If Dir$(App.Path & "\Panel.panpus") <> "" Then
        LabelPeriksa.Caption = "Konfigurasi Tersedia"
    Else
        LabelPeriksa.Caption = "Konfigurasi Tidak Tersedia"
        LabelPeriksa.ForeColor = &HFF&
        MsgBox "Silahkan Periksa Tables Dalam Database Atau Hubungi Pengembang Aplikasi", vbInformation, "Hubungi Pengembang Aplikasi"
    End If
ElseIf ProgressBarSplash.Value = 70 Then
    If Dir$(App.Path & "\Log.txt") <> "" Then
        LabelPeriksa.Caption = "File Log.txt Tersedia"
    Else
        LabelPeriksa.Caption = "File Log.txt Tidak Tersedia"
        LabelPeriksa.ForeColor = &HFF&
        MsgBox "Silahkan Cari File Log Atau Hubungi Pengembang Aplikasi", vbInformation, "Hubungi Pengembang Aplikasi"
    End If
ElseIf ProgressBarSplash.Value = 80 Then
    If Dir$(App.Path & "\Panduan\Panduan.pdf") <> "" Then
        LabelPeriksa.Caption = "File Panduan Tersedia"
    Else
        LabelPeriksa.Caption = "File Panduan Tidak Tersedia"
        LabelPeriksa.ForeColor = &HFF&
        MsgBox "Silahkan Cari File Panduan Atau Hubungi Pengembang Program", vbInformation, "Hubungi Pengembang Aplikasi"
    End If
ElseIf ProgressBarSplash.Value = 90 Then
    LabelPeriksa.Caption = "Mempersiapkan Aplikasi.."
ElseIf ProgressBarSplash.Value = 100 Then
    Unload Me
    FormLogin.Show
End If
End Sub
