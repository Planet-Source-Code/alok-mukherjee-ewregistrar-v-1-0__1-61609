VERSION 5.00
Begin VB.Form EWabout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   " About Us"
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7650
   Icon            =   "EWabout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLogo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3855
      Left            =   60
      ScaleHeight     =   3855
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   450
      Width           =   315
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   4860
      Top             =   0
   End
   Begin VB.Image Image34 
      Height          =   1380
      Left            =   2040
      Picture         =   "EWabout.frx":08CA
      Top             =   1290
      Width           =   840
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Author : Alok Mukherjee"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   1320
      TabIndex        =   5
      Top             =   2820
      Width           =   2295
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   630
      Picture         =   "EWabout.frx":1CB1
      Top             =   3090
      Width           =   3660
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   600
      Left            =   690
      Top             =   3150
      Width           =   3660
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About Us"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   30
      Width           =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   390
      X2              =   4650
      Y1              =   2730
      Y2              =   2730
   End
   Begin VB.Line Line1 
      X1              =   390
      X2              =   4650
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Image Image33 
      Height          =   3765
      Left            =   4650
      Picture         =   "EWabout.frx":8F53
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2880
   End
   Begin VB.Image Image32 
      Height          =   1050
      Left            =   7560
      Picture         =   "EWabout.frx":BB56
      Stretch         =   -1  'True
      Top             =   3060
      Width           =   75
   End
   Begin VB.Image Image31 
      Height          =   1050
      Left            =   7560
      Picture         =   "EWabout.frx":BF58
      Stretch         =   -1  'True
      Top             =   2550
      Width           =   75
   End
   Begin VB.Image Image30 
      Height          =   1050
      Left            =   7560
      Picture         =   "EWabout.frx":C35A
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   75
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ActiveX EXE / DLL / OCX / TLB / OLB Registration Utility"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   555
      TabIndex        =   1
      Top             =   660
      Width           =   3825
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail : rch_castwood@sancharnet.in"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   870
      TabIndex        =   4
      Top             =   3780
      Width           =   3240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile : +919431101704, +919334426626"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   690
      TabIndex        =   3
      Top             =   4020
      Width           =   3600
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   690
      Left            =   570
      Top             =   510
      Width           =   3795
   End
   Begin VB.Image Image29 
      Height          =   135
      Left            =   6360
      Picture         =   "EWabout.frx":C75C
      Top             =   4260
      Width           =   1125
   End
   Begin VB.Image Image28 
      Height          =   135
      Left            =   5730
      Picture         =   "EWabout.frx":CFA2
      Top             =   4260
      Width           =   1125
   End
   Begin VB.Image Image24 
      Height          =   480
      Left            =   4650
      Picture         =   "EWabout.frx":D7E8
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image23 
      Height          =   480
      Left            =   3750
      Picture         =   "EWabout.frx":EEAA
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image22 
      Height          =   885
      Left            =   0
      Picture         =   "EWabout.frx":1056C
      Top             =   3240
      Width           =   90
   End
   Begin VB.Image Image21 
      Height          =   885
      Left            =   0
      Picture         =   "EWabout.frx":10A4A
      Top             =   2430
      Width           =   90
   End
   Begin VB.Image Image20 
      Height          =   885
      Left            =   0
      Picture         =   "EWabout.frx":10F28
      Top             =   1560
      Width           =   90
   End
   Begin VB.Image Image19 
      Height          =   1260
      Left            =   90
      Picture         =   "EWabout.frx":11406
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   5625
   End
   Begin VB.Image Image18 
      Height          =   1260
      Left            =   90
      Picture         =   "EWabout.frx":17D38
      Stretch         =   -1  'True
      Top             =   1740
      Width           =   5625
   End
   Begin VB.Image Image2 
      Height          =   690
      Left            =   0
      Picture         =   "EWabout.frx":1E66A
      Top             =   0
      Width           =   465
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   450
      Picture         =   "EWabout.frx":1F7EC
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   5550
      Picture         =   "EWabout.frx":20EAE
      Top             =   0
      Width           =   2100
   End
   Begin VB.Image Image5 
      Height          =   1050
      Left            =   7560
      Picture         =   "EWabout.frx":24370
      Stretch         =   -1  'True
      Top             =   450
      Width           =   75
   End
   Begin VB.Image Image6 
      Height          =   885
      Left            =   0
      Picture         =   "EWabout.frx":24772
      Top             =   690
      Width           =   90
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   1350
      Picture         =   "EWabout.frx":24C50
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image8 
      Height          =   285
      Left            =   0
      Picture         =   "EWabout.frx":26312
      Top             =   4110
      Width           =   195
   End
   Begin VB.Image Image9 
      Height          =   135
      Left            =   180
      Picture         =   "EWabout.frx":2664C
      Top             =   4260
      Width           =   1125
   End
   Begin VB.Image Image12 
      Height          =   135
      Left            =   1290
      Picture         =   "EWabout.frx":26E92
      Top             =   4260
      Width           =   1125
   End
   Begin VB.Image Image13 
      Height          =   135
      Left            =   2400
      Picture         =   "EWabout.frx":276D8
      Top             =   4260
      Width           =   1125
   End
   Begin VB.Image Image14 
      Height          =   135
      Left            =   3510
      Picture         =   "EWabout.frx":27F1E
      Top             =   4260
      Width           =   1125
   End
   Begin VB.Image Image15 
      Height          =   135
      Left            =   4620
      Picture         =   "EWabout.frx":28764
      Top             =   4260
      Width           =   1125
   End
   Begin VB.Image Image16 
      Height          =   480
      Left            =   2250
      Picture         =   "EWabout.frx":28FAA
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image17 
      Height          =   480
      Left            =   2850
      Picture         =   "EWabout.frx":2A66C
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image10 
      Height          =   330
      Left            =   7440
      Picture         =   "EWabout.frx":2BD2E
      Top             =   4050
      Width           =   210
   End
   Begin VB.Image Image11 
      Height          =   1260
      Left            =   90
      Picture         =   "EWabout.frx":2C138
      Stretch         =   -1  'True
      Top             =   480
      Width           =   5625
   End
   Begin VB.Image Image27 
      Height          =   1260
      Left            =   2010
      Picture         =   "EWabout.frx":32A6A
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   5625
   End
   Begin VB.Image Image26 
      Height          =   1260
      Left            =   2010
      Picture         =   "EWabout.frx":3939C
      Stretch         =   -1  'True
      Top             =   1740
      Width           =   5625
   End
   Begin VB.Image Image25 
      Height          =   1260
      Left            =   2010
      Picture         =   "EWabout.frx":3FCCE
      Stretch         =   -1  'True
      Top             =   480
      Width           =   5625
   End
End
Attribute VB_Name = "EWabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim cLogo As New cLogo
Private Sub Form_Load()
  ' These lines are required to match border pictures
  Image5.Left = Image5.Left + 15
  Image30.Left = Image30.Left + 15
  Image31.Left = Image31.Left + 15
  Image32.Left = Image32.Left + 15
  Image10.Top = Image10.Top + 15
  
  cLogo.DrawingObject = picLogo
  cLogo.Caption = "EWRegister Build " & App.Major & _
               "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  On Error GoTo 0
  cLogo.Draw
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Timer1.Enabled = False
  EWregMain.WindowState = vbNormal
  Set cLogo = Nothing
End Sub

Private Sub Image1_Click()
  Unload Me
End Sub

Private Sub Image11_Click()
  Unload Me
End Sub

Private Sub Image16_Click()
  Unload Me
End Sub

Private Sub Image17_Click()
  Unload Me
End Sub

Private Sub Image18_Click()
  Unload Me
End Sub

Private Sub Image19_Click()
  Unload Me
End Sub

Private Sub Image2_Click()
  Unload Me
End Sub

Private Sub Image23_Click()
  Unload Me
End Sub

Private Sub Image24_Click()
  Unload Me
End Sub

Private Sub Image3_Click()
  Unload Me
End Sub

Private Sub Image33_Click()
  Unload Me
End Sub

Private Sub Image4_Click()
  Unload Me
End Sub

Private Sub Image7_Click()
  Unload Me
End Sub

Private Sub Label1_Click()
  Unload Me
End Sub

Private Sub Label2_Click()
  Unload Me
End Sub

Private Sub Label3_Click()
  Unload Me
End Sub

Private Sub Label4_Click()
  Unload Me
End Sub

Private Sub Label5_Click()
  Unload Me
End Sub

Private Sub Label6_Click()
  Unload Me
End Sub

Private Sub picLogo_Click()
  Unload Me
End Sub

Private Sub Timer1_Timer()
  Unload Me
End Sub

