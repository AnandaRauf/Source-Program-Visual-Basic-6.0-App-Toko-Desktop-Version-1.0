VERSION 5.00
Begin VB.Form FrmGantiPassword 
   BackColor       =   &H001FD383&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ganti Password"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5040
   Icon            =   "FrmGantiPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtPswLama 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   13
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox TxtPswBaru 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   13
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1560
      Width           =   2775
   End
   Begin VB.PictureBox CmdUbah 
      Height          =   375
      Left            =   1560
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.PictureBox CmdBatal 
      Height          =   375
      Left            =   2880
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan perubahan password Anda pada kolom yang tersedia di bawah ini"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   120
      Width           =   4095
   End
   Begin VB.Image ImgHeader 
      Height          =   855
      Left            =   0
      Picture         =   "FrmGantiPassword.frx":1CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password Lama:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password Baru :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000B&
      BorderWidth     =   2
      X1              =   120
      X2              =   4320
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000B&
      BorderWidth     =   2
      X1              =   120
      X2              =   4320
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "FrmGantiPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




