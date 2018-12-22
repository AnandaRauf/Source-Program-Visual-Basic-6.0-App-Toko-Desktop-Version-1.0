VERSION 5.00
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  :: Login POS"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPwd 
      Height          =   400
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   720
      Width           =   3540
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   3540
   End
   Begin VB.PictureBox cmdLogin 
      Height          =   375
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.PictureBox cmdTutup 
      Height          =   375
      Left            =   4080
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.PictureBox cmbLevel 
      Height          =   405
      Left            =   1800
      ScaleHeight     =   345
      ScaleWidth      =   3480
      TabIndex        =   3
      Top             =   1200
      Width           =   3540
   End
   Begin VB.Label Label5 
      Caption         =   "CONTOH: User :admin, Pwd: admin , Status: Admin"
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   210
      TabIndex        =   4
      Top             =   720
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   210
      TabIndex        =   0
      Top             =   240
      Width           =   1035
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdLogin_UnknownEvent_10() '5B2190
