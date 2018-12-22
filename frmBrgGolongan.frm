VERSION 5.00
Begin VB.Form frmBrgGolongan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   ::  DATA GOLONGAN"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNama 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2280
      MaxLength       =   60
      TabIndex        =   1
      Top             =   1440
      Width           =   5535
   End
   Begin VB.TextBox txtKode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.PictureBox GridGolongan 
      Height          =   4695
      Left            =   120
      ScaleHeight     =   4635
      ScaleWidth      =   7995
      TabIndex        =   3
      ToolTipText     =   "KLIK 2x DISINI UNTUK UBAH / HAPUS DATA..!"
      Top             =   2640
      Width           =   8055
   End
   Begin VB.PictureBox CmdBaru 
      Height          =   375
      Left            =   2760
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.PictureBox CmdSimpan 
      Height          =   375
      Left            =   4080
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.PictureBox cmdKeluar 
      Height          =   375
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.PictureBox CmdHapus 
      Height          =   375
      Left            =   5400
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   15
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   7695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Golongan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "*) Harus 3 digit"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   555
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Golongan"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmBrgGolongan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub txtKode_KeyPress(KeyAscii As Integer) '5703C0
