VERSION 5.00
Begin VB.Form frmSatuanBarang 
   BackColor       =   &H001FD383&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   ::  SATUAN BARANG"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   6
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox txtKeterangan 
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
      Top             =   720
      Width           =   4935
   End
   Begin VB.TextBox txtSatuan 
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
      MaxLength       =   30
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.PictureBox GridSatuan 
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   7395
      TabIndex        =   2
      ToolTipText     =   "KLIK 2x DISINI UNTUK UBAH / HAPUS DATA..!"
      Top             =   2040
      Width           =   7455
   End
   Begin VB.PictureBox CmdBaru 
      Height          =   375
      Left            =   2280
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.PictureBox CmdSimpan 
      Height          =   375
      Left            =   3600
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.PictureBox cmdKeluar 
      Height          =   375
      Left            =   6240
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.PictureBox CmdHapus 
      Height          =   375
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   15
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   7095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Satuan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmSatuanBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CmdBaru_UnknownEvent_10() '577910
