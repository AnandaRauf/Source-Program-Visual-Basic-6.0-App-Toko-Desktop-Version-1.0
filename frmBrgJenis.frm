VERSION 5.00
Begin VB.Form frmBrgJenis 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   ::  DATA JENIS"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   9255
      Begin VB.TextBox txtKode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtNmGolongan 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1920
         TabIndex        =   1
         Top             =   960
         Width           =   6840
      End
      Begin VB.TextBox txtNama 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1920
         MaxLength       =   60
         TabIndex        =   3
         Top             =   1920
         Width           =   6855
      End
      Begin VB.PictureBox cmbGolongan 
         Height          =   405
         Left            =   1920
         ScaleHeight     =   345
         ScaleWidth      =   1515
         TabIndex        =   0
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Jenis"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   360
         TabIndex        =   12
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pilih Golongan"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Jenis"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   360
         TabIndex        =   8
         Top             =   2040
         Width           =   1185
      End
   End
   Begin VB.PictureBox GridJenis 
      Height          =   3735
      Left            =   120
      ScaleHeight     =   3675
      ScaleWidth      =   9195
      TabIndex        =   7
      ToolTipText     =   "KLIK 2x DISINI UNTUK UBAH / HAPUS DATA..!"
      Top             =   3600
      Width           =   9255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4200
      TabIndex        =   10
      Top             =   2760
      Width           =   5175
      Begin VB.PictureBox CmdBaru 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox CmdSimpan 
         Height          =   375
         Left            =   1320
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox cmdKeluar 
         Height          =   375
         Left            =   3720
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox CmdHapus 
         Height          =   375
         Left            =   2520
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   " [ Urutkan Data ] "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   3975
      Begin VB.OptionButton rbUrutan 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nama Produk"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton rbUrutan 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmBrgJenis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CmdSimpan_UnknownEvent_10() '5736B0
