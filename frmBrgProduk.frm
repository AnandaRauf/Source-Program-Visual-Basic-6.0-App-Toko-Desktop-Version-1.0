VERSION 5.00
Begin VB.Form frmBrgProduk 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   ::  DATA PRODUK"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   9255
      Begin VB.TextBox txtJenis 
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
         Left            =   3240
         TabIndex        =   3
         Top             =   840
         Width           =   5745
      End
      Begin VB.TextBox txtGolongan 
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
         Left            =   3240
         TabIndex        =   1
         Top             =   360
         Width           =   5745
      End
      Begin VB.TextBox txtKode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1800
         MaxLength       =   11
         TabIndex        =   4
         Top             =   1440
         Width           =   2175
      End
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
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   5
         Top             =   1920
         Width           =   7200
      End
      Begin VB.PictureBox cmbJenis 
         Height          =   405
         Left            =   1800
         ScaleHeight     =   345
         ScaleWidth      =   1395
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.PictureBox cmbGolongan 
         Height          =   405
         Left            =   1800
         ScaleHeight     =   345
         ScaleWidth      =   1395
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pilih Jenis"
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
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   1020
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
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Produk"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Produk"
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
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Width           =   1380
      End
   End
   Begin VB.PictureBox GridProduk 
      Height          =   4815
      Left            =   120
      ScaleHeight     =   4755
      ScaleWidth      =   9195
      TabIndex        =   8
      ToolTipText     =   "KLIK 2x DISINI UNTUK UBAH / HAPUS DATA..!"
      Top             =   3360
      Width           =   9255
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
      TabIndex        =   14
      Top             =   2520
      Width           =   3975
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
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
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
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4200
      TabIndex        =   15
      Top             =   2520
      Width           =   5175
      Begin VB.PictureBox CmdBaru 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox CmdSimpan 
         Height          =   375
         Left            =   1320
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox cmdKeluar 
         Height          =   375
         Left            =   3720
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox CmdHapus 
         Height          =   375
         Left            =   2520
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmBrgProduk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub txtKode_KeyPress(KeyAscii As Integer) '5804B0
