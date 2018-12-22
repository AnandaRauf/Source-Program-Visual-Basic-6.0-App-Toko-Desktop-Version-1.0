VERSION 5.00
Begin VB.Form frmBrgBarang 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "   ::  DATA BARANG"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   14055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   14055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCari 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9000
      TabIndex        =   13
      Top             =   3840
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   120
      TabIndex        =   14
      Top             =   -120
      Width           =   13815
      Begin VB.TextBox txtHrgJual 
         Alignment       =   1  'Right Justify
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   10
         Text            =   "0"
         Top             =   3000
         Width           =   1920
      End
      Begin VB.TextBox txtKode 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2040
         MaxLength       =   60
         TabIndex        =   6
         Top             =   1560
         Width           =   4455
      End
      Begin VB.TextBox txtJenis 
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
         Left            =   3960
         TabIndex        =   3
         Top             =   720
         Width           =   6825
      End
      Begin VB.TextBox txtStok 
         Alignment       =   1  'Right Justify
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
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   9
         Text            =   "0"
         Top             =   2520
         Width           =   1320
      End
      Begin VB.TextBox txtGolongan 
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
         Left            =   3960
         TabIndex        =   1
         Top             =   360
         Width           =   6825
      End
      Begin VB.TextBox txtHrgBeli 
         Alignment       =   1  'Right Justify
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
         Left            =   2025
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "0"
         Top             =   2520
         Width           =   1920
      End
      Begin VB.TextBox txtProduk 
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
         Left            =   3960
         TabIndex        =   5
         Top             =   1080
         Width           =   6825
      End
      Begin VB.TextBox txtNama 
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
         Left            =   2025
         TabIndex        =   7
         Top             =   2040
         Width           =   11460
      End
      Begin VB.PictureBox cmbJenis 
         Height          =   405
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   1875
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
      Begin VB.PictureBox cmbGolongan 
         Height          =   405
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   1875
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin VB.PictureBox cmbSatuan 
         Height          =   405
         Left            =   5160
         ScaleHeight     =   345
         ScaleWidth      =   1275
         TabIndex        =   11
         Top             =   3000
         Width           =   1335
      End
      Begin VB.PictureBox cmbProduk 
         Height          =   405
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   1875
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Barcode, jika tidak diisi akan diset kode automatis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   6600
         TabIndex        =   25
         Top             =   1680
         Width           =   4590
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hrg Jual (Rp)"
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
         TabIndex        =   24
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
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
         TabIndex        =   23
         Top             =   1680
         Width           =   1365
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Produk"
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
         TabIndex        =   22
         Top             =   840
         Width           =   1365
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stok "
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
         Left            =   4560
         TabIndex        =   21
         Top             =   2640
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Golongan"
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
         TabIndex        =   20
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hrg Beli (Rp)"
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
         TabIndex        =   18
         Top             =   2640
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan "
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
         Left            =   4320
         TabIndex        =   17
         Top             =   3120
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
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
         TabIndex        =   16
         Top             =   2160
         Width           =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produk Barang"
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
         TabIndex        =   15
         Top             =   1200
         Width           =   1560
      End
   End
   Begin VB.PictureBox GridBarang 
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   13755
      TabIndex        =   12
      ToolTipText     =   "KLIK 2x UNTUK UBAH / HAPUS DATA..!"
      Top             =   4320
      Width           =   13815
   End
   Begin VB.PictureBox CmdBaru 
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   26
      Top             =   3840
      Width           =   1215
   End
   Begin VB.PictureBox CmdSimpan 
      Height          =   375
      Left            =   1560
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   27
      Top             =   3840
      Width           =   1215
   End
   Begin VB.PictureBox cmdKeluar 
      Height          =   375
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   28
      Top             =   3840
      Width           =   1215
   End
   Begin VB.PictureBox CmdHapus 
      Height          =   375
      Left            =   2880
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   29
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Barang"
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
      Left            =   7680
      TabIndex        =   19
      Top             =   3960
      Width           =   1245
   End
End
Attribute VB_Name = "frmBrgBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmbJenis_UnknownEvent_C() '585DF0
