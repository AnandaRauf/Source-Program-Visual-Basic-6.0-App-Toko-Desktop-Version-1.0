VERSION 5.00
Begin VB.Form FrmHutang 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".: Pembayaran Hutang :."
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11385
   Icon            =   "FrmHutang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox DtBayar 
      Height          =   375
      Left            =   8280
      ScaleHeight     =   315
      ScaleWidth      =   1635
      TabIndex        =   31
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox TxtJumlah 
      Height          =   350
      Left            =   8280
      MaxLength       =   10
      TabIndex        =   30
      Top             =   4200
      Width           =   1600
   End
   Begin VB.TextBox TxtNoRef 
      Height          =   350
      Left            =   8280
      MaxLength       =   30
      TabIndex        =   29
      Top             =   3840
      Width           =   1605
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sisa Hutang yang akan dibayar (Rp)"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   4815
      Begin VB.Label LblSisaHutang 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   33.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sisa Hutang (Rp)"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   4815
      Begin VB.Label LblTotHutang 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   33.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.TextBox TxtNota 
      Height          =   350
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   1
      Top             =   120
      Width           =   1600
   End
   Begin VB.TextBox TxtNoPembelian 
      Height          =   350
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   0
      Top             =   480
      Width           =   1600
   End
   Begin VB.PictureBox cmdCari 
      Height          =   495
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
   Begin VB.PictureBox CmdBaru 
      Height          =   375
      Left            =   6600
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   4920
      Width           =   1215
   End
   Begin VB.PictureBox CmdSimpan 
      Height          =   375
      Left            =   7920
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   8
      Top             =   4920
      Width           =   1215
   End
   Begin VB.PictureBox cmdKeluar 
      Height          =   375
      Left            =   9240
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   9
      Top             =   4920
      Width           =   1215
   End
   Begin VB.PictureBox GridPembayaran 
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   11115
      TabIndex        =   10
      ToolTipText     =   "KLIK 2x DISINI UNTUK PILIH DATA"
      Top             =   960
      Width           =   11175
   End
   Begin VB.PictureBox CmbVia 
      Height          =   360
      Left            =   8280
      ScaleHeight     =   300
      ScaleWidth      =   1635
      TabIndex        =   26
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " [Enter] - untuk melanjutkan transaksi"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   240
      Left            =   7080
      TabIndex        =   32
      Top             =   5400
      Width           =   3045
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jumlah :"
      Height          =   345
      Left            =   6840
      TabIndex        =   28
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No. Ref :"
      Height          =   345
      Left            =   6840
      TabIndex        =   27
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pembayaran Via :"
      Height          =   345
      Left            =   6840
      TabIndex        =   25
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal Bayar :"
      Height          =   345
      Left            =   6840
      TabIndex        =   24
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tanggal Tempo"
      Height          =   345
      Left            =   6000
      TabIndex        =   23
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label LblTglTempo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   7200
      TabIndex        =   22
      Top             =   480
      Width           =   1125
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   6360
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   4335
   End
   Begin VB.Label LblTotBayar 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   9720
      TabIndex        =   21
      Top             =   480
      Width           =   1485
   End
   Begin VB.Label LblKembali1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pembayaran"
      Height          =   345
      Left            =   8400
      TabIndex        =   20
      Top             =   480
      Width           =   1275
   End
   Begin VB.Label LblDibayar1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Pembelian"
      Height          =   345
      Left            =   8400
      TabIndex        =   19
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label LblJmlbeli 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   9720
      TabIndex        =   18
      Top             =   120
      Width           =   1485
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pemasok"
      Height          =   345
      Left            =   3720
      TabIndex        =   17
      Top             =   120
      Width           =   1065
   End
   Begin VB.Label LblNota 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No. Nota"
      Height          =   345
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LblNamaPelanggan 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3720
      TabIndex        =   15
      Top             =   480
      Width           =   2085
   End
   Begin VB.Label LblKodePelanggan 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4800
      TabIndex        =   14
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label LblTglNota 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   7200
      TabIndex        =   13
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tanggal Nota"
      Height          =   345
      Left            =   6000
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No. Pembelian"
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "FrmHutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdCari_UnknownEvent_10() '5F7590
