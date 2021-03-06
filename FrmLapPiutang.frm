VERSION 5.00
Begin VB.Form FrmLapPiutang 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".: Laporan Piutang"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12165
   Icon            =   "FrmLapPiutang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   12165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   ".: Pilih Laporan berdasarkan :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Jatuh Tempo Sebelum Tanggal :"
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sisa Piutang (Rp)"
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
         Left            =   600
         TabIndex        =   6
         Top             =   4560
         Width           =   4815
         Begin VB.Label LblTotPiutang 
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
            TabIndex        =   7
            Top             =   360
            Width           =   4455
         End
      End
      Begin VB.TextBox TxtNoPenjualan 
         Height          =   350
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2160
         Width           =   1600
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Semua Piutang ( Lunas dan Belum Lunas )"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   3375
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Semua Piutang Lunas"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   3375
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Semua Piutang Belum Lunas"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   3375
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Per Nota Transaksi"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   1800
         Width           =   3375
      End
      Begin VB.PictureBox cmdCari 
         Height          =   375
         Left            =   3480
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   8
         Top             =   2160
         Width           =   375
      End
      Begin VB.PictureBox GridPembayaran 
         Height          =   1935
         Left            =   600
         ScaleHeight     =   1875
         ScaleWidth      =   11115
         TabIndex        =   9
         Top             =   2640
         Width           =   11175
      End
      Begin VB.PictureBox DtTempo 
         Height          =   300
         Left            =   3000
         ScaleHeight     =   240
         ScaleWidth      =   1395
         TabIndex        =   22
         Top             =   1440
         Width           =   1455
      End
      Begin VB.PictureBox CmdTampil 
         Height          =   615
         Left            =   5880
         ScaleHeight     =   555
         ScaleWidth      =   1395
         TabIndex        =   24
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tanggal Tempo"
         Height          =   345
         Left            =   6360
         TabIndex        =   21
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label LblTglTempo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   7680
         TabIndex        =   20
         Top             =   2160
         Width           =   1125
      End
      Begin VB.Label LblTotBayar 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   10200
         TabIndex        =   19
         Top             =   2160
         Width           =   1485
      End
      Begin VB.Label LblKembali1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Pembayaran"
         Height          =   345
         Left            =   8880
         TabIndex        =   18
         Top             =   2160
         Width           =   1275
      End
      Begin VB.Label LblDibayar1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Penjualan"
         Height          =   345
         Left            =   8880
         TabIndex        =   17
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label LblJmlbeli 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   10200
         TabIndex        =   16
         Top             =   1800
         Width           =   1485
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Pelanggan"
         Height          =   345
         Left            =   4200
         TabIndex        =   15
         Top             =   1800
         Width           =   1065
      End
      Begin VB.Label LblNota 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No. Nota"
         Height          =   345
         Left            =   600
         TabIndex        =   14
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label LblNamaPelanggan 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   4200
         TabIndex        =   13
         Top             =   2160
         Width           =   2085
      End
      Begin VB.Label LblKodePelanggan 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   5280
         TabIndex        =   12
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label LblTglNota 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   7680
         TabIndex        =   11
         Top             =   1800
         Width           =   1125
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tanggal Nota"
         Height          =   345
         Left            =   6360
         TabIndex        =   10
         Top             =   1800
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmLapPiutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




