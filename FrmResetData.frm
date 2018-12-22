VERSION 5.00
Begin VB.Form FrmResetData 
   BackColor       =   &H001FD383&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".: Reset / Hapus Data"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7560
   Icon            =   "FrmResetData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option7 
      BackColor       =   &H001FD383&
      Caption         =   "Semua Data Piutang dan Pembayaran Piutang"
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   4800
      Width           =   4335
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H001FD383&
      Caption         =   "Semua Data Hutang dan Pembayaran Hutang"
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   4200
      Width           =   4455
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H001FD383&
      Caption         =   "Semua Data Penjualan dan Retur Penjualan"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   3600
      Width           =   3495
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H001FD383&
      Caption         =   "Semua Data Pembelian dan Retur Pembelian"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   3495
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H001FD383&
      Caption         =   "Semua Data Supplier"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H001FD383&
      Caption         =   "Semua Data Pelanggan"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H001FD383&
      Caption         =   "Semua Data Barang (Golongan, Jenis, Produk, Satuan)"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   4455
   End
   Begin VB.PictureBox CmdProses 
      Height          =   615
      Left            =   4920
      ScaleHeight     =   555
      ScaleWidth      =   2235
      TabIndex        =   8
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmResetData.frx":08CA
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   960
      TabIndex        =   7
      Top             =   120
      Width           =   4815
   End
   Begin VB.Image ImgHeader 
      Height          =   855
      Left            =   0
      Picture         =   "FrmResetData.frx":0961
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "FrmResetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




