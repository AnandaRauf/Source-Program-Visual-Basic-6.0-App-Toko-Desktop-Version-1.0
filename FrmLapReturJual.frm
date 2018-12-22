VERSION 5.00
Begin VB.Form FrmLapReturJual 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".: Laporan Retur Penjulan Per Nota"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11775
   Icon            =   "FrmLapReturJual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " [ OPSI CETAK ] "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   9240
      TabIndex        =   0
      Top             =   600
      Width           =   1935
      Begin VB.PictureBox CmdTampil 
         Height          =   495
         Left            =   240
         ScaleHeight     =   435
         ScaleWidth      =   1275
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.PictureBox CmdCetak 
         Height          =   495
         Left            =   240
         ScaleHeight     =   435
         ScaleWidth      =   1275
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.PictureBox gridTransaksi 
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   8955
      TabIndex        =   1
      Top             =   480
      Width           =   9015
   End
   Begin VB.PictureBox gridTransDetail 
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   11475
      TabIndex        =   2
      Top             =   4440
      Width           =   11535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Barang"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Daftar Retur Penjualan (List Nota)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "FrmLapReturJual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




