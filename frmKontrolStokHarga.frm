VERSION 5.00
Begin VB.Form frmKontrolStokHarga 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   ::  Kontrol Stok dan Harga"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   14775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "&Keluar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13440
      TabIndex        =   4
      Top             =   8640
      Width           =   1305
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   " [ Kriteria Pencarian ] "
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
      TabIndex        =   0
      Top             =   120
      Width           =   14535
      Begin VB.TextBox txtCari 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5760
         TabIndex        =   5
         Top             =   240
         Width           =   8655
      End
      Begin VB.OptionButton rbKategori 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nama Barang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3840
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton rbKategori 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kode Barang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton rbKategori 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Jenis Barang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.PictureBox GridBarang 
      Height          =   7575
      Left            =   120
      ScaleHeight     =   7515
      ScaleWidth      =   14475
      TabIndex        =   6
      ToolTipText     =   "KLIK 2x DISINI UNTUK PILIH DATA"
      Top             =   960
      Width           =   14535
   End
End
Attribute VB_Name = "frmKontrolStokHarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



