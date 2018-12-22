VERSION 5.00
Begin VB.Form frmLapProdukJenis 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  :: Daftar Produk dan Jenis Barang"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   12390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " [ PRODUK BARANG ] "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   7815
      Begin VB.PictureBox gridProduk 
         Height          =   6255
         Left            =   120
         ScaleHeight     =   6195
         ScaleWidth      =   7515
         TabIndex        =   3
         Top             =   360
         Width           =   7575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " [ JENIS BARANG ] "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.PictureBox gridJenis 
         Height          =   6255
         Left            =   120
         ScaleHeight     =   6195
         ScaleWidth      =   3915
         TabIndex        =   1
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.PictureBox cmdTampilProduk 
      Height          =   495
      Left            =   4440
      ScaleHeight     =   435
      ScaleWidth      =   3795
      TabIndex        =   4
      Top             =   6960
      Width           =   3855
   End
   Begin VB.PictureBox cmdTampilBarang 
      Height          =   495
      Left            =   8520
      ScaleHeight     =   435
      ScaleWidth      =   3675
      TabIndex        =   5
      Top             =   6960
      Width           =   3735
   End
End
Attribute VB_Name = "frmLapProdukJenis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




