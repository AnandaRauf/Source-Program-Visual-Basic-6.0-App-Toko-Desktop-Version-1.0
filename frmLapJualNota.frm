VERSION 5.00
Begin VB.Form frmLapJualNota 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  :: Nota History Penjualan"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
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
      Height          =   1215
      Left            =   8040
      TabIndex        =   4
      Top             =   480
      Width           =   3015
      Begin VB.PictureBox CmdTampil 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   1275
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.PictureBox CmdCetak 
         Height          =   495
         Left            =   1440
         ScaleHeight     =   435
         ScaleWidth      =   1395
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.PictureBox gridTransaksi 
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   480
      Width           =   7935
   End
   Begin VB.PictureBox gridTransDetail 
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   10875
      TabIndex        =   1
      Top             =   4440
      Width           =   10935
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
      TabIndex        =   3
      Top             =   4080
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Daftar Penjualan (List Nota)"
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
      TabIndex        =   2
      Top             =   120
      Width           =   3750
   End
End
Attribute VB_Name = "frmLapJualNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




