VERSION 5.00
Begin VB.Form frmLapBeliNota 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  :: Nota History Pembelian"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox gridTransaksi 
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   480
      Width           =   8535
   End
   Begin VB.PictureBox gridTransDetail 
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   11475
      TabIndex        =   1
      Top             =   4440
      Width           =   11535
   End
   Begin VB.PictureBox CmdTampil 
      Height          =   495
      Left            =   8760
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.PictureBox CmdCetak 
      Height          =   495
      Left            =   10080
      ScaleHeight     =   435
      ScaleWidth      =   1395
      TabIndex        =   5
      Top             =   720
      Width           =   1455
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
      Caption         =   "Daftar Pembelian (List Nota)"
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
      Width           =   3795
   End
End
Attribute VB_Name = "frmLapBeliNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




