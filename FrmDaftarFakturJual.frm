VERSION 5.00
Begin VB.Form FrmDaftarFakturJual 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".: Daftar Faktur Jual :."
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10980
   Icon            =   "FrmDaftarFakturJual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
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
      Top             =   0
      Width           =   10695
      Begin VB.OptionButton rbKategori 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nama Pelanggan"
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
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton rbKategori 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No. Nota"
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
         TabIndex        =   1
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.PictureBox GridBarang 
      Height          =   5175
      Left            =   120
      ScaleHeight     =   5115
      ScaleWidth      =   10635
      TabIndex        =   4
      ToolTipText     =   "KLIK 2x DISINI UNTUK PILIH DATA"
      Top             =   840
      Width           =   10695
   End
End
Attribute VB_Name = "FrmDaftarFakturJual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




