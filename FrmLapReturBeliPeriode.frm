VERSION 5.00
Begin VB.Form FrmLapReturBeliPeriode 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".: Laporan Retur Pembelian Per Periode :."
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13440
   Icon            =   "FrmLapReturBeliPeriode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   13440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7215
      Begin VB.PictureBox TglA 
         Height          =   405
         Left            =   1800
         ScaleHeight     =   345
         ScaleWidth      =   2025
         TabIndex        =   5
         Top             =   360
         Width           =   2085
      End
      Begin VB.PictureBox TglB 
         Height          =   405
         Left            =   4560
         ScaleHeight     =   345
         ScaleWidth      =   2025
         TabIndex        =   6
         Top             =   360
         Width           =   2085
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3960
         TabIndex        =   7
         Top             =   360
         Width           =   435
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   5775
      Begin VB.PictureBox CmdTampil 
         Height          =   495
         Left            =   2400
         ScaleHeight     =   435
         ScaleWidth      =   1275
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.PictureBox CmdCetak 
         Height          =   495
         Left            =   3720
         ScaleHeight     =   435
         ScaleWidth      =   1395
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Model Laporan :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Retur Pembelian"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7440
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.Label labelBayar 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5535
      End
   End
   Begin VB.PictureBox GridPreview 
      Height          =   5535
      Left            =   120
      ScaleHeight     =   5475
      ScaleWidth      =   13155
      TabIndex        =   9
      Top             =   2160
      Width           =   13215
   End
End
Attribute VB_Name = "FrmLapReturBeliPeriode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




