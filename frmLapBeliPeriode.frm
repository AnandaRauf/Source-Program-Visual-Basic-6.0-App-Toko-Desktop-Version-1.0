VERSION 5.00
Begin VB.Form frmLapBeliPeriode 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   ::  Laporan Pembelian Periode"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   13500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Pembelian"
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
      TabIndex        =   6
      Top             =   0
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
         TabIndex        =   7
         Top             =   360
         Width           =   5535
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.PictureBox TglA 
         Height          =   405
         Left            =   1800
         ScaleHeight     =   345
         ScaleWidth      =   2025
         TabIndex        =   1
         Top             =   360
         Width           =   2085
      End
      Begin VB.PictureBox TglB 
         Height          =   405
         Left            =   4560
         ScaleHeight     =   345
         ScaleWidth      =   2025
         TabIndex        =   2
         Top             =   360
         Width           =   2085
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
         TabIndex        =   4
         Top             =   360
         Width           =   435
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
         TabIndex        =   3
         Top             =   360
         Width           =   930
      End
   End
   Begin VB.PictureBox GridPreview 
      Height          =   5775
      Left            =   120
      ScaleHeight     =   5715
      ScaleWidth      =   13155
      TabIndex        =   5
      Top             =   1800
      Width           =   13215
   End
   Begin VB.PictureBox CmdTampil 
      Height          =   495
      Left            =   5280
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.PictureBox CmdCetak 
      Height          =   495
      Left            =   6600
      ScaleHeight     =   435
      ScaleWidth      =   1395
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "frmLapBeliPeriode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




