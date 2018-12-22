VERSION 5.00
Begin VB.Form FrmLapLaba 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Laporan Laba"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11310
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6015
      Begin VB.PictureBox TglA 
         Height          =   405
         Left            =   1080
         ScaleHeight     =   345
         ScaleWidth      =   2025
         TabIndex        =   5
         Top             =   360
         Width           =   2085
      End
      Begin VB.PictureBox TglB 
         Height          =   405
         Left            =   3720
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
         Left            =   120
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
         Left            =   3240
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
      Top             =   6960
      Width           =   5055
      Begin VB.PictureBox CmdTampil 
         Height          =   495
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   1275
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.PictureBox CmdCetak 
         Height          =   495
         Left            =   3480
         ScaleHeight     =   435
         ScaleWidth      =   1395
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
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
      Caption         =   "Total Laba"
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
      Left            =   6240
      TabIndex        =   0
      Top             =   0
      Width           =   4935
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
         Width           =   4575
      End
   End
   Begin VB.PictureBox GridPreview 
      Height          =   5535
      Left            =   120
      ScaleHeight     =   5475
      ScaleWidth      =   10995
      TabIndex        =   9
      Top             =   1320
      Width           =   11055
   End
End
Attribute VB_Name = "FrmLapLaba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




