VERSION 5.00
Begin VB.Form frmPemasok 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   ::  MANAJEMEN DATA PEMASOK"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   8175
      Begin VB.PictureBox CmdBaru 
         Height          =   375
         Left            =   1440
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox CmdSimpan 
         Height          =   375
         Left            =   2640
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox cmdKeluar 
         Height          =   375
         Left            =   5040
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox CmdHapus 
         Height          =   375
         Left            =   3840
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   8175
      Begin VB.TextBox txtKontak 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   4
         Top             =   2520
         Width           =   2625
      End
      Begin VB.TextBox txtAlamat 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1320
         Width           =   5520
      End
      Begin VB.TextBox txtKode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2280
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtNama 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2280
         MaxLength       =   35
         TabIndex        =   1
         Top             =   840
         Width           =   5505
      End
      Begin VB.PictureBox txtTelepon 
         Height          =   400
         Left            =   2280
         ScaleHeight     =   345
         ScaleWidth      =   2595
         TabIndex        =   3
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kontak Person"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   480
         TabIndex        =   11
         Top             =   2565
         Width           =   1515
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   495
         TabIndex        =   10
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pemasok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   480
         TabIndex        =   9
         Top             =   885
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telepon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   480
         TabIndex        =   8
         Top             =   2160
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pemasok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   480
         TabIndex        =   7
         Top             =   405
         Width           =   1590
      End
   End
   Begin VB.PictureBox GridPsk 
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4035
      ScaleWidth      =   8115
      TabIndex        =   5
      ToolTipText     =   "KLIK GANDA DISINI UNTUK MERUBAH DATA ..!"
      Top             =   3960
      Width           =   8175
   End
End
Attribute VB_Name = "frmPemasok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




