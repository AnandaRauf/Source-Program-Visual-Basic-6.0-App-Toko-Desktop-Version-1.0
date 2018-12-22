VERSION 5.00
Begin VB.Form frmPelanggan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   ::  MANAJEMEN DATA PELANGGAN"
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
      TabIndex        =   10
      Top             =   2640
      Width           =   8175
      Begin VB.PictureBox CmdBaru 
         Height          =   375
         Left            =   1440
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox CmdSimpan 
         Height          =   375
         Left            =   2640
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox cmdKeluar 
         Height          =   375
         Left            =   5040
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox CmdHapus 
         Height          =   375
         Left            =   3840
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   8175
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
         Left            =   2040
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1200
         Width           =   5640
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
         Height          =   360
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   0
         Top             =   240
         Width           =   1350
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
         Height          =   390
         Left            =   2040
         MaxLength       =   60
         TabIndex        =   1
         Top             =   720
         Width           =   5625
      End
      Begin VB.PictureBox txtTelepon 
         Height          =   375
         Left            =   2040
         ScaleHeight     =   315
         ScaleWidth      =   2115
         TabIndex        =   3
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   390
         TabIndex        =   9
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pelanggan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   390
         TabIndex        =   8
         Top             =   885
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Telepon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   390
         TabIndex        =   7
         Top             =   2040
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pelanggan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   375
         TabIndex        =   6
         Top             =   405
         Width           =   1410
      End
   End
   Begin VB.PictureBox GridPlg 
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4515
      ScaleWidth      =   8115
      TabIndex        =   4
      ToolTipText     =   "KLIK GANDA DISINI UNTUK MERUBAH DATA ..!"
      Top             =   3480
      Width           =   8175
   End
End
Attribute VB_Name = "frmPelanggan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




