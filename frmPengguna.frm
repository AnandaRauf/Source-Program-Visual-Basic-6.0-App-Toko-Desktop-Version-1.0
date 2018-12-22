VERSION 5.00
Begin VB.Form frmPengguna 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   ::  MANAJEMEN PENGGUNA"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7335
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
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   2
         Top             =   1200
         Width           =   4935
      End
      Begin VB.TextBox txtUserId 
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
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FFFFFF&
         Height          =   400
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   30
         PasswordChar    =   "#"
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.PictureBox cmbLevel 
         Height          =   405
         Left            =   1800
         ScaleHeight     =   345
         ScaleWidth      =   2115
         TabIndex        =   3
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pemilik"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "*) Minimal 4 digit"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   7335
      Begin VB.PictureBox CmdBaru 
         Height          =   375
         Left            =   1200
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox CmdSimpan 
         Height          =   375
         Left            =   2520
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox cmdKeluar 
         Height          =   375
         Left            =   5160
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
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.PictureBox GridPengguna 
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2835
      ScaleWidth      =   7275
      TabIndex        =   4
      ToolTipText     =   "KLIK GANDA DISINI UNTUK MERUBAH DATA ..!"
      Top             =   3120
      Width           =   7335
   End
End
Attribute VB_Name = "frmPengguna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




