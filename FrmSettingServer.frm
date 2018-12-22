VERSION 5.00
Begin VB.Form FrmSettingServer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".: Setting Server :."
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7155
   Icon            =   "FrmSettingServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtServer 
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
      TabIndex        =   1
      Top             =   1080
      Width           =   3720
   End
   Begin VB.TextBox txtDbs 
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
      Top             =   1680
      Width           =   3735
   End
   Begin VB.PictureBox CmdKonek 
      Height          =   375
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   2640
      Width           =   2055
   End
   Begin VB.PictureBox CmdSimpan 
      Height          =   375
      Left            =   3600
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.PictureBox cmdKeluar 
      Height          =   375
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server Name :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   270
      Left            =   480
      TabIndex        =   5
      Top             =   1080
      Width           =   1785
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   270
      Left            =   840
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Setting SQL Server"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   2520
      Width           =   5175
   End
End
Attribute VB_Name = "FrmSettingServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'VA: 438150
