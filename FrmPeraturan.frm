VERSION 5.00
Begin VB.Form FrmPeraturan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".: Peraturan Lisensi :."
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8640
   Icon            =   "FrmPeraturan.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox cmdKeluar 
      Height          =   375
      Left            =   3600
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   ".: www.vbawam.com :."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6. Setiap pelanggan yang telah membeli produk kami, kami anggap setuju dengan semua peraturan yang kami buat."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   3600
      Width           =   7575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5. Kami menjual produk kami berdasarkan lisensi, jadi anda tidak bisa dengan bebas mengcopy produk kami tanpa otoritas dari kami."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   3000
      Width           =   7575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"FrmPeraturan.frx":0EE2
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   3
      Top             =   1920
      Width           =   7575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"FrmPeraturan.frx":0FFC
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   7575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2. Setiap pelanggan yang telah membeli produk kami tidak boleh dibatalkan dengan alasan apapun"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   7575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1. Setiap calon pelanggan yang akan membeli wajib mencoba versi trial yang telah disediakan"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "FrmPeraturan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdKeluar_UnknownEvent_10() '662AB0
