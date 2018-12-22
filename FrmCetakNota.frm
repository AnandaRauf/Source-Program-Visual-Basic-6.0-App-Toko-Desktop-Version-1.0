VERSION 5.00
Begin VB.Form FrmCetakNota 
   Caption         =   "Cetak Nota"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox CR1 
      Height          =   6975
      Left            =   -120
      ScaleHeight     =   6915
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.PictureBox CommandButton1 
      Height          =   615
      Left            =   5520
      ScaleHeight     =   555
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "FrmCetakNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_Load() '5BEF20
