VERSION 5.00
Begin VB.Form FrmCetakPiutang 
   Caption         =   ".: Cetak Pembayaran Piutang :."
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   0
   End
   Begin VB.PictureBox CR1 
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6915
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "FrmCetakPiutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_Load() '609290
