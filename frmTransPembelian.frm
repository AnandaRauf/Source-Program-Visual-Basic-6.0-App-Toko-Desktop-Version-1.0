VERSION 5.00
Begin VB.Form frmTransPembelian 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   :: Transaksi Pembelian ( Barang Masuk )"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13200
   Icon            =   "frmTransPembelian.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   13200
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtGrandTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6480
      TabIndex        =   32
      Text            =   "0"
      Top             =   6840
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.TextBox TxtSisa 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   10680
      TabIndex        =   29
      Text            =   "0"
      Top             =   6840
      Width           =   2265
   End
   Begin VB.TextBox TxtDP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   10680
      TabIndex        =   28
      Text            =   "0"
      Top             =   6480
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total (Rp)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   23
      Top             =   6360
      Width           =   6255
      Begin VB.Label labelTotal 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   6015
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   12975
      Begin VB.PictureBox DtTempo 
         Height          =   375
         Left            =   7920
         ScaleHeight     =   315
         ScaleWidth      =   1275
         TabIndex        =   31
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtnotransaksi 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10800
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtCatatan 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10800
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtNotaBeli 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   0
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtPemasok 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
      Begin VB.PictureBox cmbJenis 
         Height          =   375
         Left            =   5040
         ScaleHeight     =   315
         ScaleWidth      =   1635
         TabIndex        =   22
         Top             =   720
         Width           =   1695
      End
      Begin VB.PictureBox cmbPemasok 
         Height          =   375
         Left            =   1800
         ScaleHeight     =   315
         ScaleWidth      =   2355
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Transkasi Beli :"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9000
         TabIndex        =   33
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label LblTgl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Tempo"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6840
         TabIndex        =   26
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan :"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9960
         TabIndex        =   20
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Nota Beli"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   12
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pemasok"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   12975
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9480
         TabIndex        =   21
         Text            =   "0"
         Top             =   720
         Width           =   1785
      End
      Begin VB.TextBox txtJumlah 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "0"
         Top             =   720
         Width           =   885
      End
      Begin VB.TextBox txtHBeli 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Text            =   "0"
         Top             =   720
         Width           =   1785
      End
      Begin VB.TextBox txtNama 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   7095
      End
      Begin VB.TextBox txtHJual 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   8
         Text            =   "0"
         Top             =   720
         Width           =   1905
      End
      Begin VB.TextBox txtKode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   2865
      End
      Begin VB.PictureBox cmdCari 
         Height          =   375
         Left            =   5160
         ScaleHeight     =   315
         ScaleWidth      =   435
         TabIndex        =   37
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox CmdMasuk 
         Height          =   375
         Left            =   11400
         ScaleHeight     =   315
         ScaleWidth      =   1395
         TabIndex        =   38
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total (Rp)"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8520
         TabIndex        =   19
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7200
         TabIndex        =   18
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Beli (Rp)"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   720
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode / Scan Barcode"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Jual (Rp)"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   720
         Width           =   1365
      End
   End
   Begin VB.PictureBox GridMasuk 
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3795
      ScaleWidth      =   12915
      TabIndex        =   9
      Top             =   2520
      Width           =   12975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   510
      Top             =   9210
   End
   Begin VB.PictureBox CmdBaru 
      Height          =   375
      Left            =   9120
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   34
      Top             =   7440
      Width           =   1215
   End
   Begin VB.PictureBox CmdSimpan 
      Height          =   375
      Left            =   10440
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   35
      Top             =   7440
      Width           =   1215
   End
   Begin VB.PictureBox cmdKeluar 
      Height          =   375
      Left            =   11760
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   36
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[Enter] - untuk melanjutkan transaksi"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   9600
      TabIndex        =   39
      Top             =   7920
      Width           =   3000
   End
   Begin VB.Label LblSisa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sisa (Rp)"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9840
      TabIndex        =   25
      Top             =   6840
      Width           =   780
   End
   Begin VB.Label LblDP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DP (Rp)"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9840
      TabIndex        =   30
      Top             =   6480
      Width           =   690
   End
End
Attribute VB_Name = "frmTransPembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CmdSimpan_UnknownEvent_10() '596F90
loc_00596F90:   push ebp
loc_00596F91:   mov ebp, esp
  loc_00596F93: sub esp, 0000000Ch
  loc_00596F96: push 00405696h ; __vbaExceptHandler
loc_00596F9B:   mov eax, fs: [00000000h]
loc_00596FA1:   push eax
loc_00596FA2:   mov fs: [00000000h] , esp
  loc_00596FA9: sub esp, 00000354h
loc_00596FAF:   push ebx
loc_00596FB0:   push esi
loc_00596FB1:   push edi
loc_00596FB2:   mov var_C, esp
  loc_00596FB5: mov var_8, 004027F8h
loc_00596FBC:   mov ebx, Me
loc_00596FBF:   mov eax, ebx
  loc_00596FC1: and eax, 00000001h
loc_00596FC4:   mov var_4, eax
  loc_00596FC7: and ebx, FFFFFFFEh
loc_00596FCA:   push ebx
loc_00596FCB:   mov Me, ebx
loc_00596FCE:   mov ecx, [ebx]
loc_00596FD0:   Call [ecx+00000004h]
  loc_00596FD3: xor edi, edi
loc_00596FD5:   mov var_18, edi
loc_00596FD8:   mov var_20, edi
loc_00596FDB:   mov var_24, edi
loc_00596FDE:   mov var_2C, edi
loc_00596FE1:   mov var_30, edi
loc_00596FE4:   mov var_34, edi
loc_00596FE7:   mov var_38, edi
loc_00596FEA:   mov var_3C, edi
loc_00596FED:   mov var_40, edi
loc_00596FF0:   mov var_44, edi
loc_00596FF3:   mov var_48, edi
loc_00596FF6:   mov var_4C, edi
loc_00596FF9:   mov var_50, edi
loc_00596FFC:   mov var_54, edi
loc_00596FFF:   mov var_58, edi
loc_00597002:   mov var_5C, edi
loc_00597005:   mov var_60, edi
loc_00597008:   mov var_64, edi
loc_0059700B:   mov var_68, edi
loc_0059700E:   mov var_6C, edi
loc_00597011:   mov var_70, edi
loc_00597014:   mov var_74, edi
loc_00597017:   mov var_78, edi
loc_0059701A:   mov var_7C, edi
loc_0059701D:   mov var_80, edi
loc_00597020:   mov var_84, edi
loc_00597026:   mov var_88, edi
loc_0059702C:   mov var_8C, edi
loc_00597032:   mov var_90, edi
loc_00597038:   mov var_A0, edi
loc_0059703E:   mov var_B0, edi
loc_00597044:   mov var_C0, edi
loc_0059704A:   mov var_D0, edi
loc_00597050:   mov var_E0, edi
loc_00597056:   mov var_F0, edi
loc_0059705C:   mov var_100, edi
loc_00597062:   mov var_110, edi
loc_00597068:   mov var_120, edi
loc_0059706E:   mov var_130, edi
loc_00597074:   mov var_140, edi
loc_0059707A:   mov var_150, edi
loc_00597080:   mov var_160, edi
loc_00597086:   mov var_170, edi
loc_0059708C:   mov var_180, edi
loc_00597092:   mov var_190, edi
loc_00597098:   mov var_1A0, edi
loc_0059709E:   mov var_1B0, edi
loc_005970A4:   mov var_1C0, edi
loc_005970AA:   mov var_1D0, edi
loc_005970B0:   mov var_1E0, edi
loc_005970B6:   mov var_1F0, edi
loc_005970BC:   mov var_200, edi
loc_005970C2:   mov var_210, edi
loc_005970C8:   mov var_220, edi
loc_005970CE:   mov var_230, edi
loc_005970D4:   mov var_240, edi
loc_005970DA:   mov var_250, edi
loc_005970E0:   mov var_260, edi
loc_005970E6:   mov var_270, edi
loc_005970EC:   mov var_280, edi
loc_005970F2:   mov var_290, edi
loc_005970F8:   mov var_2A0, edi
loc_005970FE:   mov var_2B0, edi
loc_00597104:   mov var_2C0, edi
loc_0059710A:   mov var_2E0, edi
loc_00597110:   mov var_300, edi
loc_00597116:   mov edx, [ebx]
loc_00597118:   push edi
loc_00597119:   push FFFFFDFBh
loc_0059711E:   push ebx
loc_0059711F:   mov var_314, edi
loc_00597125:   Call [edx+00000384h]
loc_0059712B:   push eax
loc_0059712C:   lea eax, var_78
loc_0059712F:   push eax
  loc_00597130: call [004010B8h] ; __vbaObjSet
loc_00597136:   lea ecx, var_A0
loc_0059713C:   push eax
loc_0059713D:   push ecx
  loc_0059713E: call [0040114Ch] ; __vbaLateIdCallLd
  loc_00597144: add esp, 00000010h
loc_00597147:   push eax
  loc_00597148: call [00401034h] ; __vbaStrVarMove
  loc_0059714E: mov esi, [0040126Ch] ; __vbaStrMove
loc_00597154:   mov edx, eax
loc_00597156:   lea ecx, var_2C
  loc_00597159: call __vbaStrMove
loc_0059715B:   push eax
  loc_0059715C: push 0042DFECh
  loc_00597161: call [0040112Ch] ; __vbaStrCmp
loc_00597167:   neg eax
loc_00597169:   sbb eax, eax
loc_0059716B:   lea ecx, var_2C
loc_0059716E:   inc eax
loc_0059716F:   neg eax
loc_00597171:   mov var_318, ax
  loc_00597178: call [0040129Ch] ; __vbaFreeStr
loc_0059717E:   lea ecx, var_78
  loc_00597181: call [004012A0h] ; __vbaFreeObj
loc_00597187:   lea ecx, var_A0
  loc_0059718D: call [00401020h] ; __vbaFreeVar
loc_00597193:   cmp var_318, di
  loc_0059719A: jz 00597270h
  loc_005971A0: mov esi, [00401238h] ; __vbaVarDup
  loc_005971A6: mov ecx, 80020004h
loc_005971AB:   mov var_C8, ecx
  loc_005971B1: mov eax, 0000000Ah
loc_005971B6:   mov var_B8, ecx
loc_005971BC:   lea edx, var_230
loc_005971C2:   lea ecx, var_B0
loc_005971C8:   mov var_D0, eax
loc_005971CE:   mov var_C0, eax
  loc_005971D4: mov var_228, 0042E624h ; "Error"
  loc_005971DE: mov var_230, 00000008h
  loc_005971E8: call __vbaVarDup
loc_005971EA:   lea edx, var_220
loc_005971F0:   lea ecx, var_A0
  loc_005971F6: mov var_218, 00434338h ; "DATA PEMASOK BELUM DIPILIH"
  loc_00597200: mov var_220, 00000008h
  loc_0059720A: call __vbaVarDup
loc_0059720C:   lea edx, var_D0
loc_00597212:   lea eax, var_C0
loc_00597218:   push edx
loc_00597219:   lea ecx, var_B0
loc_0059721F:   push eax
loc_00597220:   push ecx
loc_00597221:   lea edx, var_A0
  loc_00597227: push 00000010h
loc_00597229:   push edx
  loc_0059722A: call [004010B4h] ; rtcMsgBox
loc_00597230:   lea eax, var_D0
loc_00597236:   lea ecx, var_C0
loc_0059723C:   push eax
loc_0059723D:   lea edx, var_B0
loc_00597243:   push ecx
loc_00597244:   lea eax, var_A0
loc_0059724A:   push edx
loc_0059724B:   push eax
  loc_0059724C: push 00000004h
  loc_0059724E: call [0040103Ch] ; __vbaFreeVarList
loc_00597254:   mov ecx, [ebx]
  loc_00597256: add esp, 00000014h
loc_00597259:   push edi
  loc_0059725A: push 80011000h
loc_0059725F:   push ebx
loc_00597260:   Call [ecx+00000384h]
loc_00597266:   lea edx, var_78
loc_00597269:   push eax
loc_0059726A:   push edx
  loc_0059726B: jmp 00597346h
  loc_00597270: cmp [ebx+00000034h], 0001h
  loc_00597275: jnz 0059735Bh
  loc_0059727B: mov esi, [00401238h] ; __vbaVarDup
  loc_00597281: mov ecx, 80020004h
loc_00597286:   mov var_C8, ecx
  loc_0059728C: mov eax, 0000000Ah
loc_00597291:   mov var_B8, ecx
loc_00597297:   lea edx, var_230
loc_0059729D:   lea ecx, var_B0
loc_005972A3:   mov var_D0, eax
loc_005972A9:   mov var_C0, eax
  loc_005972AF: mov var_228, 0042E624h ; "Error"
  loc_005972B9: mov var_230, 00000008h
  loc_005972C3: call __vbaVarDup
loc_005972C5:   lea edx, var_220
loc_005972CB:   lea ecx, var_A0
  loc_005972D1: mov var_218, 00434374h ; "BELUM ADA DATA BARANG YANG MASUK"
  loc_005972DB: mov var_220, 00000008h
  loc_005972E5: call __vbaVarDup
loc_005972E7:   lea eax, var_D0
loc_005972ED:   lea ecx, var_C0
loc_005972F3:   push eax
loc_005972F4:   lea edx, var_B0
loc_005972FA:   push ecx
loc_005972FB:   push edx
loc_005972FC:   lea eax, var_A0
  loc_00597302: push 00000010h
loc_00597304:   push eax
  loc_00597305: call [004010B4h] ; rtcMsgBox
loc_0059730B:   lea ecx, var_D0
loc_00597311:   lea edx, var_C0
loc_00597317:   push ecx
loc_00597318:   lea eax, var_B0
loc_0059731E:   push edx
loc_0059731F:   lea ecx, var_A0
loc_00597325:   push eax
loc_00597326:   push ecx
  loc_00597327: push 00000004h
  loc_00597329: call [0040103Ch] ; __vbaFreeVarList
loc_0059732F:   mov edx, [ebx]
  loc_00597331: add esp, 00000014h
loc_00597334:   push edi
  loc_00597335: push 80011000h
loc_0059733A:   push ebx
loc_0059733B:   Call [edx+00000388h]
loc_00597341:   push eax
loc_00597342:   lea eax, var_78
loc_00597345:   push eax
  loc_00597346: call [004010B8h] ; __vbaObjSet
loc_0059734C:   push eax
  loc_0059734D: call [0040102Ch] ; __vbaLateIdCall
  loc_00597353: add esp, 0000000Ch
  loc_00597356: jmp 00599E89h
  loc_0059735B: push 0042DF30h
  loc_00597360: call [00401168h] ; __vbaNew
loc_00597366:   push eax
  loc_00597367: push 00668060h
  loc_0059736C: call [004010B8h] ; __vbaObjSet
loc_00597372:   mov ecx, [ebx]
loc_00597374:   push ebx
loc_00597375:   Call [ecx+0000031Ch]
loc_0059737B:   lea edx, var_78
loc_0059737E:   push eax
loc_0059737F:   push edx
  loc_00597380: call [004010B8h] ; __vbaObjSet
loc_00597386:   mov ecx, [eax]
loc_00597388:   lea edx, var_2C
loc_0059738B:   push edx
loc_0059738C:   push eax
loc_0059738D:   mov var_318, eax
loc_00597393:   Call [ecx+000000A0h]
loc_00597399:   cmp eax, edi
loc_0059739B:   fnclex
  loc_0059739D: jge 005973B7h
loc_0059739F:   mov ecx, var_318
  loc_005973A5: push 000000A0h
  loc_005973AA: push 0042DFCCh
loc_005973AF:   push ecx
loc_005973B0:   push eax
  loc_005973B1: call [00401080h] ; __vbaHresultCheckObj
loc_005973B7:   cmp [0066803Ch], edi
  loc_005973BD: jnz 005973CFh
  loc_005973BF: push 0066803Ch
  loc_005973C4: push 0042DEFCh
  loc_005973C9: call [004011E8h] ; __vbaNew2
loc_005973CF:   mov eax, var_2C
loc_005973D2:   mov edx, [0066803Ch]
  loc_005973D8: mov edi, [00401070h] ; __vbaStrCat
  loc_005973DE: push 004343BCh ; "SELECT * FROM pembelian WHERE no_notabeli='"
loc_005973E3:   push eax
loc_005973E4:   mov var_218, edx
  loc_005973EA: mov var_220, 00000009h
loc_005973F4:   Call edi
loc_005973F6:   mov edx, eax
loc_005973F8:   lea ecx, var_30
  loc_005973FB: call __vbaVarDup
loc_005973FD:   push eax
  loc_005973FE: push 0042EBA8h ; "'"
loc_00597403:   Call edi
loc_00597405:   mov ecx, [00668060h]
loc_0059740B:   push FFFFFFFFh
  loc_0059740D: push 00000004h
  loc_0059740F: push 00000002h
  loc_00597411: sub esp, 00000010h
loc_00597414:   mov var_98, eax
  loc_0059741A: mov var_A0, 00000008h
loc_00597424:   mov edx, [ecx]
loc_00597426:   mov ecx, var_220
loc_0059742C:   mov eax, esp
  loc_0059742E: sub esp, 00000010h
loc_00597431:   mov [eax], ecx
loc_00597433:   mov ecx, var_21C
loc_00597439:   mov [eax+00000004h], ecx
loc_0059743C:   mov ecx, var_218
loc_00597442:   mov [eax+00000008h], ecx
loc_00597445:   mov ecx, var_214
loc_0059744B:   mov [eax+0000000Ch], ecx
loc_0059744E:   mov ecx, var_A0
loc_00597454:   mov eax, esp
loc_00597456:   mov [eax], ecx
loc_00597458:   mov ecx, var_9C
loc_0059745E:   mov [eax+00000004h], ecx
loc_00597461:   mov ecx, var_98
loc_00597467:   mov [eax+00000008h], ecx
loc_0059746A:   mov ecx, var_94
loc_00597470:   mov [eax+0000000Ch], ecx
loc_00597473:   mov eax, [00668060h]
loc_00597478:   push eax
loc_00597479:   Call [edx+000000A0h]
loc_0059747F:   test eax, eax
loc_00597481:   fnclex
  loc_00597483: jge 0059749Dh
loc_00597485:   mov ecx, [00668060h]
  loc_0059748B: push 000000A0h
  loc_00597490: push 0042DF44h
loc_00597495:   push ecx
loc_00597496:   push eax
  loc_00597497: call [00401080h] ; __vbaHresultCheckObj
loc_0059749D:   lea edx, var_30
loc_005974A0:   lea eax, var_2C
loc_005974A3:   push edx
loc_005974A4:   push eax
  loc_005974A5: push 00000002h
  loc_005974A7: call [00401204h] ; __vbaFreeStrList
  loc_005974AD: add esp, 0000000Ch
loc_005974B0:   lea ecx, var_78
  loc_005974B3: call [004012A0h] ; __vbaFreeObj
loc_005974B9:   lea ecx, var_A0
  loc_005974BF: call [00401020h] ; __vbaFreeVar
loc_005974C5:   mov eax, [00668060h]
loc_005974CA:   lea edx, var_314
loc_005974D0:   push edx
loc_005974D1:   push eax
loc_005974D2:   mov ecx, [eax]
loc_005974D4:   Call [ecx+00000050h]
loc_005974D7:   test eax, eax
loc_005974D9:   fnclex
  loc_005974DB: jge 005974F2h
loc_005974DD:   mov ecx, [00668060h]
  loc_005974E3: push 00000050h
  loc_005974E5: push 0042DF44h
loc_005974EA:   push ecx
loc_005974EB:   push eax
  loc_005974EC: call [00401080h] ; __vbaHresultCheckObj
  loc_005974F2: cmp var_314, 0000h
  loc_005974FA: jz 00599DB3h
loc_00597500:   mov edx, [ebx]
  loc_00597502: push 00000000h
loc_00597504:   push FFFFFDFBh
loc_00597509:   push ebx
loc_0059750A:   Call [edx+00000380h]
loc_00597510:   push eax
loc_00597511:   lea eax, var_78
loc_00597514:   push eax
  loc_00597515: call [004010B8h] ; __vbaObjSet
loc_0059751B:   lea ecx, var_A0
loc_00597521:   push eax
loc_00597522:   push ecx
  loc_00597523: call [0040114Ch] ; __vbaLateIdCallLd
  loc_00597529: add esp, 00000010h
loc_0059752C:   push eax
  loc_0059752D: call [00401034h] ; __vbaStrVarMove
loc_00597533:   mov edx, eax
loc_00597535:   lea ecx, var_2C
  loc_00597538: call __vbaVarDup
loc_0059753A:   push eax
  loc_0059753B: push 00433F74h ; "Tunai"
  loc_00597540: call [0040112Ch] ; __vbaStrCmp
loc_00597546:   neg eax
loc_00597548:   sbb eax, eax
loc_0059754A:   lea ecx, var_2C
loc_0059754D:   inc eax
loc_0059754E:   neg eax
loc_00597550:   mov var_318, ax
  loc_00597557: call [0040129Ch] ; __vbaFreeStr
loc_0059755D:   lea ecx, var_78
  loc_00597560: call [004012A0h] ; __vbaFreeObj
loc_00597566:   lea ecx, var_A0
  loc_0059756C: call [00401020h] ; __vbaFreeVar
  loc_00597572: cmp var_318, 0000h
  loc_0059757A: jz 00598682h
  loc_00597580: call 0055B320h
loc_00597585:   mov edx, [ebx]
loc_00597587:   push ebx
loc_00597588:   Call [edx+00000708h]
loc_0059758E:   test eax, eax
  loc_00597590: jge 005975A4h
  loc_00597592: push 00000708h
  loc_00597597: push 00433DD8h
loc_0059759C:   push ebx
loc_0059759D:   push eax
  loc_0059759E: call [00401080h] ; __vbaHresultCheckObj
loc_005975A4:   lea eax, var_A0
loc_005975AA:   push eax
  loc_005975AB: call [00401220h] ; rtcGetDateVar
loc_005975B1:   mov ecx, [ebx]
loc_005975B3:   push ebx
loc_005975B4:   Call [ecx+0000031Ch]
loc_005975BA:   lea edx, var_78
loc_005975BD:   push eax
loc_005975BE:   push edx
  loc_005975BF: call [004010B8h] ; __vbaObjSet
loc_005975C5:   mov ecx, [eax]
loc_005975C7:   lea edx, var_38
loc_005975CA:   push edx
loc_005975CB:   push eax
loc_005975CC:   mov var_318, eax
loc_005975D2:   Call [ecx+000000A0h]
loc_005975D8:   test eax, eax
loc_005975DA:   fnclex
  loc_005975DC: jge 005975F6h
loc_005975DE:   mov ecx, var_318
  loc_005975E4: push 000000A0h
  loc_005975E9: push 0042DFCCh
loc_005975EE:   push ecx
loc_005975EF:   push eax
  loc_005975F0: call [00401080h] ; __vbaHresultCheckObj
  loc_005975F6: push 00434418h ; "INSERT INTO pembelian"
  loc_005975FB: push 00434460h ; "(no_pembelian,no_notabeli, tgl_masuk,kode_pemasok,nama_pemasok,jenis,userid,catatan,tot_beli)"
loc_00597600:   Call edi
loc_00597602:   mov edx, eax
loc_00597604:   lea ecx, var_2C
  loc_00597607: call __vbaVarDup
loc_00597609:   push eax
  loc_0059760A: push 00434520h ; "VALUES ('"
loc_0059760F:   Call edi
loc_00597611:   mov edx, eax
loc_00597613:   lea ecx, var_30
  loc_00597616: call __vbaVarDup
loc_00597618:   mov edx, [ebx+00000040h]
loc_0059761B:   push eax
loc_0059761C:   push edx
loc_0059761D:   Call edi
loc_0059761F:   mov edx, eax
loc_00597621:   lea ecx, var_34
  loc_00597624: call __vbaVarDup
loc_00597626:   push eax
  loc_00597627: push 0042EC30h ; "','"
loc_0059762C:   Call edi
loc_0059762E:   mov edx, eax
loc_00597630:   lea ecx, var_3C
  loc_00597633: call __vbaVarDup
loc_00597635:   push eax
loc_00597636:   mov eax, var_38
loc_00597639:   push eax
loc_0059763A:   Call edi
loc_0059763C:   mov edx, eax
loc_0059763E:   lea ecx, var_40
  loc_00597641: call __vbaVarDup
loc_00597643:   push eax
  loc_00597644: push 0042EC30h ; "','"
loc_00597649:   Call edi
loc_0059764B:   mov var_C8, eax
  loc_00597651: mov eax, 00000008h
loc_00597656:   lea edx, var_220
loc_0059765C:   lea ecx, var_B0
loc_00597662:   mov var_D0, eax
  loc_00597668: mov var_218, 00434538h ; "yyyy/MM/dd"
loc_00597672:   mov var_220, eax
  loc_00597678: call [00401238h] ; __vbaVarDup
  loc_0059767E: push 00000001h
loc_00597680:   lea ecx, var_B0
  loc_00597686: push 00000001h
loc_00597688:   lea edx, var_A0
loc_0059768E:   push ecx
loc_0059768F:   lea eax, var_C0
loc_00597695:   push edx
loc_00597696:   push eax
  loc_00597697: call [00401078h] ; rtcVarFromFormatVar
loc_0059769D:   mov ecx, [ebx]
  loc_0059769F: push 00000000h
loc_005976A1:   push FFFFFDFBh
loc_005976A6:   push ebx
  loc_005976A7: mov var_228, 0042EC30h ; "','"
  loc_005976B1: mov var_230, 00000008h
loc_005976BB:   Call [ecx+00000384h]
loc_005976C1:   lea edx, var_7C
loc_005976C4:   push eax
loc_005976C5:   push edx
  loc_005976C6: call [004010B8h] ; __vbaObjSet
loc_005976CC:   push eax
loc_005976CD:   lea eax, var_100
loc_005976D3:   push eax
  loc_005976D4: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005976DA: add esp, 00000010h
loc_005976DD:   push eax
  loc_005976DE: call [00401034h] ; __vbaStrVarMove
loc_005976E4:   mov ecx, [ebx]
loc_005976E6:   mov var_108, eax
  loc_005976EC: mov eax, 00000008h
loc_005976F1:   push ebx
loc_005976F2:   mov var_110, eax
  loc_005976F8: mov var_238, 0042EC30h ; "','"
loc_00597702:   mov var_240, eax
loc_00597708:   Call [ecx+00000320h]
loc_0059770E:   lea edx, var_80
loc_00597711:   push eax
loc_00597712:   push edx
  loc_00597713: call [004010B8h] ; __vbaObjSet
loc_00597719:   mov ecx, [eax]
loc_0059771B:   lea edx, var_44
loc_0059771E:   push edx
loc_0059771F:   push eax
loc_00597720:   mov var_320, eax
loc_00597726:   Call [ecx+000000A0h]
loc_0059772C:   test eax, eax
loc_0059772E:   fnclex
  loc_00597730: jge 0059774Ah
loc_00597732:   mov ecx, var_320
  loc_00597738: push 000000A0h
  loc_0059773D: push 0042DFCCh
loc_00597742:   push ecx
loc_00597743:   push eax
  loc_00597744: call [00401080h] ; __vbaHresultCheckObj
loc_0059774A:   mov eax, var_44
loc_0059774D:   mov edx, [ebx]
loc_0059774F:   mov var_138, eax
  loc_00597755: push 00000000h
  loc_00597757: mov eax, 00000008h
loc_0059775C:   push FFFFFDFBh
loc_00597761:   push ebx
  loc_00597762: mov var_44, 00000000h
loc_00597769:   mov var_140, eax
  loc_0059776F: mov var_248, 0042EC30h ; "','"
loc_00597779:   mov var_250, eax
loc_0059777F:   Call [edx+00000380h]
loc_00597785:   push eax
loc_00597786:   lea eax, var_84
loc_0059778C:   push eax
  loc_0059778D: call [004010B8h] ; __vbaObjSet
loc_00597793:   lea ecx, var_170
loc_00597799:   push eax
loc_0059779A:   push ecx
  loc_0059779B: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005977A1: add esp, 00000010h
loc_005977A4:   push eax
  loc_005977A5: call [00401034h] ; __vbaStrVarMove
loc_005977AB:   mov edx, [ebx]
loc_005977AD:   mov var_178, eax
  loc_005977B3: mov eax, 00000008h
  loc_005977B8: mov ecx, 0042EC30h ; "','"
loc_005977BD:   push ebx
loc_005977BE:   mov var_180, eax
loc_005977C4:   mov var_258, ecx
loc_005977CA:   mov var_260, eax
loc_005977D0:   mov var_268, ecx
loc_005977D6:   mov var_270, eax
loc_005977DC:   Call [edx+00000318h]
loc_005977E2:   push eax
loc_005977E3:   lea eax, var_88
loc_005977E9:   push eax
  loc_005977EA: call [004010B8h] ; __vbaObjSet
loc_005977F0:   mov ecx, [eax]
loc_005977F2:   lea edx, var_48
loc_005977F5:   push edx
loc_005977F6:   push eax
loc_005977F7:   mov var_328, eax
loc_005977FD:   Call [ecx+000000A0h]
loc_00597803:   test eax, eax
loc_00597805:   fnclex
  loc_00597807: jge 00597821h
loc_00597809:   mov ecx, var_328
  loc_0059780F: push 000000A0h
  loc_00597814: push 0042DFCCh
loc_00597819:   push ecx
loc_0059781A:   push eax
  loc_0059781B: call [00401080h] ; __vbaHresultCheckObj
loc_00597821:   mov eax, var_48
loc_00597824:   mov edx, [ebx]
loc_00597826:   mov var_1C8, eax
  loc_0059782C: mov eax, 00000008h
loc_00597831:   push ebx
  loc_00597832: mov var_48, 00000000h
loc_00597839:   mov var_1D0, eax
  loc_0059783F: mov var_278, 0042EC30h ; "','"
loc_00597849:   mov var_280, eax
loc_0059784F:   Call [edx+000002FCh]
loc_00597855:   push eax
loc_00597856:   lea eax, var_8C
loc_0059785C:   push eax
  loc_0059785D: call [004010B8h] ; __vbaObjSet
loc_00597863:   mov ecx, [eax]
loc_00597865:   lea edx, var_4C
loc_00597868:   push edx
loc_00597869:   push eax
loc_0059786A:   mov var_330, eax
loc_00597870:   Call [ecx+000000A0h]
loc_00597876:   test eax, eax
loc_00597878:   fnclex
  loc_0059787A: jge 00597894h
loc_0059787C:   mov ecx, var_330
  loc_00597882: push 000000A0h
  loc_00597887: push 0042DFCCh
loc_0059788C:   push ecx
loc_0059788D:   push eax
  loc_0059788E: call [00401080h] ; __vbaHresultCheckObj
loc_00597894:   mov edx, var_4C
loc_00597897:   push edx
  loc_00597898: call [004012A4h] ; rtcR8ValFromBstr
  loc_0059789E: fstp real8 ptr var_288
loc_005978A4:   lea eax, var_D0
loc_005978AA:   lea ecx, var_C0
loc_005978B0:   push eax
loc_005978B1:   lea edx, var_E0
loc_005978B7:   push ecx
loc_005978B8:   push edx
  loc_005978B9: mov var_290, 00000005h
  loc_005978C3: mov var_298, 0042EC3Ch ; "')"
  loc_005978CD: mov var_2A0, 00000008h
  loc_005978D7: call [004011C4h] ; __vbaVarCat
loc_005978DD:   push eax
loc_005978DE:   lea eax, var_230
loc_005978E4:   lea ecx, var_F0
loc_005978EA:   push eax
loc_005978EB:   push ecx
  loc_005978EC: call [004011C4h] ; __vbaVarCat
loc_005978F2:   push eax
loc_005978F3:   lea edx, var_110
loc_005978F9:   lea eax, var_120
loc_005978FF:   push edx
loc_00597900:   push eax
  loc_00597901: call [004011C4h] ; __vbaVarCat
loc_00597907:   lea ecx, var_240
loc_0059790D:   push eax
loc_0059790E:   lea edx, var_130
loc_00597914:   push ecx
loc_00597915:   push edx
  loc_00597916: call [004011C4h] ; __vbaVarCat
loc_0059791C:   push eax
loc_0059791D:   lea eax, var_140
loc_00597923:   lea ecx, var_150
loc_00597929:   push eax
loc_0059792A:   push ecx
  loc_0059792B: call [004011C4h] ; __vbaVarCat
loc_00597931:   push eax
loc_00597932:   lea edx, var_250
loc_00597938:   lea eax, var_160
loc_0059793E:   push edx
loc_0059793F:   push eax
  loc_00597940: call [004011C4h] ; __vbaVarCat
loc_00597946:   lea ecx, var_180
loc_0059794C:   push eax
loc_0059794D:   lea edx, var_190
loc_00597953:   push ecx
loc_00597954:   push edx
  loc_00597955: call [004011C4h] ; __vbaVarCat
loc_0059795B:   push eax
loc_0059795C:   lea eax, var_260
loc_00597962:   lea ecx, var_1A0
loc_00597968:   push eax
loc_00597969:   push ecx
  loc_0059796A: call [004011C4h] ; __vbaVarCat
loc_00597970:   push eax
loc_00597971:   lea edx, var_1B0
  loc_00597977: push 006680B4h
loc_0059797C:   push edx
  loc_0059797D: call [004011C4h] ; __vbaVarCat
loc_00597983:   push eax
loc_00597984:   lea eax, var_270
loc_0059798A:   lea ecx, var_1C0
loc_00597990:   push eax
loc_00597991:   push ecx
  loc_00597992: call [004011C4h] ; __vbaVarCat
loc_00597998:   push eax
loc_00597999:   lea edx, var_1D0
loc_0059799F:   lea eax, var_1E0
loc_005979A5:   push edx
loc_005979A6:   push eax
  loc_005979A7: call [004011C4h] ; __vbaVarCat
loc_005979AD:   lea ecx, var_280
loc_005979B3:   push eax
loc_005979B4:   lea edx, var_1F0
loc_005979BA:   push ecx
loc_005979BB:   push edx
  loc_005979BC: call [004011C4h] ; __vbaVarCat
loc_005979C2:   push eax
loc_005979C3:   lea eax, var_290
loc_005979C9:   lea ecx, var_200
loc_005979CF:   push eax
loc_005979D0:   push ecx
  loc_005979D1: call [004011C4h] ; __vbaVarCat
loc_005979D7:   push eax
loc_005979D8:   lea edx, var_2A0
loc_005979DE:   lea eax, var_210
loc_005979E4:   push edx
loc_005979E5:   push eax
  loc_005979E6: call [004011C4h] ; __vbaVarCat
loc_005979EC:   push eax
  loc_005979ED: call [00401034h] ; __vbaStrVarMove
loc_005979F3:   mov edx, eax
  loc_005979F5: mov ecx, 0066809Ch
  loc_005979FA: call __vbaVarDup
loc_005979FC:   lea ecx, var_4C
loc_005979FF:   lea edx, var_40
loc_00597A02:   push ecx
loc_00597A03:   lea eax, var_38
loc_00597A06:   push edx
loc_00597A07:   lea ecx, var_3C
loc_00597A0A:   push eax
loc_00597A0B:   lea edx, var_34
loc_00597A0E:   push ecx
loc_00597A0F:   lea eax, var_30
loc_00597A12:   push edx
loc_00597A13:   lea ecx, var_2C
loc_00597A16:   push eax
loc_00597A17:   push ecx
  loc_00597A18: push 00000007h
  loc_00597A1A: call [00401204h] ; __vbaFreeStrList
loc_00597A20:   lea edx, var_8C
loc_00597A26:   lea eax, var_88
loc_00597A2C:   push edx
loc_00597A2D:   lea ecx, var_84
loc_00597A33:   push eax
loc_00597A34:   lea edx, var_80
loc_00597A37:   push ecx
loc_00597A38:   lea eax, var_7C
loc_00597A3B:   push edx
loc_00597A3C:   lea ecx, var_78
loc_00597A3F:   push eax
loc_00597A40:   push ecx
  loc_00597A41: push 00000006h
  loc_00597A43: call [0040104Ch] ; __vbaFreeObjList
loc_00597A49:   lea edx, var_210
loc_00597A4F:   lea eax, var_200
loc_00597A55:   push edx
loc_00597A56:   lea ecx, var_1F0
loc_00597A5C:   push eax
loc_00597A5D:   lea edx, var_1E0
loc_00597A63:   push ecx
loc_00597A64:   lea eax, var_1D0
loc_00597A6A:   push edx
loc_00597A6B:   lea ecx, var_1C0
loc_00597A71:   push eax
loc_00597A72:   lea edx, var_1B0
loc_00597A78:   push ecx
loc_00597A79:   lea eax, var_1A0
loc_00597A7F:   push edx
loc_00597A80:   lea ecx, var_190
loc_00597A86:   push eax
loc_00597A87:   lea edx, var_180
loc_00597A8D:   push ecx
loc_00597A8E:   lea eax, var_160
loc_00597A94:   push edx
loc_00597A95:   lea ecx, var_170
loc_00597A9B:   push eax
loc_00597A9C:   lea edx, var_150
loc_00597AA2:   push ecx
loc_00597AA3:   lea eax, var_140
loc_00597AA9:   push edx
loc_00597AAA:   lea ecx, var_130
loc_00597AB0:   push eax
loc_00597AB1:   lea edx, var_120
loc_00597AB7:   push ecx
loc_00597AB8:   lea eax, var_110
loc_00597ABE:   push edx
loc_00597ABF:   lea ecx, var_F0
loc_00597AC5:   push eax
loc_00597AC6:   push ecx
loc_00597AC7:   lea edx, var_100
loc_00597ACD:   lea eax, var_E0
loc_00597AD3:   push edx
loc_00597AD4:   lea ecx, var_C0
loc_00597ADA:   push eax
loc_00597ADB:   lea edx, var_D0
loc_00597AE1:   push ecx
loc_00597AE2:   lea eax, var_B0
loc_00597AE8:   push edx
loc_00597AE9:   lea ecx, var_A0
loc_00597AEF:   push eax
loc_00597AF0:   push ecx
  loc_00597AF1: push 00000018h
  loc_00597AF3: call [0040103Ch] ; __vbaFreeVarList
loc_00597AF9:   mov eax, [0066803Ch]
  loc_00597AFE: add esp, 000000A0h
loc_00597B04:   test eax, eax
  loc_00597B06: jnz 00597B18h
  loc_00597B08: push 0066803Ch
  loc_00597B0D: push 0042DEFCh
  loc_00597B12: call [004011E8h] ; __vbaNew2
loc_00597B18:   mov edx, [0066803Ch]
loc_00597B1E:   lea ecx, var_A0
loc_00597B24:   mov var_318, edx
  loc_00597B2A: mov var_98, 80020004h
  loc_00597B34: mov var_A0, 0000000Ah
  loc_00597B3E: call [0040123Ch] ; __vbaFreeVarg
loc_00597B44:   mov eax, var_318
loc_00597B4A:   lea edx, var_78
loc_00597B4D:   push edx
loc_00597B4E:   lea edx, var_A0
loc_00597B54:   mov ecx, [eax]
  loc_00597B56: push 00000001h
loc_00597B58:   push edx
loc_00597B59:   mov edx, [0066809Ch]
loc_00597B5F:   push edx
loc_00597B60:   push eax
loc_00597B61:   Call [ecx+00000040h]
loc_00597B64:   test eax, eax
loc_00597B66:   fnclex
  loc_00597B68: jge 00597B7Fh
loc_00597B6A:   mov ecx, var_318
  loc_00597B70: push 00000040h
  loc_00597B72: push 0042E1B0h
loc_00597B77:   push ecx
loc_00597B78:   push eax
  loc_00597B79: call [00401080h] ; __vbaHresultCheckObj
loc_00597B7F:   lea ecx, var_78
  loc_00597B82: call [004012A0h] ; __vbaFreeObj
loc_00597B88:   lea ecx, var_A0
  loc_00597B8E: call [00401020h] ; __vbaFreeVar
loc_00597B94:   mov ax, [ebx+00000034h]
  loc_00597B98: mov var_1C, 00000001h
  loc_00597B9F: sub ax, 0001h
  loc_00597BA3: jo 0059A014h
loc_00597BA9:   mov var_344, eax
loc_00597BAF:   mov dx, var_344
loc_00597BB6:   cmp var_1C, dx
  loc_00597BBA: jg 00598587h
loc_00597BC0:   movsx eax, var_1C
  loc_00597BC4: sub esp, 00000010h
  loc_00597BC7: mov ecx, 00000003h
loc_00597BCC:   mov edx, esp
loc_00597BCE:   mov var_220, ecx
loc_00597BD4:   mov var_238, ecx
loc_00597BDA:   mov var_240, ecx
loc_00597BE0:   mov [edx], ecx
loc_00597BE2:   mov ecx, var_21C
loc_00597BE8:   mov var_218, eax
  loc_00597BEE: sub esp, 00000010h
loc_00597BF1:   mov [edx+00000004h], ecx
loc_00597BF4:   mov ecx, esp
  loc_00597BF6: push 00000002h
  loc_00597BF8: push 00000041h
loc_00597BFA:   mov [edx+00000008h], eax
loc_00597BFD:   mov eax, var_214
loc_00597C03:   push ebx
loc_00597C04:   mov [edx+0000000Ch], eax
loc_00597C07:   mov edx, var_240
loc_00597C0D:   mov eax, var_23C
loc_00597C13:   mov [ecx], edx
loc_00597C15:   mov edx, var_238
loc_00597C1B:   mov [ecx+00000004h], eax
loc_00597C1E:   mov eax, var_234
loc_00597C24:   mov [ecx+00000008h], edx
loc_00597C27:   mov [ecx+0000000Ch], eax
loc_00597C2A:   mov ecx, [ebx]
loc_00597C2C:   Call [ecx+00000390h]
loc_00597C32:   lea edx, var_78
loc_00597C35:   push eax
loc_00597C36:   push edx
  loc_00597C37: call [004010B8h] ; __vbaObjSet
loc_00597C3D:   push eax
loc_00597C3E:   lea eax, var_A0
loc_00597C44:   push eax
  loc_00597C45: call [0040114Ch] ; __vbaLateIdCallLd
  loc_00597C4B: add esp, 00000030h
loc_00597C4E:   lea ecx, var_A0
  loc_00597C54: push 00000000h
loc_00597C56:   push FFFFFFFFh
  loc_00597C58: push 00000001h
  loc_00597C5A: push 0042DFECh
  loc_00597C5F: push 0043024Ch ; "."
loc_00597C64:   push ecx
  loc_00597C65: call [00401034h] ; __vbaStrVarMove
loc_00597C6B:   mov edx, eax
loc_00597C6D:   lea ecx, var_2C
  loc_00597C70: call __vbaVarDup
loc_00597C72:   push eax
  loc_00597C73: call [00401194h] ; rtcReplace
loc_00597C79:   mov edx, eax
loc_00597C7B:   lea ecx, var_18
  loc_00597C7E: call __vbaVarDup
loc_00597C80:   lea ecx, var_2C
  loc_00597C83: call [0040129Ch] ; __vbaFreeStr
loc_00597C89:   lea ecx, var_78
  loc_00597C8C: call [004012A0h] ; __vbaFreeObj
loc_00597C92:   lea ecx, var_A0
  loc_00597C98: call [00401020h] ; __vbaFreeVar
  loc_00597C9E: mov edx, 0042DFECh
  loc_00597CA3: mov ecx, 0066809Ch
  loc_00597CA8: call [004011FCh] ; __vbaStrCopy
loc_00597CAE:   movsx eax, var_1C
  loc_00597CB2: mov ecx, 00000003h
  loc_00597CB7: push 00434554h ; "INSERT INTO Pembelian_Detail"
  loc_00597CBC: push 00434594h ; "(no_pembelian,kode_barang,nama_barang,harga_beli,jumlah,sub_total)"
loc_00597CC1:   mov var_218, eax
loc_00597CC7:   mov var_220, ecx
  loc_00597CCD: mov var_238, 00000001h
loc_00597CD7:   mov var_240, ecx
loc_00597CDD:   mov var_258, eax
loc_00597CE3:   mov var_260, ecx
  loc_00597CE9: mov var_278, 00000002h
loc_00597CF3:   mov var_280, ecx
loc_00597CF9:   mov var_298, eax
loc_00597CFF:   mov var_2A0, ecx
loc_00597D05:   Call edi
loc_00597D07:   mov edx, eax
loc_00597D09:   lea ecx, var_2C
  loc_00597D0C: call __vbaVarDup
loc_00597D0E:   push eax
  loc_00597D0F: push 00434620h ; " VALUES ('"
loc_00597D14:   Call edi
loc_00597D16:   mov edx, eax
loc_00597D18:   lea ecx, var_30
  loc_00597D1B: call __vbaVarDup
loc_00597D1D:   mov edx, [ebx+00000040h]
loc_00597D20:   push eax
loc_00597D21:   push edx
loc_00597D22:   Call edi
loc_00597D24:   mov edx, eax
loc_00597D26:   lea ecx, var_34
  loc_00597D29: call __vbaVarDup
loc_00597D2B:   push eax
  loc_00597D2C: push 0042EC30h ; "','"
loc_00597D31:   Call edi
loc_00597D33:   mov edx, eax
loc_00597D35:   lea ecx, var_38
  loc_00597D38: call __vbaVarDup
loc_00597D3A:   mov ecx, var_220
loc_00597D40:   push eax
loc_00597D41:   mov edx, var_21C
  loc_00597D47: sub esp, 00000010h
loc_00597D4A:   mov eax, esp
  loc_00597D4C: sub esp, 00000010h
loc_00597D4F:   mov [eax], ecx
loc_00597D51:   mov ecx, var_218
loc_00597D57:   mov [eax+00000004h], edx
loc_00597D5A:   mov edx, var_214
loc_00597D60:   mov [eax+00000008h], ecx
loc_00597D63:   mov ecx, var_240
loc_00597D69:   mov [eax+0000000Ch], edx
loc_00597D6C:   mov edx, var_23C
loc_00597D72:   mov eax, esp
  loc_00597D74: push 00000002h
  loc_00597D76: push 00000041h
loc_00597D78:   push ebx
loc_00597D79:   mov [eax], ecx
loc_00597D7B:   mov ecx, var_238
loc_00597D81:   mov [eax+00000004h], edx
loc_00597D84:   mov edx, var_234
loc_00597D8A:   mov [eax+00000008h], ecx
loc_00597D8D:   mov [eax+0000000Ch], edx
loc_00597D90:   mov eax, [ebx]
loc_00597D92:   Call [eax+00000390h]
loc_00597D98:   lea ecx, var_78
loc_00597D9B:   push eax
loc_00597D9C:   push ecx
  loc_00597D9D: call [004010B8h] ; __vbaObjSet
loc_00597DA3:   lea edx, var_A0
loc_00597DA9:   push eax
loc_00597DAA:   push edx
  loc_00597DAB: call [0040114Ch] ; __vbaLateIdCallLd
  loc_00597DB1: add esp, 00000030h
loc_00597DB4:   push eax
  loc_00597DB5: call [00401034h] ; __vbaStrVarMove
loc_00597DBB:   mov edx, eax
loc_00597DBD:   lea ecx, var_3C
  loc_00597DC0: call __vbaVarDup
loc_00597DC2:   push eax
loc_00597DC3:   Call edi
loc_00597DC5:   mov edx, eax
loc_00597DC7:   lea ecx, var_40
  loc_00597DCA: call __vbaVarDup
loc_00597DCC:   push eax
  loc_00597DCD: push 0042EC30h ; "','"
loc_00597DD2:   Call edi
loc_00597DD4:   mov edx, eax
loc_00597DD6:   lea ecx, var_44
  loc_00597DD9: call __vbaVarDup
loc_00597DDB:   mov ecx, var_260
loc_00597DE1:   push eax
loc_00597DE2:   mov edx, var_25C
  loc_00597DE8: sub esp, 00000010h
loc_00597DEB:   mov eax, esp
loc_00597DED:   mov [eax], ecx
loc_00597DEF:   mov ecx, var_258
loc_00597DF5:   mov [eax+00000004h], edx
loc_00597DF8:   mov edx, var_254
loc_00597DFE:   mov [eax+00000008h], ecx
loc_00597E01:   mov [eax+0000000Ch], edx
loc_00597E04:   mov ecx, var_280
loc_00597E0A:   mov edx, var_27C
  loc_00597E10: sub esp, 00000010h
loc_00597E13:   mov eax, esp
  loc_00597E15: push 00000002h
  loc_00597E17: push 00000041h
loc_00597E19:   mov [eax], ecx
loc_00597E1B:   mov ecx, var_278
loc_00597E21:   push ebx
loc_00597E22:   mov [eax+00000004h], edx
loc_00597E25:   mov edx, var_274
loc_00597E2B:   mov [eax+00000008h], ecx
loc_00597E2E:   mov [eax+0000000Ch], edx
loc_00597E31:   mov eax, [ebx]
loc_00597E33:   Call [eax+00000390h]
loc_00597E39:   lea ecx, var_7C
loc_00597E3C:   push eax
loc_00597E3D:   push ecx
  loc_00597E3E: call [004010B8h] ; __vbaObjSet
loc_00597E44:   lea edx, var_B0
loc_00597E4A:   push eax
loc_00597E4B:   push edx
  loc_00597E4C: call [0040114Ch] ; __vbaLateIdCallLd
  loc_00597E52: add esp, 00000030h
loc_00597E55:   push eax
  loc_00597E56: call [00401034h] ; __vbaStrVarMove
loc_00597E5C:   mov edx, eax
loc_00597E5E:   lea ecx, var_48
  loc_00597E61: call __vbaVarDup
loc_00597E63:   push eax
loc_00597E64:   Call edi
loc_00597E66:   mov edx, eax
loc_00597E68:   lea ecx, var_4C
  loc_00597E6B: call __vbaVarDup
loc_00597E6D:   push eax
  loc_00597E6E: push 0042EC30h ; "','"
loc_00597E73:   Call edi
loc_00597E75:   mov edx, eax
loc_00597E77:   lea ecx, var_50
  loc_00597E7A: call __vbaVarDup
loc_00597E7C:   push eax
loc_00597E7D:   mov eax, var_18
loc_00597E80:   push eax
  loc_00597E81: call [004012A4h] ; rtcR8ValFromBstr
  loc_00597E87: sub esp, 00000008h
  loc_00597E8A: fstp real8 ptr [esp]
  loc_00597E8D: call [0040115Ch] ; __vbaStrR8
loc_00597E93:   mov edx, eax
loc_00597E95:   lea ecx, var_54
  loc_00597E98: call __vbaVarDup
loc_00597E9A:   push eax
loc_00597E9B:   Call edi
loc_00597E9D:   mov edx, eax
loc_00597E9F:   lea ecx, var_58
  loc_00597EA2: call __vbaVarDup
loc_00597EA4:   push eax
  loc_00597EA5: push 0042EC30h ; "','"
loc_00597EAA:   Call edi
loc_00597EAC:   mov edx, eax
loc_00597EAE:   lea ecx, var_60
  loc_00597EB1: call __vbaVarDup
loc_00597EB3:   mov edx, var_2A0
loc_00597EB9:   push eax
loc_00597EBA:   mov eax, var_29C
  loc_00597EC0: sub esp, 00000010h
loc_00597EC3:   mov ecx, esp
  loc_00597EC5: sub esp, 00000010h
loc_00597EC8:   mov [ecx], edx
loc_00597ECA:   mov edx, var_298
loc_00597ED0:   mov [ecx+00000004h], eax
loc_00597ED3:   mov eax, var_294
loc_00597ED9:   mov [ecx+00000008h], edx
loc_00597EDC:   mov edx, var_2BC
loc_00597EE2:   mov [ecx+0000000Ch], eax
loc_00597EE5:   mov ecx, esp
  loc_00597EE7: mov eax, 00000003h
loc_00597EEC:   mov [ecx], eax
  loc_00597EEE: mov eax, 00000005h
loc_00597EF3:   mov [ecx+00000004h], edx
loc_00597EF6:   mov [ecx+00000008h], eax
loc_00597EF9:   mov eax, var_2B4
loc_00597EFF:   mov [ecx+0000000Ch], eax
loc_00597F02:   mov ecx, [ebx]
  loc_00597F04: push 00000002h
  loc_00597F06: push 00000041h
loc_00597F08:   push ebx
loc_00597F09:   Call [ecx+00000390h]
loc_00597F0F:   lea edx, var_80
loc_00597F12:   push eax
loc_00597F13:   push edx
  loc_00597F14: call [004010B8h] ; __vbaObjSet
loc_00597F1A:   push eax
loc_00597F1B:   lea eax, var_C0
loc_00597F21:   push eax
  loc_00597F22: call [0040114Ch] ; __vbaLateIdCallLd
  loc_00597F28: add esp, 00000030h
loc_00597F2B:   push eax
  loc_00597F2C: call [00401034h] ; __vbaStrVarMove
loc_00597F32:   mov edx, eax
loc_00597F34:   lea ecx, var_5C
  loc_00597F37: call __vbaVarDup
loc_00597F39:   push eax
  loc_00597F3A: call [004012A4h] ; rtcR8ValFromBstr
  loc_00597F40: sub esp, 00000008h
  loc_00597F43: fstp real8 ptr [esp]
  loc_00597F46: call [0040115Ch] ; __vbaStrR8
loc_00597F4C:   mov edx, eax
loc_00597F4E:   lea ecx, var_64
  loc_00597F51: call __vbaVarDup
loc_00597F53:   push eax
loc_00597F54:   Call edi
loc_00597F56:   mov edx, eax
loc_00597F58:   lea ecx, var_68
  loc_00597F5B: call __vbaVarDup
loc_00597F5D:   push eax
  loc_00597F5E: push 0042EC30h ; "','"
loc_00597F63:   Call edi
loc_00597F65:   mov edx, eax
loc_00597F67:   lea ecx, var_6C
  loc_00597F6A: call __vbaVarDup
loc_00597F6C:   push eax
loc_00597F6D:   mov edx, var_2DC
  loc_00597F73: sub esp, 00000010h
  loc_00597F76: mov eax, 00000003h
loc_00597F7B:   mov ecx, esp
  loc_00597F7D: sub esp, 00000010h
loc_00597F80:   mov [ecx], eax
loc_00597F82:   movsx eax, var_1C
loc_00597F86:   mov [ecx+00000004h], edx
loc_00597F89:   mov edx, var_2FC
loc_00597F8F:   mov [ecx+00000008h], eax
loc_00597F92:   mov eax, var_2D4
loc_00597F98:   mov [ecx+0000000Ch], eax
loc_00597F9B:   mov ecx, esp
  loc_00597F9D: mov eax, 00000003h
  loc_00597FA2: push 00000002h
loc_00597FA4:   mov [ecx], eax
  loc_00597FA6: mov eax, 00000006h
  loc_00597FAB: push 00000041h
loc_00597FAD:   push ebx
loc_00597FAE:   mov [ecx+00000004h], edx
loc_00597FB1:   mov [ecx+00000008h], eax
loc_00597FB4:   mov eax, var_2F4
loc_00597FBA:   mov [ecx+0000000Ch], eax
loc_00597FBD:   mov ecx, [ebx]
loc_00597FBF:   Call [ecx+00000390h]
loc_00597FC5:   lea edx, var_84
loc_00597FCB:   push eax
loc_00597FCC:   push edx
  loc_00597FCD: call [004010B8h] ; __vbaObjSet
loc_00597FD3:   push eax
loc_00597FD4:   lea eax, var_D0
loc_00597FDA:   push eax
  loc_00597FDB: call [0040114Ch] ; __vbaLateIdCallLd
  loc_00597FE1: add esp, 00000030h
loc_00597FE4:   push eax
  loc_00597FE5: call [00401034h] ; __vbaStrVarMove
loc_00597FEB:   mov edx, eax
loc_00597FED:   lea ecx, var_70
  loc_00597FF0: call __vbaVarDup
loc_00597FF2:   push eax
loc_00597FF3:   Call edi
loc_00597FF5:   mov edx, eax
loc_00597FF7:   lea ecx, var_74
  loc_00597FFA: call __vbaVarDup
loc_00597FFC:   push eax
  loc_00597FFD: push 0042EC3Ch ; "')"
loc_00598002:   Call edi
loc_00598004:   mov edx, eax
  loc_00598006: mov ecx, 0066809Ch
  loc_0059800B: call __vbaVarDup
loc_0059800D:   lea ecx, var_74
loc_00598010:   lea edx, var_70
loc_00598013:   push ecx
loc_00598014:   lea eax, var_6C
loc_00598017:   push edx
loc_00598018:   lea ecx, var_68
loc_0059801B:   push eax
loc_0059801C:   lea edx, var_64
loc_0059801F:   push ecx
loc_00598020:   lea eax, var_60
loc_00598023:   push edx
loc_00598024:   lea ecx, var_5C
loc_00598027:   push eax
loc_00598028:   lea edx, var_58
loc_0059802B:   push ecx
loc_0059802C:   lea eax, var_54
loc_0059802F:   push edx
loc_00598030:   lea ecx, var_50
loc_00598033:   push eax
loc_00598034:   lea edx, var_4C
loc_00598037:   push ecx
loc_00598038:   lea eax, var_48
loc_0059803B:   push edx
loc_0059803C:   lea ecx, var_44
loc_0059803F:   push eax
loc_00598040:   lea edx, var_40
loc_00598043:   push ecx
loc_00598044:   lea eax, var_3C
loc_00598047:   push edx
loc_00598048:   lea ecx, var_38
loc_0059804B:   push eax
loc_0059804C:   lea edx, var_34
loc_0059804F:   push ecx
loc_00598050:   lea eax, var_30
loc_00598053:   push edx
loc_00598054:   lea ecx, var_2C
loc_00598057:   push eax
loc_00598058:   push ecx
  loc_00598059: push 00000013h
  loc_0059805B: call [00401204h] ; __vbaFreeStrList
loc_00598061:   lea edx, var_84
loc_00598067:   lea eax, var_80
loc_0059806A:   push edx
loc_0059806B:   lea ecx, var_7C
loc_0059806E:   push eax
loc_0059806F:   lea edx, var_78
loc_00598072:   push ecx
loc_00598073:   push edx
  loc_00598074: push 00000004h
  loc_00598076: call [0040104Ch] ; __vbaFreeObjList
  loc_0059807C: add esp, 00000064h
loc_0059807F:   lea eax, var_D0
loc_00598085:   lea ecx, var_C0
loc_0059808B:   lea edx, var_B0
loc_00598091:   push eax
loc_00598092:   push ecx
loc_00598093:   lea eax, var_A0
loc_00598099:   push edx
loc_0059809A:   push eax
  loc_0059809B: push 00000004h
  loc_0059809D: call [0040103Ch] ; __vbaFreeVarList
loc_005980A3:   mov eax, [0066803Ch]
  loc_005980A8: add esp, 00000014h
loc_005980AB:   test eax, eax
  loc_005980AD: jnz 005980BFh
  loc_005980AF: push 0066803Ch
  loc_005980B4: push 0042DEFCh
  loc_005980B9: call [004011E8h] ; __vbaNew2
loc_005980BF:   mov ecx, [0066803Ch]
  loc_005980C5: mov var_98, 80020004h
loc_005980CF:   mov var_318, ecx
loc_005980D5:   lea ecx, var_A0
  loc_005980DB: mov var_A0, 0000000Ah
  loc_005980E5: call [0040123Ch] ; __vbaFreeVarg
loc_005980EB:   mov eax, var_318
loc_005980F1:   lea ecx, var_78
loc_005980F4:   push ecx
loc_005980F5:   lea ecx, var_A0
loc_005980FB:   mov edx, [eax]
  loc_005980FD: push 00000001h
loc_005980FF:   push ecx
loc_00598100:   mov ecx, [0066809Ch]
loc_00598106:   push ecx
loc_00598107:   push eax
loc_00598108:   Call [edx+00000040h]
loc_0059810B:   test eax, eax
loc_0059810D:   fnclex
  loc_0059810F: jge 00598126h
loc_00598111:   mov edx, var_318
  loc_00598117: push 00000040h
  loc_00598119: push 0042E1B0h
loc_0059811E:   push edx
loc_0059811F:   push eax
  loc_00598120: call [00401080h] ; __vbaHresultCheckObj
loc_00598126:   lea ecx, var_78
  loc_00598129: call [004012A0h] ; __vbaFreeObj
loc_0059812F:   lea ecx, var_A0
  loc_00598135: call [00401020h] ; __vbaFreeVar
loc_0059813B:   movsx eax, var_1C
  loc_0059813F: sub esp, 00000010h
  loc_00598142: mov ecx, 00000003h
loc_00598147:   mov edx, esp
loc_00598149:   mov var_220, ecx
loc_0059814F:   mov var_240, ecx
loc_00598155:   mov var_218, eax
loc_0059815B:   mov [edx], ecx
loc_0059815D:   mov ecx, var_21C
  loc_00598163: sub esp, 00000010h
  loc_00598166: mov var_238, 00000005h
loc_00598170:   mov [edx+00000004h], ecx
loc_00598173:   mov ecx, esp
  loc_00598175: push 00000002h
  loc_00598177: push 00000041h
loc_00598179:   mov [edx+00000008h], eax
loc_0059817C:   mov eax, var_214
loc_00598182:   push ebx
loc_00598183:   mov [edx+0000000Ch], eax
loc_00598186:   mov edx, var_240
loc_0059818C:   mov eax, var_23C
loc_00598192:   mov [ecx], edx
loc_00598194:   mov edx, var_238
loc_0059819A:   mov [ecx+00000004h], eax
loc_0059819D:   mov eax, var_234
loc_005981A3:   mov [ecx+00000008h], edx
loc_005981A6:   mov [ecx+0000000Ch], eax
loc_005981A9:   mov ecx, [ebx]
loc_005981AB:   Call [ecx+00000390h]
loc_005981B1:   lea edx, var_78
loc_005981B4:   push eax
loc_005981B5:   push edx
  loc_005981B6: call [004010B8h] ; __vbaObjSet
loc_005981BC:   push eax
loc_005981BD:   lea eax, var_A0
loc_005981C3:   push eax
  loc_005981C4: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005981CA: add esp, 00000030h
loc_005981CD:   push eax
  loc_005981CE: call [00401034h] ; __vbaStrVarMove
loc_005981D4:   mov edx, eax
loc_005981D6:   lea ecx, var_2C
  loc_005981D9: call __vbaVarDup
loc_005981DB:   push eax
  loc_005981DC: call [004012A4h] ; rtcR8ValFromBstr
  loc_005981E2: call [00401244h] ; __vbaFpI2
loc_005981E8:   lea ecx, var_2C
loc_005981EB:   mov var_28, eax
  loc_005981EE: call [0040129Ch] ; __vbaFreeStr
loc_005981F4:   lea ecx, var_78
  loc_005981F7: call [004012A0h] ; __vbaFreeObj
loc_005981FD:   lea ecx, var_A0
  loc_00598203: call [00401020h] ; __vbaFreeVar
loc_00598209:   movsx eax, var_1C
  loc_0059820D: sub esp, 00000010h
  loc_00598210: mov ecx, 00000003h
loc_00598215:   mov edx, esp
loc_00598217:   mov var_220, ecx
loc_0059821D:   mov var_240, ecx
loc_00598223:   mov var_218, eax
loc_00598229:   mov [edx], ecx
loc_0059822B:   mov ecx, var_21C
  loc_00598231: sub esp, 00000010h
  loc_00598234: mov var_238, 00000001h
loc_0059823E:   mov [edx+00000004h], ecx
loc_00598241:   mov ecx, esp
loc_00598243:   mov [edx+00000008h], eax
loc_00598246:   mov eax, var_214
loc_0059824C:   mov [edx+0000000Ch], eax
loc_0059824F:   mov edx, var_240
loc_00598255:   mov eax, var_23C
loc_0059825B:   mov [ecx], edx
loc_0059825D:   mov edx, var_238
loc_00598263:   mov [ecx+00000004h], eax
loc_00598266:   mov eax, var_234
loc_0059826C:   mov [ecx+00000008h], edx
loc_0059826F:   mov [ecx+0000000Ch], eax
loc_00598272:   mov ecx, [ebx]
  loc_00598274: push 00000002h
  loc_00598276: push 00000041h
loc_00598278:   push ebx
loc_00598279:   Call [ecx+00000390h]
loc_0059827F:   lea edx, var_78
loc_00598282:   push eax
loc_00598283:   push edx
  loc_00598284: call [004010B8h] ; __vbaObjSet
loc_0059828A:   push eax
loc_0059828B:   lea eax, var_A0
loc_00598291:   push eax
  loc_00598292: call [0040114Ch] ; __vbaLateIdCallLd
  loc_00598298: add esp, 00000030h
loc_0059829B:   push eax
  loc_0059829C: call [00401034h] ; __vbaStrVarMove
loc_005982A2:   mov edx, eax
loc_005982A4:   lea ecx, var_20
  loc_005982A7: call __vbaVarDup
loc_005982A9:   lea ecx, var_78
  loc_005982AC: call [004012A0h] ; __vbaFreeObj
loc_005982B2:   lea ecx, var_A0
  loc_005982B8: call [00401020h] ; __vbaFreeVar
loc_005982BE:   movsx eax, var_1C
  loc_005982C2: sub esp, 00000010h
  loc_005982C5: mov ecx, 00000003h
loc_005982CA:   mov edx, esp
loc_005982CC:   mov var_220, ecx
loc_005982D2:   mov var_240, ecx
loc_005982D8:   mov var_218, eax
loc_005982DE:   mov [edx], ecx
loc_005982E0:   mov ecx, var_21C
  loc_005982E6: sub esp, 00000010h
  loc_005982E9: mov var_238, 00000004h
loc_005982F3:   mov [edx+00000004h], ecx
loc_005982F6:   mov ecx, esp
  loc_005982F8: push 00000002h
  loc_005982FA: push 00000041h
loc_005982FC:   mov [edx+00000008h], eax
loc_005982FF:   mov eax, var_214
loc_00598305:   push ebx
loc_00598306:   mov [edx+0000000Ch], eax
loc_00598309:   mov edx, var_240
loc_0059830F:   mov eax, var_23C
loc_00598315:   mov [ecx], edx
loc_00598317:   mov edx, var_238
loc_0059831D:   mov [ecx+00000004h], eax
loc_00598320:   mov eax, var_234
loc_00598326:   mov [ecx+00000008h], edx
loc_00598329:   mov [ecx+0000000Ch], eax
loc_0059832C:   mov ecx, [ebx]
loc_0059832E:   Call [ecx+00000390h]
loc_00598334:   lea edx, var_78
loc_00598337:   push eax
loc_00598338:   push edx
  loc_00598339: call [004010B8h] ; __vbaObjSet
loc_0059833F:   push eax
loc_00598340:   lea eax, var_A0
loc_00598346:   push eax
  loc_00598347: call [0040114Ch] ; __vbaLateIdCallLd
  loc_0059834D: add esp, 00000030h
loc_00598350:   lea ecx, var_A0
  loc_00598356: push 00000000h
loc_00598358:   push FFFFFFFFh
  loc_0059835A: push 00000001h
  loc_0059835C: push 0042DFECh
  loc_00598361: push 0043024Ch ; "."
loc_00598366:   push ecx
  loc_00598367: call [00401034h] ; __vbaStrVarMove
loc_0059836D:   mov edx, eax
loc_0059836F:   lea ecx, var_2C
  loc_00598372: call __vbaVarDup
loc_00598374:   push eax
  loc_00598375: call [00401194h] ; rtcReplace
loc_0059837B:   mov edx, eax
loc_0059837D:   lea ecx, var_24
  loc_00598380: call __vbaVarDup
loc_00598382:   lea ecx, var_2C
  loc_00598385: call [0040129Ch] ; __vbaFreeStr
loc_0059838B:   lea ecx, var_78
  loc_0059838E: call [004012A0h] ; __vbaFreeObj
loc_00598394:   lea ecx, var_A0
  loc_0059839A: call [00401020h] ; __vbaFreeVar
  loc_005983A0: push 0043463Ch ; "UPDATE brg_barang SET "
  loc_005983A5: push 00434670h ; " stok=stok + "
loc_005983AA:   Call edi
loc_005983AC:   mov edx, eax
loc_005983AE:   lea ecx, var_2C
  loc_005983B1: call __vbaVarDup
loc_005983B3:   mov edx, var_28
loc_005983B6:   push eax
loc_005983B7:   push edx
  loc_005983B8: call [0040100Ch] ; __vbaStrI2
loc_005983BE:   mov edx, eax
loc_005983C0:   lea ecx, var_30
  loc_005983C3: call __vbaVarDup
loc_005983C5:   push eax
loc_005983C6:   Call edi
loc_005983C8:   mov edx, eax
loc_005983CA:   lea ecx, var_34
  loc_005983CD: call __vbaVarDup
loc_005983CF:   push eax
  loc_005983D0: push 00433074h
loc_005983D5:   Call edi
loc_005983D7:   mov edx, eax
loc_005983D9:   lea ecx, var_38
  loc_005983DC: call __vbaVarDup
loc_005983DE:   push eax
  loc_005983DF: push 00434690h ; " harga_beli='"
loc_005983E4:   Call edi
loc_005983E6:   mov edx, eax
loc_005983E8:   lea ecx, var_3C
  loc_005983EB: call __vbaVarDup
loc_005983ED:   push eax
loc_005983EE:   mov eax, var_18
loc_005983F1:   push eax
  loc_005983F2: call [004012A4h] ; rtcR8ValFromBstr
  loc_005983F8: sub esp, 00000008h
  loc_005983FB: fstp real8 ptr [esp]
  loc_005983FE: call [0040115Ch] ; __vbaStrR8
loc_00598404:   mov edx, eax
loc_00598406:   lea ecx, var_40
  loc_00598409: call __vbaVarDup
loc_0059840B:   push eax
loc_0059840C:   Call edi
loc_0059840E:   mov edx, eax
loc_00598410:   lea ecx, var_44
  loc_00598413: call __vbaVarDup
loc_00598415:   push eax
  loc_00598416: push 004346B0h ; "',"
loc_0059841B:   Call edi
loc_0059841D:   mov edx, eax
loc_0059841F:   lea ecx, var_48
  loc_00598422: call __vbaVarDup
loc_00598424:   push eax
  loc_00598425: push 00432938h ; " harga_jual='"
loc_0059842A:   Call edi
loc_0059842C:   mov edx, eax
loc_0059842E:   lea ecx, var_4C
  loc_00598431: call __vbaVarDup
loc_00598433:   mov ecx, var_24
loc_00598436:   push eax
loc_00598437:   push ecx
  loc_00598438: call [004012A4h] ; rtcR8ValFromBstr
  loc_0059843E: sub esp, 00000008h
  loc_00598441: fstp real8 ptr [esp]
  loc_00598444: call [0040115Ch] ; __vbaStrR8
loc_0059844A:   mov edx, eax
loc_0059844C:   lea ecx, var_50
  loc_0059844F: call __vbaVarDup
loc_00598451:   push eax
loc_00598452:   Call edi
loc_00598454:   mov edx, eax
loc_00598456:   lea ecx, var_54
  loc_00598459: call __vbaVarDup
loc_0059845B:   push eax
  loc_0059845C: push 0042EBA8h ; "'"
loc_00598461:   Call edi
loc_00598463:   mov edx, eax
loc_00598465:   lea ecx, var_58
  loc_00598468: call __vbaVarDup
loc_0059846A:   push eax
  loc_0059846B: push 00433824h ; " WHERE kode_barang='"
loc_00598470:   Call edi
loc_00598472:   mov edx, eax
loc_00598474:   lea ecx, var_5C
  loc_00598477: call __vbaVarDup
loc_00598479:   mov edx, var_20
loc_0059847C:   push eax
loc_0059847D:   push edx
loc_0059847E:   Call edi
loc_00598480:   mov edx, eax
loc_00598482:   lea ecx, var_60
  loc_00598485: call __vbaVarDup
loc_00598487:   push eax
  loc_00598488: push 0042EBA8h ; "'"
loc_0059848D:   Call edi
loc_0059848F:   mov edx, eax
  loc_00598491: mov ecx, 0066809Ch
  loc_00598496: call __vbaVarDup
loc_00598498:   lea eax, var_60
loc_0059849B:   lea ecx, var_5C
loc_0059849E:   push eax
loc_0059849F:   lea edx, var_58
loc_005984A2:   push ecx
loc_005984A3:   lea eax, var_54
loc_005984A6:   push edx
loc_005984A7:   lea ecx, var_50
loc_005984AA:   push eax
loc_005984AB:   lea edx, var_4C
loc_005984AE:   push ecx
loc_005984AF:   lea eax, var_48
loc_005984B2:   push edx
loc_005984B3:   lea ecx, var_44
loc_005984B6:   push eax
loc_005984B7:   lea edx, var_40
loc_005984BA:   push ecx
loc_005984BB:   lea eax, var_3C
loc_005984BE:   push edx
loc_005984BF:   lea ecx, var_38
loc_005984C2:   push eax
loc_005984C3:   lea edx, var_34
loc_005984C6:   push ecx
loc_005984C7:   lea eax, var_30
loc_005984CA:   push edx
loc_005984CB:   lea ecx, var_2C
loc_005984CE:   push eax
loc_005984CF:   push ecx
  loc_005984D0: push 0000000Eh
  loc_005984D2: call [00401204h] ; __vbaFreeStrList
loc_005984D8:   mov eax, [0066803Ch]
  loc_005984DD: add esp, 0000003Ch
loc_005984E0:   test eax, eax
  loc_005984E2: jnz 005984F4h
  loc_005984E4: push 0066803Ch
  loc_005984E9: push 0042DEFCh
  loc_005984EE: call [004011E8h] ; __vbaNew2
loc_005984F4:   mov edx, [0066803Ch]
loc_005984FA:   lea ecx, var_A0
loc_00598500:   mov var_318, edx
  loc_00598506: mov var_98, 80020004h
  loc_00598510: mov var_A0, 0000000Ah
  loc_0059851A: call [0040123Ch] ; __vbaFreeVarg
loc_00598520:   mov eax, var_318
loc_00598526:   lea edx, var_78
loc_00598529:   push edx
loc_0059852A:   lea edx, var_A0
loc_00598530:   mov ecx, [eax]
  loc_00598532: push 00000001h
loc_00598534:   push edx
loc_00598535:   mov edx, [0066809Ch]
loc_0059853B:   push edx
loc_0059853C:   push eax
loc_0059853D:   Call [ecx+00000040h]
loc_00598540:   test eax, eax
loc_00598542:   fnclex
  loc_00598544: jge 0059855Bh
loc_00598546:   mov ecx, var_318
  loc_0059854C: push 00000040h
  loc_0059854E: push 0042E1B0h
loc_00598553:   push ecx
loc_00598554:   push eax
  loc_00598555: call [00401080h] ; __vbaHresultCheckObj
loc_0059855B:   lea ecx, var_78
  loc_0059855E: call [004012A0h] ; __vbaFreeObj
loc_00598564:   lea ecx, var_A0
  loc_0059856A: call [00401020h] ; __vbaFreeVar
  loc_00598570: mov eax, 00000001h
loc_00598575:   Add ax, var_1C
  loc_00598579: jo 0059A014h
loc_0059857F:   mov var_1C, eax
  loc_00598582: jmp 00597BAFh
  loc_00598587: mov esi, [00401238h] ; __vbaVarDup
  loc_0059858D: mov ecx, 80020004h
loc_00598592:   mov var_C8, ecx
  loc_00598598: mov eax, 0000000Ah
loc_0059859D:   mov var_B8, ecx
  loc_005985A3: mov edi, 00000008h
loc_005985A8:   lea edx, var_230
loc_005985AE:   lea ecx, var_B0
loc_005985B4:   mov var_D0, eax
loc_005985BA:   mov var_C0, eax
  loc_005985C0: mov var_228, 0042DF8Ch ; "Info"
loc_005985CA:   mov var_230, edi
  loc_005985D0: call __vbaVarDup
loc_005985D2:   lea edx, var_220
loc_005985D8:   lea ecx, var_A0
  loc_005985DE: mov var_218, 004346BCh ; "DATA TRANSAKSI TELAH TERSIMPAN !"
loc_005985E8:   mov var_220, edi
  loc_005985EE: call __vbaVarDup
loc_005985F0:   lea edx, var_D0
loc_005985F6:   lea eax, var_C0
loc_005985FC:   push edx
loc_005985FD:   lea ecx, var_B0
loc_00598603:   push eax
loc_00598604:   push ecx
loc_00598605:   lea edx, var_A0
  loc_0059860B: push 00000040h
loc_0059860D:   push edx
  loc_0059860E: call [004010B4h] ; rtcMsgBox
loc_00598614:   lea eax, var_D0
loc_0059861A:   lea ecx, var_C0
loc_00598620:   push eax
loc_00598621:   lea edx, var_B0
loc_00598627:   push ecx
loc_00598628:   lea eax, var_A0
loc_0059862E:   push edx
loc_0059862F:   push eax
  loc_00598630: push 00000004h
  loc_00598632: call [0040103Ch] ; __vbaFreeVarList
loc_00598638:   mov ecx, [ebx]
  loc_0059863A: add esp, 00000014h
loc_0059863D:   push ebx
loc_0059863E:   Call [ecx+00000704h]
loc_00598644:   test eax, eax
  loc_00598646: jge 0059865Ah
  loc_00598648: push 00000704h
  loc_0059864D: push 00433DD8h
loc_00598652:   push ebx
loc_00598653:   push eax
  loc_00598654: call [00401080h] ; __vbaHresultCheckObj
loc_0059865A:   mov edx, [ebx]
loc_0059865C:   push ebx
loc_0059865D:   Call [edx+00000728h]
loc_00598663:   test eax, eax
  loc_00598665: jge 00599E92h
  loc_0059866B: push 00000728h
  loc_00598670: push 00433DD8h
loc_00598675:   push ebx
loc_00598676:   push eax
  loc_00598677: call [00401080h] ; __vbaHresultCheckObj
  loc_0059867D: jmp 00599E92h
  loc_00598682: call 0055B320h
loc_00598687:   mov eax, [ebx]
loc_00598689:   push ebx
loc_0059868A:   Call [eax+00000708h]
loc_00598690:   test eax, eax
  loc_00598692: jge 005986A6h
  loc_00598694: push 00000708h
  loc_00598699: push 00433DD8h
loc_0059869E:   push ebx
loc_0059869F:   push eax
  loc_005986A0: call [00401080h] ; __vbaHresultCheckObj
loc_005986A6:   lea ecx, var_A0
loc_005986AC:   push ecx
  loc_005986AD: call [00401220h] ; rtcGetDateVar
loc_005986B3:   mov edx, [ebx]
loc_005986B5:   push ebx
loc_005986B6:   Call [edx+0000031Ch]
loc_005986BC:   push eax
loc_005986BD:   lea eax, var_78
loc_005986C0:   push eax
  loc_005986C1: call [004010B8h] ; __vbaObjSet
loc_005986C7:   mov ecx, [eax]
loc_005986C9:   lea edx, var_38
loc_005986CC:   push edx
loc_005986CD:   push eax
loc_005986CE:   mov var_318, eax
loc_005986D4:   Call [ecx+000000A0h]
loc_005986DA:   test eax, eax
loc_005986DC:   fnclex
  loc_005986DE: jge 005986F8h
loc_005986E0:   mov ecx, var_318
  loc_005986E6: push 000000A0h
  loc_005986EB: push 0042DFCCh
loc_005986F0:   push ecx
loc_005986F1:   push eax
  loc_005986F2: call [00401080h] ; __vbaHresultCheckObj
  loc_005986F8: push 00434418h ; "INSERT INTO pembelian"
  loc_005986FD: push 00434460h ; "(no_pembelian,no_notabeli, tgl_masuk,kode_pemasok,nama_pemasok,jenis,userid,catatan,tot_beli)"
loc_00598702:   Call edi
loc_00598704:   mov edx, eax
loc_00598706:   lea ecx, var_2C
  loc_00598709: call __vbaVarDup
loc_0059870B:   push eax
  loc_0059870C: push 00434520h ; "VALUES ('"
loc_00598711:   Call edi
loc_00598713:   mov edx, eax
loc_00598715:   lea ecx, var_30
  loc_00598718: call __vbaVarDup
loc_0059871A:   mov edx, [ebx+00000040h]
loc_0059871D:   push eax
loc_0059871E:   push edx
loc_0059871F:   Call edi
loc_00598721:   mov edx, eax
loc_00598723:   lea ecx, var_34
  loc_00598726: call __vbaVarDup
loc_00598728:   push eax
  loc_00598729: push 0042EC30h ; "','"
loc_0059872E:   Call edi
loc_00598730:   mov edx, eax
loc_00598732:   lea ecx, var_3C
  loc_00598735: call __vbaVarDup
loc_00598737:   push eax
loc_00598738:   mov eax, var_38
loc_0059873B:   push eax
loc_0059873C:   Call edi
loc_0059873E:   mov edx, eax
loc_00598740:   lea ecx, var_40
  loc_00598743: call __vbaVarDup
loc_00598745:   push eax
  loc_00598746: push 0042EC30h ; "','"
loc_0059874B:   Call edi
loc_0059874D:   mov var_C8, eax
  loc_00598753: mov eax, 00000008h
loc_00598758:   lea edx, var_220
loc_0059875E:   lea ecx, var_B0
loc_00598764:   mov var_D0, eax
  loc_0059876A: mov var_218, 00434538h ; "yyyy/MM/dd"
loc_00598774:   mov var_220, eax
  loc_0059877A: call [00401238h] ; __vbaVarDup
  loc_00598780: push 00000001h
loc_00598782:   lea ecx, var_B0
  loc_00598788: push 00000001h
loc_0059878A:   lea edx, var_A0
loc_00598790:   push ecx
loc_00598791:   lea eax, var_C0
loc_00598797:   push edx
loc_00598798:   push eax
  loc_00598799: call [00401078h] ; rtcVarFromFormatVar
loc_0059879F:   mov ecx, [ebx]
  loc_005987A1: push 00000000h
loc_005987A3:   push FFFFFDFBh
loc_005987A8:   push ebx
  loc_005987A9: mov var_228, 0042EC30h ; "','"
  loc_005987B3: mov var_230, 00000008h
loc_005987BD:   Call [ecx+00000384h]
loc_005987C3:   lea edx, var_7C
loc_005987C6:   push eax
loc_005987C7:   push edx
  loc_005987C8: call [004010B8h] ; __vbaObjSet
loc_005987CE:   push eax
loc_005987CF:   lea eax, var_100
loc_005987D5:   push eax
  loc_005987D6: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005987DC: add esp, 00000010h
loc_005987DF:   push eax
  loc_005987E0: call [00401034h] ; __vbaStrVarMove
loc_005987E6:   mov ecx, [ebx]
loc_005987E8:   mov var_108, eax
  loc_005987EE: mov eax, 00000008h
loc_005987F3:   push ebx
loc_005987F4:   mov var_110, eax
  loc_005987FA: mov var_238, 0042EC30h ; "','"
loc_00598804:   mov var_240, eax
loc_0059880A:   Call [ecx+00000320h]
loc_00598810:   lea edx, var_80
loc_00598813:   push eax
loc_00598814:   push edx
  loc_00598815: call [004010B8h] ; __vbaObjSet
loc_0059881B:   mov ecx, [eax]
loc_0059881D:   lea edx, var_44
loc_00598820:   push edx
loc_00598821:   push eax
loc_00598822:   mov var_320, eax
loc_00598828:   Call [ecx+000000A0h]
loc_0059882E:   test eax, eax
loc_00598830:   fnclex
  loc_00598832: jge 0059884Ch
loc_00598834:   mov ecx, var_320
  loc_0059883A: push 000000A0h
  loc_0059883F: push 0042DFCCh
loc_00598844:   push ecx
loc_00598845:   push eax
  loc_00598846: call [00401080h] ; __vbaHresultCheckObj
loc_0059884C:   mov eax, var_44
loc_0059884F:   mov edx, [ebx]
loc_00598851:   mov var_138, eax
  loc_00598857: push 00000000h
  loc_00598859: mov eax, 00000008h
loc_0059885E:   push FFFFFDFBh
loc_00598863:   push ebx
  loc_00598864: mov var_44, 00000000h
loc_0059886B:   mov var_140, eax
  loc_00598871: mov var_248, 0042EC30h ; "','"
loc_0059887B:   mov var_250, eax
loc_00598881:   Call [edx+00000380h]
loc_00598887:   push eax
loc_00598888:   lea eax, var_84
loc_0059888E:   push eax
  loc_0059888F: call [004010B8h] ; __vbaObjSet
loc_00598895:   lea ecx, var_170
loc_0059889B:   push eax
loc_0059889C:   push ecx
  loc_0059889D: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005988A3: add esp, 00000010h
loc_005988A6:   push eax
  loc_005988A7: call [00401034h] ; __vbaStrVarMove
loc_005988AD:   mov edx, [ebx]
loc_005988AF:   mov var_178, eax
  loc_005988B5: mov eax, 00000008h
  loc_005988BA: mov ecx, 0042EC30h ; "','"
loc_005988BF:   push ebx
loc_005988C0:   mov var_180, eax
loc_005988C6:   mov var_258, ecx
loc_005988CC:   mov var_260, eax
loc_005988D2:   mov var_268, ecx
loc_005988D8:   mov var_270, eax
loc_005988DE:   Call [edx+00000318h]
loc_005988E4:   push eax
loc_005988E5:   lea eax, var_88
loc_005988EB:   push eax
  loc_005988EC: call [004010B8h] ; __vbaObjSet
loc_005988F2:   mov ecx, [eax]
loc_005988F4:   lea edx, var_48
loc_005988F7:   push edx
loc_005988F8:   push eax
loc_005988F9:   mov var_328, eax
loc_005988FF:   Call [ecx+000000A0h]
loc_00598905:   test eax, eax
loc_00598907:   fnclex
  loc_00598909: jge 00598923h
loc_0059890B:   mov ecx, var_328
  loc_00598911: push 000000A0h
  loc_00598916: push 0042DFCCh
loc_0059891B:   push ecx
loc_0059891C:   push eax
  loc_0059891D: call [00401080h] ; __vbaHresultCheckObj
loc_00598923:   mov eax, var_48
loc_00598926:   mov edx, [ebx]
loc_00598928:   mov var_1C8, eax
  loc_0059892E: mov eax, 00000008h
loc_00598933:   push ebx
  loc_00598934: mov var_48, 00000000h
loc_0059893B:   mov var_1D0, eax
  loc_00598941: mov var_278, 0042EC30h ; "','"
loc_0059894B:   mov var_280, eax
loc_00598951:   Call [edx+000002FCh]
loc_00598957:   push eax
loc_00598958:   lea eax, var_8C
loc_0059895E:   push eax
  loc_0059895F: call [004010B8h] ; __vbaObjSet
loc_00598965:   mov ecx, [eax]
loc_00598967:   lea edx, var_4C
loc_0059896A:   push edx
loc_0059896B:   push eax
loc_0059896C:   mov var_330, eax
loc_00598972:   Call [ecx+000000A0h]
loc_00598978:   test eax, eax
loc_0059897A:   fnclex
  loc_0059897C: jge 00598996h
loc_0059897E:   mov ecx, var_330
  loc_00598984: push 000000A0h
  loc_00598989: push 0042DFCCh
loc_0059898E:   push ecx
loc_0059898F:   push eax
  loc_00598990: call [00401080h] ; __vbaHresultCheckObj
loc_00598996:   mov edx, var_4C
loc_00598999:   push edx
  loc_0059899A: call [004012A4h] ; rtcR8ValFromBstr
  loc_005989A0: fstp real8 ptr var_288
loc_005989A6:   lea eax, var_D0
loc_005989AC:   lea ecx, var_C0
loc_005989B2:   push eax
loc_005989B3:   lea edx, var_E0
loc_005989B9:   push ecx
loc_005989BA:   push edx
  loc_005989BB: mov var_290, 00000005h
  loc_005989C5: mov var_298, 0042EC3Ch ; "')"
  loc_005989CF: mov var_2A0, 00000008h
  loc_005989D9: call [004011C4h] ; __vbaVarCat
loc_005989DF:   push eax
loc_005989E0:   lea eax, var_230
loc_005989E6:   lea ecx, var_F0
loc_005989EC:   push eax
loc_005989ED:   push ecx
  loc_005989EE: call [004011C4h] ; __vbaVarCat
loc_005989F4:   push eax
loc_005989F5:   lea edx, var_110
loc_005989FB:   lea eax, var_120
loc_00598A01:   push edx
loc_00598A02:   push eax
  loc_00598A03: call [004011C4h] ; __vbaVarCat
loc_00598A09:   lea ecx, var_240
loc_00598A0F:   push eax
loc_00598A10:   lea edx, var_130
loc_00598A16:   push ecx
loc_00598A17:   push edx
  loc_00598A18: call [004011C4h] ; __vbaVarCat
loc_00598A1E:   push eax
loc_00598A1F:   lea eax, var_140
loc_00598A25:   lea ecx, var_150
loc_00598A2B:   push eax
loc_00598A2C:   push ecx
  loc_00598A2D: call [004011C4h] ; __vbaVarCat
loc_00598A33:   push eax
loc_00598A34:   lea edx, var_250
loc_00598A3A:   lea eax, var_160
loc_00598A40:   push edx
loc_00598A41:   push eax
  loc_00598A42: call [004011C4h] ; __vbaVarCat
loc_00598A48:   lea ecx, var_180
loc_00598A4E:   push eax
loc_00598A4F:   lea edx, var_190
loc_00598A55:   push ecx
loc_00598A56:   push edx
  loc_00598A57: call [004011C4h] ; __vbaVarCat
loc_00598A5D:   push eax
loc_00598A5E:   lea eax, var_260
loc_00598A64:   lea ecx, var_1A0
loc_00598A6A:   push eax
loc_00598A6B:   push ecx
  loc_00598A6C: call [004011C4h] ; __vbaVarCat
loc_00598A72:   push eax
loc_00598A73:   lea edx, var_1B0
  loc_00598A79: push 006680B4h
loc_00598A7E:   push edx
  loc_00598A7F: call [004011C4h] ; __vbaVarCat
loc_00598A85:   push eax
loc_00598A86:   lea eax, var_270
loc_00598A8C:   lea ecx, var_1C0
loc_00598A92:   push eax
loc_00598A93:   push ecx
  loc_00598A94: call [004011C4h] ; __vbaVarCat
loc_00598A9A:   push eax
loc_00598A9B:   lea edx, var_1D0
loc_00598AA1:   lea eax, var_1E0
loc_00598AA7:   push edx
loc_00598AA8:   push eax
  loc_00598AA9: call [004011C4h] ; __vbaVarCat
loc_00598AAF:   lea ecx, var_280
loc_00598AB5:   push eax
loc_00598AB6:   lea edx, var_1F0
loc_00598ABC:   push ecx
loc_00598ABD:   push edx
  loc_00598ABE: call [004011C4h] ; __vbaVarCat
loc_00598AC4:   push eax
loc_00598AC5:   lea eax, var_290
loc_00598ACB:   lea ecx, var_200
loc_00598AD1:   push eax
loc_00598AD2:   push ecx
  loc_00598AD3: call [004011C4h] ; __vbaVarCat
loc_00598AD9:   push eax
loc_00598ADA:   lea edx, var_2A0
loc_00598AE0:   lea eax, var_210
loc_00598AE6:   push edx
loc_00598AE7:   push eax
  loc_00598AE8: call [004011C4h] ; __vbaVarCat
loc_00598AEE:   push eax
  loc_00598AEF: call [00401034h] ; __vbaStrVarMove
loc_00598AF5:   mov edx, eax
  loc_00598AF7: mov ecx, 0066809Ch
  loc_00598AFC: call __vbaVarDup
loc_00598AFE:   lea ecx, var_4C
loc_00598B01:   lea edx, var_40
loc_00598B04:   push ecx
loc_00598B05:   lea eax, var_38
loc_00598B08:   push edx
loc_00598B09:   lea ecx, var_3C
loc_00598B0C:   push eax
loc_00598B0D:   lea edx, var_34
loc_00598B10:   push ecx
loc_00598B11:   lea eax, var_30
loc_00598B14:   push edx
loc_00598B15:   lea ecx, var_2C
loc_00598B18:   push eax
loc_00598B19:   push ecx
  loc_00598B1A: push 00000007h
  loc_00598B1C: call [00401204h] ; __vbaFreeStrList
loc_00598B22:   lea edx, var_8C
loc_00598B28:   lea eax, var_88
loc_00598B2E:   push edx
loc_00598B2F:   lea ecx, var_84
loc_00598B35:   push eax
loc_00598B36:   lea edx, var_80
loc_00598B39:   push ecx
loc_00598B3A:   lea eax, var_7C
loc_00598B3D:   push edx
loc_00598B3E:   lea ecx, var_78
loc_00598B41:   push eax
loc_00598B42:   push ecx
  loc_00598B43: push 00000006h
  loc_00598B45: call [0040104Ch] ; __vbaFreeObjList
loc_00598B4B:   lea edx, var_210
loc_00598B51:   lea eax, var_200
loc_00598B57:   push edx
loc_00598B58:   lea ecx, var_1F0
loc_00598B5E:   push eax
loc_00598B5F:   lea edx, var_1E0
loc_00598B65:   push ecx
loc_00598B66:   lea eax, var_1D0
loc_00598B6C:   push edx
loc_00598B6D:   lea ecx, var_1C0
loc_00598B73:   push eax
loc_00598B74:   lea edx, var_1B0
loc_00598B7A:   push ecx
loc_00598B7B:   lea eax, var_1A0
loc_00598B81:   push edx
loc_00598B82:   lea ecx, var_190
loc_00598B88:   push eax
loc_00598B89:   lea edx, var_180
loc_00598B8F:   push ecx
loc_00598B90:   lea eax, var_160
loc_00598B96:   push edx
loc_00598B97:   lea ecx, var_170
loc_00598B9D:   push eax
loc_00598B9E:   lea edx, var_150
loc_00598BA4:   push ecx
loc_00598BA5:   lea eax, var_140
loc_00598BAB:   push edx
loc_00598BAC:   lea ecx, var_130
loc_00598BB2:   push eax
loc_00598BB3:   lea edx, var_120
loc_00598BB9:   push ecx
loc_00598BBA:   lea eax, var_110
loc_00598BC0:   push edx
loc_00598BC1:   lea ecx, var_F0
loc_00598BC7:   push eax
loc_00598BC8:   push ecx
loc_00598BC9:   lea edx, var_100
loc_00598BCF:   lea eax, var_E0
loc_00598BD5:   push edx
loc_00598BD6:   lea ecx, var_C0
loc_00598BDC:   push eax
loc_00598BDD:   lea edx, var_D0
loc_00598BE3:   push ecx
loc_00598BE4:   lea eax, var_B0
loc_00598BEA:   push edx
loc_00598BEB:   lea ecx, var_A0
loc_00598BF1:   push eax
loc_00598BF2:   push ecx
  loc_00598BF3: push 00000018h
  loc_00598BF5: call [0040103Ch] ; __vbaFreeVarList
loc_00598BFB:   mov eax, [0066803Ch]
  loc_00598C00: add esp, 000000A0h
loc_00598C06:   test eax, eax
  loc_00598C08: jnz 00598C1Ah
  loc_00598C0A: push 0066803Ch
  loc_00598C0F: push 0042DEFCh
  loc_00598C14: call [004011E8h] ; __vbaNew2
loc_00598C1A:   mov edx, [0066803Ch]
loc_00598C20:   lea ecx, var_A0
loc_00598C26:   mov var_318, edx
  loc_00598C2C: mov var_98, 80020004h
  loc_00598C36: mov var_A0, 0000000Ah
  loc_00598C40: call [0040123Ch] ; __vbaFreeVarg
loc_00598C46:   mov eax, var_318
loc_00598C4C:   lea edx, var_78
loc_00598C4F:   push edx
loc_00598C50:   lea edx, var_A0
loc_00598C56:   mov ecx, [eax]
  loc_00598C58: push 00000001h
loc_00598C5A:   push edx
loc_00598C5B:   mov edx, [0066809Ch]
loc_00598C61:   push edx
loc_00598C62:   push eax
loc_00598C63:   Call [ecx+00000040h]
loc_00598C66:   test eax, eax
loc_00598C68:   fnclex
  loc_00598C6A: jge 00598C81h
loc_00598C6C:   mov ecx, var_318
  loc_00598C72: push 00000040h
  loc_00598C74: push 0042E1B0h
loc_00598C79:   push ecx
loc_00598C7A:   push eax
  loc_00598C7B: call [00401080h] ; __vbaHresultCheckObj
loc_00598C81:   lea ecx, var_78
  loc_00598C84: call [004012A0h] ; __vbaFreeObj
loc_00598C8A:   lea ecx, var_A0
  loc_00598C90: call [00401020h] ; __vbaFreeVar
loc_00598C96:   lea edx, var_B0
loc_00598C9C:   push edx
  loc_00598C9D: call [00401220h] ; rtcGetDateVar
loc_00598CA3:   mov eax, [ebx]
loc_00598CA5:   push ebx
loc_00598CA6:   Call [eax+0000037Ch]
loc_00598CAC:   lea ecx, var_90
loc_00598CB2:   push eax
loc_00598CB3:   push ecx
  loc_00598CB4: call [004010B8h] ; __vbaObjSet
loc_00598CBA:   mov edx, [ebx]
loc_00598CBC:   push ebx
loc_00598CBD:   Call [edx+0000031Ch]
loc_00598CC3:   push eax
loc_00598CC4:   lea eax, var_78
loc_00598CC7:   push eax
  loc_00598CC8: call [004010B8h] ; __vbaObjSet
loc_00598CCE:   mov ecx, [eax]
loc_00598CD0:   lea edx, var_38
loc_00598CD3:   push edx
loc_00598CD4:   push eax
loc_00598CD5:   mov var_318, eax
loc_00598CDB:   Call [ecx+000000A0h]
loc_00598CE1:   test eax, eax
loc_00598CE3:   fnclex
  loc_00598CE5: jge 00598CFFh
loc_00598CE7:   mov ecx, var_318
  loc_00598CED: push 000000A0h
  loc_00598CF2: push 0042DFCCh
loc_00598CF7:   push ecx
loc_00598CF8:   push eax
  loc_00598CF9: call [00401080h] ; __vbaHresultCheckObj
loc_00598CFF:   mov edx, [ebx]
loc_00598D01:   push ebx
loc_00598D02:   Call [edx+00000320h]
loc_00598D08:   push eax
loc_00598D09:   lea eax, var_80
loc_00598D0C:   push eax
  loc_00598D0D: call [004010B8h] ; __vbaObjSet
loc_00598D13:   mov ecx, [eax]
loc_00598D15:   lea edx, var_50
loc_00598D18:   push edx
loc_00598D19:   push eax
loc_00598D1A:   mov var_320, eax
loc_00598D20:   Call [ecx+000000A0h]
loc_00598D26:   test eax, eax
loc_00598D28:   fnclex
  loc_00598D2A: jge 00598D44h
loc_00598D2C:   mov ecx, var_320
  loc_00598D32: push 000000A0h
  loc_00598D37: push 0042DFCCh
loc_00598D3C:   push ecx
loc_00598D3D:   push eax
  loc_00598D3E: call [00401080h] ; __vbaHresultCheckObj
  loc_00598D44: push 00434704h ; "INSERT INTO hutang"
  loc_00598D49: push 00434730h ; "(no_pembelian,no_notabeli,kode_pemasok,nama_pemasok,tgl_masuk,tgl_tempo,tot_beli,pembayaran,sisa_hutang)"
loc_00598D4E:   Call edi
loc_00598D50:   mov edx, eax
loc_00598D52:   lea ecx, var_2C
  loc_00598D55: call __vbaVarDup
loc_00598D57:   push eax
  loc_00598D58: push 00434520h ; "VALUES ('"
loc_00598D5D:   Call edi
loc_00598D5F:   mov edx, eax
loc_00598D61:   lea ecx, var_30
  loc_00598D64: call __vbaVarDup
loc_00598D66:   mov edx, [ebx+00000040h]
loc_00598D69:   push eax
loc_00598D6A:   push edx
loc_00598D6B:   Call edi
loc_00598D6D:   mov edx, eax
loc_00598D6F:   lea ecx, var_34
  loc_00598D72: call __vbaVarDup
loc_00598D74:   push eax
  loc_00598D75: push 0042EC30h ; "','"
loc_00598D7A:   Call edi
loc_00598D7C:   mov edx, eax
loc_00598D7E:   lea ecx, var_3C
  loc_00598D81: call __vbaVarDup
loc_00598D83:   push eax
loc_00598D84:   mov eax, var_38
loc_00598D87:   push eax
loc_00598D88:   Call edi
loc_00598D8A:   mov edx, eax
loc_00598D8C:   lea ecx, var_40
  loc_00598D8F: call __vbaVarDup
loc_00598D91:   push eax
  loc_00598D92: push 0042EC30h ; "','"
loc_00598D97:   Call edi
loc_00598D99:   mov edx, eax
loc_00598D9B:   lea ecx, var_44
  loc_00598D9E: call __vbaVarDup
loc_00598DA0:   mov ecx, [ebx]
loc_00598DA2:   push eax
  loc_00598DA3: push 00000000h
loc_00598DA5:   push FFFFFDFBh
loc_00598DAA:   push ebx
loc_00598DAB:   Call [ecx+00000384h]
loc_00598DB1:   lea edx, var_7C
loc_00598DB4:   push eax
loc_00598DB5:   push edx
  loc_00598DB6: call [004010B8h] ; __vbaObjSet
loc_00598DBC:   push eax
loc_00598DBD:   lea eax, var_A0
loc_00598DC3:   push eax
  loc_00598DC4: call [0040114Ch] ; __vbaLateIdCallLd
  loc_00598DCA: add esp, 00000010h
loc_00598DCD:   push eax
  loc_00598DCE: call [00401034h] ; __vbaStrVarMove
loc_00598DD4:   mov edx, eax
loc_00598DD6:   lea ecx, var_48
  loc_00598DD9: call __vbaVarDup
loc_00598DDB:   push eax
loc_00598DDC:   Call edi
loc_00598DDE:   mov edx, eax
loc_00598DE0:   lea ecx, var_4C
  loc_00598DE3: call __vbaVarDup
loc_00598DE5:   push eax
  loc_00598DE6: push 0042EC30h ; "','"
loc_00598DEB:   Call edi
loc_00598DED:   mov edx, eax
loc_00598DEF:   lea ecx, var_54
  loc_00598DF2: call __vbaVarDup
loc_00598DF4:   mov ecx, var_50
loc_00598DF7:   push eax
loc_00598DF8:   push ecx
loc_00598DF9:   Call edi
loc_00598DFB:   mov edx, eax
loc_00598DFD:   lea ecx, var_58
  loc_00598E00: call __vbaVarDup
loc_00598E02:   push eax
  loc_00598E03: push 0042EC30h ; "','"
loc_00598E08:   Call edi
loc_00598E0A:   mov var_D8, eax
  loc_00598E10: mov eax, 00000008h
loc_00598E15:   lea edx, var_220
loc_00598E1B:   lea ecx, var_C0
loc_00598E21:   mov var_E0, eax
  loc_00598E27: mov var_218, 00434538h ; "yyyy/MM/dd"
loc_00598E31:   mov var_220, eax
  loc_00598E37: call [00401238h] ; __vbaVarDup
  loc_00598E3D: push 00000001h
loc_00598E3F:   lea edx, var_C0
  loc_00598E45: push 00000001h
loc_00598E47:   lea eax, var_B0
loc_00598E4D:   push edx
loc_00598E4E:   lea ecx, var_D0
loc_00598E54:   push eax
loc_00598E55:   push ecx
  loc_00598E56: call [00401078h] ; rtcVarFromFormatVar
  loc_00598E5C: mov eax, 00000008h
loc_00598E61:   lea edx, var_240
loc_00598E67:   lea ecx, var_120
  loc_00598E6D: mov var_228, 0042EC30h ; "','"
loc_00598E77:   mov var_230, eax
  loc_00598E7D: mov var_238, 00434538h ; "yyyy/MM/dd"
loc_00598E87:   mov var_240, eax
  loc_00598E8D: call [00401238h] ; __vbaVarDup
loc_00598E93:   mov eax, var_90
  loc_00598E99: push 00000001h
loc_00598E9B:   mov var_108, eax
loc_00598EA1:   lea edx, var_120
  loc_00598EA7: push 00000001h
loc_00598EA9:   lea eax, var_110
loc_00598EAF:   push edx
loc_00598EB0:   lea ecx, var_130
loc_00598EB6:   push eax
loc_00598EB7:   push ecx
  loc_00598EB8: mov var_90, 00000000h
  loc_00598EC2: mov var_110, 00000009h
  loc_00598ECC: call [00401078h] ; rtcVarFromFormatVar
loc_00598ED2:   mov edx, [ebx]
loc_00598ED4:   push ebx
  loc_00598ED5: mov var_248, 0042EC30h ; "','"
  loc_00598EDF: mov var_250, 00000008h
loc_00598EE9:   Call [edx+000002FCh]
loc_00598EEF:   push eax
loc_00598EF0:   lea eax, var_84
loc_00598EF6:   push eax
  loc_00598EF7: call [004010B8h] ; __vbaObjSet
loc_00598EFD:   mov ecx, [eax]
loc_00598EFF:   lea edx, var_5C
loc_00598F02:   push edx
loc_00598F03:   push eax
loc_00598F04:   mov var_328, eax
loc_00598F0A:   Call [ecx+000000A0h]
loc_00598F10:   test eax, eax
loc_00598F12:   fnclex
  loc_00598F14: jge 00598F2Eh
loc_00598F16:   mov ecx, var_328
  loc_00598F1C: push 000000A0h
  loc_00598F21: push 0042DFCCh
loc_00598F26:   push ecx
loc_00598F27:   push eax
  loc_00598F28: call [00401080h] ; __vbaHresultCheckObj
loc_00598F2E:   mov edx, var_5C
loc_00598F31:   push edx
  loc_00598F32: call [004012A4h] ; rtcR8ValFromBstr
loc_00598F38:   mov eax, [ebx]
loc_00598F3A:   push ebx
  loc_00598F3B: fstp real8 ptr var_258
  loc_00598F41: mov var_260, 00000005h
  loc_00598F4B: mov var_268, 0042EC30h ; "','"
  loc_00598F55: mov var_270, 00000008h
loc_00598F5F:   Call [eax+00000304h]
loc_00598F65:   lea ecx, var_88
loc_00598F6B:   push eax
loc_00598F6C:   push ecx
  loc_00598F6D: call [004010B8h] ; __vbaObjSet
loc_00598F73:   mov edx, [eax]
loc_00598F75:   lea ecx, var_60
loc_00598F78:   push ecx
loc_00598F79:   push eax
loc_00598F7A:   mov var_330, eax
loc_00598F80:   Call [edx+000000A0h]
loc_00598F86:   test eax, eax
loc_00598F88:   fnclex
  loc_00598F8A: jge 00598FA4h
loc_00598F8C:   mov edx, var_330
  loc_00598F92: push 000000A0h
  loc_00598F97: push 0042DFCCh
loc_00598F9C:   push edx
loc_00598F9D:   push eax
  loc_00598F9E: call [00401080h] ; __vbaHresultCheckObj
loc_00598FA4:   mov eax, var_60
loc_00598FA7:   push eax
  loc_00598FA8: call [004012A4h] ; rtcR8ValFromBstr
loc_00598FAE:   mov ecx, [ebx]
loc_00598FB0:   push ebx
  loc_00598FB1: fstp real8 ptr var_278
  loc_00598FB7: mov var_280, 00000005h
  loc_00598FC1: mov var_288, 0042EC30h ; "','"
  loc_00598FCB: mov var_290, 00000008h
loc_00598FD5:   Call [ecx+00000300h]
loc_00598FDB:   lea edx, var_8C
loc_00598FE1:   push eax
loc_00598FE2:   push edx
  loc_00598FE3: call [004010B8h] ; __vbaObjSet
loc_00598FE9:   mov ecx, [eax]
loc_00598FEB:   lea edx, var_64
loc_00598FEE:   push edx
loc_00598FEF:   push eax
loc_00598FF0:   mov var_338, eax
loc_00598FF6:   Call [ecx+000000A0h]
loc_00598FFC:   test eax, eax
loc_00598FFE:   fnclex
  loc_00599000: jge 0059901Ah
loc_00599002:   mov ecx, var_338
  loc_00599008: push 000000A0h
  loc_0059900D: push 0042DFCCh
loc_00599012:   push ecx
loc_00599013:   push eax
  loc_00599014: call [00401080h] ; __vbaHresultCheckObj
loc_0059901A:   mov edx, var_64
loc_0059901D:   push edx
  loc_0059901E: call [004012A4h] ; rtcR8ValFromBstr
  loc_00599024: fstp real8 ptr var_298
loc_0059902A:   lea eax, var_E0
loc_00599030:   lea ecx, var_D0
loc_00599036:   push eax
loc_00599037:   lea edx, var_F0
loc_0059903D:   push ecx
loc_0059903E:   push edx
  loc_0059903F: mov var_2A0, 00000005h
  loc_00599049: mov var_2A8, 0042EC3Ch ; "')"
  loc_00599053: mov var_2B0, 00000008h
  loc_0059905D: call [004011C4h] ; __vbaVarCat
loc_00599063:   push eax
loc_00599064:   lea eax, var_230
loc_0059906A:   lea ecx, var_100
loc_00599070:   push eax
loc_00599071:   push ecx
  loc_00599072: call [004011C4h] ; __vbaVarCat
loc_00599078:   push eax
loc_00599079:   lea edx, var_130
loc_0059907F:   lea eax, var_140
loc_00599085:   push edx
loc_00599086:   push eax
  loc_00599087: call [004011C4h] ; __vbaVarCat
loc_0059908D:   lea ecx, var_250
loc_00599093:   push eax
loc_00599094:   lea edx, var_150
loc_0059909A:   push ecx
loc_0059909B:   push edx
  loc_0059909C: call [004011C4h] ; __vbaVarCat
loc_005990A2:   push eax
loc_005990A3:   lea eax, var_260
loc_005990A9:   lea ecx, var_160
loc_005990AF:   push eax
loc_005990B0:   push ecx
  loc_005990B1: call [004011C4h] ; __vbaVarCat
loc_005990B7:   push eax
loc_005990B8:   lea edx, var_270
loc_005990BE:   lea eax, var_170
loc_005990C4:   push edx
loc_005990C5:   push eax
  loc_005990C6: call [004011C4h] ; __vbaVarCat
loc_005990CC:   lea ecx, var_280
loc_005990D2:   push eax
loc_005990D3:   lea edx, var_180
loc_005990D9:   push ecx
loc_005990DA:   push edx
  loc_005990DB: call [004011C4h] ; __vbaVarCat
loc_005990E1:   push eax
loc_005990E2:   lea eax, var_290
loc_005990E8:   lea ecx, var_190
loc_005990EE:   push eax
loc_005990EF:   push ecx
  loc_005990F0: call [004011C4h] ; __vbaVarCat
loc_005990F6:   push eax
loc_005990F7:   lea edx, var_2A0
loc_005990FD:   lea eax, var_1A0
loc_00599103:   push edx
loc_00599104:   push eax
  loc_00599105: call [004011C4h] ; __vbaVarCat
loc_0059910B:   lea ecx, var_2B0
loc_00599111:   push eax
loc_00599112:   lea edx, var_1B0
loc_00599118:   push ecx
loc_00599119:   push edx
  loc_0059911A: call [004011C4h] ; __vbaVarCat
loc_00599120:   push eax
  loc_00599121: call [00401034h] ; __vbaStrVarMove
loc_00599127:   mov edx, eax
  loc_00599129: mov ecx, 0066809Ch
  loc_0059912E: call __vbaVarDup
loc_00599130:   lea eax, var_64
loc_00599133:   lea ecx, var_60
loc_00599136:   push eax
loc_00599137:   lea edx, var_5C
loc_0059913A:   push ecx
loc_0059913B:   lea eax, var_58
loc_0059913E:   push edx
loc_0059913F:   push eax
loc_00599140:   lea ecx, var_50
loc_00599143:   lea edx, var_54
loc_00599146:   push ecx
loc_00599147:   lea eax, var_4C
loc_0059914A:   push edx
loc_0059914B:   lea ecx, var_48
loc_0059914E:   push eax
loc_0059914F:   lea edx, var_44
loc_00599152:   push ecx
loc_00599153:   lea eax, var_40
loc_00599156:   push edx
loc_00599157:   lea ecx, var_38
loc_0059915A:   push eax
loc_0059915B:   lea edx, var_3C
loc_0059915E:   push ecx
loc_0059915F:   lea eax, var_34
loc_00599162:   push edx
loc_00599163:   lea ecx, var_30
loc_00599166:   push eax
loc_00599167:   lea edx, var_2C
loc_0059916A:   push ecx
loc_0059916B:   push edx
  loc_0059916C: push 0000000Fh
  loc_0059916E: call [00401204h] ; __vbaFreeStrList
loc_00599174:   lea eax, var_90
loc_0059917A:   lea ecx, var_8C
loc_00599180:   push eax
loc_00599181:   lea edx, var_88
loc_00599187:   push ecx
loc_00599188:   lea eax, var_84
loc_0059918E:   push edx
loc_0059918F:   lea ecx, var_80
loc_00599192:   push eax
loc_00599193:   lea edx, var_7C
loc_00599196:   push ecx
loc_00599197:   lea eax, var_78
loc_0059919A:   push edx
loc_0059919B:   push eax
  loc_0059919C: push 00000007h
  loc_0059919E: call [0040104Ch] ; __vbaFreeObjList
  loc_005991A4: add esp, 00000060h
loc_005991A7:   lea ecx, var_1B0
loc_005991AD:   lea edx, var_1A0
loc_005991B3:   lea eax, var_190
loc_005991B9:   push ecx
loc_005991BA:   push edx
loc_005991BB:   lea ecx, var_180
loc_005991C1:   push eax
loc_005991C2:   lea edx, var_170
loc_005991C8:   push ecx
loc_005991C9:   lea eax, var_160
loc_005991CF:   push edx
loc_005991D0:   lea ecx, var_150
loc_005991D6:   push eax
loc_005991D7:   lea edx, var_140
loc_005991DD:   push ecx
loc_005991DE:   lea eax, var_130
loc_005991E4:   push edx
loc_005991E5:   lea ecx, var_100
loc_005991EB:   push eax
loc_005991EC:   lea edx, var_120
loc_005991F2:   push ecx
loc_005991F3:   lea eax, var_110
loc_005991F9:   push edx
loc_005991FA:   lea ecx, var_F0
loc_00599200:   push eax
loc_00599201:   lea edx, var_D0
loc_00599207:   push ecx
loc_00599208:   lea eax, var_E0
loc_0059920E:   push edx
loc_0059920F:   lea ecx, var_C0
loc_00599215:   push eax
loc_00599216:   lea edx, var_B0
loc_0059921C:   push ecx
loc_0059921D:   lea eax, var_A0
loc_00599223:   push edx
loc_00599224:   push eax
  loc_00599225: push 00000012h
  loc_00599227: call [0040103Ch] ; __vbaFreeVarList
loc_0059922D:   mov eax, [0066803Ch]
  loc_00599232: add esp, 0000004Ch
loc_00599235:   test eax, eax
  loc_00599237: jnz 00599249h
  loc_00599239: push 0066803Ch
  loc_0059923E: push 0042DEFCh
  loc_00599243: call [004011E8h] ; __vbaNew2
loc_00599249:   mov ecx, [0066803Ch]
  loc_0059924F: mov var_98, 80020004h
loc_00599259:   mov var_318, ecx
loc_0059925F:   lea ecx, var_A0
  loc_00599265: mov var_A0, 0000000Ah
  loc_0059926F: call [0040123Ch] ; __vbaFreeVarg
loc_00599275:   mov eax, var_318
loc_0059927B:   lea ecx, var_78
loc_0059927E:   push ecx
loc_0059927F:   lea ecx, var_A0
loc_00599285:   mov edx, [eax]
  loc_00599287: push 00000001h
loc_00599289:   push ecx
loc_0059928A:   mov ecx, [0066809Ch]
loc_00599290:   push ecx
loc_00599291:   push eax
loc_00599292:   Call [edx+00000040h]
loc_00599295:   test eax, eax
loc_00599297:   fnclex
  loc_00599299: jge 005992B0h
loc_0059929B:   mov edx, var_318
  loc_005992A1: push 00000040h
  loc_005992A3: push 0042E1B0h
loc_005992A8:   push edx
loc_005992A9:   push eax
  loc_005992AA: call [00401080h] ; __vbaHresultCheckObj
loc_005992B0:   lea ecx, var_78
  loc_005992B3: call [004012A0h] ; __vbaFreeObj
loc_005992B9:   lea ecx, var_A0
  loc_005992BF: call [00401020h] ; __vbaFreeVar
loc_005992C5:   mov ax, [ebx+00000034h]
  loc_005992C9: mov var_1C, 00000001h
  loc_005992D0: sub ax, 0001h
  loc_005992D4: jo 0059A014h
loc_005992DA:   mov var_34C, eax
loc_005992E0:   mov ax, var_34C
loc_005992E7:   cmp var_1C, ax
  loc_005992EB: jg 00599CB8h
loc_005992F1:   movsx eax, var_1C
  loc_005992F5: sub esp, 00000010h
  loc_005992F8: mov ecx, 00000003h
loc_005992FD:   mov edx, esp
loc_005992FF:   mov var_220, ecx
loc_00599305:   mov var_238, ecx
loc_0059930B:   mov var_240, ecx
loc_00599311:   mov [edx], ecx
loc_00599313:   mov ecx, var_21C
loc_00599319:   mov var_218, eax
  loc_0059931F: sub esp, 00000010h
loc_00599322:   mov [edx+00000004h], ecx
loc_00599325:   mov ecx, esp
  loc_00599327: push 00000002h
  loc_00599329: push 00000041h
loc_0059932B:   mov [edx+00000008h], eax
loc_0059932E:   mov eax, var_214
loc_00599334:   push ebx
loc_00599335:   mov [edx+0000000Ch], eax
loc_00599338:   mov edx, var_240
loc_0059933E:   mov eax, var_23C
loc_00599344:   mov [ecx], edx
loc_00599346:   mov edx, var_238
loc_0059934C:   mov [ecx+00000004h], eax
loc_0059934F:   mov eax, var_234
loc_00599355:   mov [ecx+00000008h], edx
loc_00599358:   mov [ecx+0000000Ch], eax
loc_0059935B:   mov ecx, [ebx]
loc_0059935D:   Call [ecx+00000390h]
loc_00599363:   lea edx, var_78
loc_00599366:   push eax
loc_00599367:   push edx
  loc_00599368: call [004010B8h] ; __vbaObjSet
loc_0059936E:   push eax
loc_0059936F:   lea eax, var_A0
loc_00599375:   push eax
  loc_00599376: call [0040114Ch] ; __vbaLateIdCallLd
  loc_0059937C: add esp, 00000030h
loc_0059937F:   lea ecx, var_A0
  loc_00599385: push 00000000h
loc_00599387:   push FFFFFFFFh
  loc_00599389: push 00000001h
  loc_0059938B: push 0042DFECh
  loc_00599390: push 0043024Ch ; "."
loc_00599395:   push ecx
  loc_00599396: call [00401034h] ; __vbaStrVarMove
loc_0059939C:   mov edx, eax
loc_0059939E:   lea ecx, var_2C
  loc_005993A1: call __vbaVarDup
loc_005993A3:   push eax
  loc_005993A4: call [00401194h] ; rtcReplace
loc_005993AA:   mov edx, eax
loc_005993AC:   lea ecx, var_18
  loc_005993AF: call __vbaVarDup
loc_005993B1:   lea ecx, var_2C
  loc_005993B4: call [0040129Ch] ; __vbaFreeStr
loc_005993BA:   lea ecx, var_78
  loc_005993BD: call [004012A0h] ; __vbaFreeObj
loc_005993C3:   lea ecx, var_A0
  loc_005993C9: call [00401020h] ; __vbaFreeVar
  loc_005993CF: mov edx, 0042DFECh
  loc_005993D4: mov ecx, 0066809Ch
  loc_005993D9: call [004011FCh] ; __vbaStrCopy
loc_005993DF:   movsx eax, var_1C
  loc_005993E3: mov ecx, 00000003h
  loc_005993E8: push 00434554h ; "INSERT INTO Pembelian_Detail"
  loc_005993ED: push 00434594h ; "(no_pembelian,kode_barang,nama_barang,harga_beli,jumlah,sub_total)"
loc_005993F2:   mov var_218, eax
loc_005993F8:   mov var_220, ecx
  loc_005993FE: mov var_238, 00000001h
loc_00599408:   mov var_240, ecx
loc_0059940E:   mov var_258, eax
loc_00599414:   mov var_260, ecx
  loc_0059941A: mov var_278, 00000002h
loc_00599424:   mov var_280, ecx
loc_0059942A:   mov var_298, eax
loc_00599430:   mov var_2A0, ecx
loc_00599436:   Call edi
loc_00599438:   mov edx, eax
loc_0059943A:   lea ecx, var_2C
  loc_0059943D: call __vbaVarDup
loc_0059943F:   push eax
  loc_00599440: push 00434620h ; " VALUES ('"
loc_00599445:   Call edi
loc_00599447:   mov edx, eax
loc_00599449:   lea ecx, var_30
  loc_0059944C: call __vbaVarDup
loc_0059944E:   mov edx, [ebx+00000040h]
loc_00599451:   push eax
loc_00599452:   push edx
loc_00599453:   Call edi
loc_00599455:   mov edx, eax
loc_00599457:   lea ecx, var_34
  loc_0059945A: call __vbaVarDup
loc_0059945C:   push eax
  loc_0059945D: push 0042EC30h ; "','"
loc_00599462:   Call edi
loc_00599464:   mov edx, eax
loc_00599466:   lea ecx, var_38
  loc_00599469: call __vbaVarDup
loc_0059946B:   mov ecx, var_220
loc_00599471:   push eax
loc_00599472:   mov edx, var_21C
  loc_00599478: sub esp, 00000010h
loc_0059947B:   mov eax, esp
  loc_0059947D: sub esp, 00000010h
loc_00599480:   mov [eax], ecx
loc_00599482:   mov ecx, var_218
loc_00599488:   mov [eax+00000004h], edx
loc_0059948B:   mov edx, var_214
loc_00599491:   mov [eax+00000008h], ecx
loc_00599494:   mov ecx, var_240
loc_0059949A:   mov [eax+0000000Ch], edx
loc_0059949D:   mov edx, var_23C
loc_005994A3:   mov eax, esp
  loc_005994A5: push 00000002h
  loc_005994A7: push 00000041h
loc_005994A9:   push ebx
loc_005994AA:   mov [eax], ecx
loc_005994AC:   mov ecx, var_238
loc_005994B2:   mov [eax+00000004h], edx
loc_005994B5:   mov edx, var_234
loc_005994BB:   mov [eax+00000008h], ecx
loc_005994BE:   mov [eax+0000000Ch], edx
loc_005994C1:   mov eax, [ebx]
loc_005994C3:   Call [eax+00000390h]
loc_005994C9:   lea ecx, var_78
loc_005994CC:   push eax
loc_005994CD:   push ecx
  loc_005994CE: call [004010B8h] ; __vbaObjSet
loc_005994D4:   lea edx, var_A0
loc_005994DA:   push eax
loc_005994DB:   push edx
  loc_005994DC: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005994E2: add esp, 00000030h
loc_005994E5:   push eax
  loc_005994E6: call [00401034h] ; __vbaStrVarMove
loc_005994EC:   mov edx, eax
loc_005994EE:   lea ecx, var_3C
  loc_005994F1: call __vbaVarDup
loc_005994F3:   push eax
loc_005994F4:   Call edi
loc_005994F6:   mov edx, eax
loc_005994F8:   lea ecx, var_40
  loc_005994FB: call __vbaVarDup
loc_005994FD:   push eax
  loc_005994FE: push 0042EC30h ; "','"
loc_00599503:   Call edi
loc_00599505:   mov edx, eax
loc_00599507:   lea ecx, var_44
  loc_0059950A: call __vbaVarDup
loc_0059950C:   mov ecx, var_260
loc_00599512:   push eax
loc_00599513:   mov edx, var_25C
  loc_00599519: sub esp, 00000010h
loc_0059951C:   mov eax, esp
loc_0059951E:   mov [eax], ecx
loc_00599520:   mov ecx, var_258
loc_00599526:   mov [eax+00000004h], edx
loc_00599529:   mov edx, var_254
loc_0059952F:   mov [eax+00000008h], ecx
loc_00599532:   mov [eax+0000000Ch], edx
loc_00599535:   mov ecx, var_280
loc_0059953B:   mov edx, var_27C
  loc_00599541: sub esp, 00000010h
loc_00599544:   mov eax, esp
  loc_00599546: push 00000002h
  loc_00599548: push 00000041h
loc_0059954A:   mov [eax], ecx
loc_0059954C:   mov ecx, var_278
loc_00599552:   push ebx
loc_00599553:   mov [eax+00000004h], edx
loc_00599556:   mov edx, var_274
loc_0059955C:   mov [eax+00000008h], ecx
loc_0059955F:   mov [eax+0000000Ch], edx
loc_00599562:   mov eax, [ebx]
loc_00599564:   Call [eax+00000390h]
loc_0059956A:   lea ecx, var_7C
loc_0059956D:   push eax
loc_0059956E:   push ecx
  loc_0059956F: call [004010B8h] ; __vbaObjSet
loc_00599575:   lea edx, var_B0
loc_0059957B:   push eax
loc_0059957C:   push edx
  loc_0059957D: call [0040114Ch] ; __vbaLateIdCallLd
  loc_00599583: add esp, 00000030h
loc_00599586:   push eax
  loc_00599587: call [00401034h] ; __vbaStrVarMove
loc_0059958D:   mov edx, eax
loc_0059958F:   lea ecx, var_48
  loc_00599592: call __vbaVarDup
loc_00599594:   push eax
loc_00599595:   Call edi
loc_00599597:   mov edx, eax
loc_00599599:   lea ecx, var_4C
  loc_0059959C: call __vbaVarDup
loc_0059959E:   push eax
  loc_0059959F: push 0042EC30h ; "','"
loc_005995A4:   Call edi
loc_005995A6:   mov edx, eax
loc_005995A8:   lea ecx, var_50
  loc_005995AB: call __vbaVarDup
loc_005995AD:   push eax
loc_005995AE:   mov eax, var_18
loc_005995B1:   push eax
  loc_005995B2: call [004012A4h] ; rtcR8ValFromBstr
  loc_005995B8: sub esp, 00000008h
  loc_005995BB: fstp real8 ptr [esp]
  loc_005995BE: call [0040115Ch] ; __vbaStrR8
loc_005995C4:   mov edx, eax
loc_005995C6:   lea ecx, var_54
  loc_005995C9: call __vbaVarDup
loc_005995CB:   push eax
loc_005995CC:   Call edi
loc_005995CE:   mov edx, eax
loc_005995D0:   lea ecx, var_58
  loc_005995D3: call __vbaVarDup
loc_005995D5:   push eax
  loc_005995D6: push 0042EC30h ; "','"
loc_005995DB:   Call edi
loc_005995DD:   mov edx, eax
loc_005995DF:   lea ecx, var_60
  loc_005995E2: call __vbaVarDup
loc_005995E4:   mov edx, var_2A0
loc_005995EA:   push eax
loc_005995EB:   mov eax, var_29C
  loc_005995F1: sub esp, 00000010h
loc_005995F4:   mov ecx, esp
  loc_005995F6: sub esp, 00000010h
loc_005995F9:   mov [ecx], edx
loc_005995FB:   mov edx, var_298
loc_00599601:   mov [ecx+00000004h], eax
loc_00599604:   mov eax, var_294
loc_0059960A:   mov [ecx+00000008h], edx
loc_0059960D:   mov edx, var_2BC
loc_00599613:   mov [ecx+0000000Ch], eax
loc_00599616:   mov ecx, esp
  loc_00599618: mov eax, 00000003h
loc_0059961D:   mov [ecx], eax
  loc_0059961F: mov eax, 00000005h
loc_00599624:   mov [ecx+00000004h], edx
loc_00599627:   mov [ecx+00000008h], eax
loc_0059962A:   mov eax, var_2B4
loc_00599630:   mov [ecx+0000000Ch], eax
loc_00599633:   mov ecx, [ebx]
  loc_00599635: push 00000002h
  loc_00599637: push 00000041h
loc_00599639:   push ebx
loc_0059963A:   Call [ecx+00000390h]
loc_00599640:   lea edx, var_80
loc_00599643:   push eax
loc_00599644:   push edx
  loc_00599645: call [004010B8h] ; __vbaObjSet
loc_0059964B:   push eax
loc_0059964C:   lea eax, var_C0
loc_00599652:   push eax
  loc_00599653: call [0040114Ch] ; __vbaLateIdCallLd
  loc_00599659: add esp, 00000030h
loc_0059965C:   push eax
  loc_0059965D: call [00401034h] ; __vbaStrVarMove
loc_00599663:   mov edx, eax
loc_00599665:   lea ecx, var_5C
  loc_00599668: call __vbaVarDup
loc_0059966A:   push eax
  loc_0059966B: call [004012A4h] ; rtcR8ValFromBstr
  loc_00599671: sub esp, 00000008h
  loc_00599674: fstp real8 ptr [esp]
  loc_00599677: call [0040115Ch] ; __vbaStrR8
loc_0059967D:   mov edx, eax
loc_0059967F:   lea ecx, var_64
  loc_00599682: call __vbaVarDup
loc_00599684:   push eax
loc_00599685:   Call edi
loc_00599687:   mov edx, eax
loc_00599689:   lea ecx, var_68
  loc_0059968C: call __vbaVarDup
loc_0059968E:   push eax
  loc_0059968F: push 0042EC30h ; "','"
loc_00599694:   Call edi
loc_00599696:   mov edx, eax
loc_00599698:   lea ecx, var_6C
  loc_0059969B: call __vbaVarDup
loc_0059969D:   push eax
loc_0059969E:   mov edx, var_2DC
  loc_005996A4: sub esp, 00000010h
  loc_005996A7: mov eax, 00000003h
loc_005996AC:   mov ecx, esp
  loc_005996AE: sub esp, 00000010h
loc_005996B1:   mov [ecx], eax
loc_005996B3:   movsx eax, var_1C
loc_005996B7:   mov [ecx+00000004h], edx
loc_005996BA:   mov edx, var_2FC
loc_005996C0:   mov [ecx+00000008h], eax
loc_005996C3:   mov eax, var_2D4
loc_005996C9:   mov [ecx+0000000Ch], eax
loc_005996CC:   mov ecx, esp
  loc_005996CE: mov eax, 00000003h
  loc_005996D3: push 00000002h
loc_005996D5:   mov [ecx], eax
  loc_005996D7: mov eax, 00000006h
  loc_005996DC: push 00000041h
loc_005996DE:   push ebx
loc_005996DF:   mov [ecx+00000004h], edx
loc_005996E2:   mov [ecx+00000008h], eax
loc_005996E5:   mov eax, var_2F4
loc_005996EB:   mov [ecx+0000000Ch], eax
loc_005996EE:   mov ecx, [ebx]
loc_005996F0:   Call [ecx+00000390h]
loc_005996F6:   lea edx, var_84
loc_005996FC:   push eax
loc_005996FD:   push edx
  loc_005996FE: call [004010B8h] ; __vbaObjSet
loc_00599704:   push eax
loc_00599705:   lea eax, var_D0
loc_0059970B:   push eax
  loc_0059970C: call [0040114Ch] ; __vbaLateIdCallLd
  loc_00599712: add esp, 00000030h
loc_00599715:   push eax
  loc_00599716: call [00401034h] ; __vbaStrVarMove
loc_0059971C:   mov edx, eax
loc_0059971E:   lea ecx, var_70
  loc_00599721: call __vbaVarDup
loc_00599723:   push eax
loc_00599724:   Call edi
loc_00599726:   mov edx, eax
loc_00599728:   lea ecx, var_74
  loc_0059972B: call __vbaVarDup
loc_0059972D:   push eax
  loc_0059972E: push 0042EC3Ch ; "')"
loc_00599733:   Call edi
loc_00599735:   mov edx, eax
  loc_00599737: mov ecx, 0066809Ch
  loc_0059973C: call __vbaVarDup
loc_0059973E:   lea ecx, var_74
loc_00599741:   lea edx, var_70
loc_00599744:   push ecx
loc_00599745:   lea eax, var_6C
loc_00599748:   push edx
loc_00599749:   lea ecx, var_68
loc_0059974C:   push eax
loc_0059974D:   lea edx, var_64
loc_00599750:   push ecx
loc_00599751:   lea eax, var_60
loc_00599754:   push edx
loc_00599755:   lea ecx, var_5C
loc_00599758:   push eax
loc_00599759:   lea edx, var_58
loc_0059975C:   push ecx
loc_0059975D:   lea eax, var_54
loc_00599760:   push edx
loc_00599761:   lea ecx, var_50
loc_00599764:   push eax
loc_00599765:   lea edx, var_4C
loc_00599768:   push ecx
loc_00599769:   lea eax, var_48
loc_0059976C:   push edx
loc_0059976D:   lea ecx, var_44
loc_00599770:   push eax
loc_00599771:   lea edx, var_40
loc_00599774:   push ecx
loc_00599775:   lea eax, var_3C
loc_00599778:   push edx
loc_00599779:   lea ecx, var_38
loc_0059977C:   push eax
loc_0059977D:   lea edx, var_34
loc_00599780:   push ecx
loc_00599781:   lea eax, var_30
loc_00599784:   push edx
loc_00599785:   lea ecx, var_2C
loc_00599788:   push eax
loc_00599789:   push ecx
  loc_0059978A: push 00000013h
  loc_0059978C: call [00401204h] ; __vbaFreeStrList
loc_00599792:   lea edx, var_84
loc_00599798:   lea eax, var_80
loc_0059979B:   push edx
loc_0059979C:   lea ecx, var_7C
loc_0059979F:   push eax
loc_005997A0:   lea edx, var_78
loc_005997A3:   push ecx
loc_005997A4:   push edx
  loc_005997A5: push 00000004h
  loc_005997A7: call [0040104Ch] ; __vbaFreeObjList
  loc_005997AD: add esp, 00000064h
loc_005997B0:   lea eax, var_D0
loc_005997B6:   lea ecx, var_C0
loc_005997BC:   lea edx, var_B0
loc_005997C2:   push eax
loc_005997C3:   push ecx
loc_005997C4:   lea eax, var_A0
loc_005997CA:   push edx
loc_005997CB:   push eax
  loc_005997CC: push 00000004h
  loc_005997CE: call [0040103Ch] ; __vbaFreeVarList
loc_005997D4:   mov eax, [0066803Ch]
  loc_005997D9: add esp, 00000014h
loc_005997DC:   test eax, eax
  loc_005997DE: jnz 005997F0h
  loc_005997E0: push 0066803Ch
  loc_005997E5: push 0042DEFCh
  loc_005997EA: call [004011E8h] ; __vbaNew2
loc_005997F0:   mov ecx, [0066803Ch]
  loc_005997F6: mov var_98, 80020004h
loc_00599800:   mov var_318, ecx
loc_00599806:   lea ecx, var_A0
  loc_0059980C: mov var_A0, 0000000Ah
  loc_00599816: call [0040123Ch] ; __vbaFreeVarg
loc_0059981C:   mov eax, var_318
loc_00599822:   lea ecx, var_78
loc_00599825:   push ecx
loc_00599826:   lea ecx, var_A0
loc_0059982C:   mov edx, [eax]
  loc_0059982E: push 00000001h
loc_00599830:   push ecx
loc_00599831:   mov ecx, [0066809Ch]
loc_00599837:   push ecx
loc_00599838:   push eax
loc_00599839:   Call [edx+00000040h]
loc_0059983C:   test eax, eax
loc_0059983E:   fnclex
  loc_00599840: jge 00599857h
loc_00599842:   mov edx, var_318
  loc_00599848: push 00000040h
  loc_0059984A: push 0042E1B0h
loc_0059984F:   push edx
loc_00599850:   push eax
  loc_00599851: call [00401080h] ; __vbaHresultCheckObj
loc_00599857:   lea ecx, var_78
  loc_0059985A: call [004012A0h] ; __vbaFreeObj
loc_00599860:   lea ecx, var_A0
  loc_00599866: call [00401020h] ; __vbaFreeVar
loc_0059986C:   movsx eax, var_1C
  loc_00599870: sub esp, 00000010h
  loc_00599873: mov ecx, 00000003h
loc_00599878:   mov edx, esp
loc_0059987A:   mov var_220, ecx
loc_00599880:   mov var_240, ecx
loc_00599886:   mov var_218, eax
loc_0059988C:   mov [edx], ecx
loc_0059988E:   mov ecx, var_21C
  loc_00599894: sub esp, 00000010h
  loc_00599897: mov var_238, 00000005h
loc_005998A1:   mov [edx+00000004h], ecx
loc_005998A4:   mov ecx, esp
  loc_005998A6: push 00000002h
  loc_005998A8: push 00000041h
loc_005998AA:   mov [edx+00000008h], eax
loc_005998AD:   mov eax, var_214
loc_005998B3:   push ebx
loc_005998B4:   mov [edx+0000000Ch], eax
loc_005998B7:   mov edx, var_240
loc_005998BD:   mov eax, var_23C
loc_005998C3:   mov [ecx], edx
loc_005998C5:   mov edx, var_238
loc_005998CB:   mov [ecx+00000004h], eax
loc_005998CE:   mov eax, var_234
loc_005998D4:   mov [ecx+00000008h], edx
loc_005998D7:   mov [ecx+0000000Ch], eax
loc_005998DA:   mov ecx, [ebx]
loc_005998DC:   Call [ecx+00000390h]
loc_005998E2:   lea edx, var_78
loc_005998E5:   push eax
loc_005998E6:   push edx
  loc_005998E7: call [004010B8h] ; __vbaObjSet
loc_005998ED:   push eax
loc_005998EE:   lea eax, var_A0
loc_005998F4:   push eax
  loc_005998F5: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005998FB: add esp, 00000030h
loc_005998FE:   push eax
  loc_005998FF: call [00401034h] ; __vbaStrVarMove
loc_00599905:   mov edx, eax
loc_00599907:   lea ecx, var_2C
  loc_0059990A: call __vbaVarDup
loc_0059990C:   push eax
  loc_0059990D: call [004012A4h] ; rtcR8ValFromBstr
  loc_00599913: call [00401244h] ; __vbaFpI2
loc_00599919:   lea ecx, var_2C
loc_0059991C:   mov var_28, eax
  loc_0059991F: call [0040129Ch] ; __vbaFreeStr
loc_00599925:   lea ecx, var_78
  loc_00599928: call [004012A0h] ; __vbaFreeObj
loc_0059992E:   lea ecx, var_A0
  loc_00599934: call [00401020h] ; __vbaFreeVar
loc_0059993A:   movsx eax, var_1C
  loc_0059993E: sub esp, 00000010h
  loc_00599941: mov ecx, 00000003h
loc_00599946:   mov edx, esp
loc_00599948:   mov var_220, ecx
loc_0059994E:   mov var_240, ecx
loc_00599954:   mov var_218, eax
loc_0059995A:   mov [edx], ecx
loc_0059995C:   mov ecx, var_21C
  loc_00599962: sub esp, 00000010h
  loc_00599965: mov var_238, 00000001h
loc_0059996F:   mov [edx+00000004h], ecx
loc_00599972:   mov ecx, esp
loc_00599974:   mov [edx+00000008h], eax
loc_00599977:   mov eax, var_214
loc_0059997D:   mov [edx+0000000Ch], eax
loc_00599980:   mov edx, var_240
loc_00599986:   mov eax, var_23C
loc_0059998C:   mov [ecx], edx
loc_0059998E:   mov edx, var_238
loc_00599994:   mov [ecx+00000004h], eax
loc_00599997:   mov eax, var_234
loc_0059999D:   mov [ecx+00000008h], edx
loc_005999A0:   mov [ecx+0000000Ch], eax
loc_005999A3:   mov ecx, [ebx]
  loc_005999A5: push 00000002h
  loc_005999A7: push 00000041h
loc_005999A9:   push ebx
loc_005999AA:   Call [ecx+00000390h]
loc_005999B0:   lea edx, var_78
loc_005999B3:   push eax
loc_005999B4:   push edx
  loc_005999B5: call [004010B8h] ; __vbaObjSet
loc_005999BB:   push eax
loc_005999BC:   lea eax, var_A0
loc_005999C2:   push eax
  loc_005999C3: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005999C9: add esp, 00000030h
loc_005999CC:   push eax
  loc_005999CD: call [00401034h] ; __vbaStrVarMove
loc_005999D3:   mov edx, eax
loc_005999D5:   lea ecx, var_20
  loc_005999D8: call __vbaVarDup
loc_005999DA:   lea ecx, var_78
  loc_005999DD: call [004012A0h] ; __vbaFreeObj
loc_005999E3:   lea ecx, var_A0
  loc_005999E9: call [00401020h] ; __vbaFreeVar
loc_005999EF:   movsx eax, var_1C
  loc_005999F3: sub esp, 00000010h
  loc_005999F6: mov ecx, 00000003h
loc_005999FB:   mov edx, esp
loc_005999FD:   mov var_220, ecx
loc_00599A03:   mov var_240, ecx
loc_00599A09:   mov var_218, eax
loc_00599A0F:   mov [edx], ecx
loc_00599A11:   mov ecx, var_21C
  loc_00599A17: sub esp, 00000010h
  loc_00599A1A: mov var_238, 00000004h
loc_00599A24:   mov [edx+00000004h], ecx
loc_00599A27:   mov ecx, esp
  loc_00599A29: push 00000002h
  loc_00599A2B: push 00000041h
loc_00599A2D:   mov [edx+00000008h], eax
loc_00599A30:   mov eax, var_214
loc_00599A36:   push ebx
loc_00599A37:   mov [edx+0000000Ch], eax
loc_00599A3A:   mov edx, var_240
loc_00599A40:   mov eax, var_23C
loc_00599A46:   mov [ecx], edx
loc_00599A48:   mov edx, var_238
loc_00599A4E:   mov [ecx+00000004h], eax
loc_00599A51:   mov eax, var_234
loc_00599A57:   mov [ecx+00000008h], edx
loc_00599A5A:   mov [ecx+0000000Ch], eax
loc_00599A5D:   mov ecx, [ebx]
loc_00599A5F:   Call [ecx+00000390h]
loc_00599A65:   lea edx, var_78
loc_00599A68:   push eax
loc_00599A69:   push edx
  loc_00599A6A: call [004010B8h] ; __vbaObjSet
loc_00599A70:   push eax
loc_00599A71:   lea eax, var_A0
loc_00599A77:   push eax
  loc_00599A78: call [0040114Ch] ; __vbaLateIdCallLd
  loc_00599A7E: add esp, 00000030h
loc_00599A81:   lea ecx, var_A0
  loc_00599A87: push 00000000h
loc_00599A89:   push FFFFFFFFh
  loc_00599A8B: push 00000001h
  loc_00599A8D: push 0042DFECh
  loc_00599A92: push 0043024Ch ; "."
loc_00599A97:   push ecx
  loc_00599A98: call [00401034h] ; __vbaStrVarMove
loc_00599A9E:   mov edx, eax
loc_00599AA0:   lea ecx, var_2C
  loc_00599AA3: call __vbaVarDup
loc_00599AA5:   push eax
  loc_00599AA6: call [00401194h] ; rtcReplace
loc_00599AAC:   mov edx, eax
loc_00599AAE:   lea ecx, var_24
  loc_00599AB1: call __vbaVarDup
loc_00599AB3:   lea ecx, var_2C
  loc_00599AB6: call [0040129Ch] ; __vbaFreeStr
loc_00599ABC:   lea ecx, var_78
  loc_00599ABF: call [004012A0h] ; __vbaFreeObj
loc_00599AC5:   lea ecx, var_A0
  loc_00599ACB: call [00401020h] ; __vbaFreeVar
  loc_00599AD1: push 0043463Ch ; "UPDATE brg_barang SET "
  loc_00599AD6: push 00434670h ; " stok=stok + "
loc_00599ADB:   Call edi
loc_00599ADD:   mov edx, eax
loc_00599ADF:   lea ecx, var_2C
  loc_00599AE2: call __vbaVarDup
loc_00599AE4:   mov edx, var_28
loc_00599AE7:   push eax
loc_00599AE8:   push edx
  loc_00599AE9: call [0040100Ch] ; __vbaStrI2
loc_00599AEF:   mov edx, eax
loc_00599AF1:   lea ecx, var_30
  loc_00599AF4: call __vbaVarDup
loc_00599AF6:   push eax
loc_00599AF7:   Call edi
loc_00599AF9:   mov edx, eax
loc_00599AFB:   lea ecx, var_34
  loc_00599AFE: call __vbaVarDup
loc_00599B00:   push eax
  loc_00599B01: push 00433074h
loc_00599B06:   Call edi
loc_00599B08:   mov edx, eax
loc_00599B0A:   lea ecx, var_38
  loc_00599B0D: call __vbaVarDup
loc_00599B0F:   push eax
  loc_00599B10: push 00434690h ; " harga_beli='"
loc_00599B15:   Call edi
loc_00599B17:   mov edx, eax
loc_00599B19:   lea ecx, var_3C
  loc_00599B1C: call __vbaVarDup
loc_00599B1E:   push eax
loc_00599B1F:   mov eax, var_18
loc_00599B22:   push eax
  loc_00599B23: call [004012A4h] ; rtcR8ValFromBstr
  loc_00599B29: sub esp, 00000008h
  loc_00599B2C: fstp real8 ptr [esp]
  loc_00599B2F: call [0040115Ch] ; __vbaStrR8
loc_00599B35:   mov edx, eax
loc_00599B37:   lea ecx, var_40
  loc_00599B3A: call __vbaVarDup
loc_00599B3C:   push eax
loc_00599B3D:   Call edi
loc_00599B3F:   mov edx, eax
loc_00599B41:   lea ecx, var_44
  loc_00599B44: call __vbaVarDup
loc_00599B46:   push eax
  loc_00599B47: push 004346B0h ; "',"
loc_00599B4C:   Call edi
loc_00599B4E:   mov edx, eax
loc_00599B50:   lea ecx, var_48
  loc_00599B53: call __vbaVarDup
loc_00599B55:   push eax
  loc_00599B56: push 00432938h ; " harga_jual='"
loc_00599B5B:   Call edi
loc_00599B5D:   mov edx, eax
loc_00599B5F:   lea ecx, var_4C
  loc_00599B62: call __vbaVarDup
loc_00599B64:   mov ecx, var_24
loc_00599B67:   push eax
loc_00599B68:   push ecx
  loc_00599B69: call [004012A4h] ; rtcR8ValFromBstr
  loc_00599B6F: sub esp, 00000008h
  loc_00599B72: fstp real8 ptr [esp]
  loc_00599B75: call [0040115Ch] ; __vbaStrR8
loc_00599B7B:   mov edx, eax
loc_00599B7D:   lea ecx, var_50
  loc_00599B80: call __vbaVarDup
loc_00599B82:   push eax
loc_00599B83:   Call edi
loc_00599B85:   mov edx, eax
loc_00599B87:   lea ecx, var_54
  loc_00599B8A: call __vbaVarDup
loc_00599B8C:   push eax
  loc_00599B8D: push 0042EBA8h ; "'"
loc_00599B92:   Call edi
loc_00599B94:   mov edx, eax
loc_00599B96:   lea ecx, var_58
  loc_00599B99: call __vbaVarDup
loc_00599B9B:   push eax
  loc_00599B9C: push 00433824h ; " WHERE kode_barang='"
loc_00599BA1:   Call edi
loc_00599BA3:   mov edx, eax
loc_00599BA5:   lea ecx, var_5C
  loc_00599BA8: call __vbaVarDup
loc_00599BAA:   mov edx, var_20
loc_00599BAD:   push eax
loc_00599BAE:   push edx
loc_00599BAF:   Call edi
loc_00599BB1:   mov edx, eax
loc_00599BB3:   lea ecx, var_60
  loc_00599BB6: call __vbaVarDup
loc_00599BB8:   push eax
  loc_00599BB9: push 0042EBA8h ; "'"
loc_00599BBE:   Call edi
loc_00599BC0:   mov edx, eax
  loc_00599BC2: mov ecx, 0066809Ch
  loc_00599BC7: call __vbaVarDup
loc_00599BC9:   lea eax, var_60
loc_00599BCC:   lea ecx, var_5C
loc_00599BCF:   push eax
loc_00599BD0:   lea edx, var_58
loc_00599BD3:   push ecx
loc_00599BD4:   lea eax, var_54
loc_00599BD7:   push edx
loc_00599BD8:   lea ecx, var_50
loc_00599BDB:   push eax
loc_00599BDC:   lea edx, var_4C
loc_00599BDF:   push ecx
loc_00599BE0:   lea eax, var_48
loc_00599BE3:   push edx
loc_00599BE4:   lea ecx, var_44
loc_00599BE7:   push eax
loc_00599BE8:   lea edx, var_40
loc_00599BEB:   push ecx
loc_00599BEC:   lea eax, var_3C
loc_00599BEF:   push edx
loc_00599BF0:   lea ecx, var_38
loc_00599BF3:   push eax
loc_00599BF4:   lea edx, var_34
loc_00599BF7:   push ecx
loc_00599BF8:   lea eax, var_30
loc_00599BFB:   push edx
loc_00599BFC:   lea ecx, var_2C
loc_00599BFF:   push eax
loc_00599C00:   push ecx
  loc_00599C01: push 0000000Eh
  loc_00599C03: call [00401204h] ; __vbaFreeStrList
loc_00599C09:   mov eax, [0066803Ch]
  loc_00599C0E: add esp, 0000003Ch
loc_00599C11:   test eax, eax
  loc_00599C13: jnz 00599C25h
  loc_00599C15: push 0066803Ch
  loc_00599C1A: push 0042DEFCh
  loc_00599C1F: call [004011E8h] ; __vbaNew2
loc_00599C25:   mov edx, [0066803Ch]
loc_00599C2B:   lea ecx, var_A0
loc_00599C31:   mov var_318, edx
  loc_00599C37: mov var_98, 80020004h
  loc_00599C41: mov var_A0, 0000000Ah
  loc_00599C4B: call [0040123Ch] ; __vbaFreeVarg
loc_00599C51:   mov eax, var_318
loc_00599C57:   lea edx, var_78
loc_00599C5A:   push edx
loc_00599C5B:   lea edx, var_A0
loc_00599C61:   mov ecx, [eax]
  loc_00599C63: push 00000001h
loc_00599C65:   push edx
loc_00599C66:   mov edx, [0066809Ch]
loc_00599C6C:   push edx
loc_00599C6D:   push eax
loc_00599C6E:   Call [ecx+00000040h]
loc_00599C71:   test eax, eax
loc_00599C73:   fnclex
  loc_00599C75: jge 00599C8Ch
loc_00599C77:   mov ecx, var_318
  loc_00599C7D: push 00000040h
  loc_00599C7F: push 0042E1B0h
loc_00599C84:   push ecx
loc_00599C85:   push eax
  loc_00599C86: call [00401080h] ; __vbaHresultCheckObj
loc_00599C8C:   lea ecx, var_78
  loc_00599C8F: call [004012A0h] ; __vbaFreeObj
loc_00599C95:   lea ecx, var_A0
  loc_00599C9B: call [00401020h] ; __vbaFreeVar
  loc_00599CA1: mov eax, 00000001h
loc_00599CA6:   Add ax, var_1C
  loc_00599CAA: jo 0059A014h
loc_00599CB0:   mov var_1C, eax
  loc_00599CB3: jmp 005992E0h
  loc_00599CB8: mov esi, [00401238h] ; __vbaVarDup
  loc_00599CBE: mov ecx, 80020004h
loc_00599CC3:   mov var_C8, ecx
  loc_00599CC9: mov eax, 0000000Ah
loc_00599CCE:   mov var_B8, ecx
  loc_00599CD4: mov edi, 00000008h
loc_00599CD9:   lea edx, var_230
loc_00599CDF:   lea ecx, var_B0
loc_00599CE5:   mov var_D0, eax
loc_00599CEB:   mov var_C0, eax
  loc_00599CF1: mov var_228, 0042DF8Ch ; "Info"
loc_00599CFB:   mov var_230, edi
  loc_00599D01: call __vbaVarDup
loc_00599D03:   lea edx, var_220
loc_00599D09:   lea ecx, var_A0
  loc_00599D0F: mov var_218, 004346BCh ; "DATA TRANSAKSI TELAH TERSIMPAN !"
loc_00599D19:   mov var_220, edi
  loc_00599D1F: call __vbaVarDup
loc_00599D21:   lea edx, var_D0
loc_00599D27:   lea eax, var_C0
loc_00599D2D:   push edx
loc_00599D2E:   lea ecx, var_B0
loc_00599D34:   push eax
loc_00599D35:   push ecx
loc_00599D36:   lea edx, var_A0
  loc_00599D3C: push 00000040h
loc_00599D3E:   push edx
  loc_00599D3F: call [004010B4h] ; rtcMsgBox
loc_00599D45:   lea eax, var_D0
loc_00599D4B:   lea ecx, var_C0
loc_00599D51:   push eax
loc_00599D52:   lea edx, var_B0
loc_00599D58:   push ecx
loc_00599D59:   lea eax, var_A0
loc_00599D5F:   push edx
loc_00599D60:   push eax
  loc_00599D61: push 00000004h
  loc_00599D63: call [0040103Ch] ; __vbaFreeVarList
loc_00599D69:   mov ecx, [ebx]
  loc_00599D6B: add esp, 00000014h
loc_00599D6E:   push ebx
loc_00599D6F:   Call [ecx+00000704h]
loc_00599D75:   test eax, eax
  loc_00599D77: jge 00599D8Bh
  loc_00599D79: push 00000704h
  loc_00599D7E: push 00433DD8h
loc_00599D83:   push ebx
loc_00599D84:   push eax
  loc_00599D85: call [00401080h] ; __vbaHresultCheckObj
loc_00599D8B:   mov edx, [ebx]
loc_00599D8D:   push ebx
loc_00599D8E:   Call [edx+00000728h]
loc_00599D94:   test eax, eax
  loc_00599D96: jge 00599E92h
  loc_00599D9C: push 00000728h
  loc_00599DA1: push 00433DD8h
loc_00599DA6:   push ebx
loc_00599DA7:   push eax
  loc_00599DA8: call [00401080h] ; __vbaHresultCheckObj
  loc_00599DAE: jmp 00599E92h
  loc_00599DB3: mov ecx, 80020004h
  loc_00599DB8: mov eax, 0000000Ah
loc_00599DBD:   mov var_C8, ecx
loc_00599DC3:   mov var_B8, ecx
loc_00599DC9:   mov var_A8, ecx
loc_00599DCF:   lea edx, var_220
loc_00599DD5:   lea ecx, var_A0
loc_00599DDB:   mov var_D0, eax
loc_00599DE1:   mov var_C0, eax
loc_00599DE7:   mov var_B0, eax
  loc_00599DED: mov var_218, 00434818h ; "nomor faktur ganda, ganti dengan nomor lain"
  loc_00599DF7: mov var_220, 00000008h
  loc_00599E01: call [00401238h] ; __vbaVarDup
loc_00599E07:   lea eax, var_D0
loc_00599E0D:   lea ecx, var_C0
loc_00599E13:   push eax
loc_00599E14:   lea edx, var_B0
loc_00599E1A:   push ecx
loc_00599E1B:   push edx
loc_00599E1C:   lea eax, var_A0
  loc_00599E22: push 00000000h
loc_00599E24:   push eax
  loc_00599E25: call [004010B4h] ; rtcMsgBox
loc_00599E2B:   lea ecx, var_D0
loc_00599E31:   lea edx, var_C0
loc_00599E37:   push ecx
loc_00599E38:   lea eax, var_B0
loc_00599E3E:   push edx
loc_00599E3F:   lea ecx, var_A0
loc_00599E45:   push eax
loc_00599E46:   push ecx
  loc_00599E47: push 00000004h
  loc_00599E49: call [0040103Ch] ; __vbaFreeVarList
loc_00599E4F:   mov edx, [ebx]
  loc_00599E51: add esp, 00000014h
loc_00599E54:   push ebx
loc_00599E55:   Call [edx+0000031Ch]
loc_00599E5B:   push eax
loc_00599E5C:   lea eax, var_78
loc_00599E5F:   push eax
  loc_00599E60: call [004010B8h] ; __vbaObjSet
loc_00599E66:   mov esi, eax
loc_00599E68:   push esi
loc_00599E69:   mov ecx, [esi]
loc_00599E6B:   Call [ecx+00000204h]
loc_00599E71:   test eax, eax
loc_00599E73:   fnclex
  loc_00599E75: jge 00599E89h
  loc_00599E77: push 00000204h
  loc_00599E7C: push 0042DFCCh
loc_00599E81:   push esi
loc_00599E82:   push eax
  loc_00599E83: call [00401080h] ; __vbaHresultCheckObj
loc_00599E89:   lea ecx, var_78
  loc_00599E8C: call [004012A0h] ; __vbaFreeObj
  loc_00599E92: mov var_4, 00000000h
loc_00599E99:   fwait
  loc_00599E9A: push 00599FF5h
  loc_00599E9F: jmp 00599FDFh
loc_00599EA4:   lea edx, var_74
loc_00599EA7:   lea eax, var_70
loc_00599EAA:   push edx
loc_00599EAB:   lea ecx, var_6C
loc_00599EAE:   push eax
loc_00599EAF:   lea edx, var_68
loc_00599EB2:   push ecx
loc_00599EB3:   lea eax, var_64
loc_00599EB6:   push edx
loc_00599EB7:   lea ecx, var_60
loc_00599EBA:   push eax
loc_00599EBB:   lea edx, var_5C
loc_00599EBE:   push ecx
loc_00599EBF:   lea eax, var_58
loc_00599EC2:   push edx
loc_00599EC3:   lea ecx, var_54
loc_00599EC6:   push eax
loc_00599EC7:   lea edx, var_50
loc_00599ECA:   push ecx
loc_00599ECB:   lea eax, var_4C
loc_00599ECE:   push edx
loc_00599ECF:   lea ecx, var_48
loc_00599ED2:   push eax
loc_00599ED3:   lea edx, var_44
loc_00599ED6:   push ecx
loc_00599ED7:   lea eax, var_40
loc_00599EDA:   push edx
loc_00599EDB:   lea ecx, var_3C
loc_00599EDE:   push eax
loc_00599EDF:   lea edx, var_38
loc_00599EE2:   push ecx
loc_00599EE3:   lea eax, var_34
loc_00599EE6:   push edx
loc_00599EE7:   lea ecx, var_30
loc_00599EEA:   push eax
loc_00599EEB:   lea edx, var_2C
loc_00599EEE:   push ecx
loc_00599EEF:   push edx
  loc_00599EF0: push 00000013h
  loc_00599EF2: call [00401204h] ; __vbaFreeStrList
loc_00599EF8:   lea eax, var_90
loc_00599EFE:   lea ecx, var_8C
loc_00599F04:   push eax
loc_00599F05:   lea edx, var_88
loc_00599F0B:   push ecx
loc_00599F0C:   lea eax, var_84
loc_00599F12:   push edx
loc_00599F13:   lea ecx, var_80
loc_00599F16:   push eax
loc_00599F17:   lea edx, var_7C
loc_00599F1A:   push ecx
loc_00599F1B:   lea eax, var_78
loc_00599F1E:   push edx
loc_00599F1F:   push eax
  loc_00599F20: push 00000007h
  loc_00599F22: call [0040104Ch] ; __vbaFreeObjList
  loc_00599F28: add esp, 00000070h
loc_00599F2B:   lea ecx, var_210
loc_00599F31:   lea edx, var_200
loc_00599F37:   lea eax, var_1F0
loc_00599F3D:   push ecx
loc_00599F3E:   push edx
loc_00599F3F:   lea ecx, var_1E0
loc_00599F45:   push eax
loc_00599F46:   lea edx, var_1D0
loc_00599F4C:   push ecx
loc_00599F4D:   lea eax, var_1C0
loc_00599F53:   push edx
loc_00599F54:   lea ecx, var_1B0
loc_00599F5A:   push eax
loc_00599F5B:   lea edx, var_1A0
loc_00599F61:   push ecx
loc_00599F62:   lea eax, var_190
loc_00599F68:   push edx
loc_00599F69:   lea ecx, var_180
loc_00599F6F:   push eax
loc_00599F70:   lea edx, var_170
loc_00599F76:   push ecx
loc_00599F77:   push edx
loc_00599F78:   lea eax, var_160
loc_00599F7E:   lea ecx, var_150
loc_00599F84:   push eax
loc_00599F85:   lea edx, var_140
loc_00599F8B:   push ecx
loc_00599F8C:   lea eax, var_130
loc_00599F92:   push edx
loc_00599F93:   lea ecx, var_120
loc_00599F99:   push eax
loc_00599F9A:   lea edx, var_110
loc_00599FA0:   push ecx
loc_00599FA1:   lea eax, var_100
loc_00599FA7:   push edx
loc_00599FA8:   lea ecx, var_F0
loc_00599FAE:   push eax
loc_00599FAF:   lea edx, var_E0
loc_00599FB5:   push ecx
loc_00599FB6:   lea eax, var_D0
loc_00599FBC:   push edx
loc_00599FBD:   lea ecx, var_C0
loc_00599FC3:   push eax
loc_00599FC4:   lea edx, var_B0
loc_00599FCA:   push ecx
loc_00599FCB:   lea eax, var_A0
loc_00599FD1:   push edx
loc_00599FD2:   push eax
  loc_00599FD3: push 00000018h
  loc_00599FD5: call [0040103Ch] ; __vbaFreeVarList
  loc_00599FDB: add esp, 00000064h
loc_00599FDE:   ret
  loc_00599FDF: mov esi, [0040129Ch] ; __vbaFreeStr
loc_00599FE5:   lea ecx, var_18
  loc_00599FE8: call __vbaFreeStr
loc_00599FEA:   lea ecx, var_20
  loc_00599FED: call __vbaFreeStr
loc_00599FEF:   lea ecx, var_24
  loc_00599FF2: call __vbaFreeStr
loc_00599FF4:   ret
loc_00599FF5:   mov eax, Me
loc_00599FF8:   push eax
loc_00599FF9:   mov ecx, [eax]
loc_00599FFB:   Call [ecx+00000008h]
loc_00599FFE:   mov eax, var_4
loc_0059A001:   mov ecx, var_14
loc_0059A004:   pop edi
loc_0059A005:   pop esi
loc_0059A006:   mov fs: [00000000h] , ecx
loc_0059A00D:   pop ebx
loc_0059A00E:   mov esp, ebp
loc_0059A010:   pop ebp
  loc_0059A011: retn 0004h
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer) '59CE70
loc_0059CE70:   push ebp
loc_0059CE71:   mov ebp, esp
  loc_0059CE73: sub esp, 0000000Ch
  loc_0059CE76: push 00405696h ; __vbaExceptHandler
loc_0059CE7B:   mov eax, fs: [00000000h]
loc_0059CE81:   push eax
loc_0059CE82:   mov fs: [00000000h] , esp
  loc_0059CE89: sub esp, 00000028h
loc_0059CE8C:   push ebx
loc_0059CE8D:   push esi
loc_0059CE8E:   push edi
loc_0059CE8F:   mov var_C, esp
  loc_0059CE92: mov var_8, 00402960h
loc_0059CE99:   mov eax, Me
loc_0059CE9C:   mov ecx, eax
  loc_0059CE9E: and ecx, 00000001h
loc_0059CEA1:   mov var_4, ecx
  loc_0059CEA4: and al, FEh
loc_0059CEA6:   push eax
loc_0059CEA7:   mov Me, eax
loc_0059CEAA:   mov edx, [eax]
loc_0059CEAC:   Call [edx+00000004h]
loc_0059CEAF:   mov esi, KeyAscii
  loc_0059CEB2: xor edi, edi
loc_0059CEB4:   mov var_24, edi
  loc_0059CEB7: cmp [esi], 000Dh
  loc_0059CEBB: jnz 0059CEE6h
loc_0059CEBD:   lea eax, var_24
  loc_0059CEC0: mov var_1C, 80020004h
loc_0059CEC7:   push eax
  loc_0059CEC8: push 004300CCh ; "{tab}"
  loc_0059CECD: mov var_24, 0000000Ah
  loc_0059CED4: call [004010D4h] ; rtcSendKeys
loc_0059CEDA:   lea ecx, var_24
  loc_0059CEDD: call [00401020h] ; __vbaFreeVar
loc_0059CEE3:   mov [esi], di
loc_0059CEE6:   mov var_4, edi
  loc_0059CEE9: push 0059CEFBh
  loc_0059CEEE: jmp 0059CEFAh
loc_0059CEF0:   lea ecx, var_24
  loc_0059CEF3: call [00401020h] ; __vbaFreeVar
loc_0059CEF9:   ret
loc_0059CEFA:   ret
loc_0059CEFB:   mov eax, Me
loc_0059CEFE:   push eax
loc_0059CEFF:   mov ecx, [eax]
loc_0059CF01:   Call [ecx+00000008h]
loc_0059CF04:   mov eax, var_4
loc_0059CF07:   mov ecx, var_14
loc_0059CF0A:   pop edi
loc_0059CF0B:   pop esi
loc_0059CF0C:   mov fs: [00000000h] , ecx
loc_0059CF13:   pop ebx
loc_0059CF14:   mov esp, ebp
loc_0059CF16:   pop ebp
  loc_0059CF17: retn 0008h
End Sub

Private Sub CmdBaru_UnknownEvent_10() '595BC0
loc_00595BC0:   push ebp
loc_00595BC1:   mov ebp, esp
  loc_00595BC3: sub esp, 0000000Ch
  loc_00595BC6: push 00405696h ; __vbaExceptHandler
loc_00595BCB:   mov eax, fs: [00000000h]
loc_00595BD1:   push eax
loc_00595BD2:   mov fs: [00000000h] , esp
  loc_00595BD9: sub esp, 00000034h
loc_00595BDC:   push ebx
loc_00595BDD:   push esi
loc_00595BDE:   push edi
loc_00595BDF:   mov var_C, esp
  loc_00595BE2: mov var_8, 004027D0h
loc_00595BE9:   mov esi, Me
loc_00595BEC:   mov eax, esi
  loc_00595BEE: and eax, 00000001h
loc_00595BF1:   mov var_4, eax
  loc_00595BF4: and esi, FFFFFFFEh
loc_00595BF7:   push esi
loc_00595BF8:   mov Me, esi
loc_00595BFB:   mov ecx, [esi]
loc_00595BFD:   Call [ecx+00000004h]
loc_00595C00:   mov edx, [esi]
loc_00595C02:   push esi
  loc_00595C03: mov var_18, 00000000h
loc_00595C0A:   Call [edx+00000708h]
loc_00595C10:   test eax, eax
  loc_00595C12: jge 00595C26h
  loc_00595C14: push 00000708h
  loc_00595C19: push 00433DD8h
loc_00595C1E:   push esi
loc_00595C1F:   push eax
  loc_00595C20: call [00401080h] ; __vbaHresultCheckObj
loc_00595C26:   mov eax, [esi]
loc_00595C28:   push esi
loc_00595C29:   Call [eax+00000314h]
  loc_00595C2F: mov ebx, [004010B8h] ; __vbaObjSet
loc_00595C35:   lea ecx, var_18
loc_00595C38:   push eax
loc_00595C39:   push ecx
loc_00595C3A:   Call ebx
loc_00595C3C:   mov edi, eax
loc_00595C3E:   mov eax, [esi+00000040h]
loc_00595C41:   push eax
loc_00595C42:   push edi
loc_00595C43:   mov edx, [edi]
loc_00595C45:   Call [edx+000000A4h]
loc_00595C4B:   test eax, eax
loc_00595C4D:   fnclex
  loc_00595C4F: jge 00595C63h
  loc_00595C51: push 000000A4h
  loc_00595C56: push 0042DFCCh
loc_00595C5B:   push edi
loc_00595C5C:   push eax
  loc_00595C5D: call [00401080h] ; __vbaHresultCheckObj
  loc_00595C63: mov edi, [004012A0h] ; __vbaFreeObj
loc_00595C69:   lea ecx, var_18
loc_00595C6C:   Call edi
loc_00595C6E:   mov ecx, [esi]
loc_00595C70:   push esi
loc_00595C71:   Call [ecx+00000700h]
loc_00595C77:   test eax, eax
  loc_00595C79: jge 00595C8Dh
  loc_00595C7B: push 00000700h
  loc_00595C80: push 00433DD8h
loc_00595C85:   push esi
loc_00595C86:   push eax
  loc_00595C87: call [00401080h] ; __vbaHresultCheckObj
loc_00595C8D:   mov edx, [esi]
loc_00595C8F:   push esi
loc_00595C90:   Call [edx+0000070Ch]
loc_00595C96:   test eax, eax
  loc_00595C98: jge 00595CACh
  loc_00595C9A: push 0000070Ch
  loc_00595C9F: push 00433DD8h
loc_00595CA4:   push esi
loc_00595CA5:   push eax
  loc_00595CA6: call [00401080h] ; __vbaHresultCheckObj
  loc_00595CAC: sub esp, 00000010h
  loc_00595CAF: mov ecx, 0000000Bh
loc_00595CB4:   mov edx, esp
  loc_00595CB6: xor eax, eax
  loc_00595CB8: push 8001000Dh
loc_00595CBD:   push esi
loc_00595CBE:   mov [edx], ecx
loc_00595CC0:   mov ecx, var_24
loc_00595CC3:   mov [edx+00000004h], ecx
loc_00595CC6:   mov ecx, [esi]
loc_00595CC8:   mov [edx+00000008h], eax
loc_00595CCB:   mov eax, var_1C
loc_00595CCE:   mov [edx+0000000Ch], eax
loc_00595CD1:   Call [ecx+00000394h]
loc_00595CD7:   lea edx, var_18
loc_00595CDA:   push eax
loc_00595CDB:   push edx
loc_00595CDC:   Call ebx
loc_00595CDE:   push eax
  loc_00595CDF: call [00401280h] ; __vbaLateIdSt
loc_00595CE5:   lea ecx, var_18
loc_00595CE8:   Call edi
  loc_00595CEA: sub esp, 00000010h
  loc_00595CED: mov ecx, 0000000Bh
loc_00595CF2:   mov edx, esp
  loc_00595CF4: or eax, FFFFFFFFh
  loc_00595CF7: push 8001000Dh
loc_00595CFC:   push esi
loc_00595CFD:   mov [edx], ecx
loc_00595CFF:   mov ecx, var_24
loc_00595D02:   mov [edx+00000004h], ecx
loc_00595D05:   mov ecx, [esi]
loc_00595D07:   mov [edx+00000008h], eax
loc_00595D0A:   mov eax, var_1C
loc_00595D0D:   mov [edx+0000000Ch], eax
loc_00595D10:   Call [ecx+00000388h]
loc_00595D16:   lea edx, var_18
loc_00595D19:   push eax
loc_00595D1A:   push edx
loc_00595D1B:   Call ebx
loc_00595D1D:   push eax
  loc_00595D1E: call [00401280h] ; __vbaLateIdSt
loc_00595D24:   lea ecx, var_18
loc_00595D27:   Call edi
  loc_00595D29: sub esp, 00000010h
  loc_00595D2C: mov ecx, 00000008h
loc_00595D31:   mov edx, esp
  loc_00595D33: mov eax, 004341FCh ; "&Batal"
loc_00595D38:   push FFFFFDFAh
loc_00595D3D:   push esi
loc_00595D3E:   mov [edx], ecx
loc_00595D40:   mov ecx, var_24
loc_00595D43:   mov [edx+00000004h], ecx
loc_00595D46:   mov ecx, [esi]
loc_00595D48:   mov [edx+00000008h], eax
loc_00595D4B:   mov eax, var_1C
loc_00595D4E:   mov [edx+0000000Ch], eax
loc_00595D51:   Call [ecx+0000039Ch]
loc_00595D57:   lea edx, var_18
loc_00595D5A:   push eax
loc_00595D5B:   push edx
loc_00595D5C:   Call ebx
loc_00595D5E:   push eax
  loc_00595D5F: call [00401280h] ; __vbaLateIdSt
loc_00595D65:   lea ecx, var_18
loc_00595D68:   Call edi
  loc_00595D6A: sub esp, 00000010h
  loc_00595D6D: mov ecx, 0000000Bh
loc_00595D72:   mov edx, esp
  loc_00595D74: or eax, FFFFFFFFh
  loc_00595D77: push 8001000Dh
loc_00595D7C:   push esi
loc_00595D7D:   mov [edx], ecx
loc_00595D7F:   mov ecx, var_24
loc_00595D82:   mov [edx+00000004h], ecx
loc_00595D85:   mov ecx, [esi]
loc_00595D87:   mov [edx+00000008h], eax
loc_00595D8A:   mov eax, var_1C
loc_00595D8D:   mov [edx+0000000Ch], eax
loc_00595D90:   Call [ecx+00000398h]
loc_00595D96:   push eax
loc_00595D97:   lea edx, var_18
loc_00595D9A:   push edx
loc_00595D9B:   Call ebx
loc_00595D9D:   push eax
  loc_00595D9E: call [00401280h] ; __vbaLateIdSt
loc_00595DA4:   lea ecx, var_18
loc_00595DA7:   Call edi
  loc_00595DA9: sub esp, 00000010h
  loc_00595DAC: mov ecx, 0000000Bh
loc_00595DB1:   mov edx, esp
  loc_00595DB3: or eax, FFFFFFFFh
  loc_00595DB6: push 8001000Dh
loc_00595DBB:   push esi
loc_00595DBC:   mov [edx], ecx
loc_00595DBE:   mov ecx, var_24
loc_00595DC1:   mov [edx+00000004h], ecx
loc_00595DC4:   mov ecx, [esi]
loc_00595DC6:   mov [edx+00000008h], eax
loc_00595DC9:   mov eax, var_1C
loc_00595DCC:   mov [edx+0000000Ch], eax
loc_00595DCF:   Call [ecx+0000038Ch]
loc_00595DD5:   lea edx, var_18
loc_00595DD8:   push eax
loc_00595DD9:   push edx
loc_00595DDA:   Call ebx
loc_00595DDC:   push eax
  loc_00595DDD: call [00401280h] ; __vbaLateIdSt
loc_00595DE3:   lea ecx, var_18
loc_00595DE6:   Call edi
loc_00595DE8:   mov eax, [esi]
  loc_00595DEA: push 00000000h
  loc_00595DEC: push 80011000h
loc_00595DF1:   push esi
loc_00595DF2:   Call [eax+00000384h]
loc_00595DF8:   lea ecx, var_18
loc_00595DFB:   push eax
loc_00595DFC:   push ecx
loc_00595DFD:   Call ebx
loc_00595DFF:   push eax
  loc_00595E00: call [0040102Ch] ; __vbaLateIdCall
  loc_00595E06: add esp, 0000000Ch
loc_00595E09:   lea ecx, var_18
loc_00595E0C:   Call edi
  loc_00595E0E: mov [esi+00000034h], 0001h
  loc_00595E14: mov var_4, 00000000h
  loc_00595E1B: push 00595E2Dh
  loc_00595E20: jmp 00595E2Ch
loc_00595E22:   lea ecx, var_18
  loc_00595E25: call [004012A0h] ; __vbaFreeObj
loc_00595E2B:   ret
loc_00595E2C:   ret
loc_00595E2D:   mov eax, Me
loc_00595E30:   push eax
loc_00595E31:   mov edx, [eax]
loc_00595E33:   Call [edx+00000008h]
loc_00595E36:   mov eax, var_4
loc_00595E39:   mov ecx, var_14
loc_00595E3C:   pop edi
loc_00595E3D:   pop esi
loc_00595E3E:   mov fs: [00000000h] , ecx
loc_00595E45:   pop ebx
loc_00595E46:   mov esp, ebp
loc_00595E48:   pop ebp
  loc_00595E49: retn 0004h
End Sub

Private Sub cmdKeluar_UnknownEvent_10() '59CCB0
loc_0059CCB0:   push ebp
loc_0059CCB1:   mov ebp, esp
  loc_0059CCB3: sub esp, 0000000Ch
  loc_0059CCB6: push 00405696h ; __vbaExceptHandler
loc_0059CCBB:   mov eax, fs: [00000000h]
loc_0059CCC1:   push eax
loc_0059CCC2:   mov fs: [00000000h] , esp
  loc_0059CCC9: sub esp, 0000002Ch
loc_0059CCCC:   push ebx
loc_0059CCCD:   push esi
loc_0059CCCE:   push edi
loc_0059CCCF:   mov var_C, esp
  loc_0059CCD2: mov var_8, 00402950h
loc_0059CCD9:   mov edi, Me
loc_0059CCDC:   mov eax, edi
  loc_0059CCDE: and eax, 00000001h
loc_0059CCE1:   mov var_4, eax
  loc_0059CCE4: and edi, FFFFFFFEh
loc_0059CCE7:   push edi
loc_0059CCE8:   mov Me, edi
loc_0059CCEB:   mov ecx, [edi]
loc_0059CCED:   Call [ecx+00000004h]
loc_0059CCF0:   mov edx, [edi]
  loc_0059CCF2: xor ebx, ebx
loc_0059CCF4:   push ebx
loc_0059CCF5:   push FFFFFDFAh
loc_0059CCFA:   push edi
loc_0059CCFB:   mov var_18, ebx
loc_0059CCFE:   mov var_1C, ebx
loc_0059CD01:   mov var_2C, ebx
loc_0059CD04:   Call [edx+0000039Ch]
loc_0059CD0A:   push eax
loc_0059CD0B:   lea eax, var_1C
loc_0059CD0E:   push eax
  loc_0059CD0F: call [004010B8h] ; __vbaObjSet
loc_0059CD15:   lea ecx, var_2C
loc_0059CD18:   push eax
loc_0059CD19:   push ecx
  loc_0059CD1A: call [0040114Ch] ; __vbaLateIdCallLd
  loc_0059CD20: add esp, 00000010h
loc_0059CD23:   push eax
  loc_0059CD24: call [00401034h] ; __vbaStrVarMove
loc_0059CD2A:   mov edx, eax
loc_0059CD2C:   lea ecx, var_18
  loc_0059CD2F: call [0040126Ch] ; __vbaStrMove
loc_0059CD35:   push eax
  loc_0059CD36: push 0042E930h ; "&Keluar"
  loc_0059CD3B: call [0040112Ch] ; __vbaStrCmp
loc_0059CD41:   mov esi, eax
loc_0059CD43:   lea ecx, var_18
loc_0059CD46:   neg esi
loc_0059CD48:   sbb esi, esi
loc_0059CD4A:   inc esi
loc_0059CD4B:   neg esi
  loc_0059CD4D: call [0040129Ch] ; __vbaFreeStr
loc_0059CD53:   lea ecx, var_1C
  loc_0059CD56: call [004012A0h] ; __vbaFreeObj
loc_0059CD5C:   lea ecx, var_2C
  loc_0059CD5F: call [00401020h] ; __vbaFreeVar
loc_0059CD65:   cmp si, bx
  loc_0059CD68: jz 0059CE05h
loc_0059CD6E:   cmp [00668228h], ebx
  loc_0059CD74: jnz 0059CD86h
  loc_0059CD76: push 00668228h
  loc_0059CD7B: push 004296BCh
  loc_0059CD80: call [004011E8h] ; __vbaNew2
loc_0059CD86:   mov esi, [00668228h]
loc_0059CD8C:   push FFFFFFFFh
loc_0059CD8E:   push esi
loc_0059CD8F:   mov edx, [esi]
loc_0059CD91:   Call [edx+00000094h]
loc_0059CD97:   cmp eax, ebx
loc_0059CD99:   fnclex
  loc_0059CD9B: jge 0059CDAFh
  loc_0059CD9D: push 00000094h
  loc_0059CDA2: push 00433880h
loc_0059CDA7:   push esi
loc_0059CDA8:   push eax
  loc_0059CDA9: call [00401080h] ; __vbaHresultCheckObj
loc_0059CDAF:   cmp [0066A318h], ebx
  loc_0059CDB5: jnz 0059CDC7h
  loc_0059CDB7: push 0066A318h
  loc_0059CDBC: push 0042E390h
  loc_0059CDC1: call [004011E8h] ; __vbaNew2
loc_0059CDC7:   mov esi, [0066A318h]
loc_0059CDCD:   lea eax, var_1C
loc_0059CDD0:   push edi
loc_0059CDD1:   push eax
loc_0059CDD2:   mov edx, [esi]
loc_0059CDD4:   mov var_40, edx
  loc_0059CDD7: call [004010C8h] ; __vbaObjSetAddref
loc_0059CDDD:   mov ecx, var_40
loc_0059CDE0:   push eax
loc_0059CDE1:   push esi
loc_0059CDE2:   Call [ecx+00000010h]
loc_0059CDE5:   cmp eax, ebx
loc_0059CDE7:   fnclex
  loc_0059CDE9: jge 0059CDFAh
  loc_0059CDEB: push 00000010h
  loc_0059CDED: push 0042E380h
loc_0059CDF2:   push esi
loc_0059CDF3:   push eax
  loc_0059CDF4: call [00401080h] ; __vbaHresultCheckObj
loc_0059CDFA:   lea ecx, var_1C
  loc_0059CDFD: call [004012A0h] ; __vbaFreeObj
  loc_0059CE03: jmp 0059CE24h
loc_0059CE05:   mov edx, [edi]
loc_0059CE07:   push edi
loc_0059CE08:   Call [edx+00000704h]
loc_0059CE0E:   cmp eax, ebx
  loc_0059CE10: jge 0059CE24h
  loc_0059CE12: push 00000704h
  loc_0059CE17: push 00433DD8h
loc_0059CE1C:   push edi
loc_0059CE1D:   push eax
  loc_0059CE1E: call [00401080h] ; __vbaHresultCheckObj
loc_0059CE24:   mov var_4, ebx
  loc_0059CE27: push 0059CE4Bh
  loc_0059CE2C: jmp 0059CE4Ah
loc_0059CE2E:   lea ecx, var_18
  loc_0059CE31: call [0040129Ch] ; __vbaFreeStr
loc_0059CE37:   lea ecx, var_1C
  loc_0059CE3A: call [004012A0h] ; __vbaFreeObj
loc_0059CE40:   lea ecx, var_2C
  loc_0059CE43: call [00401020h] ; __vbaFreeVar
loc_0059CE49:   ret
loc_0059CE4A:   ret
loc_0059CE4B:   mov eax, Me
loc_0059CE4E:   push eax
loc_0059CE4F:   mov ecx, [eax]
loc_0059CE51:   Call [ecx+00000008h]
loc_0059CE54:   mov eax, var_4
loc_0059CE57:   mov ecx, var_14
loc_0059CE5A:   pop edi
loc_0059CE5B:   pop esi
loc_0059CE5C:   mov fs: [00000000h] , ecx
loc_0059CE63:   pop ebx
loc_0059CE64:   mov esp, ebp
loc_0059CE66:   pop ebp
  loc_0059CE67: retn 0004h
End Sub

Private Sub Form_Load() '592C40
loc_00592C40:   push ebp
loc_00592C41:   mov ebp, esp
  loc_00592C43: sub esp, 0000000Ch
  loc_00592C46: push 00405696h ; __vbaExceptHandler
loc_00592C4B:   mov eax, fs: [00000000h]
loc_00592C51:   push eax
loc_00592C52:   mov fs: [00000000h] , esp
  loc_00592C59: sub esp, 00000034h
loc_00592C5C:   push ebx
loc_00592C5D:   push esi
loc_00592C5E:   push edi
loc_00592C5F:   mov var_C, esp
  loc_00592C62: mov var_8, 00402730h
loc_00592C69:   mov esi, Me
loc_00592C6C:   mov eax, esi
  loc_00592C6E: and eax, 00000001h
loc_00592C71:   mov var_4, eax
  loc_00592C74: and esi, FFFFFFFEh
loc_00592C77:   push esi
loc_00592C78:   mov Me, esi
loc_00592C7B:   mov ecx, [esi]
loc_00592C7D:   Call [ecx+00000004h]
  loc_00592C80: mov var_18, 00000000h
  loc_00592C87: call 0055B320h
loc_00592C8C:   mov edx, [esi]
loc_00592C8E:   push esi
loc_00592C8F:   Call [edx+000006FCh]
loc_00592C95:   test eax, eax
  loc_00592C97: jge 00592CABh
  loc_00592C99: push 000006FCh
  loc_00592C9E: push 00433DD8h
loc_00592CA3:   push esi
loc_00592CA4:   push eax
  loc_00592CA5: call [00401080h] ; __vbaHresultCheckObj
loc_00592CAB:   mov eax, [esi]
loc_00592CAD:   push esi
loc_00592CAE:   Call [eax+00000714h]
loc_00592CB4:   test eax, eax
  loc_00592CB6: jge 00592CCAh
  loc_00592CB8: push 00000714h
  loc_00592CBD: push 00433DD8h
loc_00592CC2:   push esi
loc_00592CC3:   push eax
  loc_00592CC4: call [00401080h] ; __vbaHresultCheckObj
  loc_00592CCA: sub esp, 00000010h
  loc_00592CCD: mov ecx, 0000000Bh
loc_00592CD2:   mov edx, esp
  loc_00592CD4: xor eax, eax
  loc_00592CD6: push 8001000Dh
loc_00592CDB:   push esi
loc_00592CDC:   mov [edx], ecx
loc_00592CDE:   mov ecx, var_24
loc_00592CE1:   mov [edx+00000004h], ecx
loc_00592CE4:   mov ecx, [esi]
loc_00592CE6:   mov [edx+00000008h], eax
loc_00592CE9:   mov eax, var_1C
loc_00592CEC:   mov [edx+0000000Ch], eax
loc_00592CEF:   Call [ecx+00000398h]
  loc_00592CF5: mov edi, [004010B8h] ; __vbaObjSet
loc_00592CFB:   lea edx, var_18
loc_00592CFE:   push eax
loc_00592CFF:   push edx
loc_00592D00:   Call edi
loc_00592D02:   push eax
  loc_00592D03: call [00401280h] ; __vbaLateIdSt
  loc_00592D09: mov ebx, [004012A0h] ; __vbaFreeObj
loc_00592D0F:   lea ecx, var_18
loc_00592D12:   Call ebx
  loc_00592D14: sub esp, 00000010h
  loc_00592D17: mov ecx, 0000000Bh
loc_00592D1C:   mov edx, esp
  loc_00592D1E: xor eax, eax
  loc_00592D20: push 8001000Dh
loc_00592D25:   push esi
loc_00592D26:   mov [edx], ecx
loc_00592D28:   mov ecx, var_24
loc_00592D2B:   mov [edx+00000004h], ecx
loc_00592D2E:   mov ecx, [esi]
loc_00592D30:   mov [edx+00000008h], eax
loc_00592D33:   mov eax, var_1C
loc_00592D36:   mov [edx+0000000Ch], eax
loc_00592D39:   Call [ecx+00000388h]
loc_00592D3F:   lea edx, var_18
loc_00592D42:   push eax
loc_00592D43:   push edx
loc_00592D44:   Call edi
loc_00592D46:   push eax
  loc_00592D47: call [00401280h] ; __vbaLateIdSt
loc_00592D4D:   lea ecx, var_18
loc_00592D50:   Call ebx
  loc_00592D52: sub esp, 00000010h
  loc_00592D55: mov ecx, 0000000Bh
loc_00592D5A:   mov edx, esp
  loc_00592D5C: xor eax, eax
  loc_00592D5E: push 8001000Dh
loc_00592D63:   push esi
loc_00592D64:   mov [edx], ecx
loc_00592D66:   mov ecx, var_24
loc_00592D69:   mov [edx+00000004h], ecx
loc_00592D6C:   mov ecx, [esi]
loc_00592D6E:   mov [edx+00000008h], eax
loc_00592D71:   mov eax, var_1C
loc_00592D74:   mov [edx+0000000Ch], eax
loc_00592D77:   Call [ecx+0000038Ch]
loc_00592D7D:   lea edx, var_18
loc_00592D80:   push eax
loc_00592D81:   push edx
loc_00592D82:   Call edi
loc_00592D84:   push eax
  loc_00592D85: call [00401280h] ; __vbaLateIdSt
loc_00592D8B:   lea ecx, var_18
loc_00592D8E:   Call ebx
loc_00592D90:   mov eax, [esi]
loc_00592D92:   push esi
loc_00592D93:   Call [eax+00000314h]
loc_00592D99:   lea ecx, var_18
loc_00592D9C:   push eax
loc_00592D9D:   push ecx
loc_00592D9E:   Call edi
loc_00592DA0:   mov edx, [eax]
loc_00592DA2:   push FFFFFFFFh
loc_00592DA4:   push eax
loc_00592DA5:   mov var_3C, eax
loc_00592DA8:   Call [edx+000001B4h]
loc_00592DAE:   fnclex
loc_00592DB0:   test eax, eax
  loc_00592DB2: jge 00592DC9h
loc_00592DB4:   mov ecx, var_3C
  loc_00592DB7: push 000001B4h
  loc_00592DBC: push 0042DFCCh
loc_00592DC1:   push ecx
loc_00592DC2:   push eax
  loc_00592DC3: call [00401080h] ; __vbaHresultCheckObj
loc_00592DC9:   lea ecx, var_18
loc_00592DCC:   Call ebx
  loc_00592DCE: sub esp, 00000010h
  loc_00592DD1: mov ecx, 00000008h
loc_00592DD6:   mov edx, esp
  loc_00592DD8: mov eax, 00433F74h ; "Tunai"
  loc_00592DDD: push 00000001h
loc_00592DDF:   push FFFFFDD7h
loc_00592DE4:   mov [edx], ecx
loc_00592DE6:   mov ecx, var_24
loc_00592DE9:   push esi
loc_00592DEA:   mov [edx+00000004h], ecx
loc_00592DED:   mov ecx, [esi]
loc_00592DEF:   mov [edx+00000008h], eax
loc_00592DF2:   mov eax, var_1C
loc_00592DF5:   mov [edx+0000000Ch], eax
loc_00592DF8:   Call [ecx+00000380h]
loc_00592DFE:   lea edx, var_18
loc_00592E01:   push eax
loc_00592E02:   push edx
loc_00592E03:   Call edi
loc_00592E05:   push eax
  loc_00592E06: call [0040102Ch] ; __vbaLateIdCall
  loc_00592E0C: add esp, 0000001Ch
loc_00592E0F:   lea ecx, var_18
loc_00592E12:   Call ebx
  loc_00592E14: sub esp, 00000010h
  loc_00592E17: mov ecx, 00000008h
loc_00592E1C:   mov edx, esp
  loc_00592E1E: mov eax, 00433F94h ; "Kredit"
  loc_00592E23: push 00000001h
loc_00592E25:   push FFFFFDD7h
loc_00592E2A:   mov [edx], ecx
loc_00592E2C:   mov ecx, var_24
loc_00592E2F:   push esi
loc_00592E30:   mov [edx+00000004h], ecx
loc_00592E33:   mov ecx, [esi]
loc_00592E35:   mov [edx+00000008h], eax
loc_00592E38:   mov eax, var_1C
loc_00592E3B:   mov [edx+0000000Ch], eax
loc_00592E3E:   Call [ecx+00000380h]
loc_00592E44:   lea edx, var_18
loc_00592E47:   push eax
loc_00592E48:   push edx
loc_00592E49:   Call edi
loc_00592E4B:   push eax
  loc_00592E4C: call [0040102Ch] ; __vbaLateIdCall
  loc_00592E52: add esp, 0000001Ch
loc_00592E55:   lea ecx, var_18
loc_00592E58:   Call ebx
  loc_00592E5A: mov var_4, 00000000h
  loc_00592E61: push 00592E73h
  loc_00592E66: jmp 00592E72h
loc_00592E68:   lea ecx, var_18
  loc_00592E6B: call [004012A0h] ; __vbaFreeObj
loc_00592E71:   ret
loc_00592E72:   ret
loc_00592E73:   mov eax, Me
loc_00592E76:   push eax
loc_00592E77:   mov ecx, [eax]
loc_00592E79:   Call [ecx+00000008h]
loc_00592E7C:   mov eax, var_4
loc_00592E7F:   mov ecx, var_14
loc_00592E82:   pop edi
loc_00592E83:   pop esi
loc_00592E84:   mov fs: [00000000h] , ecx
loc_00592E8B:   pop ebx
loc_00592E8C:   mov esp, ebp
loc_00592E8E:   pop ebp
  loc_00592E8F: retn 0004h
End Sub

Private Sub Form_Unload(Cancel As Integer) '59A020
loc_0059A020:   push ebp
loc_0059A021:   mov ebp, esp
  loc_0059A023: sub esp, 0000000Ch
  loc_0059A026: push 00405696h ; __vbaExceptHandler
loc_0059A02B:   mov eax, fs: [00000000h]
loc_0059A031:   push eax
loc_0059A032:   mov fs: [00000000h] , esp
  loc_0059A039: sub esp, 00000098h
loc_0059A03F:   push ebx
loc_0059A040:   push esi
loc_0059A041:   push edi
loc_0059A042:   mov var_C, esp
  loc_0059A045: mov var_8, 00402808h
loc_0059A04C:   mov ebx, Me
loc_0059A04F:   mov eax, ebx
  loc_0059A051: and eax, 00000001h
loc_0059A054:   mov var_4, eax
  loc_0059A057: and ebx, FFFFFFFEh
loc_0059A05A:   push ebx
loc_0059A05B:   mov Me, ebx
loc_0059A05E:   mov ecx, [ebx]
loc_0059A060:   Call [ecx+00000004h]
  loc_0059A063: mov edi, [00401238h] ; __vbaVarDup
  loc_0059A069: mov ecx, 80020004h
  loc_0059A06E: xor esi, esi
loc_0059A070:   mov var_50, ecx
  loc_0059A073: mov eax, 0000000Ah
loc_0059A078:   mov var_40, ecx
loc_0059A07B:   mov var_48, esi
loc_0059A07E:   mov var_58, esi
loc_0059A081:   mov var_78, esi
loc_0059A084:   lea edx, var_78
loc_0059A087:   lea ecx, var_38
loc_0059A08A:   mov var_18, esi
loc_0059A08D:   mov var_28, esi
loc_0059A090:   mov var_38, esi
loc_0059A093:   mov var_68, esi
loc_0059A096:   mov var_58, eax
loc_0059A099:   mov var_48, eax
  loc_0059A09C: mov var_70, 0042ED7Ch ; "Konfirmasi"
  loc_0059A0A3: mov var_78, 00000008h
loc_0059A0AA:   Call edi
loc_0059A0AC:   lea edx, var_68
loc_0059A0AF:   lea ecx, var_28
  loc_0059A0B2: mov var_60, 00434874h ; "APAKAH KAMU INGIN MENUTUP APLIKASI INI..?"
  loc_0059A0B9: mov var_68, 00000008h
loc_0059A0C0:   Call edi
loc_0059A0C2:   lea edx, var_58
loc_0059A0C5:   lea eax, var_48
loc_0059A0C8:   push edx
loc_0059A0C9:   lea ecx, var_38
loc_0059A0CC:   push eax
loc_0059A0CD:   push ecx
loc_0059A0CE:   lea edx, var_28
  loc_0059A0D1: push 00000024h
loc_0059A0D3:   push edx
  loc_0059A0D4: call [004010B4h] ; rtcMsgBox
  loc_0059A0DA: xor ecx, ecx
  loc_0059A0DC: cmp eax, 00000007h
loc_0059A0DF:   setz cl
loc_0059A0E2:   neg ecx
loc_0059A0E4:   lea edx, var_58
loc_0059A0E7:   mov di, cx
loc_0059A0EA:   lea eax, var_48
loc_0059A0ED:   push edx
loc_0059A0EE:   lea ecx, var_38
loc_0059A0F1:   push eax
loc_0059A0F2:   lea edx, var_28
loc_0059A0F5:   push ecx
loc_0059A0F6:   push edx
  loc_0059A0F7: push 00000004h
  loc_0059A0F9: call [0040103Ch] ; __vbaFreeVarList
  loc_0059A0FF: add esp, 00000014h
loc_0059A102:   cmp di, si
  loc_0059A105: jz 0059A114h
loc_0059A107:   mov eax, Cancel
  loc_0059A10A: mov [eax], 0001h
  loc_0059A10F: jmp 0059A1AFh
loc_0059A114:   cmp [00668228h], esi
  loc_0059A11A: jnz 0059A12Ch
  loc_0059A11C: push 00668228h
  loc_0059A121: push 004296BCh
  loc_0059A126: call [004011E8h] ; __vbaNew2
loc_0059A12C:   mov edi, [00668228h]
loc_0059A132:   push FFFFFFFFh
loc_0059A134:   push edi
loc_0059A135:   mov ecx, [edi]
loc_0059A137:   Call [ecx+00000094h]
loc_0059A13D:   cmp eax, esi
loc_0059A13F:   fnclex
  loc_0059A141: jge 0059A155h
  loc_0059A143: push 00000094h
  loc_0059A148: push 00433880h
loc_0059A14D:   push edi
loc_0059A14E:   push eax
  loc_0059A14F: call [00401080h] ; __vbaHresultCheckObj
loc_0059A155:   cmp [0066A318h], esi
  loc_0059A15B: jnz 0059A16Dh
  loc_0059A15D: push 0066A318h
  loc_0059A162: push 0042E390h
  loc_0059A167: call [004011E8h] ; __vbaNew2
loc_0059A16D:   mov edi, [0066A318h]
loc_0059A173:   lea eax, var_18
loc_0059A176:   push ebx
loc_0059A177:   push eax
loc_0059A178:   mov edx, [edi]
loc_0059A17A:   mov var_AC, edx
  loc_0059A180: call [004010C8h] ; __vbaObjSetAddref
loc_0059A186:   mov ecx, var_AC
loc_0059A18C:   push eax
loc_0059A18D:   push edi
loc_0059A18E:   Call [ecx+00000010h]
loc_0059A191:   cmp eax, esi
loc_0059A193:   fnclex
  loc_0059A195: jge 0059A1A6h
  loc_0059A197: push 00000010h
  loc_0059A199: push 0042E380h
loc_0059A19E:   push edi
loc_0059A19F:   push eax
  loc_0059A1A0: call [00401080h] ; __vbaHresultCheckObj
loc_0059A1A6:   lea ecx, var_18
  loc_0059A1A9: call [004012A0h] ; __vbaFreeObj
loc_0059A1AF:   mov var_4, esi
  loc_0059A1B2: push 0059A1DFh
  loc_0059A1B7: jmp 0059A1DEh
loc_0059A1B9:   lea ecx, var_18
  loc_0059A1BC: call [004012A0h] ; __vbaFreeObj
loc_0059A1C2:   lea edx, var_58
loc_0059A1C5:   lea eax, var_48
loc_0059A1C8:   push edx
loc_0059A1C9:   lea ecx, var_38
loc_0059A1CC:   push eax
loc_0059A1CD:   lea edx, var_28
loc_0059A1D0:   push ecx
loc_0059A1D1:   push edx
  loc_0059A1D2: push 00000004h
  loc_0059A1D4: call [0040103Ch] ; __vbaFreeVarList
  loc_0059A1DA: add esp, 00000014h
loc_0059A1DD:   ret
loc_0059A1DE:   ret
loc_0059A1DF:   mov eax, Me
loc_0059A1E2:   push eax
loc_0059A1E3:   mov ecx, [eax]
loc_0059A1E5:   Call [ecx+00000008h]
loc_0059A1E8:   mov eax, var_4
loc_0059A1EB:   mov ecx, var_14
loc_0059A1EE:   pop edi
loc_0059A1EF:   pop esi
loc_0059A1F0:   mov fs: [00000000h] , ecx
loc_0059A1F7:   pop ebx
loc_0059A1F8:   mov esp, ebp
loc_0059A1FA:   pop ebp
  loc_0059A1FB: retn 0008h
End Sub

Private Sub cmdCari_UnknownEvent_10() '595E50
loc_00595E50:   push ebp
loc_00595E51:   mov ebp, esp
  loc_00595E53: sub esp, 0000000Ch
  loc_00595E56: push 00405696h ; __vbaExceptHandler
loc_00595E5B:   mov eax, fs: [00000000h]
loc_00595E61:   push eax
loc_00595E62:   mov fs: [00000000h] , esp
  loc_00595E69: sub esp, 00000030h
loc_00595E6C:   push ebx
loc_00595E6D:   push esi
loc_00595E6E:   push edi
loc_00595E6F:   mov var_C, esp
  loc_00595E72: mov var_8, 004027E0h
loc_00595E79:   mov eax, Me
loc_00595E7C:   mov ecx, eax
  loc_00595E7E: and ecx, 00000001h
loc_00595E81:   mov var_4, ecx
  loc_00595E84: and al, FEh
loc_00595E86:   push eax
loc_00595E87:   mov Me, eax
loc_00595E8A:   mov edx, [eax]
loc_00595E8C:   Call [edx+00000004h]
loc_00595E8F:   mov eax, [00668228h]
loc_00595E94:   test eax, eax
  loc_00595E96: jnz 00595EACh
  loc_00595E98: mov ebx, [004011E8h] ; __vbaNew2
  loc_00595E9E: push 00668228h
  loc_00595EA3: push 004296BCh
loc_00595EA8:   Call ebx
  loc_00595EAA: jmp 00595EB2h
  loc_00595EAC: mov ebx, [004011E8h] ; __vbaNew2
loc_00595EB2:   mov esi, [00668228h]
  loc_00595EB8: push 00000000h
loc_00595EBA:   push esi
loc_00595EBB:   mov eax, [esi]
loc_00595EBD:   Call [eax+00000094h]
loc_00595EC3:   test eax, eax
loc_00595EC5:   fnclex
  loc_00595EC7: jge 00595EDFh
  loc_00595EC9: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_00595ECF: push 00000094h
  loc_00595ED4: push 00433880h
loc_00595ED9:   push esi
loc_00595EDA:   push eax
loc_00595EDB:   Call edi
  loc_00595EDD: jmp 00595EE5h
  loc_00595EDF: mov edi, [00401080h] ; __vbaHresultCheckObj
loc_00595EE5:   mov eax, [0066823Ch]
loc_00595EEA:   test eax, eax
  loc_00595EEC: jnz 00595EFAh
  loc_00595EEE: push 0066823Ch
  loc_00595EF3: push 00427B78h
loc_00595EF8:   Call ebx
loc_00595EFA:   mov esi, [0066823Ch]
  loc_00595F00: push 00000000h
loc_00595F02:   push esi
loc_00595F03:   mov ecx, [esi]
loc_00595F05:   Call [ecx+00000094h]
loc_00595F0B:   test eax, eax
loc_00595F0D:   fnclex
  loc_00595F0F: jge 00595F1Fh
  loc_00595F11: push 00000094h
  loc_00595F16: push 00433DA8h
loc_00595F1B:   push esi
loc_00595F1C:   push eax
loc_00595F1D:   Call edi
loc_00595F1F:   mov eax, [00668250h]
loc_00595F24:   test eax, eax
  loc_00595F26: jnz 00595F34h
  loc_00595F28: push 00668250h
  loc_00595F2D: push 0040A9ECh
loc_00595F32:   Call ebx
loc_00595F34:   mov esi, [00668250h]
  loc_00595F3A: push 00434210h ; "PEMBELIAN"
loc_00595F3F:   push esi
loc_00595F40:   mov edx, [esi]
loc_00595F42:   Call [edx+000006FCh]
loc_00595F48:   test eax, eax
loc_00595F4A:   fnclex
  loc_00595F4C: jge 00595F5Ch
  loc_00595F4E: push 000006FCh
  loc_00595F53: push 00434254h
loc_00595F58:   push esi
loc_00595F59:   push eax
loc_00595F5A:   Call edi
loc_00595F5C:   mov eax, [00668250h]
loc_00595F61:   test eax, eax
  loc_00595F63: jnz 00595F71h
  loc_00595F65: push 00668250h
  loc_00595F6A: push 0040A9ECh
loc_00595F6F:   Call ebx
  loc_00595F71: sub esp, 00000010h
  loc_00595F74: mov ecx, 0000000Ah
loc_00595F79:   mov ebx, esp
  loc_00595F7B: mov eax, 80020004h
  loc_00595F80: sub esp, 00000010h
loc_00595F83:   mov esi, [00668250h]
loc_00595F89:   mov [ebx], ecx
loc_00595F8B:   mov ecx, var_30
loc_00595F8E:   mov edi, [esi]
  loc_00595F90: mov edx, 00000001h
loc_00595F95:   mov [ebx+00000004h], ecx
loc_00595F98:   mov ecx, esp
loc_00595F9A:   push esi
loc_00595F9B:   mov [ebx+00000008h], eax
loc_00595F9E:   mov eax, var_28
loc_00595FA1:   mov [ebx+0000000Ch], eax
  loc_00595FA4: mov eax, 00000002h
loc_00595FA9:   mov [ecx], eax
loc_00595FAB:   mov eax, var_20
loc_00595FAE:   mov [ecx+00000004h], eax
loc_00595FB1:   mov [ecx+00000008h], edx
loc_00595FB4:   mov edx, var_18
loc_00595FB7:   mov [ecx+0000000Ch], edx
loc_00595FBA:   Call [edi+000002B0h]
loc_00595FC0:   test eax, eax
loc_00595FC2:   fnclex
  loc_00595FC4: jge 00595FD8h
  loc_00595FC6: push 000002B0h
  loc_00595FCB: push 00434224h
loc_00595FD0:   push esi
loc_00595FD1:   push eax
  loc_00595FD2: call [00401080h] ; __vbaHresultCheckObj
  loc_00595FD8: mov var_4, 00000000h
loc_00595FDF:   mov eax, Me
loc_00595FE2:   push eax
loc_00595FE3:   mov ecx, [eax]
loc_00595FE5:   Call [ecx+00000008h]
loc_00595FE8:   mov eax, var_4
loc_00595FEB:   mov ecx, var_14
loc_00595FEE:   pop edi
loc_00595FEF:   pop esi
loc_00595FF0:   mov fs: [00000000h] , ecx
loc_00595FF7:   pop ebx
loc_00595FF8:   mov esp, ebp
loc_00595FFA:   pop ebp
  loc_00595FFB: retn 0004h
End Sub

Private Sub CmdMasuk_UnknownEvent_10() '596000
loc_00596000:   push ebp
loc_00596001:   mov ebp, esp
  loc_00596003: sub esp, 0000000Ch
  loc_00596006: push 00405696h ; __vbaExceptHandler
loc_0059600B:   mov eax, fs: [00000000h]
loc_00596011:   push eax
loc_00596012:   mov fs: [00000000h] , esp
  loc_00596019: sub esp, 00000108h
loc_0059601F:   push ebx
loc_00596020:   push esi
loc_00596021:   push edi
loc_00596022:   mov var_C, esp
  loc_00596025: mov var_8, 004027E8h
loc_0059602C:   mov esi, Me
loc_0059602F:   mov eax, esi
  loc_00596031: and eax, 00000001h
loc_00596034:   mov var_4, eax
  loc_00596037: and esi, FFFFFFFEh
loc_0059603A:   push esi
loc_0059603B:   mov Me, esi
loc_0059603E:   mov ecx, [esi]
loc_00596040:   Call [ecx+00000004h]
loc_00596043:   mov edx, [esi]
  loc_00596045: xor ebx, ebx
loc_00596047:   push esi
loc_00596048:   mov var_24, ebx
loc_0059604B:   mov var_28, ebx
loc_0059604E:   mov var_38, ebx
loc_00596051:   mov var_3C, ebx
loc_00596054:   mov var_40, ebx
loc_00596057:   mov var_44, ebx
loc_0059605A:   mov var_48, ebx
loc_0059605D:   mov var_58, ebx
loc_00596060:   mov var_68, ebx
loc_00596063:   mov var_78, ebx
loc_00596066:   mov var_88, ebx
loc_0059606C:   mov var_98, ebx
loc_00596072:   mov var_A8, ebx
loc_00596078:   mov var_B8, ebx
loc_0059607E:   Call [edx+00000354h]
  loc_00596084: mov edi, [004010B8h] ; __vbaObjSet
loc_0059608A:   push eax
loc_0059608B:   lea eax, var_44
loc_0059608E:   push eax
loc_0059608F:   Call edi
loc_00596091:   mov ecx, [eax]
loc_00596093:   lea edx, var_3C
loc_00596096:   push edx
loc_00596097:   push eax
loc_00596098:   mov var_E4, eax
loc_0059609E:   Call [ecx+000000A0h]
loc_005960A4:   cmp eax, ebx
loc_005960A6:   fnclex
  loc_005960A8: jge 005960C2h
loc_005960AA:   mov ecx, var_E4
  loc_005960B0: push 000000A0h
  loc_005960B5: push 0042DFCCh
loc_005960BA:   push ecx
loc_005960BB:   push eax
  loc_005960BC: call [00401080h] ; __vbaHresultCheckObj
loc_005960C2:   mov edx, var_3C
loc_005960C5:   push edx
  loc_005960C6: push 0042DFECh
  loc_005960CB: call [0040112Ch] ; __vbaStrCmp
loc_005960D1:   mov ebx, eax
loc_005960D3:   lea ecx, var_3C
loc_005960D6:   neg ebx
loc_005960D8:   sbb ebx, ebx
loc_005960DA:   inc ebx
loc_005960DB:   neg ebx
  loc_005960DD: call [0040129Ch] ; __vbaFreeStr
loc_005960E3:   lea ecx, var_44
  loc_005960E6: call [004012A0h] ; __vbaFreeObj
loc_005960EC:   test bx, bx
  loc_005960EF: jz 005961C7h
  loc_005960F5: mov ebx, [00401238h] ; __vbaVarDup
  loc_005960FB: mov ecx, 80020004h
loc_00596100:   mov var_80, ecx
  loc_00596103: mov eax, 0000000Ah
loc_00596108:   mov var_70, ecx
loc_0059610B:   lea edx, var_A8
loc_00596111:   lea ecx, var_68
loc_00596114:   mov var_88, eax
loc_0059611A:   mov var_78, eax
  loc_0059611D: mov var_A0, 0042E624h ; "Error"
  loc_00596127: mov var_A8, 00000008h
loc_00596131:   Call ebx
loc_00596133:   lea edx, var_98
loc_00596139:   lea ecx, var_58
  loc_0059613C: mov var_90, 00434288h ; "KODE BARANG BELUM LENGKAP"
  loc_00596146: mov var_98, 00000008h
loc_00596150:   Call ebx
loc_00596152:   lea eax, var_88
loc_00596158:   lea ecx, var_78
loc_0059615B:   push eax
loc_0059615C:   lea edx, var_68
loc_0059615F:   push ecx
loc_00596160:   push edx
loc_00596161:   lea eax, var_58
  loc_00596164: push 00000010h
loc_00596166:   push eax
  loc_00596167: call [004010B4h] ; rtcMsgBox
loc_0059616D:   lea ecx, var_88
loc_00596173:   lea edx, var_78
loc_00596176:   push ecx
loc_00596177:   lea eax, var_68
loc_0059617A:   push edx
loc_0059617B:   lea ecx, var_58
loc_0059617E:   push eax
loc_0059617F:   push ecx
  loc_00596180: push 00000004h
  loc_00596182: call [0040103Ch] ; __vbaFreeVarList
loc_00596188:   mov edx, [esi]
  loc_0059618A: add esp, 00000014h
loc_0059618D:   push esi
loc_0059618E:   Call [edx+00000354h]
loc_00596194:   push eax
loc_00596195:   lea eax, var_44
loc_00596198:   push eax
loc_00596199:   Call edi
loc_0059619B:   mov esi, eax
loc_0059619D:   push esi
loc_0059619E:   mov ecx, [esi]
loc_005961A0:   Call [ecx+00000204h]
loc_005961A6:   test eax, eax
loc_005961A8:   fnclex
  loc_005961AA: jge 00596EF7h
  loc_005961B0: push 00000204h
  loc_005961B5: push 0042DFCCh
loc_005961BA:   push esi
loc_005961BB:   push eax
  loc_005961BC: call [00401080h] ; __vbaHresultCheckObj
  loc_005961C2: jmp 00596EF7h
loc_005961C7:   mov edx, [esi]
loc_005961C9:   push esi
loc_005961CA:   Call [edx+00000344h]
loc_005961D0:   push eax
loc_005961D1:   lea eax, var_44
loc_005961D4:   push eax
loc_005961D5:   Call edi
loc_005961D7:   mov ebx, eax
loc_005961D9:   lea edx, var_3C
loc_005961DC:   push edx
loc_005961DD:   push ebx
loc_005961DE:   mov ecx, [ebx]
loc_005961E0:   Call [ecx+000000A0h]
loc_005961E6:   test eax, eax
loc_005961E8:   fnclex
  loc_005961EA: jge 005961FEh
  loc_005961EC: push 000000A0h
  loc_005961F1: push 0042DFCCh
loc_005961F6:   push ebx
loc_005961F7:   push eax
  loc_005961F8: call [00401080h] ; __vbaHresultCheckObj
loc_005961FE:   mov eax, [esi]
loc_00596200:   push esi
loc_00596201:   Call [eax+00000344h]
loc_00596207:   lea ecx, var_48
loc_0059620A:   push eax
loc_0059620B:   push ecx
loc_0059620C:   Call edi
loc_0059620E:   mov ebx, eax
loc_00596210:   lea eax, var_40
loc_00596213:   push eax
loc_00596214:   push ebx
loc_00596215:   mov edx, [ebx]
loc_00596217:   Call [edx+000000A0h]
loc_0059621D:   test eax, eax
loc_0059621F:   fnclex
  loc_00596221: jge 00596235h
  loc_00596223: push 000000A0h
  loc_00596228: push 0042DFCCh
loc_0059622D:   push ebx
loc_0059622E:   push eax
  loc_0059622F: call [00401080h] ; __vbaHresultCheckObj
loc_00596235:   mov ecx, var_40
loc_00596238:   push ecx
  loc_00596239: push 0042E3A4h
  loc_0059623E: call [0040112Ch] ; __vbaStrCmp
loc_00596244:   mov edx, var_3C
loc_00596247:   mov ebx, eax
loc_00596249:   neg ebx
loc_0059624B:   sbb ebx, ebx
loc_0059624D:   push edx
loc_0059624E:   inc ebx
  loc_0059624F: push 0042DFECh
loc_00596254:   neg ebx
  loc_00596256: call [0040112Ch] ; __vbaStrCmp
loc_0059625C:   neg eax
loc_0059625E:   sbb eax, eax
loc_00596260:   lea ecx, var_3C
loc_00596263:   inc eax
loc_00596264:   neg eax
  loc_00596266: or ebx, eax
loc_00596268:   lea eax, var_40
loc_0059626B:   push eax
loc_0059626C:   push ecx
  loc_0059626D: push 00000002h
  loc_0059626F: call [00401204h] ; __vbaFreeStrList
loc_00596275:   lea edx, var_48
loc_00596278:   lea eax, var_44
loc_0059627B:   push edx
loc_0059627C:   push eax
  loc_0059627D: push 00000002h
  loc_0059627F: call [0040104Ch] ; __vbaFreeObjList
  loc_00596285: add esp, 00000018h
loc_00596288:   test bx, bx
  loc_0059628B: jz 00596363h
  loc_00596291: mov ebx, [00401238h] ; __vbaVarDup
  loc_00596297: mov ecx, 80020004h
loc_0059629C:   mov var_80, ecx
  loc_0059629F: mov eax, 0000000Ah
loc_005962A4:   mov var_70, ecx
loc_005962A7:   lea edx, var_A8
loc_005962AD:   lea ecx, var_68
loc_005962B0:   mov var_88, eax
loc_005962B6:   mov var_78, eax
  loc_005962B9: mov var_A0, 0042E624h ; "Error"
  loc_005962C3: mov var_A8, 00000008h
loc_005962CD:   Call ebx
loc_005962CF:   lea edx, var_98
loc_005962D5:   lea ecx, var_58
  loc_005962D8: mov var_90, 004342C0h ; "JUMLAH BARANG MASIH KOSONG"
  loc_005962E2: mov var_98, 00000008h
loc_005962EC:   Call ebx
loc_005962EE:   lea ecx, var_88
loc_005962F4:   lea edx, var_78
loc_005962F7:   push ecx
loc_005962F8:   lea eax, var_68
loc_005962FB:   push edx
loc_005962FC:   push eax
loc_005962FD:   lea ecx, var_58
  loc_00596300: push 00000010h
loc_00596302:   push ecx
  loc_00596303: call [004010B4h] ; rtcMsgBox
loc_00596309:   lea edx, var_88
loc_0059630F:   lea eax, var_78
loc_00596312:   push edx
loc_00596313:   lea ecx, var_68
loc_00596316:   push eax
loc_00596317:   lea edx, var_58
loc_0059631A:   push ecx
loc_0059631B:   push edx
  loc_0059631C: push 00000004h
  loc_0059631E: call [0040103Ch] ; __vbaFreeVarList
loc_00596324:   mov eax, [esi]
  loc_00596326: add esp, 00000014h
loc_00596329:   push esi
loc_0059632A:   Call [eax+00000344h]
loc_00596330:   lea ecx, var_44
loc_00596333:   push eax
loc_00596334:   push ecx
loc_00596335:   Call edi
loc_00596337:   mov esi, eax
loc_00596339:   push esi
loc_0059633A:   mov edx, [esi]
loc_0059633C:   Call [edx+00000204h]
loc_00596342:   test eax, eax
loc_00596344:   fnclex
  loc_00596346: jge 00596EF7h
  loc_0059634C: push 00000204h
  loc_00596351: push 0042DFCCh
loc_00596356:   push esi
loc_00596357:   push eax
  loc_00596358: call [00401080h] ; __vbaHresultCheckObj
  loc_0059635E: jmp 00596EF7h
loc_00596363:   mov eax, [esi]
loc_00596365:   push esi
loc_00596366:   Call [eax+00000350h]
loc_0059636C:   lea ecx, var_48
loc_0059636F:   push eax
loc_00596370:   push ecx
loc_00596371:   Call edi
loc_00596373:   mov ebx, eax
loc_00596375:   lea eax, var_40
loc_00596378:   push eax
loc_00596379:   push ebx
loc_0059637A:   mov edx, [ebx]
loc_0059637C:   Call [edx+000000A0h]
loc_00596382:   test eax, eax
loc_00596384:   fnclex
  loc_00596386: jge 0059639Ah
  loc_00596388: push 000000A0h
  loc_0059638D: push 0042DFCCh
loc_00596392:   push ebx
loc_00596393:   push eax
  loc_00596394: call [00401080h] ; __vbaHresultCheckObj
loc_0059639A:   mov ecx, var_40
loc_0059639D:   push ecx
  loc_0059639E: call [004012A4h] ; rtcR8ValFromBstr
loc_005963A4:   mov edx, [esi]
loc_005963A6:   push esi
  loc_005963A7: fstp real8 ptr var_E0
loc_005963AD:   Call [edx+00000348h]
loc_005963B3:   push eax
loc_005963B4:   lea eax, var_44
loc_005963B7:   push eax
loc_005963B8:   Call edi
loc_005963BA:   mov ebx, eax
loc_005963BC:   lea edx, var_3C
loc_005963BF:   push edx
loc_005963C0:   push ebx
loc_005963C1:   mov ecx, [ebx]
loc_005963C3:   Call [ecx+000000A0h]
loc_005963C9:   test eax, eax
loc_005963CB:   fnclex
  loc_005963CD: jge 005963E1h
  loc_005963CF: push 000000A0h
  loc_005963D4: push 0042DFCCh
loc_005963D9:   push ebx
loc_005963DA:   push eax
  loc_005963DB: call [00401080h] ; __vbaHresultCheckObj
loc_005963E1:   mov eax, var_3C
loc_005963E4:   push eax
  loc_005963E5: call [004012A4h] ; rtcR8ValFromBstr
  loc_005963EB: call [004010ECh] ; __vbaFpR8
  loc_005963F1: fstp real8 ptr var_11C
  loc_005963F7: fld real8 ptr var_E0
  loc_005963FD: call [004010ECh] ; __vbaFpR8
  loc_00596403: fcomp real8 ptr var_11C
loc_00596409:   fnstsw ax
  loc_0059640B: test ah, 41h
  loc_0059640E: jz 00596417h
  loc_00596410: mov eax, 00000001h
  loc_00596415: jmp 00596419h
  loc_00596417: xor eax, eax
loc_00596419:   lea ecx, var_40
loc_0059641C:   lea edx, var_3C
loc_0059641F:   push ecx
loc_00596420:   push edx
loc_00596421:   neg eax
  loc_00596423: push 00000002h
loc_00596425:   mov ebx, eax
  loc_00596427: call [00401204h] ; __vbaFreeStrList
loc_0059642D:   lea eax, var_48
loc_00596430:   lea ecx, var_44
loc_00596433:   push eax
loc_00596434:   push ecx
  loc_00596435: push 00000002h
  loc_00596437: call [0040104Ch] ; __vbaFreeObjList
  loc_0059643D: add esp, 00000018h
loc_00596440:   test bx, bx
  loc_00596443: jz 0059651Bh
  loc_00596449: mov ebx, [00401238h] ; __vbaVarDup
  loc_0059644F: mov ecx, 80020004h
loc_00596454:   mov var_80, ecx
  loc_00596457: mov eax, 0000000Ah
loc_0059645C:   mov var_70, ecx
loc_0059645F:   lea edx, var_A8
loc_00596465:   lea ecx, var_68
loc_00596468:   mov var_88, eax
loc_0059646E:   mov var_78, eax
  loc_00596471: mov var_A0, 0042E624h ; "Error"
  loc_0059647B: mov var_A8, 00000008h
loc_00596485:   Call ebx
loc_00596487:   lea edx, var_98
loc_0059648D:   lea ecx, var_58
  loc_00596490: mov var_90, 004342FCh ; "HARGA JUAL BELUM DITENTUKAN"
  loc_0059649A: mov var_98, 00000008h
loc_005964A4:   Call ebx
loc_005964A6:   lea edx, var_88
loc_005964AC:   lea eax, var_78
loc_005964AF:   push edx
loc_005964B0:   lea ecx, var_68
loc_005964B3:   push eax
loc_005964B4:   push ecx
loc_005964B5:   lea edx, var_58
  loc_005964B8: push 00000010h
loc_005964BA:   push edx
  loc_005964BB: call [004010B4h] ; rtcMsgBox
loc_005964C1:   lea eax, var_88
loc_005964C7:   lea ecx, var_78
loc_005964CA:   push eax
loc_005964CB:   lea edx, var_68
loc_005964CE:   push ecx
loc_005964CF:   lea eax, var_58
loc_005964D2:   push edx
loc_005964D3:   push eax
  loc_005964D4: push 00000004h
  loc_005964D6: call [0040103Ch] ; __vbaFreeVarList
loc_005964DC:   mov ecx, [esi]
  loc_005964DE: add esp, 00000014h
loc_005964E1:   push esi
loc_005964E2:   Call [ecx+00000350h]
loc_005964E8:   lea edx, var_44
loc_005964EB:   push eax
loc_005964EC:   push edx
loc_005964ED:   Call edi
loc_005964EF:   mov esi, eax
loc_005964F1:   push esi
loc_005964F2:   mov eax, [esi]
loc_005964F4:   Call [eax+00000204h]
loc_005964FA:   test eax, eax
loc_005964FC:   fnclex
  loc_005964FE: jge 00596EF7h
  loc_00596504: push 00000204h
  loc_00596509: push 0042DFCCh
loc_0059650E:   push esi
loc_0059650F:   push eax
  loc_00596510: call [00401080h] ; __vbaHresultCheckObj
  loc_00596516: jmp 00596EF7h
loc_0059651B:   mov ecx, [esi]
loc_0059651D:   push esi
loc_0059651E:   Call [ecx+00000348h]
loc_00596524:   lea edx, var_44
loc_00596527:   push eax
loc_00596528:   push edx
loc_00596529:   Call edi
loc_0059652B:   mov ebx, eax
loc_0059652D:   lea ecx, var_3C
loc_00596530:   push ecx
loc_00596531:   push ebx
loc_00596532:   mov eax, [ebx]
loc_00596534:   Call [eax+000000A0h]
loc_0059653A:   test eax, eax
loc_0059653C:   fnclex
  loc_0059653E: jge 00596552h
  loc_00596540: push 000000A0h
  loc_00596545: push 0042DFCCh
loc_0059654A:   push ebx
loc_0059654B:   push eax
  loc_0059654C: call [00401080h] ; __vbaHresultCheckObj
  loc_00596552: mov ebx, 00000008h
loc_00596557:   lea edx, var_98
loc_0059655D:   lea ecx, var_68
  loc_00596560: mov var_90, 004340F0h ; "#,#"
loc_0059656A:   mov var_98, ebx
  loc_00596570: call [00401238h] ; __vbaVarDup
loc_00596576:   mov eax, var_3C
  loc_00596579: push 00000001h
loc_0059657B:   mov var_50, eax
loc_0059657E:   lea edx, var_68
  loc_00596581: push 00000001h
loc_00596583:   lea eax, var_58
loc_00596586:   push edx
loc_00596587:   lea ecx, var_78
loc_0059658A:   push eax
loc_0059658B:   push ecx
  loc_0059658C: mov var_3C, 00000000h
loc_00596593:   mov var_58, ebx
  loc_00596596: call [00401078h] ; rtcVarFromFormatVar
loc_0059659C:   lea edx, var_78
loc_0059659F:   lea ecx, var_24
  loc_005965A2: call [0040101Ch] ; __vbaVarMove
  loc_005965A8: mov ebx, [004012A0h] ; __vbaFreeObj
loc_005965AE:   lea ecx, var_44
loc_005965B1:   Call ebx
loc_005965B3:   lea edx, var_68
loc_005965B6:   lea eax, var_58
loc_005965B9:   push edx
loc_005965BA:   push eax
  loc_005965BB: push 00000002h
  loc_005965BD: call [0040103Ch] ; __vbaFreeVarList
loc_005965C3:   mov ecx, [esi]
  loc_005965C5: add esp, 0000000Ch
loc_005965C8:   push esi
loc_005965C9:   Call [ecx+00000350h]
loc_005965CF:   lea edx, var_44
loc_005965D2:   push eax
loc_005965D3:   push edx
loc_005965D4:   Call edi
loc_005965D6:   mov ecx, [eax]
loc_005965D8:   lea edx, var_3C
loc_005965DB:   push edx
loc_005965DC:   push eax
loc_005965DD:   mov var_E4, eax
loc_005965E3:   Call [ecx+000000A0h]
loc_005965E9:   test eax, eax
loc_005965EB:   fnclex
  loc_005965ED: jge 00596607h
loc_005965EF:   mov ecx, var_E4
  loc_005965F5: push 000000A0h
  loc_005965FA: push 0042DFCCh
loc_005965FF:   push ecx
loc_00596600:   push eax
  loc_00596601: call [00401080h] ; __vbaHresultCheckObj
loc_00596607:   lea edx, var_98
loc_0059660D:   lea ecx, var_68
  loc_00596610: mov var_90, 004340F0h ; "#,#"
  loc_0059661A: mov var_98, 00000008h
  loc_00596624: call [00401238h] ; __vbaVarDup
loc_0059662A:   mov eax, var_3C
  loc_0059662D: push 00000001h
loc_0059662F:   mov var_50, eax
loc_00596632:   lea edx, var_68
  loc_00596635: push 00000001h
loc_00596637:   lea eax, var_58
loc_0059663A:   push edx
loc_0059663B:   lea ecx, var_78
loc_0059663E:   push eax
loc_0059663F:   push ecx
  loc_00596640: mov var_3C, 00000000h
  loc_00596647: mov var_58, 00000008h
  loc_0059664E: call [00401078h] ; rtcVarFromFormatVar
loc_00596654:   lea edx, var_78
loc_00596657:   lea ecx, var_38
  loc_0059665A: call [0040101Ch] ; __vbaVarMove
loc_00596660:   lea ecx, var_44
loc_00596663:   Call ebx
loc_00596665:   lea edx, var_68
loc_00596668:   lea eax, var_58
loc_0059666B:   push edx
loc_0059666C:   push eax
  loc_0059666D: push 00000002h
  loc_0059666F: call [0040103Ch] ; __vbaFreeVarList
loc_00596675:   mov cx, [esi+00000034h]
  loc_00596679: add cx, 0001h
  loc_0059667D: jo 00596F87h
loc_00596683:   movsx eax, cx
  loc_00596686: mov ecx, 00000003h
loc_0059668B:   mov var_90, eax
loc_00596691:   push ecx
loc_00596692:   mov var_98, ecx
loc_00596698:   mov edx, esp
  loc_0059669A: push 00000004h
loc_0059669C:   push esi
loc_0059669D:   mov [edx], ecx
loc_0059669F:   mov ecx, var_94
loc_005966A5:   mov [edx+00000004h], ecx
loc_005966A8:   mov ecx, [esi]
loc_005966AA:   mov [edx+00000008h], eax
loc_005966AD:   mov eax, var_8C
loc_005966B3:   mov [edx+0000000Ch], eax
loc_005966B6:   Call [ecx+00000390h]
loc_005966BC:   lea edx, var_44
loc_005966BF:   push eax
loc_005966C0:   push edx
loc_005966C1:   Call edi
loc_005966C3:   push eax
  loc_005966C4: call [00401280h] ; __vbaLateIdSt
loc_005966CA:   lea ecx, var_44
loc_005966CD:   Call ebx
loc_005966CF:   mov ax, [esi+00000034h]
  loc_005966D3: mov ebx, 00000003h
loc_005966D8:   movsx ecx, ax
loc_005966DB:   push eax
loc_005966DC:   mov var_90, ecx
loc_005966E2:   mov var_98, ebx
  loc_005966E8: call [0040100Ch] ; __vbaStrI2
loc_005966EE:   mov ecx, var_98
  loc_005966F4: sub esp, 00000010h
loc_005966F7:   mov edx, esp
  loc_005966F9: sub esp, 00000010h
loc_005966FC:   mov var_50, eax
  loc_005966FF: mov var_58, 00000008h
loc_00596706:   mov [edx], ecx
loc_00596708:   mov ecx, var_94
loc_0059670E:   mov [edx+00000004h], ecx
loc_00596711:   mov ecx, var_90
loc_00596717:   mov [edx+00000008h], ecx
loc_0059671A:   mov ecx, var_8C
loc_00596720:   mov [edx+0000000Ch], ecx
loc_00596723:   mov edx, esp
  loc_00596725: xor ecx, ecx
loc_00596727:   mov [edx], ebx
loc_00596729:   mov ebx, var_B4
loc_0059672F:   mov [edx+00000004h], ebx
loc_00596732:   mov [edx+00000008h], ecx
loc_00596735:   mov ecx, var_AC
  loc_0059673B: sub esp, 00000010h
loc_0059673E:   mov [edx+0000000Ch], ecx
loc_00596741:   mov ecx, var_58
loc_00596744:   mov edx, esp
  loc_00596746: push 00000002h
  loc_00596748: push 00000041h
loc_0059674A:   push esi
loc_0059674B:   mov [edx], ecx
loc_0059674D:   mov ecx, var_54
loc_00596750:   mov [edx+00000004h], ecx
loc_00596753:   mov ecx, [esi]
loc_00596755:   mov [edx+00000008h], eax
loc_00596758:   mov eax, var_4C
loc_0059675B:   mov [edx+0000000Ch], eax
loc_0059675E:   Call [ecx+00000390h]
loc_00596764:   lea edx, var_44
loc_00596767:   push eax
loc_00596768:   push edx
loc_00596769:   Call edi
loc_0059676B:   push eax
  loc_0059676C: call [0040117Ch] ; __vbaLateIdCallSt
  loc_00596772: add esp, 0000003Ch
loc_00596775:   lea ecx, var_44
  loc_00596778: call [004012A0h] ; __vbaFreeObj
loc_0059677E:   lea ecx, var_58
  loc_00596781: call [00401020h] ; __vbaFreeVar
loc_00596787:   movsx eax, [esi+00000034h]
loc_0059678B:   mov ecx, [esi]
loc_0059678D:   push esi
loc_0059678E:   mov var_90, eax
  loc_00596794: mov var_98, 00000003h
loc_0059679E:   Call [ecx+00000354h]
loc_005967A4:   lea edx, var_44
loc_005967A7:   push eax
loc_005967A8:   push edx
loc_005967A9:   Call edi
loc_005967AB:   mov ecx, [eax]
loc_005967AD:   lea edx, var_3C
loc_005967B0:   push edx
loc_005967B1:   push eax
loc_005967B2:   mov var_E4, eax
loc_005967B8:   Call [ecx+000000A0h]
loc_005967BE:   test eax, eax
loc_005967C0:   fnclex
  loc_005967C2: jge 005967DCh
loc_005967C4:   mov ecx, var_E4
  loc_005967CA: push 000000A0h
  loc_005967CF: push 0042DFCCh
loc_005967D4:   push ecx
loc_005967D5:   push eax
  loc_005967D6: call [00401080h] ; __vbaHresultCheckObj
loc_005967DC:   mov ecx, var_98
  loc_005967E2: sub esp, 00000010h
loc_005967E5:   mov edx, esp
  loc_005967E7: sub esp, 00000010h
loc_005967EA:   mov eax, var_3C
  loc_005967ED: mov var_58, 00000008h
loc_005967F4:   mov [edx], ecx
loc_005967F6:   mov ecx, var_94
loc_005967FC:   mov var_50, eax
  loc_005967FF: mov var_3C, 00000000h
loc_00596806:   mov [edx+00000004h], ecx
loc_00596809:   mov ecx, var_90
loc_0059680F:   mov [edx+00000008h], ecx
loc_00596812:   mov ecx, var_8C
loc_00596818:   mov [edx+0000000Ch], ecx
loc_0059681B:   mov edx, esp
  loc_0059681D: mov ecx, 00000003h
  loc_00596822: sub esp, 00000010h
loc_00596825:   mov [edx], ecx
  loc_00596827: mov ecx, 00000001h
loc_0059682C:   mov [edx+00000004h], ebx
loc_0059682F:   mov [edx+00000008h], ecx
loc_00596832:   mov ecx, var_AC
loc_00596838:   mov [edx+0000000Ch], ecx
loc_0059683B:   mov ecx, var_58
loc_0059683E:   mov edx, esp
  loc_00596840: push 00000002h
  loc_00596842: push 00000041h
loc_00596844:   push esi
loc_00596845:   mov [edx], ecx
loc_00596847:   mov ecx, var_54
loc_0059684A:   mov [edx+00000004h], ecx
loc_0059684D:   mov ecx, [esi]
loc_0059684F:   mov [edx+00000008h], eax
loc_00596852:   mov eax, var_4C
loc_00596855:   mov [edx+0000000Ch], eax
loc_00596858:   Call [ecx+00000390h]
loc_0059685E:   lea edx, var_48
loc_00596861:   push eax
loc_00596862:   push edx
loc_00596863:   Call edi
loc_00596865:   push eax
  loc_00596866: call [0040117Ch] ; __vbaLateIdCallSt
loc_0059686C:   lea eax, var_48
loc_0059686F:   lea ecx, var_44
loc_00596872:   push eax
loc_00596873:   push ecx
  loc_00596874: push 00000002h
  loc_00596876: call [0040104Ch] ; __vbaFreeObjList
  loc_0059687C: add esp, 00000048h
loc_0059687F:   lea ecx, var_58
  loc_00596882: call [00401020h] ; __vbaFreeVar
loc_00596888:   movsx edx, [esi+00000034h]
loc_0059688C:   mov eax, [esi]
loc_0059688E:   push esi
loc_0059688F:   mov var_90, edx
  loc_00596895: mov var_98, 00000003h
loc_0059689F:   Call [eax+0000034Ch]
loc_005968A5:   lea ecx, var_44
loc_005968A8:   push eax
loc_005968A9:   push ecx
loc_005968AA:   Call edi
loc_005968AC:   mov edx, [eax]
loc_005968AE:   lea ecx, var_3C
loc_005968B1:   push ecx
loc_005968B2:   push eax
loc_005968B3:   mov var_E4, eax
loc_005968B9:   Call [edx+000000A0h]
loc_005968BF:   test eax, eax
loc_005968C1:   fnclex
  loc_005968C3: jge 005968DDh
loc_005968C5:   mov edx, var_E4
  loc_005968CB: push 000000A0h
  loc_005968D0: push 0042DFCCh
loc_005968D5:   push edx
loc_005968D6:   push eax
  loc_005968D7: call [00401080h] ; __vbaHresultCheckObj
loc_005968DD:   mov edx, var_98
  loc_005968E3: sub esp, 00000010h
loc_005968E6:   mov ecx, esp
  loc_005968E8: sub esp, 00000010h
loc_005968EB:   mov eax, var_3C
  loc_005968EE: mov var_58, 00000008h
loc_005968F5:   mov [ecx], edx
loc_005968F7:   mov edx, var_94
loc_005968FD:   mov var_50, eax
  loc_00596900: mov var_3C, 00000000h
loc_00596907:   mov [ecx+00000004h], edx
loc_0059690A:   mov edx, var_90
loc_00596910:   mov [ecx+00000008h], edx
loc_00596913:   mov edx, var_8C
loc_00596919:   mov [ecx+0000000Ch], edx
loc_0059691C:   mov edx, esp
  loc_0059691E: mov ecx, 00000003h
  loc_00596923: sub esp, 00000010h
loc_00596926:   mov [edx], ecx
  loc_00596928: mov ecx, 00000002h
loc_0059692D:   mov [edx+00000004h], ebx
loc_00596930:   mov [edx+00000008h], ecx
loc_00596933:   mov ecx, var_AC
loc_00596939:   mov [edx+0000000Ch], ecx
loc_0059693C:   mov ecx, var_58
loc_0059693F:   mov edx, esp
  loc_00596941: push 00000002h
  loc_00596943: push 00000041h
loc_00596945:   push esi
loc_00596946:   mov [edx], ecx
loc_00596948:   mov ecx, var_54
loc_0059694B:   mov [edx+00000004h], ecx
loc_0059694E:   mov ecx, [esi]
loc_00596950:   mov [edx+00000008h], eax
loc_00596953:   mov eax, var_4C
loc_00596956:   mov [edx+0000000Ch], eax
loc_00596959:   Call [ecx+00000390h]
loc_0059695F:   lea edx, var_48
loc_00596962:   push eax
loc_00596963:   push edx
loc_00596964:   Call edi
loc_00596966:   push eax
  loc_00596967: call [0040117Ch] ; __vbaLateIdCallSt
loc_0059696D:   lea eax, var_48
loc_00596970:   lea ecx, var_44
loc_00596973:   push eax
loc_00596974:   push ecx
  loc_00596975: push 00000002h
  loc_00596977: call [0040104Ch] ; __vbaFreeObjList
  loc_0059697D: add esp, 00000048h
loc_00596980:   lea ecx, var_58
  loc_00596983: call [00401020h] ; __vbaFreeVar
loc_00596989:   movsx edx, [esi+00000034h]
loc_0059698D:   lea eax, var_24
loc_00596990:   mov var_90, edx
loc_00596996:   push eax
  loc_00596997: mov var_98, 00000003h
  loc_005969A1: call [00401270h] ; __vbaStrVarCopy
loc_005969A7:   mov edx, var_98
  loc_005969AD: sub esp, 00000010h
loc_005969B0:   mov ecx, esp
  loc_005969B2: sub esp, 00000010h
  loc_005969B5: mov var_58, 00000008h
loc_005969BC:   mov var_50, eax
loc_005969BF:   mov [ecx], edx
loc_005969C1:   mov edx, var_94
loc_005969C7:   mov [ecx+00000004h], edx
loc_005969CA:   mov edx, var_90
loc_005969D0:   mov [ecx+00000008h], edx
loc_005969D3:   mov edx, var_8C
loc_005969D9:   mov [ecx+0000000Ch], edx
loc_005969DC:   mov edx, esp
  loc_005969DE: mov ecx, 00000003h
  loc_005969E3: sub esp, 00000010h
loc_005969E6:   mov [edx], ecx
loc_005969E8:   mov [edx+00000004h], ebx
loc_005969EB:   mov [edx+00000008h], ecx
loc_005969EE:   mov ecx, var_AC
loc_005969F4:   mov [edx+0000000Ch], ecx
loc_005969F7:   mov ecx, var_58
loc_005969FA:   mov edx, esp
loc_005969FC:   mov [edx], ecx
loc_005969FE:   mov ecx, var_54
  loc_00596A01: push 00000002h
  loc_00596A03: push 00000041h
loc_00596A05:   mov [edx+00000004h], ecx
loc_00596A08:   mov ecx, [esi]
loc_00596A0A:   push esi
loc_00596A0B:   mov [edx+00000008h], eax
loc_00596A0E:   mov eax, var_4C
loc_00596A11:   mov [edx+0000000Ch], eax
loc_00596A14:   Call [ecx+00000390h]
loc_00596A1A:   lea edx, var_44
loc_00596A1D:   push eax
loc_00596A1E:   push edx
loc_00596A1F:   Call edi
loc_00596A21:   push eax
  loc_00596A22: call [0040117Ch] ; __vbaLateIdCallSt
  loc_00596A28: add esp, 0000003Ch
loc_00596A2B:   lea ecx, var_44
  loc_00596A2E: call [004012A0h] ; __vbaFreeObj
loc_00596A34:   lea ecx, var_58
  loc_00596A37: call [00401020h] ; __vbaFreeVar
loc_00596A3D:   movsx eax, [esi+00000034h]
loc_00596A41:   lea ecx, var_38
loc_00596A44:   mov var_90, eax
loc_00596A4A:   push ecx
  loc_00596A4B: mov var_98, 00000003h
  loc_00596A55: mov var_B0, 00000004h
  loc_00596A5F: call [00401270h] ; __vbaStrVarCopy
loc_00596A65:   mov ecx, var_98
  loc_00596A6B: sub esp, 00000010h
loc_00596A6E:   mov edx, esp
  loc_00596A70: sub esp, 00000010h
  loc_00596A73: mov var_58, 00000008h
loc_00596A7A:   mov var_50, eax
loc_00596A7D:   mov [edx], ecx
loc_00596A7F:   mov ecx, var_94
loc_00596A85:   mov [edx+00000004h], ecx
loc_00596A88:   mov ecx, var_90
loc_00596A8E:   mov [edx+00000008h], ecx
loc_00596A91:   mov ecx, var_8C
loc_00596A97:   mov [edx+0000000Ch], ecx
loc_00596A9A:   mov edx, esp
  loc_00596A9C: mov ecx, 00000003h
  loc_00596AA1: sub esp, 00000010h
loc_00596AA4:   mov [edx], ecx
loc_00596AA6:   mov ecx, var_B0
loc_00596AAC:   mov [edx+00000004h], ebx
loc_00596AAF:   mov [edx+00000008h], ecx
loc_00596AB2:   mov ecx, var_AC
loc_00596AB8:   mov [edx+0000000Ch], ecx
loc_00596ABB:   mov ecx, var_58
loc_00596ABE:   mov edx, esp
  loc_00596AC0: push 00000002h
  loc_00596AC2: push 00000041h
loc_00596AC4:   push esi
loc_00596AC5:   mov [edx], ecx
loc_00596AC7:   mov ecx, var_54
loc_00596ACA:   mov [edx+00000004h], ecx
loc_00596ACD:   mov ecx, [esi]
loc_00596ACF:   mov [edx+00000008h], eax
loc_00596AD2:   mov eax, var_4C
loc_00596AD5:   mov [edx+0000000Ch], eax
loc_00596AD8:   Call [ecx+00000390h]
loc_00596ADE:   lea edx, var_44
loc_00596AE1:   push eax
loc_00596AE2:   push edx
loc_00596AE3:   Call edi
loc_00596AE5:   push eax
  loc_00596AE6: call [0040117Ch] ; __vbaLateIdCallSt
  loc_00596AEC: add esp, 0000003Ch
loc_00596AEF:   lea ecx, var_44
  loc_00596AF2: call [004012A0h] ; __vbaFreeObj
loc_00596AF8:   lea ecx, var_58
  loc_00596AFB: call [00401020h] ; __vbaFreeVar
loc_00596B01:   movsx eax, [esi+00000034h]
loc_00596B05:   mov ecx, [esi]
loc_00596B07:   push esi
loc_00596B08:   mov var_90, eax
  loc_00596B0E: mov var_98, 00000003h
loc_00596B18:   Call [ecx+00000344h]
loc_00596B1E:   lea edx, var_44
loc_00596B21:   push eax
loc_00596B22:   push edx
loc_00596B23:   Call edi
loc_00596B25:   mov ecx, [eax]
loc_00596B27:   lea edx, var_3C
loc_00596B2A:   push edx
loc_00596B2B:   push eax
loc_00596B2C:   mov var_E4, eax
loc_00596B32:   Call [ecx+000000A0h]
loc_00596B38:   test eax, eax
loc_00596B3A:   fnclex
  loc_00596B3C: jge 00596B56h
loc_00596B3E:   mov ecx, var_E4
  loc_00596B44: push 000000A0h
  loc_00596B49: push 0042DFCCh
loc_00596B4E:   push ecx
loc_00596B4F:   push eax
  loc_00596B50: call [00401080h] ; __vbaHresultCheckObj
loc_00596B56:   mov ecx, var_98
  loc_00596B5C: sub esp, 00000010h
loc_00596B5F:   mov edx, esp
  loc_00596B61: sub esp, 00000010h
loc_00596B64:   mov eax, var_3C
  loc_00596B67: mov var_58, 00000008h
loc_00596B6E:   mov [edx], ecx
loc_00596B70:   mov ecx, var_94
loc_00596B76:   mov var_50, eax
  loc_00596B79: mov var_3C, 00000000h
loc_00596B80:   mov [edx+00000004h], ecx
loc_00596B83:   mov ecx, var_90
loc_00596B89:   mov [edx+00000008h], ecx
loc_00596B8C:   mov ecx, var_8C
loc_00596B92:   mov [edx+0000000Ch], ecx
loc_00596B95:   mov edx, esp
  loc_00596B97: mov ecx, 00000003h
  loc_00596B9C: sub esp, 00000010h
loc_00596B9F:   mov [edx], ecx
  loc_00596BA1: mov ecx, 00000005h
loc_00596BA6:   mov [edx+00000004h], ebx
loc_00596BA9:   mov [edx+00000008h], ecx
loc_00596BAC:   mov ecx, var_AC
loc_00596BB2:   mov [edx+0000000Ch], ecx
loc_00596BB5:   mov ecx, var_58
loc_00596BB8:   mov edx, esp
  loc_00596BBA: push 00000002h
  loc_00596BBC: push 00000041h
loc_00596BBE:   push esi
loc_00596BBF:   mov [edx], ecx
loc_00596BC1:   mov ecx, var_54
loc_00596BC4:   mov [edx+00000004h], ecx
loc_00596BC7:   mov ecx, [esi]
loc_00596BC9:   mov [edx+00000008h], eax
loc_00596BCC:   mov eax, var_4C
loc_00596BCF:   mov [edx+0000000Ch], eax
loc_00596BD2:   Call [ecx+00000390h]
loc_00596BD8:   lea edx, var_48
loc_00596BDB:   push eax
loc_00596BDC:   push edx
loc_00596BDD:   Call edi
loc_00596BDF:   push eax
  loc_00596BE0: call [0040117Ch] ; __vbaLateIdCallSt
loc_00596BE6:   lea eax, var_48
loc_00596BE9:   lea ecx, var_44
loc_00596BEC:   push eax
loc_00596BED:   push ecx
  loc_00596BEE: push 00000002h
  loc_00596BF0: call [0040104Ch] ; __vbaFreeObjList
  loc_00596BF6: add esp, 00000048h
loc_00596BF9:   lea ecx, var_58
  loc_00596BFC: call [00401020h] ; __vbaFreeVar
loc_00596C02:   movsx edx, [esi+00000034h]
  loc_00596C06: mov eax, 00000003h
loc_00596C0B:   push esi
loc_00596C0C:   mov var_98, eax
loc_00596C12:   mov var_B8, eax
loc_00596C18:   mov eax, [esi]
loc_00596C1A:   mov var_90, edx
loc_00596C20:   Call [eax+00000340h]
loc_00596C26:   lea ecx, var_44
loc_00596C29:   push eax
loc_00596C2A:   push ecx
loc_00596C2B:   Call edi
loc_00596C2D:   mov edx, [eax]
loc_00596C2F:   lea ecx, var_3C
loc_00596C32:   push ecx
loc_00596C33:   push eax
loc_00596C34:   mov var_E4, eax
loc_00596C3A:   Call [edx+000000A0h]
loc_00596C40:   test eax, eax
loc_00596C42:   fnclex
  loc_00596C44: jge 00596C5Eh
loc_00596C46:   mov edx, var_E4
  loc_00596C4C: push 000000A0h
  loc_00596C51: push 0042DFCCh
loc_00596C56:   push edx
loc_00596C57:   push eax
  loc_00596C58: call [00401080h] ; __vbaHresultCheckObj
loc_00596C5E:   mov edx, var_98
  loc_00596C64: sub esp, 00000010h
loc_00596C67:   mov ecx, esp
  loc_00596C69: sub esp, 00000010h
loc_00596C6C:   mov eax, var_3C
  loc_00596C6F: mov var_58, 00000008h
loc_00596C76:   mov [ecx], edx
loc_00596C78:   mov edx, var_94
loc_00596C7E:   mov var_50, eax
  loc_00596C81: mov var_3C, 00000000h
loc_00596C88:   mov [ecx+00000004h], edx
loc_00596C8B:   mov edx, var_90
loc_00596C91:   mov [ecx+00000008h], edx
loc_00596C94:   mov edx, var_8C
loc_00596C9A:   mov [ecx+0000000Ch], edx
loc_00596C9D:   mov ecx, var_B8
loc_00596CA3:   mov edx, esp
  loc_00596CA5: sub esp, 00000010h
loc_00596CA8:   mov [edx], ecx
  loc_00596CAA: mov ecx, 00000006h
loc_00596CAF:   mov [edx+00000004h], ebx
loc_00596CB2:   mov [edx+00000008h], ecx
loc_00596CB5:   mov ecx, var_AC
loc_00596CBB:   mov [edx+0000000Ch], ecx
loc_00596CBE:   mov ecx, var_58
loc_00596CC1:   mov edx, esp
  loc_00596CC3: push 00000002h
  loc_00596CC5: push 00000041h
loc_00596CC7:   push esi
loc_00596CC8:   mov [edx], ecx
loc_00596CCA:   mov ecx, var_54
loc_00596CCD:   mov [edx+00000004h], ecx
loc_00596CD0:   mov ecx, [esi]
loc_00596CD2:   mov [edx+00000008h], eax
loc_00596CD5:   mov eax, var_4C
loc_00596CD8:   mov [edx+0000000Ch], eax
loc_00596CDB:   Call [ecx+00000390h]
loc_00596CE1:   lea edx, var_48
loc_00596CE4:   push eax
loc_00596CE5:   push edx
loc_00596CE6:   Call edi
loc_00596CE8:   push eax
  loc_00596CE9: call [0040117Ch] ; __vbaLateIdCallSt
loc_00596CEF:   lea eax, var_48
loc_00596CF2:   lea ecx, var_44
loc_00596CF5:   push eax
loc_00596CF6:   push ecx
  loc_00596CF7: push 00000002h
  loc_00596CF9: call [0040104Ch] ; __vbaFreeObjList
  loc_00596CFF: add esp, 00000048h
loc_00596D02:   lea ecx, var_58
  loc_00596D05: call [00401020h] ; __vbaFreeVar
loc_00596D0B:   mov dx, [esi+00000034h]
loc_00596D0F:   mov eax, [esi]
  loc_00596D11: add dx, 0001h
loc_00596D15:   push esi
  loc_00596D16: jo 00596F87h
loc_00596D1C:   mov [esi+00000034h], dx
loc_00596D20:   Call [eax+00000354h]
loc_00596D26:   lea ecx, var_44
loc_00596D29:   push eax
loc_00596D2A:   push ecx
loc_00596D2B:   Call edi
loc_00596D2D:   mov ebx, eax
  loc_00596D2F: push 0042DFECh
loc_00596D34:   push ebx
loc_00596D35:   mov edx, [ebx]
loc_00596D37:   Call [edx+000000A4h]
loc_00596D3D:   test eax, eax
loc_00596D3F:   fnclex
  loc_00596D41: jge 00596D55h
  loc_00596D43: push 000000A4h
  loc_00596D48: push 0042DFCCh
loc_00596D4D:   push ebx
loc_00596D4E:   push eax
  loc_00596D4F: call [00401080h] ; __vbaHresultCheckObj
loc_00596D55:   lea ecx, var_44
  loc_00596D58: call [004012A0h] ; __vbaFreeObj
loc_00596D5E:   mov eax, [esi]
loc_00596D60:   push esi
loc_00596D61:   Call [eax+0000034Ch]
loc_00596D67:   lea ecx, var_44
loc_00596D6A:   push eax
loc_00596D6B:   push ecx
loc_00596D6C:   Call edi
loc_00596D6E:   mov ebx, eax
  loc_00596D70: push 0042DFECh
loc_00596D75:   push ebx
loc_00596D76:   mov edx, [ebx]
loc_00596D78:   Call [edx+000000A4h]
loc_00596D7E:   test eax, eax
loc_00596D80:   fnclex
  loc_00596D82: jge 00596D96h
  loc_00596D84: push 000000A4h
  loc_00596D89: push 0042DFCCh
loc_00596D8E:   push ebx
loc_00596D8F:   push eax
  loc_00596D90: call [00401080h] ; __vbaHresultCheckObj
loc_00596D96:   lea ecx, var_44
  loc_00596D99: call [004012A0h] ; __vbaFreeObj
loc_00596D9F:   mov eax, [esi]
loc_00596DA1:   push esi
loc_00596DA2:   Call [eax+00000348h]
loc_00596DA8:   lea ecx, var_44
loc_00596DAB:   push eax
loc_00596DAC:   push ecx
loc_00596DAD:   Call edi
loc_00596DAF:   mov ebx, eax
  loc_00596DB1: push 0042E3A4h
loc_00596DB6:   push ebx
loc_00596DB7:   mov edx, [ebx]
loc_00596DB9:   Call [edx+000000A4h]
loc_00596DBF:   test eax, eax
loc_00596DC1:   fnclex
  loc_00596DC3: jge 00596DD7h
  loc_00596DC5: push 000000A4h
  loc_00596DCA: push 0042DFCCh
loc_00596DCF:   push ebx
loc_00596DD0:   push eax
  loc_00596DD1: call [00401080h] ; __vbaHresultCheckObj
loc_00596DD7:   lea ecx, var_44
  loc_00596DDA: call [004012A0h] ; __vbaFreeObj
loc_00596DE0:   mov eax, [esi]
loc_00596DE2:   push esi
loc_00596DE3:   Call [eax+00000350h]
loc_00596DE9:   lea ecx, var_44
loc_00596DEC:   push eax
loc_00596DED:   push ecx
loc_00596DEE:   Call edi
loc_00596DF0:   mov ebx, eax
  loc_00596DF2: push 0042E3A4h
loc_00596DF7:   push ebx
loc_00596DF8:   mov edx, [ebx]
loc_00596DFA:   Call [edx+000000A4h]
loc_00596E00:   test eax, eax
loc_00596E02:   fnclex
  loc_00596E04: jge 00596E18h
  loc_00596E06: push 000000A4h
  loc_00596E0B: push 0042DFCCh
loc_00596E10:   push ebx
loc_00596E11:   push eax
  loc_00596E12: call [00401080h] ; __vbaHresultCheckObj
loc_00596E18:   lea ecx, var_44
  loc_00596E1B: call [004012A0h] ; __vbaFreeObj
loc_00596E21:   mov eax, [esi]
loc_00596E23:   push esi
loc_00596E24:   Call [eax+00000344h]
loc_00596E2A:   lea ecx, var_44
loc_00596E2D:   push eax
loc_00596E2E:   push ecx
loc_00596E2F:   Call edi
loc_00596E31:   mov ebx, eax
  loc_00596E33: push 0042E3A4h
loc_00596E38:   push ebx
loc_00596E39:   mov edx, [ebx]
loc_00596E3B:   Call [edx+000000A4h]
loc_00596E41:   test eax, eax
loc_00596E43:   fnclex
  loc_00596E45: jge 00596E59h
  loc_00596E47: push 000000A4h
  loc_00596E4C: push 0042DFCCh
loc_00596E51:   push ebx
loc_00596E52:   push eax
  loc_00596E53: call [00401080h] ; __vbaHresultCheckObj
loc_00596E59:   lea ecx, var_44
  loc_00596E5C: call [004012A0h] ; __vbaFreeObj
loc_00596E62:   mov eax, [esi]
loc_00596E64:   push esi
loc_00596E65:   Call [eax+00000340h]
loc_00596E6B:   lea ecx, var_44
loc_00596E6E:   push eax
loc_00596E6F:   push ecx
loc_00596E70:   Call edi
loc_00596E72:   mov ebx, eax
  loc_00596E74: push 0042E3A4h
loc_00596E79:   push ebx
loc_00596E7A:   mov edx, [ebx]
loc_00596E7C:   Call [edx+000000A4h]
loc_00596E82:   test eax, eax
loc_00596E84:   fnclex
  loc_00596E86: jge 00596E9Eh
  loc_00596E88: push 000000A4h
  loc_00596E8D: push 0042DFCCh
loc_00596E92:   push ebx
  loc_00596E93: mov ebx, [00401080h] ; __vbaHresultCheckObj
loc_00596E99:   push eax
loc_00596E9A:   Call ebx
  loc_00596E9C: jmp 00596EA4h
  loc_00596E9E: mov ebx, [00401080h] ; __vbaHresultCheckObj
loc_00596EA4:   lea ecx, var_44
  loc_00596EA7: call [004012A0h] ; __vbaFreeObj
loc_00596EAD:   mov eax, [esi]
loc_00596EAF:   push esi
loc_00596EB0:   Call [eax+00000710h]
loc_00596EB6:   test eax, eax
  loc_00596EB8: jge 00596EC8h
  loc_00596EBA: push 00000710h
  loc_00596EBF: push 00433DD8h
loc_00596EC4:   push esi
loc_00596EC5:   push eax
loc_00596EC6:   Call ebx
loc_00596EC8:   mov ecx, [esi]
loc_00596ECA:   push esi
loc_00596ECB:   Call [ecx+00000354h]
loc_00596ED1:   lea edx, var_44
loc_00596ED4:   push eax
loc_00596ED5:   push edx
loc_00596ED6:   Call edi
loc_00596ED8:   mov esi, eax
loc_00596EDA:   push esi
loc_00596EDB:   mov eax, [esi]
loc_00596EDD:   Call [eax+00000204h]
loc_00596EE3:   test eax, eax
loc_00596EE5:   fnclex
  loc_00596EE7: jge 00596EF7h
  loc_00596EE9: push 00000204h
  loc_00596EEE: push 0042DFCCh
loc_00596EF3:   push esi
loc_00596EF4:   push eax
loc_00596EF5:   Call ebx
loc_00596EF7:   lea ecx, var_44
  loc_00596EFA: call [004012A0h] ; __vbaFreeObj
  loc_00596F00: mov var_4, 00000000h
loc_00596F07:   fwait
  loc_00596F08: push 00596F68h
  loc_00596F0D: jmp 00596F4Eh
loc_00596F0F:   lea ecx, var_40
loc_00596F12:   lea edx, var_3C
loc_00596F15:   push ecx
loc_00596F16:   push edx
  loc_00596F17: push 00000002h
  loc_00596F19: call [00401204h] ; __vbaFreeStrList
loc_00596F1F:   lea eax, var_48
loc_00596F22:   lea ecx, var_44
loc_00596F25:   push eax
loc_00596F26:   push ecx
  loc_00596F27: push 00000002h
  loc_00596F29: call [0040104Ch] ; __vbaFreeObjList
loc_00596F2F:   lea edx, var_88
loc_00596F35:   lea eax, var_78
loc_00596F38:   push edx
loc_00596F39:   lea ecx, var_68
loc_00596F3C:   push eax
loc_00596F3D:   lea edx, var_58
loc_00596F40:   push ecx
loc_00596F41:   push edx
  loc_00596F42: push 00000004h
  loc_00596F44: call [0040103Ch] ; __vbaFreeVarList
  loc_00596F4A: add esp, 0000002Ch
loc_00596F4D:   ret
  loc_00596F4E: mov esi, [00401020h] ; __vbaFreeVar
loc_00596F54:   lea ecx, var_24
  loc_00596F57: call __vbaFreeVar
loc_00596F59:   lea ecx, var_28
  loc_00596F5C: call [0040129Ch] ; __vbaFreeStr
loc_00596F62:   lea ecx, var_38
  loc_00596F65: call __vbaFreeVar
loc_00596F67:   ret
loc_00596F68:   mov eax, Me
loc_00596F6B:   push eax
loc_00596F6C:   mov ecx, [eax]
loc_00596F6E:   Call [ecx+00000008h]
loc_00596F71:   mov eax, var_4
loc_00596F74:   mov ecx, var_14
loc_00596F77:   pop edi
loc_00596F78:   pop esi
loc_00596F79:   mov fs: [00000000h] , ecx
loc_00596F80:   pop ebx
loc_00596F81:   mov esp, ebp
loc_00596F83:   pop ebp
  loc_00596F84: retn 0004h
End Sub

Private Sub txtNotaBeli_KeyPress(KeyAscii As Integer) '59CF20
loc_0059CF20:   push ebp
loc_0059CF21:   mov ebp, esp
  loc_0059CF23: sub esp, 0000000Ch
  loc_0059CF26: push 00405696h ; __vbaExceptHandler
loc_0059CF2B:   mov eax, fs: [00000000h]
loc_0059CF31:   push eax
loc_0059CF32:   mov fs: [00000000h] , esp
  loc_0059CF39: sub esp, 000000A4h
loc_0059CF3F:   push ebx
loc_0059CF40:   push esi
loc_0059CF41:   push edi
loc_0059CF42:   mov var_C, esp
  loc_0059CF45: mov var_8, 00402970h
loc_0059CF4C:   mov esi, Me
loc_0059CF4F:   mov eax, esi
  loc_0059CF51: and eax, 00000001h
loc_0059CF54:   mov var_4, eax
  loc_0059CF57: and esi, FFFFFFFEh
loc_0059CF5A:   push esi
loc_0059CF5B:   mov Me, esi
loc_0059CF5E:   mov ecx, [esi]
loc_0059CF60:   Call [ecx+00000004h]
loc_0059CF63:   mov ebx, KeyAscii
  loc_0059CF66: xor edi, edi
loc_0059CF68:   mov var_18, edi
loc_0059CF6B:   mov var_1C, edi
  loc_0059CF6E: cmp [ebx], 000Dh
loc_0059CF72:   mov var_20, edi
loc_0059CF75:   mov var_30, edi
loc_0059CF78:   mov var_40, edi
loc_0059CF7B:   mov var_50, edi
loc_0059CF7E:   mov var_60, edi
loc_0059CF81:   mov var_70, edi
loc_0059CF84:   mov var_A4, edi
  loc_0059CF8A: jnz 0059D2F4h
loc_0059CF90:   mov edx, [esi]
loc_0059CF92:   push esi
loc_0059CF93:   Call [edx+0000031Ch]
  loc_0059CF99: mov ebx, [004010B8h] ; __vbaObjSet
loc_0059CF9F:   push eax
loc_0059CFA0:   lea eax, var_20
loc_0059CFA3:   push eax
loc_0059CFA4:   Call ebx
loc_0059CFA6:   mov edi, eax
loc_0059CFA8:   lea edx, var_18
loc_0059CFAB:   push edx
loc_0059CFAC:   push edi
loc_0059CFAD:   mov ecx, [edi]
loc_0059CFAF:   Call [ecx+000000A0h]
loc_0059CFB5:   test eax, eax
loc_0059CFB7:   fnclex
  loc_0059CFB9: jge 0059CFCDh
  loc_0059CFBB: push 000000A0h
  loc_0059CFC0: push 0042DFCCh
loc_0059CFC5:   push edi
loc_0059CFC6:   push eax
  loc_0059CFC7: call [00401080h] ; __vbaHresultCheckObj
loc_0059CFCD:   mov eax, var_18
loc_0059CFD0:   push eax
  loc_0059CFD1: push 0042DFECh
  loc_0059CFD6: call [0040112Ch] ; __vbaStrCmp
loc_0059CFDC:   mov edi, eax
loc_0059CFDE:   lea ecx, var_18
loc_0059CFE1:   neg edi
loc_0059CFE3:   sbb edi, edi
loc_0059CFE5:   inc edi
loc_0059CFE6:   neg edi
  loc_0059CFE8: call [0040129Ch] ; __vbaFreeStr
loc_0059CFEE:   lea ecx, var_20
  loc_0059CFF1: call [004012A0h] ; __vbaFreeObj
loc_0059CFF7:   test di, di
  loc_0059CFFA: jz 0059D0ACh
  loc_0059D000: mov ecx, 80020004h
  loc_0059D005: mov eax, 0000000Ah
loc_0059D00A:   mov var_58, ecx
loc_0059D00D:   mov var_48, ecx
loc_0059D010:   mov var_38, ecx
loc_0059D013:   lea edx, var_70
loc_0059D016:   lea ecx, var_30
loc_0059D019:   mov var_60, eax
loc_0059D01C:   mov var_50, eax
loc_0059D01F:   mov var_40, eax
  loc_0059D022: mov var_68, 004349DCh ; "Nota Beli wajib diisi"
  loc_0059D029: mov var_70, 00000008h
  loc_0059D030: call [00401238h] ; __vbaVarDup
loc_0059D036:   lea ecx, var_60
loc_0059D039:   lea edx, var_50
loc_0059D03C:   push ecx
loc_0059D03D:   lea eax, var_40
loc_0059D040:   push edx
loc_0059D041:   push eax
loc_0059D042:   lea ecx, var_30
  loc_0059D045: push 00000000h
loc_0059D047:   push ecx
  loc_0059D048: call [004010B4h] ; rtcMsgBox
loc_0059D04E:   lea edx, var_60
loc_0059D051:   lea eax, var_50
loc_0059D054:   push edx
loc_0059D055:   lea ecx, var_40
loc_0059D058:   push eax
loc_0059D059:   lea edx, var_30
loc_0059D05C:   push ecx
loc_0059D05D:   push edx
  loc_0059D05E: push 00000004h
  loc_0059D060: call [0040103Ch] ; __vbaFreeVarList
loc_0059D066:   mov eax, [esi]
  loc_0059D068: add esp, 00000014h
loc_0059D06B:   push esi
loc_0059D06C:   Call [eax+0000031Ch]
loc_0059D072:   lea ecx, var_20
loc_0059D075:   push eax
loc_0059D076:   push ecx
loc_0059D077:   Call ebx
loc_0059D079:   mov esi, eax
loc_0059D07B:   push esi
loc_0059D07C:   mov edx, [esi]
loc_0059D07E:   Call [edx+00000204h]
loc_0059D084:   test eax, eax
loc_0059D086:   fnclex
  loc_0059D088: jge 0059D09Ch
  loc_0059D08A: push 00000204h
  loc_0059D08F: push 0042DFCCh
loc_0059D094:   push esi
loc_0059D095:   push eax
  loc_0059D096: call [00401080h] ; __vbaHresultCheckObj
loc_0059D09C:   lea ecx, var_20
  loc_0059D09F: call [004012A0h] ; __vbaFreeObj
  loc_0059D0A5: xor edi, edi
  loc_0059D0A7: jmp 0059D344h
  loc_0059D0AC: push 0042DF30h
  loc_0059D0B1: call [00401168h] ; __vbaNew
loc_0059D0B7:   push eax
  loc_0059D0B8: push 00668060h
loc_0059D0BD:   Call ebx
loc_0059D0BF:   mov eax, [esi]
loc_0059D0C1:   push esi
loc_0059D0C2:   Call [eax+0000031Ch]
loc_0059D0C8:   lea ecx, var_20
loc_0059D0CB:   push eax
loc_0059D0CC:   push ecx
loc_0059D0CD:   Call ebx
loc_0059D0CF:   mov edi, eax
loc_0059D0D1:   lea eax, var_18
loc_0059D0D4:   push eax
loc_0059D0D5:   push edi
loc_0059D0D6:   mov edx, [edi]
loc_0059D0D8:   Call [edx+000000A0h]
loc_0059D0DE:   test eax, eax
loc_0059D0E0:   fnclex
  loc_0059D0E2: jge 0059D0F6h
  loc_0059D0E4: push 000000A0h
  loc_0059D0E9: push 0042DFCCh
loc_0059D0EE:   push edi
loc_0059D0EF:   push eax
  loc_0059D0F0: call [00401080h] ; __vbaHresultCheckObj
loc_0059D0F6:   mov eax, [0066803Ch]
loc_0059D0FB:   test eax, eax
  loc_0059D0FD: jnz 0059D10Fh
  loc_0059D0FF: push 0066803Ch
  loc_0059D104: push 0042DEFCh
  loc_0059D109: call [004011E8h] ; __vbaNew2
loc_0059D10F:   mov edx, var_18
loc_0059D112:   mov ecx, [0066803Ch]
  loc_0059D118: mov edi, [00401070h] ; __vbaStrCat
  loc_0059D11E: push 004343BCh ; "SELECT * FROM pembelian WHERE no_notabeli='"
loc_0059D123:   push edx
loc_0059D124:   mov var_68, ecx
  loc_0059D127: mov var_70, 00000009h
loc_0059D12E:   Call edi
loc_0059D130:   mov edx, eax
loc_0059D132:   lea ecx, var_1C
  loc_0059D135: call [0040126Ch] ; __vbaStrMove
loc_0059D13B:   push eax
  loc_0059D13C: push 0042EBA8h ; "'"
loc_0059D141:   Call edi
loc_0059D143:   mov ebx, var_70
loc_0059D146:   push FFFFFFFFh
  loc_0059D148: push 00000004h
  loc_0059D14A: push 00000002h
  loc_0059D14C: sub esp, 00000010h
loc_0059D14F:   mov edx, [00668060h]
loc_0059D155:   mov edi, esp
  loc_0059D157: sub esp, 00000010h
  loc_0059D15A: mov ecx, 00000008h
loc_0059D15F:   mov var_28, eax
loc_0059D162:   mov [edi], ebx
loc_0059D164:   mov ebx, var_6C
loc_0059D167:   mov var_30, ecx
loc_0059D16A:   mov edx, [edx]
loc_0059D16C:   mov [edi+00000004h], ebx
loc_0059D16F:   mov ebx, var_68
loc_0059D172:   mov [edi+00000008h], ebx
loc_0059D175:   mov ebx, var_64
loc_0059D178:   mov [edi+0000000Ch], ebx
loc_0059D17B:   mov edi, esp
loc_0059D17D:   mov [edi], ecx
loc_0059D17F:   mov ecx, var_2C
loc_0059D182:   mov [edi+00000004h], ecx
loc_0059D185:   mov ecx, [00668060h]
loc_0059D18B:   push ecx
loc_0059D18C:   mov [edi+00000008h], eax
loc_0059D18F:   mov eax, var_24
loc_0059D192:   mov [edi+0000000Ch], eax
loc_0059D195:   Call [edx+000000A0h]
loc_0059D19B:   test eax, eax
loc_0059D19D:   fnclex
  loc_0059D19F: jge 0059D1BDh
loc_0059D1A1:   mov edx, [00668060h]
  loc_0059D1A7: mov ebx, [00401080h] ; __vbaHresultCheckObj
  loc_0059D1AD: push 000000A0h
  loc_0059D1B2: push 0042DF44h
loc_0059D1B7:   push edx
loc_0059D1B8:   push eax
loc_0059D1B9:   Call ebx
  loc_0059D1BB: jmp 0059D1C3h
  loc_0059D1BD: mov ebx, [00401080h] ; __vbaHresultCheckObj
loc_0059D1C3:   lea eax, var_1C
loc_0059D1C6:   lea ecx, var_18
loc_0059D1C9:   push eax
loc_0059D1CA:   push ecx
  loc_0059D1CB: push 00000002h
  loc_0059D1CD: call [00401204h] ; __vbaFreeStrList
  loc_0059D1D3: mov edi, [004012A0h] ; __vbaFreeObj
  loc_0059D1D9: add esp, 0000000Ch
loc_0059D1DC:   lea ecx, var_20
loc_0059D1DF:   Call edi
loc_0059D1E1:   lea ecx, var_30
  loc_0059D1E4: call [00401020h] ; __vbaFreeVar
loc_0059D1EA:   mov eax, [00668060h]
loc_0059D1EF:   lea ecx, var_A4
loc_0059D1F5:   push ecx
loc_0059D1F6:   push eax
loc_0059D1F7:   mov edx, [eax]
loc_0059D1F9:   Call [edx+00000050h]
loc_0059D1FC:   test eax, eax
loc_0059D1FE:   fnclex
  loc_0059D200: jge 0059D213h
loc_0059D202:   mov edx, [00668060h]
  loc_0059D208: push 00000050h
  loc_0059D20A: push 0042DF44h
loc_0059D20F:   push edx
loc_0059D210:   push eax
loc_0059D211:   Call ebx
  loc_0059D213: cmp var_A4, 0000h
  loc_0059D21B: jz 0059D24Eh
loc_0059D21D:   mov eax, [esi]
  loc_0059D21F: push 00000000h
  loc_0059D221: push 80011000h
loc_0059D226:   push esi
loc_0059D227:   Call [eax+00000380h]
loc_0059D22D:   lea ecx, var_20
loc_0059D230:   push eax
loc_0059D231:   push ecx
  loc_0059D232: call [004010B8h] ; __vbaObjSet
loc_0059D238:   push eax
  loc_0059D239: call [0040102Ch] ; __vbaLateIdCall
  loc_0059D23F: add esp, 0000000Ch
loc_0059D242:   lea ecx, var_20
loc_0059D245:   Call edi
  loc_0059D247: xor edi, edi
  loc_0059D249: jmp 0059D344h
  loc_0059D24E: mov ecx, 80020004h
  loc_0059D253: mov eax, 0000000Ah
loc_0059D258:   mov var_58, ecx
loc_0059D25B:   mov var_48, ecx
loc_0059D25E:   mov var_38, ecx
loc_0059D261:   lea edx, var_70
loc_0059D264:   lea ecx, var_30
loc_0059D267:   mov var_60, eax
loc_0059D26A:   mov var_50, eax
loc_0059D26D:   mov var_40, eax
  loc_0059D270: mov var_68, 00434818h ; "nomor faktur ganda, ganti dengan nomor lain"
  loc_0059D277: mov var_70, 00000008h
  loc_0059D27E: call [00401238h] ; __vbaVarDup
loc_0059D284:   lea edx, var_60
loc_0059D287:   lea eax, var_50
loc_0059D28A:   push edx
loc_0059D28B:   lea ecx, var_40
loc_0059D28E:   push eax
loc_0059D28F:   push ecx
loc_0059D290:   lea edx, var_30
  loc_0059D293: push 00000000h
loc_0059D295:   push edx
  loc_0059D296: call [004010B4h] ; rtcMsgBox
loc_0059D29C:   lea eax, var_60
loc_0059D29F:   lea ecx, var_50
loc_0059D2A2:   push eax
loc_0059D2A3:   lea edx, var_40
loc_0059D2A6:   push ecx
loc_0059D2A7:   lea eax, var_30
loc_0059D2AA:   push edx
loc_0059D2AB:   push eax
  loc_0059D2AC: push 00000004h
  loc_0059D2AE: call [0040103Ch] ; __vbaFreeVarList
loc_0059D2B4:   mov ecx, [esi]
  loc_0059D2B6: add esp, 00000014h
loc_0059D2B9:   push esi
loc_0059D2BA:   Call [ecx+0000031Ch]
loc_0059D2C0:   lea edx, var_20
loc_0059D2C3:   push eax
loc_0059D2C4:   push edx
  loc_0059D2C5: call [004010B8h] ; __vbaObjSet
loc_0059D2CB:   mov esi, eax
loc_0059D2CD:   push esi
loc_0059D2CE:   mov eax, [esi]
loc_0059D2D0:   Call [eax+00000204h]
loc_0059D2D6:   test eax, eax
loc_0059D2D8:   fnclex
  loc_0059D2DA: jge 0059D2EAh
  loc_0059D2DC: push 00000204h
  loc_0059D2E1: push 0042DFCCh
loc_0059D2E6:   push esi
loc_0059D2E7:   push eax
loc_0059D2E8:   Call ebx
loc_0059D2EA:   lea ecx, var_20
loc_0059D2ED:   Call edi
loc_0059D2EF:   mov ebx, KeyAscii
  loc_0059D2F2: xor edi, edi
loc_0059D2F4:   movsx ecx, [ebx]
loc_0059D2F7:   lea edx, var_30
loc_0059D2FA:   push ecx
loc_0059D2FB:   push edx
  loc_0059D2FC: call [004011B0h] ; rtcVarBstrFromAnsi
loc_0059D302:   lea eax, var_30
loc_0059D305:   lea ecx, var_40
loc_0059D308:   push eax
loc_0059D309:   push ecx
  loc_0059D30A: call [00401124h] ; rtcUpperCaseVar
loc_0059D310:   lea edx, var_40
loc_0059D313:   lea eax, var_18
loc_0059D316:   push edx
loc_0059D317:   push eax
  loc_0059D318: call [004011C0h] ; __vbaStrVarVal
loc_0059D31E:   push eax
  loc_0059D31F: call [00401050h] ; rtcAnsiValueBstr
loc_0059D325:   lea ecx, var_18
loc_0059D328:   mov [ebx], ax
  loc_0059D32B: call [0040129Ch] ; __vbaFreeStr
loc_0059D331:   lea ecx, var_40
loc_0059D334:   lea edx, var_30
loc_0059D337:   push ecx
loc_0059D338:   push edx
  loc_0059D339: push 00000002h
  loc_0059D33B: call [0040103Ch] ; __vbaFreeVarList
  loc_0059D341: add esp, 0000000Ch
loc_0059D344:   mov var_4, edi
  loc_0059D347: push 0059D387h
  loc_0059D34C: jmp 0059D386h
loc_0059D34E:   lea eax, var_1C
loc_0059D351:   lea ecx, var_18
loc_0059D354:   push eax
loc_0059D355:   push ecx
  loc_0059D356: push 00000002h
  loc_0059D358: call [00401204h] ; __vbaFreeStrList
  loc_0059D35E: add esp, 0000000Ch
loc_0059D361:   lea ecx, var_20
  loc_0059D364: call [004012A0h] ; __vbaFreeObj
loc_0059D36A:   lea edx, var_60
loc_0059D36D:   lea eax, var_50
loc_0059D370:   push edx
loc_0059D371:   lea ecx, var_40
loc_0059D374:   push eax
loc_0059D375:   lea edx, var_30
loc_0059D378:   push ecx
loc_0059D379:   push edx
  loc_0059D37A: push 00000004h
  loc_0059D37C: call [0040103Ch] ; __vbaFreeVarList
  loc_0059D382: add esp, 00000014h
loc_0059D385:   ret
loc_0059D386:   ret
loc_0059D387:   mov eax, Me
loc_0059D38A:   push eax
loc_0059D38B:   mov ecx, [eax]
loc_0059D38D:   Call [ecx+00000008h]
loc_0059D390:   mov eax, var_4
loc_0059D393:   mov ecx, var_14
loc_0059D396:   pop edi
loc_0059D397:   pop esi
loc_0059D398:   mov fs: [00000000h] , ecx
loc_0059D39F:   pop ebx
loc_0059D3A0:   mov esp, ebp
loc_0059D3A2:   pop ebp
  loc_0059D3A3: retn 0008h
End Sub

Private Sub txtPemasok_KeyPress(KeyAscii As Integer) '59D3B0
loc_0059D3B0:   push ebp
