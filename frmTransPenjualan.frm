VERSION 5.00
Begin VB.Form frmTransPenjualan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  :: Transaksi Penjualan Retail"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13500
   Icon            =   "frmTransPenjualan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   13500
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bayar (Rp) : "
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   13800
      TabIndex        =   37
      Top             =   240
      Width           =   5055
      Begin VB.Label LabelBeli 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   38.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.TextBox TxtHargaBeli 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   15000
      TabIndex        =   35
      Top             =   2160
      Width           =   1335
   End
   Begin VB.PictureBox CmdBaru 
      Height          =   375
      Left            =   9360
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   27
      Top             =   7560
      Width           =   1215
   End
   Begin VB.PictureBox cmdCari 
      Height          =   375
      Left            =   5040
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   26
      Top             =   2040
      Width           =   495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   6135
      Begin VB.TextBox txtNoNota 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtPelanggan 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   15
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtCatatan 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4080
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.PictureBox DtTempo 
         Height          =   375
         Left            =   4080
         ScaleHeight     =   315
         ScaleWidth      =   1635
         TabIndex        =   21
         Top             =   600
         Width           =   1695
      End
      Begin VB.PictureBox cmbJenis 
         Height          =   360
         Left            =   4080
         ScaleHeight     =   300
         ScaleWidth      =   1635
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
      Begin VB.PictureBox cmbPelanggan 
         Height          =   360
         Left            =   1200
         ScaleHeight     =   300
         ScaleWidth      =   1635
         TabIndex        =   18
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama :"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   480
         TabIndex        =   25
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3480
         TabIndex        =   24
         Top             =   240
         Width           =   450
      End
      Begin VB.Label LblTgl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Tempo"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3000
         TabIndex        =   22
         Top             =   600
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Nota :"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   270
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lblpel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pelanggan :"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   270
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3240
         TabIndex        =   17
         Top             =   960
         Width           =   690
      End
   End
   Begin VB.TextBox txtDiskon 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11160
      TabIndex        =   10
      Top             =   2100
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bayar (Rp) : "
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6600
      TabIndex        =   9
      Top             =   120
      Width           =   6255
      Begin VB.Label labelBayar 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   38.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   6015
      End
   End
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   2100
      Width           =   855
   End
   Begin VB.TextBox txtHarga 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12120
      TabIndex        =   3
      Top             =   2100
      Width           =   1335
   End
   Begin VB.TextBox txtNamaBarang 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5640
      TabIndex        =   2
      Top             =   2100
      Width           =   5415
   End
   Begin VB.TextBox txtKode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Top             =   2100
      Width           =   3855
   End
   Begin VB.PictureBox GridJual 
      Height          =   4575
      Left            =   0
      ScaleHeight     =   4515
      ScaleWidth      =   13395
      TabIndex        =   4
      Top             =   2520
      Width           =   13455
   End
   Begin VB.PictureBox CmdSimpan 
      Height          =   375
      Left            =   10680
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   28
      Top             =   7560
      Width           =   1215
   End
   Begin VB.PictureBox cmdKeluar 
      Height          =   375
      Left            =   12000
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   29
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Beli"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   270
      Left            =   13800
      TabIndex        =   36
      Top             =   2160
      Width           =   960
   End
   Begin VB.Label Label10 
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
      ForeColor       =   &H0080FFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   34
      Top             =   7320
      Width           =   3000
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kasir :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Left            =   360
      TabIndex        =   33
      Top             =   7680
      Width           =   570
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Left            =   2160
      TabIndex        =   32
      Top             =   7680
      Width           =   825
   End
   Begin VB.Label TxtKasir 
      BackStyle       =   0  'Transparent
      Caption         =   "Kasir"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   1080
      TabIndex        =   31
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label TxtTgl 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   3000
      TabIndex        =   30
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   7200
      Width           =   4455
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disc(%)"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   270
      Left            =   11160
      TabIndex        =   11
      Top             =   1800
      Width           =   705
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Harga"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   270
      Left            =   12120
      TabIndex        =   8
      Top             =   1800
      Width           =   555
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Barang :"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   270
      Left            =   5640
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode / Scan Barcode :"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   270
      Left            =   1200
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty :"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   420
   End
End
Attribute VB_Name = "frmTransPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdCari_UnknownEvent_10() '5A9430
loc_005A9430:   push ebp
loc_005A9431:   mov ebp, esp
  loc_005A9433: sub esp, 0000000Ch
  loc_005A9436: push 00405696h ; __vbaExceptHandler
loc_005A943B:   mov eax, fs: [00000000h]
loc_005A9441:   push eax
loc_005A9442:   mov fs: [00000000h] , esp
  loc_005A9449: sub esp, 00000030h
loc_005A944C:   push ebx
loc_005A944D:   push esi
loc_005A944E:   push edi
loc_005A944F:   mov var_C, esp
  loc_005A9452: mov var_8, 00402DF0h
loc_005A9459:   mov eax, Me
loc_005A945C:   mov ecx, eax
  loc_005A945E: and ecx, 00000001h
loc_005A9461:   mov var_4, ecx
  loc_005A9464: and al, FEh
loc_005A9466:   push eax
loc_005A9467:   mov Me, eax
loc_005A946A:   mov edx, [eax]
loc_005A946C:   Call [edx+00000004h]
loc_005A946F:   mov eax, [00668250h]
loc_005A9474:   test eax, eax
  loc_005A9476: jnz 005A9488h
  loc_005A9478: push 00668250h
  loc_005A947D: push 0040A9ECh
  loc_005A9482: call [004011E8h] ; __vbaNew2
loc_005A9488:   mov esi, [00668250h]
  loc_005A948E: push 00434A7Ch ; "PENJUALAN"
loc_005A9493:   push esi
loc_005A9494:   mov eax, [esi]
loc_005A9496:   Call [eax+000006FCh]
loc_005A949C:   test eax, eax
loc_005A949E:   fnclex
  loc_005A94A0: jge 005A94B4h
  loc_005A94A2: push 000006FCh
  loc_005A94A7: push 00434254h
loc_005A94AC:   push esi
loc_005A94AD:   push eax
  loc_005A94AE: call [00401080h] ; __vbaHresultCheckObj
loc_005A94B4:   mov eax, [00668250h]
loc_005A94B9:   test eax, eax
  loc_005A94BB: jnz 005A94CDh
  loc_005A94BD: push 00668250h
  loc_005A94C2: push 0040A9ECh
  loc_005A94C7: call [004011E8h] ; __vbaNew2
  loc_005A94CD: sub esp, 00000010h
  loc_005A94D0: mov ecx, 0000000Ah
loc_005A94D5:   mov ebx, esp
  loc_005A94D7: mov eax, 80020004h
  loc_005A94DC: sub esp, 00000010h
loc_005A94DF:   mov esi, [00668250h]
loc_005A94E5:   mov [ebx], ecx
loc_005A94E7:   mov ecx, var_30
loc_005A94EA:   mov edi, [esi]
  loc_005A94EC: mov edx, 00000001h
loc_005A94F1:   mov [ebx+00000004h], ecx
loc_005A94F4:   mov ecx, esp
loc_005A94F6:   push esi
loc_005A94F7:   mov [ebx+00000008h], eax
loc_005A94FA:   mov eax, var_28
loc_005A94FD:   mov [ebx+0000000Ch], eax
  loc_005A9500: mov eax, 00000002h
loc_005A9505:   mov [ecx], eax
loc_005A9507:   mov eax, var_20
loc_005A950A:   mov [ecx+00000004h], eax
loc_005A950D:   mov [ecx+00000008h], edx
loc_005A9510:   mov edx, var_18
loc_005A9513:   mov [ecx+0000000Ch], edx
loc_005A9516:   Call [edi+000002B0h]
loc_005A951C:   test eax, eax
loc_005A951E:   fnclex
  loc_005A9520: jge 005A9534h
  loc_005A9522: push 000002B0h
  loc_005A9527: push 00434224h
loc_005A952C:   push esi
loc_005A952D:   push eax
  loc_005A952E: call [00401080h] ; __vbaHresultCheckObj
  loc_005A9534: mov var_4, 00000000h
loc_005A953B:   mov eax, Me
loc_005A953E:   push eax
loc_005A953F:   mov ecx, [eax]
loc_005A9541:   Call [ecx+00000008h]
loc_005A9544:   mov eax, var_4
loc_005A9547:   mov ecx, var_14
loc_005A954A:   pop edi
loc_005A954B:   pop esi
loc_005A954C:   mov fs: [00000000h] , ecx
loc_005A9553:   pop ebx
loc_005A9554:   mov esp, ebp
loc_005A9556:   pop ebp
  loc_005A9557: retn 0004h
End Sub

Private Sub txtKode_Change() '5AF270
loc_005AF270:   push ebp
loc_005AF271:   mov ebp, esp
  loc_005AF273: sub esp, 0000000Ch
  loc_005AF276: push 00405696h ; __vbaExceptHandler
loc_005AF27B:   mov eax, fs: [00000000h]
loc_005AF281:   push eax
loc_005AF282:   mov fs: [00000000h] , esp
  loc_005AF289: sub esp, 00000160h
loc_005AF28F:   push ebx
loc_005AF290:   push esi
loc_005AF291:   push edi
loc_005AF292:   mov var_C, esp
  loc_005AF295: mov var_8, 00402F78h
loc_005AF29C:   mov esi, Me
loc_005AF29F:   mov eax, esi
  loc_005AF2A1: and eax, 00000001h
loc_005AF2A4:   mov var_4, eax
  loc_005AF2A7: and esi, FFFFFFFEh
loc_005AF2AA:   push esi
loc_005AF2AB:   mov Me, esi
loc_005AF2AE:   mov ecx, [esi]
loc_005AF2B0:   Call [ecx+00000004h]
  loc_005AF2B3: xor ebx, ebx
  loc_005AF2B5: push 0042DF30h
loc_005AF2BA:   mov var_18, ebx
loc_005AF2BD:   mov var_1C, ebx
loc_005AF2C0:   mov var_20, ebx
loc_005AF2C3:   mov var_24, ebx
loc_005AF2C6:   mov var_28, ebx
loc_005AF2C9:   mov var_2C, ebx
loc_005AF2CC:   mov var_30, ebx
loc_005AF2CF:   mov var_34, ebx
loc_005AF2D2:   mov var_38, ebx
loc_005AF2D5:   mov var_3C, ebx
loc_005AF2D8:   mov var_40, ebx
loc_005AF2DB:   mov var_50, ebx
loc_005AF2DE:   mov var_60, ebx
loc_005AF2E1:   mov var_70, ebx
loc_005AF2E4:   mov var_80, ebx
loc_005AF2E7:   mov var_90, ebx
loc_005AF2ED:   mov var_A0, ebx
loc_005AF2F3:   mov var_B0, ebx
loc_005AF2F9:   mov var_C0, ebx
loc_005AF2FF:   mov var_D0, ebx
loc_005AF305:   mov var_F4, ebx
loc_005AF30B:   mov var_148, ebx
  loc_005AF311: call [00401168h] ; __vbaNew
  loc_005AF317: mov edi, [004010B8h] ; __vbaObjSet
loc_005AF31D:   push eax
  loc_005AF31E: push 00668050h
loc_005AF323:   Call edi
loc_005AF325:   mov edx, [esi]
loc_005AF327:   push esi
loc_005AF328:   Call [edx+00000348h]
loc_005AF32E:   push eax
loc_005AF32F:   lea eax, var_30
loc_005AF332:   push eax
loc_005AF333:   Call edi
loc_005AF335:   mov ecx, [eax]
loc_005AF337:   lea edx, var_18
loc_005AF33A:   push edx
loc_005AF33B:   push eax
loc_005AF33C:   mov var_118, eax
loc_005AF342:   Call [ecx+000000A0h]
loc_005AF348:   cmp eax, ebx
loc_005AF34A:   fnclex
  loc_005AF34C: jge 005AF366h
loc_005AF34E:   mov ecx, var_118
  loc_005AF354: push 000000A0h
  loc_005AF359: push 0042DFCCh
loc_005AF35E:   push ecx
loc_005AF35F:   push eax
  loc_005AF360: call [00401080h] ; __vbaHresultCheckObj
loc_005AF366:   cmp [0066803Ch], ebx
  loc_005AF36C: jnz 005AF37Eh
  loc_005AF36E: push 0066803Ch
  loc_005AF373: push 0042DEFCh
  loc_005AF378: call [004011E8h] ; __vbaNew2
loc_005AF37E:   mov eax, var_18
loc_005AF381:   mov edx, [0066803Ch]
  loc_005AF387: push 004348CCh ; "SELECT * FROM brg_barang WHERE kode_barang='"
loc_005AF38C:   push eax
loc_005AF38D:   mov var_98, edx
  loc_005AF393: mov var_A0, 00000009h
  loc_005AF39D: call [00401070h] ; __vbaStrCat
loc_005AF3A3:   mov edx, eax
loc_005AF3A5:   lea ecx, var_1C
  loc_005AF3A8: call [0040126Ch] ; __vbaStrMove
loc_005AF3AE:   push eax
  loc_005AF3AF: push 0042E0E0h ; "' "
  loc_005AF3B4: call [00401070h] ; __vbaStrCat
loc_005AF3BA:   mov ecx, [00668050h]
loc_005AF3C0:   push FFFFFFFFh
  loc_005AF3C2: push 00000004h
  loc_005AF3C4: push 00000002h
  loc_005AF3C6: sub esp, 00000010h
loc_005AF3C9:   mov var_48, eax
  loc_005AF3CC: mov var_50, 00000008h
loc_005AF3D3:   mov edx, [ecx]
loc_005AF3D5:   mov ecx, var_A0
loc_005AF3DB:   mov eax, esp
  loc_005AF3DD: sub esp, 00000010h
loc_005AF3E0:   mov [eax], ecx
loc_005AF3E2:   mov ecx, var_9C
loc_005AF3E8:   mov [eax+00000004h], ecx
loc_005AF3EB:   mov ecx, var_98
loc_005AF3F1:   mov [eax+00000008h], ecx
loc_005AF3F4:   mov ecx, var_94
loc_005AF3FA:   mov [eax+0000000Ch], ecx
loc_005AF3FD:   mov ecx, var_50
loc_005AF400:   mov eax, esp
loc_005AF402:   mov [eax], ecx
loc_005AF404:   mov ecx, var_4C
loc_005AF407:   mov [eax+00000004h], ecx
loc_005AF40A:   mov ecx, var_48
loc_005AF40D:   mov [eax+00000008h], ecx
loc_005AF410:   mov ecx, var_44
loc_005AF413:   mov [eax+0000000Ch], ecx
loc_005AF416:   mov eax, [00668050h]
loc_005AF41B:   push eax
loc_005AF41C:   Call [edx+000000A0h]
loc_005AF422:   cmp eax, ebx
loc_005AF424:   fnclex
  loc_005AF426: jge 005AF440h
loc_005AF428:   mov ecx, [00668050h]
  loc_005AF42E: push 000000A0h
  loc_005AF433: push 0042DF44h
loc_005AF438:   push ecx
loc_005AF439:   push eax
  loc_005AF43A: call [00401080h] ; __vbaHresultCheckObj
loc_005AF440:   lea edx, var_1C
loc_005AF443:   lea eax, var_18
loc_005AF446:   push edx
loc_005AF447:   push eax
  loc_005AF448: push 00000002h
  loc_005AF44A: call [00401204h] ; __vbaFreeStrList
  loc_005AF450: add esp, 0000000Ch
loc_005AF453:   lea ecx, var_30
  loc_005AF456: call [004012A0h] ; __vbaFreeObj
loc_005AF45C:   lea ecx, var_50
  loc_005AF45F: call [00401020h] ; __vbaFreeVar
loc_005AF465:   mov eax, [00668050h]
loc_005AF46A:   lea edx, var_F4
loc_005AF470:   push edx
loc_005AF471:   push eax
loc_005AF472:   mov ecx, [eax]
loc_005AF474:   Call [ecx+00000034h]
loc_005AF477:   cmp eax, ebx
loc_005AF479:   fnclex
  loc_005AF47B: jge 005AF492h
loc_005AF47D:   mov ecx, [00668050h]
  loc_005AF483: push 00000034h
  loc_005AF485: push 0042DF44h
loc_005AF48A:   push ecx
loc_005AF48B:   push eax
  loc_005AF48C: call [00401080h] ; __vbaHresultCheckObj
loc_005AF492:   cmp var_F4, bx
  loc_005AF499: jnz 005B0E4Dh
loc_005AF49F:   mov edx, [00668050h]
loc_005AF4A5:   lea eax, var_148
loc_005AF4AB:   push edx
loc_005AF4AC:   push eax
  loc_005AF4AD: call [004010C8h] ; __vbaObjSetAddref
loc_005AF4B3:   mov ecx, [esi]
loc_005AF4B5:   push esi
loc_005AF4B6:   Call [ecx+00000344h]
loc_005AF4BC:   lea edx, var_40
loc_005AF4BF:   push eax
loc_005AF4C0:   push edx
loc_005AF4C1:   Call edi
loc_005AF4C3:   mov var_140, eax
loc_005AF4C9:   mov eax, var_148
loc_005AF4CF:   lea edx, var_30
loc_005AF4D2:   mov ecx, [eax]
loc_005AF4D4:   push edx
loc_005AF4D5:   push eax
loc_005AF4D6:   Call [ecx+00000054h]
loc_005AF4D9:   test eax, eax
loc_005AF4DB:   fnclex
  loc_005AF4DD: jge 005AF4F4h
loc_005AF4DF:   mov ecx, var_148
  loc_005AF4E5: push 00000054h
  loc_005AF4E7: push 0042DF44h
loc_005AF4EC:   push ecx
loc_005AF4ED:   push eax
  loc_005AF4EE: call [00401080h] ; __vbaHresultCheckObj
loc_005AF4F4:   lea ebx, var_34
loc_005AF4F7:   mov eax, var_30
loc_005AF4FA:   push ebx
  loc_005AF4FB: mov ecx, 00000008h
  loc_005AF500: sub esp, 00000010h
loc_005AF503:   mov var_A0, ecx
loc_005AF509:   mov ebx, esp
  loc_005AF50B: mov var_98, 00431ED0h ; "nama_barang"
loc_005AF515:   mov edx, [eax]
loc_005AF517:   push eax
loc_005AF518:   mov [ebx], ecx
loc_005AF51A:   mov ecx, var_9C
loc_005AF520:   mov var_11C, eax
loc_005AF526:   mov [ebx+00000004h], ecx
loc_005AF529:   mov ecx, var_98
loc_005AF52F:   mov [ebx+00000008h], ecx
loc_005AF532:   mov ecx, var_94
loc_005AF538:   mov [ebx+0000000Ch], ecx
loc_005AF53B:   Call [edx+00000028h]
loc_005AF53E:   test eax, eax
loc_005AF540:   fnclex
  loc_005AF542: jge 005AF559h
loc_005AF544:   mov edx, var_11C
  loc_005AF54A: push 00000028h
  loc_005AF54C: push 0042DFACh
loc_005AF551:   push edx
loc_005AF552:   push eax
  loc_005AF553: call [00401080h] ; __vbaHresultCheckObj
loc_005AF559:   mov eax, var_34
loc_005AF55C:   lea edx, var_50
loc_005AF55F:   push edx
loc_005AF560:   push eax
loc_005AF561:   mov ecx, [eax]
loc_005AF563:   mov ebx, eax
loc_005AF565:   Call [ecx+00000034h]
loc_005AF568:   test eax, eax
loc_005AF56A:   fnclex
  loc_005AF56C: jge 005AF57Dh
  loc_005AF56E: push 00000034h
  loc_005AF570: push 0042DFBCh
loc_005AF575:   push ebx
loc_005AF576:   push eax
  loc_005AF577: call [00401080h] ; __vbaHresultCheckObj
loc_005AF57D:   mov eax, var_148
loc_005AF583:   lea edx, var_38
  loc_005AF586: mov var_A8, 0043492Ch ; " ( "
  loc_005AF590: mov var_B0, 00000008h
loc_005AF59A:   mov ecx, [eax]
loc_005AF59C:   push edx
loc_005AF59D:   push eax
loc_005AF59E:   Call [ecx+00000054h]
loc_005AF5A1:   test eax, eax
loc_005AF5A3:   fnclex
  loc_005AF5A5: jge 005AF5BCh
loc_005AF5A7:   mov ecx, var_148
  loc_005AF5AD: push 00000054h
  loc_005AF5AF: push 0042DF44h
loc_005AF5B4:   push ecx
loc_005AF5B5:   push eax
  loc_005AF5B6: call [00401080h] ; __vbaHresultCheckObj
loc_005AF5BC:   lea ebx, var_3C
loc_005AF5BF:   mov eax, var_38
loc_005AF5C2:   push ebx
  loc_005AF5C3: mov ecx, 00000008h
  loc_005AF5C8: sub esp, 00000010h
loc_005AF5CB:   mov edx, [eax]
loc_005AF5CD:   mov ebx, esp
loc_005AF5CF:   mov var_130, eax
loc_005AF5D5:   push eax
loc_005AF5D6:   mov [ebx], ecx
loc_005AF5D8:   mov ecx, var_BC
loc_005AF5DE:   mov [ebx+00000004h], ecx
  loc_005AF5E1: mov ecx, 00430B44h ; "satuan"
loc_005AF5E6:   mov [ebx+00000008h], ecx
loc_005AF5E9:   mov ecx, var_B4
loc_005AF5EF:   mov [ebx+0000000Ch], ecx
loc_005AF5F2:   Call [edx+00000028h]
loc_005AF5F5:   test eax, eax
loc_005AF5F7:   fnclex
  loc_005AF5F9: jge 005AF610h
loc_005AF5FB:   mov edx, var_130
  loc_005AF601: push 00000028h
  loc_005AF603: push 0042DFACh
loc_005AF608:   push edx
loc_005AF609:   push eax
  loc_005AF60A: call [00401080h] ; __vbaHresultCheckObj
loc_005AF610:   mov eax, var_3C
loc_005AF613:   lea edx, var_70
loc_005AF616:   push edx
loc_005AF617:   push eax
loc_005AF618:   mov ecx, [eax]
loc_005AF61A:   mov ebx, eax
loc_005AF61C:   Call [ecx+00000034h]
loc_005AF61F:   test eax, eax
loc_005AF621:   fnclex
  loc_005AF623: jge 005AF634h
  loc_005AF625: push 00000034h
  loc_005AF627: push 0042DFBCh
loc_005AF62C:   push ebx
loc_005AF62D:   push eax
  loc_005AF62E: call [00401080h] ; __vbaHresultCheckObj
loc_005AF634:   mov eax, var_140
  loc_005AF63A: mov var_C8, 00434938h ; " )"
  loc_005AF644: mov var_D0, 00000008h
loc_005AF64E:   lea ecx, var_50
loc_005AF651:   mov ebx, [eax]
loc_005AF653:   lea edx, var_B0
loc_005AF659:   push ecx
loc_005AF65A:   lea eax, var_60
loc_005AF65D:   push edx
loc_005AF65E:   push eax
  loc_005AF65F: call [004011C4h] ; __vbaVarCat
loc_005AF665:   lea ecx, var_70
loc_005AF668:   push eax
loc_005AF669:   lea edx, var_80
loc_005AF66C:   push ecx
loc_005AF66D:   push edx
  loc_005AF66E: call [004011C4h] ; __vbaVarCat
loc_005AF674:   push eax
loc_005AF675:   lea eax, var_D0
loc_005AF67B:   lea ecx, var_90
loc_005AF681:   push eax
loc_005AF682:   push ecx
  loc_005AF683: call [004011C4h] ; __vbaVarCat
loc_005AF689:   lea edx, var_18
loc_005AF68C:   push eax
loc_005AF68D:   push edx
  loc_005AF68E: call [004011C0h] ; __vbaStrVarVal
loc_005AF694:   mov var_168, ebx
loc_005AF69A:   mov ebx, var_140
loc_005AF6A0:   push eax
loc_005AF6A1:   mov eax, var_168
loc_005AF6A7:   push ebx
loc_005AF6A8:   Call [eax+000000A4h]
loc_005AF6AE:   test eax, eax
loc_005AF6B0:   fnclex
  loc_005AF6B2: jge 005AF6C6h
  loc_005AF6B4: push 000000A4h
  loc_005AF6B9: push 0042DFCCh
loc_005AF6BE:   push ebx
loc_005AF6BF:   push eax
  loc_005AF6C0: call [00401080h] ; __vbaHresultCheckObj
loc_005AF6C6:   lea ecx, var_18
  loc_005AF6C9: call [0040129Ch] ; __vbaFreeStr
loc_005AF6CF:   lea ecx, var_40
loc_005AF6D2:   lea edx, var_3C
loc_005AF6D5:   push ecx
loc_005AF6D6:   lea eax, var_38
loc_005AF6D9:   push edx
loc_005AF6DA:   lea ecx, var_34
loc_005AF6DD:   push eax
loc_005AF6DE:   lea edx, var_30
loc_005AF6E1:   push ecx
loc_005AF6E2:   push edx
  loc_005AF6E3: push 00000005h
  loc_005AF6E5: call [0040104Ch] ; __vbaFreeObjList
loc_005AF6EB:   lea eax, var_90
loc_005AF6F1:   lea ecx, var_80
loc_005AF6F4:   push eax
loc_005AF6F5:   lea edx, var_70
loc_005AF6F8:   push ecx
loc_005AF6F9:   lea eax, var_60
loc_005AF6FC:   push edx
loc_005AF6FD:   lea ecx, var_50
loc_005AF700:   push eax
loc_005AF701:   push ecx
  loc_005AF702: push 00000005h
  loc_005AF704: call [0040103Ch] ; __vbaFreeVarList
loc_005AF70A:   mov edx, [esi]
  loc_005AF70C: add esp, 00000030h
loc_005AF70F:   push esi
loc_005AF710:   Call [edx+00000330h]
loc_005AF716:   push eax
loc_005AF717:   lea eax, var_38
loc_005AF71A:   push eax
loc_005AF71B:   Call edi
loc_005AF71D:   mov var_12C, eax
loc_005AF723:   mov eax, var_148
loc_005AF729:   lea edx, var_30
loc_005AF72C:   mov ecx, [eax]
loc_005AF72E:   push edx
loc_005AF72F:   push eax
loc_005AF730:   Call [ecx+00000054h]
loc_005AF733:   test eax, eax
loc_005AF735:   fnclex
  loc_005AF737: jge 005AF74Eh
loc_005AF739:   mov ecx, var_148
  loc_005AF73F: push 00000054h
  loc_005AF741: push 0042DF44h
loc_005AF746:   push ecx
loc_005AF747:   push eax
  loc_005AF748: call [00401080h] ; __vbaHresultCheckObj
loc_005AF74E:   lea ebx, var_34
loc_005AF751:   mov eax, var_30
loc_005AF754:   push ebx
  loc_005AF755: mov ecx, 00000008h
  loc_005AF75A: sub esp, 00000010h
loc_005AF75D:   mov var_A0, ecx
loc_005AF763:   mov ebx, esp
  loc_005AF765: mov var_98, 0043317Ch ; "Diskon"
loc_005AF76F:   mov edx, [eax]
loc_005AF771:   push eax
loc_005AF772:   mov [ebx], ecx
loc_005AF774:   mov ecx, var_9C
loc_005AF77A:   mov var_11C, eax
loc_005AF780:   mov [ebx+00000004h], ecx
loc_005AF783:   mov ecx, var_98
loc_005AF789:   mov [ebx+00000008h], ecx
loc_005AF78C:   mov ecx, var_94
loc_005AF792:   mov [ebx+0000000Ch], ecx
loc_005AF795:   Call [edx+00000028h]
loc_005AF798:   test eax, eax
loc_005AF79A:   fnclex
  loc_005AF79C: jge 005AF7B3h
loc_005AF79E:   mov edx, var_11C
  loc_005AF7A4: push 00000028h
  loc_005AF7A6: push 0042DFACh
loc_005AF7AB:   push edx
loc_005AF7AC:   push eax
  loc_005AF7AD: call [00401080h] ; __vbaHresultCheckObj
loc_005AF7B3:   mov eax, var_34
loc_005AF7B6:   lea edx, var_50
loc_005AF7B9:   push edx
loc_005AF7BA:   push eax
loc_005AF7BB:   mov ecx, [eax]
loc_005AF7BD:   mov ebx, eax
loc_005AF7BF:   Call [ecx+00000034h]
loc_005AF7C2:   test eax, eax
loc_005AF7C4:   fnclex
  loc_005AF7C6: jge 005AF7D7h
  loc_005AF7C8: push 00000034h
  loc_005AF7CA: push 0042DFBCh
loc_005AF7CF:   push ebx
loc_005AF7D0:   push eax
  loc_005AF7D1: call [00401080h] ; __vbaHresultCheckObj
loc_005AF7D7:   mov eax, var_12C
loc_005AF7DD:   lea ecx, var_50
loc_005AF7E0:   push ecx
loc_005AF7E1:   mov ebx, [eax]
  loc_005AF7E3: call [00401034h] ; __vbaStrVarMove
loc_005AF7E9:   mov edx, eax
loc_005AF7EB:   lea ecx, var_18
  loc_005AF7EE: call [0040126Ch] ; __vbaStrMove
loc_005AF7F4:   mov edx, ebx
loc_005AF7F6:   mov ebx, var_12C
loc_005AF7FC:   push eax
loc_005AF7FD:   push ebx
loc_005AF7FE:   Call [edx+000000A4h]
loc_005AF804:   test eax, eax
loc_005AF806:   fnclex
  loc_005AF808: jge 005AF81Ch
  loc_005AF80A: push 000000A4h
  loc_005AF80F: push 0042DFCCh
loc_005AF814:   push ebx
loc_005AF815:   push eax
  loc_005AF816: call [00401080h] ; __vbaHresultCheckObj
loc_005AF81C:   lea ecx, var_18
  loc_005AF81F: call [0040129Ch] ; __vbaFreeStr
loc_005AF825:   lea eax, var_38
loc_005AF828:   lea ecx, var_34
loc_005AF82B:   push eax
loc_005AF82C:   lea edx, var_30
loc_005AF82F:   push ecx
loc_005AF830:   push edx
  loc_005AF831: push 00000003h
  loc_005AF833: call [0040104Ch] ; __vbaFreeObjList
  loc_005AF839: add esp, 00000010h
loc_005AF83C:   lea ecx, var_50
  loc_005AF83F: call [00401020h] ; __vbaFreeVar
loc_005AF845:   mov eax, [esi]
loc_005AF847:   push esi
loc_005AF848:   Call [eax+00000340h]
loc_005AF84E:   lea ecx, var_38
loc_005AF851:   push eax
loc_005AF852:   push ecx
loc_005AF853:   Call edi
loc_005AF855:   mov var_12C, eax
loc_005AF85B:   mov eax, var_148
loc_005AF861:   lea ecx, var_30
loc_005AF864:   mov edx, [eax]
loc_005AF866:   push ecx
loc_005AF867:   push eax
loc_005AF868:   Call [edx+00000054h]
loc_005AF86B:   test eax, eax
loc_005AF86D:   fnclex
  loc_005AF86F: jge 005AF886h
loc_005AF871:   mov edx, var_148
  loc_005AF877: push 00000054h
  loc_005AF879: push 0042DF44h
loc_005AF87E:   push edx
loc_005AF87F:   push eax
  loc_005AF880: call [00401080h] ; __vbaHresultCheckObj
loc_005AF886:   lea ebx, var_34
loc_005AF889:   mov eax, var_30
loc_005AF88C:   push ebx
  loc_005AF88D: mov ecx, 00000008h
  loc_005AF892: sub esp, 00000010h
loc_005AF895:   mov var_A0, ecx
loc_005AF89B:   mov ebx, esp
  loc_005AF89D: mov var_98, 00431F08h ; "harga_jual"
loc_005AF8A7:   mov edx, [eax]
loc_005AF8A9:   push eax
loc_005AF8AA:   mov [ebx], ecx
loc_005AF8AC:   mov ecx, var_9C
loc_005AF8B2:   mov var_11C, eax
loc_005AF8B8:   mov [ebx+00000004h], ecx
loc_005AF8BB:   mov ecx, var_98
loc_005AF8C1:   mov [ebx+00000008h], ecx
loc_005AF8C4:   mov ecx, var_94
loc_005AF8CA:   mov [ebx+0000000Ch], ecx
loc_005AF8CD:   Call [edx+00000028h]
loc_005AF8D0:   test eax, eax
loc_005AF8D2:   fnclex
  loc_005AF8D4: jge 005AF8EBh
loc_005AF8D6:   mov edx, var_11C
  loc_005AF8DC: push 00000028h
  loc_005AF8DE: push 0042DFACh
loc_005AF8E3:   push edx
loc_005AF8E4:   push eax
  loc_005AF8E5: call [00401080h] ; __vbaHresultCheckObj
loc_005AF8EB:   mov eax, var_34
loc_005AF8EE:   lea edx, var_50
loc_005AF8F1:   push edx
loc_005AF8F2:   push eax
loc_005AF8F3:   mov ecx, [eax]
loc_005AF8F5:   mov ebx, eax
loc_005AF8F7:   Call [ecx+00000034h]
loc_005AF8FA:   test eax, eax
loc_005AF8FC:   fnclex
  loc_005AF8FE: jge 005AF90Fh
  loc_005AF900: push 00000034h
  loc_005AF902: push 0042DFBCh
loc_005AF907:   push ebx
loc_005AF908:   push eax
  loc_005AF909: call [00401080h] ; __vbaHresultCheckObj
loc_005AF90F:   mov eax, var_12C
loc_005AF915:   lea ecx, var_50
loc_005AF918:   push ecx
loc_005AF919:   mov ebx, [eax]
  loc_005AF91B: call [00401034h] ; __vbaStrVarMove
loc_005AF921:   mov edx, eax
loc_005AF923:   lea ecx, var_18
  loc_005AF926: call [0040126Ch] ; __vbaStrMove
loc_005AF92C:   mov edx, ebx
loc_005AF92E:   mov ebx, var_12C
loc_005AF934:   push eax
loc_005AF935:   push ebx
loc_005AF936:   Call [edx+000000A4h]
loc_005AF93C:   test eax, eax
loc_005AF93E:   fnclex
  loc_005AF940: jge 005AF954h
  loc_005AF942: push 000000A4h
  loc_005AF947: push 0042DFCCh
loc_005AF94C:   push ebx
loc_005AF94D:   push eax
  loc_005AF94E: call [00401080h] ; __vbaHresultCheckObj
loc_005AF954:   lea ecx, var_18
  loc_005AF957: call [0040129Ch] ; __vbaFreeStr
loc_005AF95D:   lea eax, var_38
loc_005AF960:   lea ecx, var_34
loc_005AF963:   push eax
loc_005AF964:   lea edx, var_30
loc_005AF967:   push ecx
loc_005AF968:   push edx
  loc_005AF969: push 00000003h
  loc_005AF96B: call [0040104Ch] ; __vbaFreeObjList
  loc_005AF971: add esp, 00000010h
loc_005AF974:   lea ecx, var_50
  loc_005AF977: call [00401020h] ; __vbaFreeVar
loc_005AF97D:   mov eax, [esi]
loc_005AF97F:   push esi
loc_005AF980:   Call [eax+00000304h]
loc_005AF986:   lea ecx, var_38
loc_005AF989:   push eax
loc_005AF98A:   push ecx
loc_005AF98B:   Call edi
loc_005AF98D:   mov var_12C, eax
loc_005AF993:   mov eax, var_148
loc_005AF999:   lea ecx, var_30
loc_005AF99C:   mov edx, [eax]
loc_005AF99E:   push ecx
loc_005AF99F:   push eax
loc_005AF9A0:   Call [edx+00000054h]
loc_005AF9A3:   test eax, eax
loc_005AF9A5:   fnclex
  loc_005AF9A7: jge 005AF9BEh
loc_005AF9A9:   mov edx, var_148
  loc_005AF9AF: push 00000054h
  loc_005AF9B1: push 0042DF44h
loc_005AF9B6:   push edx
loc_005AF9B7:   push eax
  loc_005AF9B8: call [00401080h] ; __vbaHresultCheckObj
loc_005AF9BE:   lea ebx, var_34
loc_005AF9C1:   mov eax, var_30
loc_005AF9C4:   push ebx
  loc_005AF9C5: mov ecx, 00000008h
  loc_005AF9CA: sub esp, 00000010h
loc_005AF9CD:   mov var_A0, ecx
loc_005AF9D3:   mov ebx, esp
  loc_005AF9D5: mov var_98, 00431EECh ; "harga_beli"
loc_005AF9DF:   mov edx, [eax]
loc_005AF9E1:   push eax
loc_005AF9E2:   mov [ebx], ecx
loc_005AF9E4:   mov ecx, var_9C
loc_005AF9EA:   mov var_11C, eax
loc_005AF9F0:   mov [ebx+00000004h], ecx
loc_005AF9F3:   mov ecx, var_98
loc_005AF9F9:   mov [ebx+00000008h], ecx
loc_005AF9FC:   mov ecx, var_94
loc_005AFA02:   mov [ebx+0000000Ch], ecx
loc_005AFA05:   Call [edx+00000028h]
loc_005AFA08:   test eax, eax
loc_005AFA0A:   fnclex
  loc_005AFA0C: jge 005AFA23h
loc_005AFA0E:   mov edx, var_11C
  loc_005AFA14: push 00000028h
  loc_005AFA16: push 0042DFACh
loc_005AFA1B:   push edx
loc_005AFA1C:   push eax
  loc_005AFA1D: call [00401080h] ; __vbaHresultCheckObj
loc_005AFA23:   mov eax, var_34
loc_005AFA26:   lea edx, var_50
loc_005AFA29:   push edx
loc_005AFA2A:   push eax
loc_005AFA2B:   mov ecx, [eax]
loc_005AFA2D:   mov ebx, eax
loc_005AFA2F:   Call [ecx+00000034h]
loc_005AFA32:   test eax, eax
loc_005AFA34:   fnclex
  loc_005AFA36: jge 005AFA47h
  loc_005AFA38: push 00000034h
  loc_005AFA3A: push 0042DFBCh
loc_005AFA3F:   push ebx
loc_005AFA40:   push eax
  loc_005AFA41: call [00401080h] ; __vbaHresultCheckObj
loc_005AFA47:   mov eax, var_12C
loc_005AFA4D:   lea ecx, var_50
loc_005AFA50:   push ecx
loc_005AFA51:   mov ebx, [eax]
  loc_005AFA53: call [00401034h] ; __vbaStrVarMove
loc_005AFA59:   mov edx, eax
loc_005AFA5B:   lea ecx, var_18
  loc_005AFA5E: call [0040126Ch] ; __vbaStrMove
loc_005AFA64:   mov edx, ebx
loc_005AFA66:   mov ebx, var_12C
loc_005AFA6C:   push eax
loc_005AFA6D:   push ebx
loc_005AFA6E:   Call [edx+000000A4h]
loc_005AFA74:   test eax, eax
loc_005AFA76:   fnclex
  loc_005AFA78: jge 005AFA8Ch
  loc_005AFA7A: push 000000A4h
  loc_005AFA7F: push 0042DFCCh
loc_005AFA84:   push ebx
loc_005AFA85:   push eax
  loc_005AFA86: call [00401080h] ; __vbaHresultCheckObj
loc_005AFA8C:   lea ecx, var_18
  loc_005AFA8F: call [0040129Ch] ; __vbaFreeStr
loc_005AFA95:   lea eax, var_38
loc_005AFA98:   lea ecx, var_34
loc_005AFA9B:   push eax
loc_005AFA9C:   lea edx, var_30
loc_005AFA9F:   push ecx
loc_005AFAA0:   push edx
  loc_005AFAA1: push 00000003h
  loc_005AFAA3: call [0040104Ch] ; __vbaFreeObjList
  loc_005AFAA9: add esp, 00000010h
loc_005AFAAC:   lea ecx, var_50
  loc_005AFAAF: call [00401020h] ; __vbaFreeVar
loc_005AFAB5:   lea eax, var_148
  loc_005AFABB: push 00000000h
loc_005AFABD:   push eax
  loc_005AFABE: call [004010C8h] ; __vbaObjSetAddref
  loc_005AFAC4: push 0042DF30h
  loc_005AFAC9: call [00401168h] ; __vbaNew
loc_005AFACF:   push eax
  loc_005AFAD0: push 00668050h
loc_005AFAD5:   Call edi
loc_005AFAD7:   mov ecx, [esi]
loc_005AFAD9:   push esi
loc_005AFADA:   Call [ecx+00000348h]
loc_005AFAE0:   lea edx, var_30
loc_005AFAE3:   push eax
loc_005AFAE4:   push edx
loc_005AFAE5:   Call edi
loc_005AFAE7:   mov ebx, eax
loc_005AFAE9:   lea ecx, var_24
loc_005AFAEC:   push ecx
loc_005AFAED:   push ebx
loc_005AFAEE:   mov eax, [ebx]
loc_005AFAF0:   Call [eax+000000A0h]
loc_005AFAF6:   test eax, eax
loc_005AFAF8:   fnclex
  loc_005AFAFA: jge 005AFB0Eh
  loc_005AFAFC: push 000000A0h
  loc_005AFB01: push 0042DFCCh
loc_005AFB06:   push ebx
loc_005AFB07:   push eax
  loc_005AFB08: call [00401080h] ; __vbaHresultCheckObj
loc_005AFB0E:   mov eax, [0066803Ch]
loc_005AFB13:   test eax, eax
  loc_005AFB15: jnz 005AFB27h
  loc_005AFB17: push 0066803Ch
  loc_005AFB1C: push 0042DEFCh
  loc_005AFB21: call [004011E8h] ; __vbaNew2
loc_005AFB27:   mov edx, [0066803Ch]
  loc_005AFB2D: mov ebx, [00401070h] ; __vbaStrCat
  loc_005AFB33: push 004330A0h ; "SELECT brg_barang.*, brg_jenis.* "
  loc_005AFB38: push 00436144h ; " FROM brg_jenis, brg_produk, brg_barang WHERE"
loc_005AFB3D:   mov var_98, edx
  loc_005AFB43: mov var_A0, 00000009h
loc_005AFB4D:   Call ebx
loc_005AFB4F:   mov edx, eax
loc_005AFB51:   lea ecx, var_18
  loc_005AFB54: call [0040126Ch] ; __vbaStrMove
loc_005AFB5A:   push eax
  loc_005AFB5B: push 004361B4h ; " brg_jenis.Kode_Jenis=brg_produk.Kode_Jenis AND"
loc_005AFB60:   Call ebx
loc_005AFB62:   mov edx, eax
loc_005AFB64:   lea ecx, var_1C
  loc_005AFB67: call [0040126Ch] ; __vbaStrMove
loc_005AFB6D:   push eax
  loc_005AFB6E: push 00436218h ; " brg_produk.Kode_Produk=brg_barang.Kode_Produk "
loc_005AFB73:   Call ebx
loc_005AFB75:   mov edx, eax
loc_005AFB77:   lea ecx, var_20
  loc_005AFB7A: call [0040126Ch] ; __vbaStrMove
loc_005AFB80:   push eax
  loc_005AFB81: push 00432DACh ; " AND brg_barang.kode_barang='"
loc_005AFB86:   Call ebx
loc_005AFB88:   mov edx, eax
loc_005AFB8A:   lea ecx, var_28
  loc_005AFB8D: call [0040126Ch] ; __vbaStrMove
loc_005AFB93:   push eax
loc_005AFB94:   mov eax, var_24
loc_005AFB97:   push eax
loc_005AFB98:   Call ebx
loc_005AFB9A:   mov edx, eax
loc_005AFB9C:   lea ecx, var_2C
  loc_005AFB9F: call [0040126Ch] ; __vbaStrMove
loc_005AFBA5:   push eax
  loc_005AFBA6: push 0042EBA8h ; "'"
loc_005AFBAB:   Call ebx
loc_005AFBAD:   mov ecx, [00668050h]
loc_005AFBB3:   mov ebx, var_A0
loc_005AFBB9:   push FFFFFFFFh
  loc_005AFBBB: push 00000003h
  loc_005AFBBD: push 00000002h
loc_005AFBBF:   mov var_48, eax
  loc_005AFBC2: mov var_50, 00000008h
loc_005AFBC9:   mov edx, [ecx]
  loc_005AFBCB: sub esp, 00000010h
loc_005AFBCE:   mov ecx, esp
  loc_005AFBD0: sub esp, 00000010h
loc_005AFBD3:   mov [ecx], ebx
loc_005AFBD5:   mov ebx, var_9C
loc_005AFBDB:   mov [ecx+00000004h], ebx
loc_005AFBDE:   mov ebx, var_98
loc_005AFBE4:   mov [ecx+00000008h], ebx
loc_005AFBE7:   mov ebx, var_94
loc_005AFBED:   mov [ecx+0000000Ch], ebx
loc_005AFBF0:   mov ebx, var_50
loc_005AFBF3:   mov ecx, esp
loc_005AFBF5:   mov [ecx], ebx
loc_005AFBF7:   mov ebx, var_4C
loc_005AFBFA:   mov [ecx+00000004h], ebx
loc_005AFBFD:   mov [ecx+00000008h], eax
loc_005AFC00:   mov eax, var_44
loc_005AFC03:   mov [ecx+0000000Ch], eax
loc_005AFC06:   mov ecx, [00668050h]
loc_005AFC0C:   push ecx
loc_005AFC0D:   Call [edx+000000A0h]
loc_005AFC13:   test eax, eax
loc_005AFC15:   fnclex
  loc_005AFC17: jge 005AFC31h
loc_005AFC19:   mov edx, [00668050h]
  loc_005AFC1F: push 000000A0h
  loc_005AFC24: push 0042DF44h
loc_005AFC29:   push edx
loc_005AFC2A:   push eax
  loc_005AFC2B: call [00401080h] ; __vbaHresultCheckObj
loc_005AFC31:   lea eax, var_2C
loc_005AFC34:   lea ecx, var_24
loc_005AFC37:   push eax
loc_005AFC38:   lea edx, var_28
loc_005AFC3B:   push ecx
loc_005AFC3C:   lea eax, var_20
loc_005AFC3F:   push edx
loc_005AFC40:   lea ecx, var_1C
loc_005AFC43:   push eax
loc_005AFC44:   lea edx, var_18
loc_005AFC47:   push ecx
loc_005AFC48:   push edx
  loc_005AFC49: push 00000006h
  loc_005AFC4B: call [00401204h] ; __vbaFreeStrList
  loc_005AFC51: add esp, 0000001Ch
loc_005AFC54:   lea ecx, var_30
  loc_005AFC57: call [004012A0h] ; __vbaFreeObj
loc_005AFC5D:   lea ecx, var_50
  loc_005AFC60: call [00401020h] ; __vbaFreeVar
loc_005AFC66:   mov eax, [esi]
loc_005AFC68:   push esi
loc_005AFC69:   Call [eax+00000348h]
loc_005AFC6F:   lea ecx, var_30
loc_005AFC72:   push eax
loc_005AFC73:   push ecx
loc_005AFC74:   Call edi
loc_005AFC76:   mov ebx, eax
loc_005AFC78:   lea eax, var_18
loc_005AFC7B:   push eax
loc_005AFC7C:   push ebx
loc_005AFC7D:   mov edx, [ebx]
loc_005AFC7F:   Call [edx+000000A0h]
loc_005AFC85:   test eax, eax
loc_005AFC87:   fnclex
  loc_005AFC89: jge 005AFC9Dh
  loc_005AFC8B: push 000000A0h
  loc_005AFC90: push 0042DFCCh
loc_005AFC95:   push ebx
loc_005AFC96:   push eax
  loc_005AFC97: call [00401080h] ; __vbaHresultCheckObj
loc_005AFC9D:   mov ecx, var_18
loc_005AFCA0:   push ecx
  loc_005AFCA1: push 0042DFECh
  loc_005AFCA6: call [0040112Ch] ; __vbaStrCmp
loc_005AFCAC:   mov ebx, eax
loc_005AFCAE:   lea ecx, var_18
loc_005AFCB1:   neg ebx
loc_005AFCB3:   sbb ebx, ebx
loc_005AFCB5:   inc ebx
loc_005AFCB6:   neg ebx
  loc_005AFCB8: call [0040129Ch] ; __vbaFreeStr
loc_005AFCBE:   lea ecx, var_30
  loc_005AFCC1: call [004012A0h] ; __vbaFreeObj
loc_005AFCC7:   test bx, bx
  loc_005AFCCA: jz 005AFD8Ch
  loc_005AFCD0: mov ebx, [00401238h] ; __vbaVarDup
  loc_005AFCD6: mov ecx, 80020004h
loc_005AFCDB:   mov var_78, ecx
  loc_005AFCDE: mov eax, 0000000Ah
loc_005AFCE3:   mov var_68, ecx
loc_005AFCE6:   lea edx, var_B0
loc_005AFCEC:   lea ecx, var_60
loc_005AFCEF:   mov var_80, eax
loc_005AFCF2:   mov var_70, eax
  loc_005AFCF5: mov var_A8, 0042E624h ; "Error"
  loc_005AFCFF: mov var_B0, 00000008h
loc_005AFD09:   Call ebx
loc_005AFD0B:   lea edx, var_A0
loc_005AFD11:   lea ecx, var_50
  loc_005AFD14: mov var_98, 0043627Ch ; "Kode Barang kosong!"
  loc_005AFD1E: mov var_A0, 00000008h
loc_005AFD28:   Call ebx
loc_005AFD2A:   lea edx, var_80
loc_005AFD2D:   lea eax, var_70
loc_005AFD30:   push edx
loc_005AFD31:   lea ecx, var_60
loc_005AFD34:   push eax
loc_005AFD35:   push ecx
loc_005AFD36:   lea edx, var_50
  loc_005AFD39: push 00000010h
loc_005AFD3B:   push edx
  loc_005AFD3C: call [004010B4h] ; rtcMsgBox
loc_005AFD42:   lea eax, var_80
loc_005AFD45:   lea ecx, var_70
loc_005AFD48:   push eax
loc_005AFD49:   lea edx, var_60
loc_005AFD4C:   push ecx
loc_005AFD4D:   lea eax, var_50
loc_005AFD50:   push edx
loc_005AFD51:   push eax
  loc_005AFD52: push 00000004h
  loc_005AFD54: call [0040103Ch] ; __vbaFreeVarList
loc_005AFD5A:   mov ecx, [esi]
  loc_005AFD5C: add esp, 00000014h
  loc_005AFD5F: push 00000000h
  loc_005AFD61: push 80011000h
loc_005AFD66:   push esi
loc_005AFD67:   Call [ecx+00000380h]
loc_005AFD6D:   lea edx, var_30
loc_005AFD70:   push eax
loc_005AFD71:   push edx
loc_005AFD72:   Call edi
loc_005AFD74:   push eax
  loc_005AFD75: call [0040102Ch] ; __vbaLateIdCall
  loc_005AFD7B: add esp, 0000000Ch
loc_005AFD7E:   lea ecx, var_30
  loc_005AFD81: call [004012A0h] ; __vbaFreeObj
  loc_005AFD87: jmp 005B0E4Dh
loc_005AFD8C:   mov eax, [esi]
loc_005AFD8E:   push esi
loc_005AFD8F:   Call [eax+0000033Ch]
loc_005AFD95:   lea ecx, var_30
loc_005AFD98:   push eax
loc_005AFD99:   push ecx
loc_005AFD9A:   Call edi
loc_005AFD9C:   mov ebx, eax
loc_005AFD9E:   lea eax, var_18
loc_005AFDA1:   push eax
loc_005AFDA2:   push ebx
loc_005AFDA3:   mov edx, [ebx]
loc_005AFDA5:   Call [edx+000000A0h]
loc_005AFDAB:   test eax, eax
loc_005AFDAD:   fnclex
  loc_005AFDAF: jge 005AFDC3h
  loc_005AFDB1: push 000000A0h
  loc_005AFDB6: push 0042DFCCh
loc_005AFDBB:   push ebx
loc_005AFDBC:   push eax
  loc_005AFDBD: call [00401080h] ; __vbaHresultCheckObj
loc_005AFDC3:   mov ecx, [esi]
loc_005AFDC5:   push esi
loc_005AFDC6:   Call [ecx+0000033Ch]
loc_005AFDCC:   lea edx, var_34
loc_005AFDCF:   push eax
loc_005AFDD0:   push edx
loc_005AFDD1:   Call edi
loc_005AFDD3:   mov ebx, eax
loc_005AFDD5:   lea ecx, var_1C
loc_005AFDD8:   push ecx
loc_005AFDD9:   push ebx
loc_005AFDDA:   mov eax, [ebx]
loc_005AFDDC:   Call [eax+000000A0h]
loc_005AFDE2:   test eax, eax
loc_005AFDE4:   fnclex
  loc_005AFDE6: jge 005AFDFAh
  loc_005AFDE8: push 000000A0h
  loc_005AFDED: push 0042DFCCh
loc_005AFDF2:   push ebx
loc_005AFDF3:   push eax
  loc_005AFDF4: call [00401080h] ; __vbaHresultCheckObj
loc_005AFDFA:   mov edx, var_1C
loc_005AFDFD:   push edx
  loc_005AFDFE: push 0042DFECh
  loc_005AFE03: call [0040112Ch] ; __vbaStrCmp
loc_005AFE09:   mov ebx, eax
loc_005AFE0B:   mov eax, var_18
loc_005AFE0E:   neg ebx
loc_005AFE10:   sbb ebx, ebx
loc_005AFE12:   push eax
loc_005AFE13:   inc ebx
  loc_005AFE14: push 0042E3A4h
loc_005AFE19:   neg ebx
  loc_005AFE1B: call [0040112Ch] ; __vbaStrCmp
loc_005AFE21:   neg eax
loc_005AFE23:   sbb eax, eax
loc_005AFE25:   lea ecx, var_1C
loc_005AFE28:   inc eax
loc_005AFE29:   lea edx, var_18
loc_005AFE2C:   push ecx
loc_005AFE2D:   push edx
loc_005AFE2E:   neg eax
  loc_005AFE30: push 00000002h
  loc_005AFE32: or ebx, eax
  loc_005AFE34: call [00401204h] ; __vbaFreeStrList
loc_005AFE3A:   lea eax, var_34
loc_005AFE3D:   lea ecx, var_30
loc_005AFE40:   push eax
loc_005AFE41:   push ecx
  loc_005AFE42: push 00000002h
  loc_005AFE44: call [0040104Ch] ; __vbaFreeObjList
  loc_005AFE4A: add esp, 00000018h
loc_005AFE4D:   test bx, bx
  loc_005AFE50: jz 005AFEE5h
  loc_005AFE56: mov esi, [00401238h] ; __vbaVarDup
  loc_005AFE5C: mov ecx, 80020004h
loc_005AFE61:   mov var_78, ecx
  loc_005AFE64: mov eax, 0000000Ah
loc_005AFE69:   mov var_68, ecx
  loc_005AFE6C: mov edi, 00000008h
loc_005AFE71:   lea edx, var_B0
loc_005AFE77:   lea ecx, var_60
loc_005AFE7A:   mov var_80, eax
loc_005AFE7D:   mov var_70, eax
  loc_005AFE80: mov var_A8, 0042E624h ; "Error"
loc_005AFE8A:   mov var_B0, edi
  loc_005AFE90: call __vbaVarDup
loc_005AFE92:   lea edx, var_A0
loc_005AFE98:   lea ecx, var_50
  loc_005AFE9B: mov var_98, 004362A8h ; "Jumlah barang masih kosong! "
loc_005AFEA5:   mov var_A0, edi
  loc_005AFEAB: call __vbaVarDup
loc_005AFEAD:   lea edx, var_80
loc_005AFEB0:   lea eax, var_70
loc_005AFEB3:   push edx
loc_005AFEB4:   lea ecx, var_60
loc_005AFEB7:   push eax
loc_005AFEB8:   push ecx
loc_005AFEB9:   lea edx, var_50
  loc_005AFEBC: push 00000010h
loc_005AFEBE:   push edx
  loc_005AFEBF: call [004010B4h] ; rtcMsgBox
loc_005AFEC5:   lea eax, var_80
loc_005AFEC8:   lea ecx, var_70
loc_005AFECB:   push eax
loc_005AFECC:   lea edx, var_60
loc_005AFECF:   push ecx
loc_005AFED0:   lea eax, var_50
loc_005AFED3:   push edx
loc_005AFED4:   push eax
  loc_005AFED5: push 00000004h
  loc_005AFED7: call [0040103Ch] ; __vbaFreeVarList
  loc_005AFEDD: add esp, 00000014h
  loc_005AFEE0: jmp 005B0E4Dh
loc_005AFEE5:   mov eax, [00668050h]
loc_005AFEEA:   lea edx, var_30
loc_005AFEED:   push edx
loc_005AFEEE:   push eax
loc_005AFEEF:   mov ecx, [eax]
loc_005AFEF1:   Call [ecx+00000054h]
loc_005AFEF4:   test eax, eax
loc_005AFEF6:   fnclex
  loc_005AFEF8: jge 005AFF0Fh
loc_005AFEFA:   mov ecx, [00668050h]
  loc_005AFF00: push 00000054h
  loc_005AFF02: push 0042DF44h
loc_005AFF07:   push ecx
loc_005AFF08:   push eax
  loc_005AFF09: call [00401080h] ; __vbaHresultCheckObj
loc_005AFF0F:   lea ebx, var_34
loc_005AFF12:   mov eax, var_30
loc_005AFF15:   push ebx
  loc_005AFF16: mov ecx, 00000008h
  loc_005AFF1B: sub esp, 00000010h
loc_005AFF1E:   mov var_A0, ecx
loc_005AFF24:   mov ebx, esp
  loc_005AFF26: mov var_98, 00431F24h ; "stok"
loc_005AFF30:   mov edx, [eax]
loc_005AFF32:   push eax
loc_005AFF33:   mov [ebx], ecx
loc_005AFF35:   mov ecx, var_9C
loc_005AFF3B:   mov var_11C, eax
loc_005AFF41:   mov [ebx+00000004h], ecx
loc_005AFF44:   mov ecx, var_98
loc_005AFF4A:   mov [ebx+00000008h], ecx
loc_005AFF4D:   mov ecx, var_94
loc_005AFF53:   mov [ebx+0000000Ch], ecx
loc_005AFF56:   Call [edx+00000028h]
loc_005AFF59:   test eax, eax
loc_005AFF5B:   fnclex
  loc_005AFF5D: jge 005AFF74h
loc_005AFF5F:   mov edx, var_11C
  loc_005AFF65: push 00000028h
  loc_005AFF67: push 0042DFACh
loc_005AFF6C:   push edx
loc_005AFF6D:   push eax
  loc_005AFF6E: call [00401080h] ; __vbaHresultCheckObj
loc_005AFF74:   mov eax, var_34
loc_005AFF77:   lea edx, var_50
loc_005AFF7A:   push edx
loc_005AFF7B:   push eax
loc_005AFF7C:   mov ecx, [eax]
loc_005AFF7E:   mov ebx, eax
loc_005AFF80:   Call [ecx+00000034h]
loc_005AFF83:   test eax, eax
loc_005AFF85:   fnclex
  loc_005AFF87: jge 005AFF98h
  loc_005AFF89: push 00000034h
  loc_005AFF8B: push 0042DFBCh
loc_005AFF90:   push ebx
loc_005AFF91:   push eax
  loc_005AFF92: call [00401080h] ; __vbaHresultCheckObj
loc_005AFF98:   mov eax, [esi]
loc_005AFF9A:   push esi
loc_005AFF9B:   Call [eax+0000033Ch]
loc_005AFFA1:   lea ecx, var_38
loc_005AFFA4:   push eax
loc_005AFFA5:   push ecx
loc_005AFFA6:   Call edi
loc_005AFFA8:   mov ebx, eax
loc_005AFFAA:   lea eax, var_18
loc_005AFFAD:   push eax
loc_005AFFAE:   push ebx
loc_005AFFAF:   mov edx, [ebx]
loc_005AFFB1:   Call [edx+000000A0h]
loc_005AFFB7:   test eax, eax
loc_005AFFB9:   fnclex
  loc_005AFFBB: jge 005AFFCFh
  loc_005AFFBD: push 000000A0h
  loc_005AFFC2: push 0042DFCCh
loc_005AFFC7:   push ebx
loc_005AFFC8:   push eax
  loc_005AFFC9: call [00401080h] ; __vbaHresultCheckObj
loc_005AFFCF:   mov ecx, var_18
loc_005AFFD2:   push ecx
  loc_005AFFD3: call [004012A4h] ; rtcR8ValFromBstr
  loc_005AFFD9: fstp real8 ptr var_A8
loc_005AFFDF:   lea edx, var_50
loc_005AFFE2:   lea eax, var_B0
loc_005AFFE8:   push edx
loc_005AFFE9:   push eax
  loc_005AFFEA: mov var_B0, 00008005h
  loc_005AFFF4: call [004010E8h] ; __vbaVarTstLt
loc_005AFFFA:   lea ecx, var_18
loc_005AFFFD:   mov bx, ax
  loc_005B0000: call [0040129Ch] ; __vbaFreeStr
loc_005B0006:   lea ecx, var_34
loc_005B0009:   lea edx, var_38
loc_005B000C:   push ecx
loc_005B000D:   lea eax, var_30
loc_005B0010:   push edx
loc_005B0011:   push eax
  loc_005B0012: push 00000003h
  loc_005B0014: call [0040104Ch] ; __vbaFreeObjList
  loc_005B001A: add esp, 00000010h
loc_005B001D:   lea ecx, var_50
  loc_005B0020: call [00401020h] ; __vbaFreeVar
loc_005B0026:   test bx, bx
  loc_005B0029: jz 005B0102h
  loc_005B002F: mov ebx, [00401238h] ; __vbaVarDup
  loc_005B0035: mov ecx, 80020004h
loc_005B003A:   mov var_78, ecx
  loc_005B003D: mov eax, 0000000Ah
loc_005B0042:   mov var_68, ecx
loc_005B0045:   lea edx, var_B0
loc_005B004B:   lea ecx, var_60
loc_005B004E:   mov var_80, eax
loc_005B0051:   mov var_70, eax
  loc_005B0054: mov var_A8, 0042E624h ; "Error"
  loc_005B005E: mov var_B0, 00000008h
loc_005B0068:   Call ebx
loc_005B006A:   lea edx, var_A0
loc_005B0070:   lea ecx, var_50
  loc_005B0073: mov var_98, 004362E8h ; "Ma'af, Stok habis!"
  loc_005B007D: mov var_A0, 00000008h
loc_005B0087:   Call ebx
loc_005B0089:   lea ecx, var_80
loc_005B008C:   lea edx, var_70
loc_005B008F:   push ecx
loc_005B0090:   lea eax, var_60
loc_005B0093:   push edx
loc_005B0094:   push eax
loc_005B0095:   lea ecx, var_50
  loc_005B0098: push 00000010h
loc_005B009A:   push ecx
  loc_005B009B: call [004010B4h] ; rtcMsgBox
loc_005B00A1:   lea edx, var_80
loc_005B00A4:   lea eax, var_70
loc_005B00A7:   push edx
loc_005B00A8:   lea ecx, var_60
loc_005B00AB:   push eax
loc_005B00AC:   lea edx, var_50
loc_005B00AF:   push ecx
loc_005B00B0:   push edx
  loc_005B00B1: push 00000004h
  loc_005B00B3: call [0040103Ch] ; __vbaFreeVarList
loc_005B00B9:   mov eax, [esi]
  loc_005B00BB: add esp, 00000014h
loc_005B00BE:   push esi
loc_005B00BF:   Call [eax+0000033Ch]
loc_005B00C5:   lea ecx, var_30
loc_005B00C8:   push eax
loc_005B00C9:   push ecx
loc_005B00CA:   Call edi
loc_005B00CC:   mov esi, eax
  loc_005B00CE: push 0042DFECh
loc_005B00D3:   push esi
loc_005B00D4:   mov edx, [esi]
loc_005B00D6:   Call [edx+000000A4h]
loc_005B00DC:   test eax, eax
loc_005B00DE:   fnclex
  loc_005B00E0: jge 005B00F4h
  loc_005B00E2: push 000000A4h
  loc_005B00E7: push 0042DFCCh
loc_005B00EC:   push esi
loc_005B00ED:   push eax
  loc_005B00EE: call [00401080h] ; __vbaHresultCheckObj
loc_005B00F4:   lea ecx, var_30
  loc_005B00F7: call [004012A0h] ; __vbaFreeObj
  loc_005B00FD: jmp 005B0E4Dh
loc_005B0102:   mov eax, [esi]
loc_005B0104:   push esi
loc_005B0105:   Call [eax+00000304h]
loc_005B010B:   lea ecx, var_30
loc_005B010E:   push eax
loc_005B010F:   push ecx
loc_005B0110:   Call edi
loc_005B0112:   mov ebx, eax
loc_005B0114:   lea eax, var_18
loc_005B0117:   push eax
loc_005B0118:   push ebx
loc_005B0119:   mov edx, [ebx]
loc_005B011B:   Call [edx+000000A0h]
loc_005B0121:   test eax, eax
loc_005B0123:   fnclex
  loc_005B0125: jge 005B0139h
  loc_005B0127: push 000000A0h
  loc_005B012C: push 0042DFCCh
loc_005B0131:   push ebx
loc_005B0132:   push eax
  loc_005B0133: call [00401080h] ; __vbaHresultCheckObj
loc_005B0139:   mov ecx, var_18
loc_005B013C:   push ecx
  loc_005B013D: call [004012A4h] ; rtcR8ValFromBstr
loc_005B0143:   mov edx, [esi]
loc_005B0145:   push esi
  loc_005B0146: fstp real8 ptr var_FC
loc_005B014C:   Call [edx+0000033Ch]
loc_005B0152:   push eax
loc_005B0153:   lea eax, var_34
loc_005B0156:   push eax
loc_005B0157:   Call edi
loc_005B0159:   mov ebx, eax
loc_005B015B:   lea edx, var_1C
loc_005B015E:   push edx
loc_005B015F:   push ebx
loc_005B0160:   mov ecx, [ebx]
loc_005B0162:   Call [ecx+000000A0h]
loc_005B0168:   test eax, eax
loc_005B016A:   fnclex
  loc_005B016C: jge 005B0180h
  loc_005B016E: push 000000A0h
  loc_005B0173: push 0042DFCCh
loc_005B0178:   push ebx
loc_005B0179:   push eax
  loc_005B017A: call [00401080h] ; __vbaHresultCheckObj
loc_005B0180:   mov eax, var_1C
loc_005B0183:   push eax
  loc_005B0184: call [004012A4h] ; rtcR8ValFromBstr
  loc_005B018A: fstp real8 ptr var_104
  loc_005B0190: fld real8 ptr var_104
  loc_005B0196: fmul st0, real8 ptr var_FC
loc_005B019C:   lea ecx, var_50
  loc_005B019F: push 00000000h
loc_005B01A1:   lea edx, var_60
loc_005B01A4:   push ecx
  loc_005B01A5: fstp real8 ptr var_48
loc_005B01A8:   fnstsw ax
  loc_005B01AA: test al, 0Dh
  loc_005B01AC: jnz 005B0EE7h
loc_005B01B2:   push edx
  loc_005B01B3: mov var_50, 00000005h
  loc_005B01BA: call [004011ACh] ; rtcRound
loc_005B01C0:   lea ecx, [esi+000000A0h]
loc_005B01C6:   lea edx, var_60
  loc_005B01C9: call [0040101Ch] ; __vbaVarMove
loc_005B01CF:   lea eax, var_1C
loc_005B01D2:   lea ecx, var_18
loc_005B01D5:   push eax
loc_005B01D6:   push ecx
  loc_005B01D7: push 00000002h
  loc_005B01D9: call [00401204h] ; __vbaFreeStrList
loc_005B01DF:   lea edx, var_34
loc_005B01E2:   lea eax, var_30
loc_005B01E5:   push edx
loc_005B01E6:   push eax
  loc_005B01E7: push 00000002h
  loc_005B01E9: call [0040104Ch] ; __vbaFreeObjList
loc_005B01EF:   lea ecx, var_60
loc_005B01F2:   lea edx, var_50
loc_005B01F5:   push ecx
loc_005B01F6:   push edx
  loc_005B01F7: push 00000002h
  loc_005B01F9: call [0040103Ch] ; __vbaFreeVarList
loc_005B01FF:   mov eax, [esi]
  loc_005B0201: add esp, 00000024h
loc_005B0204:   push esi
loc_005B0205:   Call [eax+00000340h]
loc_005B020B:   lea ecx, var_30
loc_005B020E:   push eax
loc_005B020F:   push ecx
loc_005B0210:   Call edi
loc_005B0212:   mov ebx, eax
loc_005B0214:   lea eax, var_18
loc_005B0217:   push eax
loc_005B0218:   push ebx
loc_005B0219:   mov edx, [ebx]
loc_005B021B:   Call [edx+000000A0h]
loc_005B0221:   test eax, eax
loc_005B0223:   fnclex
  loc_005B0225: jge 005B0239h
  loc_005B0227: push 000000A0h
  loc_005B022C: push 0042DFCCh
loc_005B0231:   push ebx
loc_005B0232:   push eax
  loc_005B0233: call [00401080h] ; __vbaHresultCheckObj
loc_005B0239:   mov ecx, var_18
loc_005B023C:   push ecx
  loc_005B023D: call [004012A4h] ; rtcR8ValFromBstr
loc_005B0243:   mov edx, [esi]
loc_005B0245:   push esi
  loc_005B0246: fstp real8 ptr var_FC
loc_005B024C:   Call [edx+00000340h]
loc_005B0252:   push eax
loc_005B0253:   lea eax, var_34
loc_005B0256:   push eax
loc_005B0257:   Call edi
loc_005B0259:   mov ebx, eax
loc_005B025B:   lea edx, var_1C
loc_005B025E:   push edx
loc_005B025F:   push ebx
loc_005B0260:   mov ecx, [ebx]
loc_005B0262:   Call [ecx+000000A0h]
loc_005B0268:   test eax, eax
loc_005B026A:   fnclex
  loc_005B026C: jge 005B0280h
  loc_005B026E: push 000000A0h
  loc_005B0273: push 0042DFCCh
loc_005B0278:   push ebx
loc_005B0279:   push eax
  loc_005B027A: call [00401080h] ; __vbaHresultCheckObj
loc_005B0280:   mov eax, var_1C
loc_005B0283:   push eax
  loc_005B0284: call [004012A4h] ; rtcR8ValFromBstr
loc_005B028A:   mov ecx, [esi]
loc_005B028C:   push esi
  loc_005B028D: fstp real8 ptr var_104
loc_005B0293:   Call [ecx+00000330h]
loc_005B0299:   lea edx, var_38
loc_005B029C:   push eax
loc_005B029D:   push edx
loc_005B029E:   Call edi
loc_005B02A0:   mov ebx, eax
loc_005B02A2:   lea ecx, var_20
loc_005B02A5:   push ecx
loc_005B02A6:   push ebx
loc_005B02A7:   mov eax, [ebx]
loc_005B02A9:   Call [eax+000000A0h]
loc_005B02AF:   test eax, eax
loc_005B02B1:   fnclex
  loc_005B02B3: jge 005B02C7h
  loc_005B02B5: push 000000A0h
  loc_005B02BA: push 0042DFCCh
loc_005B02BF:   push ebx
loc_005B02C0:   push eax
  loc_005B02C1: call [00401080h] ; __vbaHresultCheckObj
loc_005B02C7:   mov edx, var_20
loc_005B02CA:   push edx
  loc_005B02CB: call [004012A4h] ; rtcR8ValFromBstr
loc_005B02D1:   mov eax, [esi]
loc_005B02D3:   push esi
  loc_005B02D4: fstp real8 ptr var_10C
loc_005B02DA:   Call [eax+0000033Ch]
loc_005B02E0:   lea ecx, var_3C
loc_005B02E3:   push eax
loc_005B02E4:   push ecx
loc_005B02E5:   Call edi
loc_005B02E7:   mov ebx, eax
loc_005B02E9:   lea eax, var_24
loc_005B02EC:   push eax
loc_005B02ED:   push ebx
loc_005B02EE:   mov edx, [ebx]
loc_005B02F0:   Call [edx+000000A0h]
loc_005B02F6:   test eax, eax
loc_005B02F8:   fnclex
  loc_005B02FA: jge 005B030Eh
  loc_005B02FC: push 000000A0h
  loc_005B0301: push 0042DFCCh
loc_005B0306:   push ebx
loc_005B0307:   push eax
  loc_005B0308: call [00401080h] ; __vbaHresultCheckObj
loc_005B030E:   mov ecx, var_24
loc_005B0311:   push ecx
  loc_005B0312: call [004012A4h] ; rtcR8ValFromBstr
  loc_005B0318: fstp real8 ptr var_114
  loc_005B031E: fld real8 ptr var_10C
  loc_005B0324: cmp [00668000h], 00000000h
  loc_005B032B: jnz 005B0335h
  loc_005B032D: fdiv st0, real8 ptr [004025F0h]
  loc_005B0333: jmp 005B0346h
loc_005B0335:   push [004025F4h]
loc_005B033B:   push [004025F0h]
  loc_005B0341: call 004056B4h ; _adj_fdiv_m64
loc_005B0346:   lea edx, var_50
  loc_005B0349: push 00000000h
loc_005B034B:   push edx
  loc_005B034C: mov var_50, 00000005h
  loc_005B0353: fmul st0, real8 ptr var_104
  loc_005B0359: fsubr st0, real8 ptr var_FC
  loc_005B035F: fmul st0, real8 ptr var_114
  loc_005B0365: fstp real8 ptr var_48
loc_005B0368:   fnstsw ax
  loc_005B036A: test al, 0Dh
  loc_005B036C: jnz 005B0EE7h
loc_005B0372:   lea eax, var_60
loc_005B0375:   push eax
  loc_005B0376: call [004011ACh] ; rtcRound
loc_005B037C:   lea ebx, [esi+00000070h]
loc_005B037F:   lea edx, var_60
loc_005B0382:   mov ecx, ebx
  loc_005B0384: call [0040101Ch] ; __vbaVarMove
loc_005B038A:   lea ecx, var_24
loc_005B038D:   lea edx, var_20
loc_005B0390:   push ecx
loc_005B0391:   lea eax, var_1C
loc_005B0394:   push edx
loc_005B0395:   lea ecx, var_18
loc_005B0398:   push eax
loc_005B0399:   push ecx
  loc_005B039A: push 00000004h
  loc_005B039C: call [00401204h] ; __vbaFreeStrList
loc_005B03A2:   lea edx, var_3C
loc_005B03A5:   lea eax, var_38
loc_005B03A8:   push edx
loc_005B03A9:   lea ecx, var_34
loc_005B03AC:   push eax
loc_005B03AD:   lea edx, var_30
loc_005B03B0:   push ecx
loc_005B03B1:   push edx
  loc_005B03B2: push 00000004h
  loc_005B03B4: call [0040104Ch] ; __vbaFreeObjList
loc_005B03BA:   lea eax, var_60
loc_005B03BD:   lea ecx, var_50
loc_005B03C0:   push eax
loc_005B03C1:   push ecx
  loc_005B03C2: push 00000002h
  loc_005B03C4: call [0040103Ch] ; __vbaFreeVarList
  loc_005B03CA: add esp, 00000034h
loc_005B03CD:   lea eax, [esi+000000A0h]
loc_005B03D3:   lea edx, var_50
loc_005B03D6:   push ebx
loc_005B03D7:   push eax
loc_005B03D8:   push edx
  loc_005B03D9: call [00401004h] ; __vbaVarSub
loc_005B03DF:   mov edx, eax
loc_005B03E1:   lea ecx, [esi+000000B0h]
  loc_005B03E7: call [0040101Ch] ; __vbaVarMove
loc_005B03ED:   mov ax, [esi+00000034h]
  loc_005B03F1: mov ebx, 00000003h
  loc_005B03F6: add ax, 0001h
loc_005B03FA:   mov ecx, ebx
  loc_005B03FC: jo 005B0EECh
  loc_005B0402: sub esp, 00000010h
loc_005B0405:   mov var_A0, ecx
loc_005B040B:   mov edx, esp
loc_005B040D:   movsx eax, ax
loc_005B0410:   mov [edx], ecx
loc_005B0412:   mov ecx, var_9C
loc_005B0418:   mov var_98, eax
  loc_005B041E: push 00000004h
loc_005B0420:   mov [edx+00000004h], ecx
loc_005B0423:   mov ecx, [esi]
loc_005B0425:   mov [edx+00000008h], eax
loc_005B0428:   mov eax, var_94
loc_005B042E:   mov [edx+0000000Ch], eax
loc_005B0431:   push esi
loc_005B0432:   Call [ecx+00000390h]
loc_005B0438:   lea edx, var_30
loc_005B043B:   push eax
loc_005B043C:   push edx
loc_005B043D:   Call edi
loc_005B043F:   push eax
  loc_005B0440: call [00401280h] ; __vbaLateIdSt
loc_005B0446:   lea ecx, var_30
  loc_005B0449: call [004012A0h] ; __vbaFreeObj
loc_005B044F:   movsx eax, [esi+00000034h]
loc_005B0453:   mov ecx, [esi]
loc_005B0455:   push esi
loc_005B0456:   mov var_98, eax
loc_005B045C:   mov var_A0, ebx
loc_005B0462:   mov var_C0, ebx
loc_005B0468:   Call [ecx+00000348h]
loc_005B046E:   lea edx, var_30
loc_005B0471:   push eax
loc_005B0472:   push edx
loc_005B0473:   Call edi
loc_005B0475:   mov ebx, eax
loc_005B0477:   lea ecx, var_18
loc_005B047A:   push ecx
loc_005B047B:   push ebx
loc_005B047C:   mov eax, [ebx]
loc_005B047E:   Call [eax+000000A0h]
loc_005B0484:   test eax, eax
loc_005B0486:   fnclex
  loc_005B0488: jge 005B049Ch
  loc_005B048A: push 000000A0h
  loc_005B048F: push 0042DFCCh
loc_005B0494:   push ebx
loc_005B0495:   push eax
  loc_005B0496: call [00401080h] ; __vbaHresultCheckObj
loc_005B049C:   mov ebx, var_A0
  loc_005B04A2: sub esp, 00000010h
loc_005B04A5:   mov edx, esp
  loc_005B04A7: sub esp, 00000010h
loc_005B04AA:   mov eax, var_18
  loc_005B04AD: mov ecx, 00000008h
loc_005B04B2:   mov [edx], ebx
loc_005B04B4:   mov ebx, var_9C
loc_005B04BA:   mov var_50, ecx
loc_005B04BD:   mov var_48, eax
loc_005B04C0:   mov [edx+00000004h], ebx
loc_005B04C3:   mov ebx, var_98
  loc_005B04C9: mov var_18, 00000000h
loc_005B04D0:   mov [edx+00000008h], ebx
loc_005B04D3:   mov ebx, var_94
loc_005B04D9:   mov [edx+0000000Ch], ebx
loc_005B04DC:   mov edx, var_C0
loc_005B04E2:   mov ebx, esp
  loc_005B04E4: sub esp, 00000010h
loc_005B04E7:   mov [ebx], edx
loc_005B04E9:   mov edx, var_BC
loc_005B04EF:   mov [ebx+00000004h], edx
  loc_005B04F2: xor edx, edx
loc_005B04F4:   mov [ebx+00000008h], edx
loc_005B04F7:   mov edx, var_B4
loc_005B04FD:   mov [ebx+0000000Ch], edx
loc_005B0500:   mov edx, esp
  loc_005B0502: push 00000002h
  loc_005B0504: push 00000041h
loc_005B0506:   mov [edx], ecx
loc_005B0508:   mov ecx, var_4C
loc_005B050B:   push esi
loc_005B050C:   mov [edx+00000004h], ecx
loc_005B050F:   mov ecx, [esi]
loc_005B0511:   mov [edx+00000008h], eax
loc_005B0514:   mov eax, var_44
loc_005B0517:   mov [edx+0000000Ch], eax
loc_005B051A:   Call [ecx+00000390h]
loc_005B0520:   lea edx, var_34
loc_005B0523:   push eax
loc_005B0524:   push edx
loc_005B0525:   Call edi
  loc_005B0527: mov ebx, [0040117Ch] ; __vbaLateIdCallSt
loc_005B052D:   push eax
loc_005B052E:   Call ebx
loc_005B0530:   lea eax, var_34
loc_005B0533:   lea ecx, var_30
loc_005B0536:   push eax
loc_005B0537:   push ecx
  loc_005B0538: push 00000002h
  loc_005B053A: call [0040104Ch] ; __vbaFreeObjList
  loc_005B0540: add esp, 00000048h
loc_005B0543:   lea ecx, var_50
  loc_005B0546: call [00401020h] ; __vbaFreeVar
loc_005B054C:   movsx edx, [esi+00000034h]
loc_005B0550:   mov eax, [esi]
loc_005B0552:   push esi
loc_005B0553:   mov var_98, edx
  loc_005B0559: mov var_A0, 00000003h
loc_005B0563:   Call [eax+00000344h]
loc_005B0569:   lea ecx, var_30
loc_005B056C:   push eax
loc_005B056D:   push ecx
loc_005B056E:   Call edi
loc_005B0570:   mov edx, [eax]
loc_005B0572:   lea ecx, var_18
loc_005B0575:   push ecx
loc_005B0576:   push eax
loc_005B0577:   mov var_118, eax
loc_005B057D:   Call [edx+000000A0h]
loc_005B0583:   test eax, eax
loc_005B0585:   fnclex
  loc_005B0587: jge 005B05A1h
loc_005B0589:   mov edx, var_118
  loc_005B058F: push 000000A0h
  loc_005B0594: push 0042DFCCh
loc_005B0599:   push edx
loc_005B059A:   push eax
  loc_005B059B: call [00401080h] ; __vbaHresultCheckObj
loc_005B05A1:   mov edx, var_A0
  loc_005B05A7: sub esp, 00000010h
loc_005B05AA:   mov ecx, esp
  loc_005B05AC: sub esp, 00000010h
loc_005B05AF:   mov eax, var_18
  loc_005B05B2: mov var_50, 00000008h
loc_005B05B9:   mov [ecx], edx
loc_005B05BB:   mov edx, var_9C
loc_005B05C1:   mov var_48, eax
  loc_005B05C4: mov var_18, 00000000h
loc_005B05CB:   mov [ecx+00000004h], edx
loc_005B05CE:   mov edx, var_98
loc_005B05D4:   mov [ecx+00000008h], edx
loc_005B05D7:   mov edx, var_94
loc_005B05DD:   mov [ecx+0000000Ch], edx
loc_005B05E0:   mov edx, esp
  loc_005B05E2: mov ecx, 00000003h
  loc_005B05E7: sub esp, 00000010h
loc_005B05EA:   mov [edx], ecx
loc_005B05EC:   mov ecx, var_BC
loc_005B05F2:   mov [edx+00000004h], ecx
  loc_005B05F5: mov ecx, 00000001h
loc_005B05FA:   mov [edx+00000008h], ecx
loc_005B05FD:   mov ecx, var_B4
loc_005B0603:   mov [edx+0000000Ch], ecx
loc_005B0606:   mov ecx, var_50
loc_005B0609:   mov edx, esp
  loc_005B060B: push 00000002h
  loc_005B060D: push 00000041h
loc_005B060F:   push esi
loc_005B0610:   mov [edx], ecx
loc_005B0612:   mov ecx, var_4C
loc_005B0615:   mov [edx+00000004h], ecx
loc_005B0618:   mov ecx, [esi]
loc_005B061A:   mov [edx+00000008h], eax
loc_005B061D:   mov eax, var_44
loc_005B0620:   mov [edx+0000000Ch], eax
loc_005B0623:   Call [ecx+00000390h]
loc_005B0629:   lea edx, var_34
loc_005B062C:   push eax
loc_005B062D:   push edx
loc_005B062E:   Call edi
loc_005B0630:   push eax
loc_005B0631:   Call ebx
loc_005B0633:   lea eax, var_34
loc_005B0636:   lea ecx, var_30
loc_005B0639:   push eax
loc_005B063A:   push ecx
  loc_005B063B: push 00000002h
  loc_005B063D: call [0040104Ch] ; __vbaFreeObjList
  loc_005B0643: add esp, 00000048h
loc_005B0646:   lea ecx, var_50
  loc_005B0649: call [00401020h] ; __vbaFreeVar
loc_005B064F:   mov edx, [esi]
loc_005B0651:   push esi
loc_005B0652:   Call [edx+00000340h]
loc_005B0658:   push eax
loc_005B0659:   lea eax, var_30
loc_005B065C:   push eax
loc_005B065D:   Call edi
loc_005B065F:   mov ecx, [eax]
loc_005B0661:   lea edx, var_18
loc_005B0664:   push edx
loc_005B0665:   push eax
loc_005B0666:   mov var_118, eax
loc_005B066C:   Call [ecx+000000A0h]
loc_005B0672:   test eax, eax
loc_005B0674:   fnclex
  loc_005B0676: jge 005B0690h
loc_005B0678:   mov ecx, var_118
  loc_005B067E: push 000000A0h
  loc_005B0683: push 0042DFCCh
loc_005B0688:   push ecx
loc_005B0689:   push eax
  loc_005B068A: call [00401080h] ; __vbaHresultCheckObj
loc_005B0690:   movsx edx, [esi+00000034h]
loc_005B0694:   mov var_A8, edx
  loc_005B069A: mov eax, 00000003h
loc_005B069F:   lea edx, var_A0
loc_005B06A5:   lea ecx, var_60
loc_005B06A8:   mov var_B0, eax
  loc_005B06AE: mov var_C8, 00000002h
loc_005B06B8:   mov var_D0, eax
  loc_005B06BE: mov var_98, 00435A2Ch ; "##,##0"
  loc_005B06C8: mov var_A0, 00000008h
  loc_005B06D2: call [00401238h] ; __vbaVarDup
loc_005B06D8:   mov eax, var_18
  loc_005B06DB: push 00000001h
loc_005B06DD:   mov var_48, eax
loc_005B06E0:   lea eax, var_60
  loc_005B06E3: push 00000001h
loc_005B06E5:   lea ecx, var_50
loc_005B06E8:   push eax
loc_005B06E9:   lea edx, var_70
loc_005B06EC:   push ecx
loc_005B06ED:   push edx
  loc_005B06EE: mov var_18, 00000000h
  loc_005B06F5: mov var_50, 00000008h
  loc_005B06FC: call [00401078h] ; rtcVarFromFormatVar
loc_005B0702:   lea eax, var_70
loc_005B0705:   push eax
  loc_005B0706: call [00401034h] ; __vbaStrVarMove
loc_005B070C:   mov edx, var_B0
  loc_005B0712: sub esp, 00000010h
loc_005B0715:   mov ecx, esp
  loc_005B0717: sub esp, 00000010h
  loc_005B071A: mov var_80, 00000008h
loc_005B0721:   mov var_78, eax
loc_005B0724:   mov [ecx], edx
loc_005B0726:   mov edx, var_AC
loc_005B072C:   mov [ecx+00000004h], edx
loc_005B072F:   mov edx, var_A8
loc_005B0735:   mov [ecx+00000008h], edx
loc_005B0738:   mov edx, var_A4
loc_005B073E:   mov [ecx+0000000Ch], edx
loc_005B0741:   mov edx, var_D0
loc_005B0747:   mov ecx, esp
  loc_005B0749: sub esp, 00000010h
loc_005B074C:   mov [ecx], edx
loc_005B074E:   mov edx, var_CC
loc_005B0754:   mov [ecx+00000004h], edx
loc_005B0757:   mov edx, var_C8
loc_005B075D:   mov [ecx+00000008h], edx
loc_005B0760:   mov edx, var_C4
loc_005B0766:   mov [ecx+0000000Ch], edx
loc_005B0769:   mov edx, var_80
loc_005B076C:   mov ecx, esp
  loc_005B076E: push 00000002h
  loc_005B0770: push 00000041h
loc_005B0772:   push esi
loc_005B0773:   mov [ecx], edx
loc_005B0775:   mov edx, var_7C
loc_005B0778:   mov [ecx+00000004h], edx
loc_005B077B:   mov [ecx+00000008h], eax
loc_005B077E:   mov eax, var_74
loc_005B0781:   mov [ecx+0000000Ch], eax
loc_005B0784:   mov ecx, [esi]
loc_005B0786:   Call [ecx+00000390h]
loc_005B078C:   lea edx, var_34
loc_005B078F:   push eax
loc_005B0790:   push edx
loc_005B0791:   Call edi
loc_005B0793:   push eax
loc_005B0794:   Call ebx
loc_005B0796:   lea eax, var_34
loc_005B0799:   lea ecx, var_30
loc_005B079C:   push eax
loc_005B079D:   push ecx
  loc_005B079E: push 00000002h
  loc_005B07A0: call [0040104Ch] ; __vbaFreeObjList
  loc_005B07A6: add esp, 00000048h
loc_005B07A9:   lea edx, var_80
loc_005B07AC:   lea eax, var_70
loc_005B07AF:   lea ecx, var_60
loc_005B07B2:   push edx
loc_005B07B3:   push eax
loc_005B07B4:   lea edx, var_50
loc_005B07B7:   push ecx
loc_005B07B8:   push edx
  loc_005B07B9: push 00000004h
  loc_005B07BB: call [0040103Ch] ; __vbaFreeVarList
loc_005B07C1:   movsx eax, [esi+00000034h]
loc_005B07C5:   mov ecx, [esi]
  loc_005B07C7: add esp, 00000014h
loc_005B07CA:   mov var_98, eax
  loc_005B07D0: mov var_A0, 00000003h
loc_005B07DA:   push esi
loc_005B07DB:   Call [ecx+00000330h]
loc_005B07E1:   lea edx, var_30
loc_005B07E4:   push eax
loc_005B07E5:   push edx
loc_005B07E6:   Call edi
loc_005B07E8:   mov ecx, [eax]
loc_005B07EA:   lea edx, var_18
loc_005B07ED:   push edx
loc_005B07EE:   push eax
loc_005B07EF:   mov var_118, eax
loc_005B07F5:   Call [ecx+000000A0h]
loc_005B07FB:   test eax, eax
loc_005B07FD:   fnclex
  loc_005B07FF: jge 005B0819h
loc_005B0801:   mov ecx, var_118
  loc_005B0807: push 000000A0h
  loc_005B080C: push 0042DFCCh
loc_005B0811:   push ecx
loc_005B0812:   push eax
  loc_005B0813: call [00401080h] ; __vbaHresultCheckObj
loc_005B0819:   mov ecx, var_A0
  loc_005B081F: sub esp, 00000010h
loc_005B0822:   mov edx, esp
  loc_005B0824: sub esp, 00000010h
loc_005B0827:   mov eax, var_18
  loc_005B082A: mov var_50, 00000008h
loc_005B0831:   mov [edx], ecx
loc_005B0833:   mov ecx, var_9C
loc_005B0839:   mov var_48, eax
  loc_005B083C: mov var_18, 00000000h
loc_005B0843:   mov [edx+00000004h], ecx
loc_005B0846:   mov ecx, var_98
loc_005B084C:   mov [edx+00000008h], ecx
loc_005B084F:   mov ecx, var_94
loc_005B0855:   mov [edx+0000000Ch], ecx
loc_005B0858:   mov edx, esp
  loc_005B085A: mov ecx, 00000003h
  loc_005B085F: sub esp, 00000010h
loc_005B0862:   mov [edx], ecx
loc_005B0864:   mov ecx, var_BC
loc_005B086A:   mov [edx+00000004h], ecx
  loc_005B086D: mov ecx, 00000003h
loc_005B0872:   mov [edx+00000008h], ecx
loc_005B0875:   mov ecx, var_B4
loc_005B087B:   mov [edx+0000000Ch], ecx
loc_005B087E:   mov ecx, var_50
loc_005B0881:   mov edx, esp
  loc_005B0883: push 00000002h
  loc_005B0885: push 00000041h
loc_005B0887:   push esi
loc_005B0888:   mov [edx], ecx
loc_005B088A:   mov ecx, var_4C
loc_005B088D:   mov [edx+00000004h], ecx
loc_005B0890:   mov ecx, [esi]
loc_005B0892:   mov [edx+00000008h], eax
loc_005B0895:   mov eax, var_44
loc_005B0898:   mov [edx+0000000Ch], eax
loc_005B089B:   Call [ecx+00000390h]
loc_005B08A1:   lea edx, var_34
loc_005B08A4:   push eax
loc_005B08A5:   push edx
loc_005B08A6:   Call edi
loc_005B08A8:   push eax
loc_005B08A9:   Call ebx
loc_005B08AB:   lea eax, var_34
loc_005B08AE:   lea ecx, var_30
loc_005B08B1:   push eax
loc_005B08B2:   push ecx
  loc_005B08B3: push 00000002h
  loc_005B08B5: call [0040104Ch] ; __vbaFreeObjList
  loc_005B08BB: add esp, 00000048h
loc_005B08BE:   lea ecx, var_50
  loc_005B08C1: call [00401020h] ; __vbaFreeVar
loc_005B08C7:   movsx edx, [esi+00000034h]
loc_005B08CB:   mov eax, [esi]
loc_005B08CD:   push esi
loc_005B08CE:   mov var_98, edx
  loc_005B08D4: mov var_A0, 00000003h
  loc_005B08DE: mov var_B8, 00000004h
loc_005B08E8:   Call [eax+0000033Ch]
loc_005B08EE:   lea ecx, var_30
loc_005B08F1:   push eax
loc_005B08F2:   push ecx
loc_005B08F3:   Call edi
loc_005B08F5:   mov edx, [eax]
loc_005B08F7:   lea ecx, var_18
loc_005B08FA:   push ecx
loc_005B08FB:   push eax
loc_005B08FC:   mov var_118, eax
loc_005B0902:   Call [edx+000000A0h]
loc_005B0908:   test eax, eax
loc_005B090A:   fnclex
  loc_005B090C: jge 005B0926h
loc_005B090E:   mov edx, var_118
  loc_005B0914: push 000000A0h
  loc_005B0919: push 0042DFCCh
loc_005B091E:   push edx
loc_005B091F:   push eax
  loc_005B0920: call [00401080h] ; __vbaHresultCheckObj
loc_005B0926:   mov edx, var_A0
  loc_005B092C: sub esp, 00000010h
loc_005B092F:   mov ecx, esp
  loc_005B0931: sub esp, 00000010h
loc_005B0934:   mov eax, var_18
  loc_005B0937: mov var_50, 00000008h
loc_005B093E:   mov [ecx], edx
loc_005B0940:   mov edx, var_9C
loc_005B0946:   mov var_48, eax
  loc_005B0949: mov var_18, 00000000h
loc_005B0950:   mov [ecx+00000004h], edx
loc_005B0953:   mov edx, var_98
loc_005B0959:   mov [ecx+00000008h], edx
loc_005B095C:   mov edx, var_94
loc_005B0962:   mov [ecx+0000000Ch], edx
loc_005B0965:   mov edx, esp
  loc_005B0967: mov ecx, 00000003h
  loc_005B096C: sub esp, 00000010h
loc_005B096F:   mov [edx], ecx
loc_005B0971:   mov ecx, var_BC
loc_005B0977:   mov [edx+00000004h], ecx
loc_005B097A:   mov ecx, var_B8
loc_005B0980:   mov [edx+00000008h], ecx
loc_005B0983:   mov ecx, var_B4
loc_005B0989:   mov [edx+0000000Ch], ecx
loc_005B098C:   mov ecx, var_50
loc_005B098F:   mov edx, esp
  loc_005B0991: push 00000002h
  loc_005B0993: push 00000041h
loc_005B0995:   push esi
loc_005B0996:   mov [edx], ecx
loc_005B0998:   mov ecx, var_4C
loc_005B099B:   mov [edx+00000004h], ecx
loc_005B099E:   mov ecx, [esi]
loc_005B09A0:   mov [edx+00000008h], eax
loc_005B09A3:   mov eax, var_44
loc_005B09A6:   mov [edx+0000000Ch], eax
loc_005B09A9:   Call [ecx+00000390h]
loc_005B09AF:   lea edx, var_34
loc_005B09B2:   push eax
loc_005B09B3:   push edx
loc_005B09B4:   Call edi
loc_005B09B6:   push eax
loc_005B09B7:   Call ebx
loc_005B09B9:   lea eax, var_34
loc_005B09BC:   lea ecx, var_30
loc_005B09BF:   push eax
loc_005B09C0:   push ecx
  loc_005B09C1: push 00000002h
  loc_005B09C3: call [0040104Ch] ; __vbaFreeObjList
  loc_005B09C9: add esp, 00000048h
loc_005B09CC:   lea ecx, var_50
  loc_005B09CF: call [00401020h] ; __vbaFreeVar
loc_005B09D5:   movsx edx, [esi+00000034h]
loc_005B09D9:   mov var_A8, edx
  loc_005B09DF: mov eax, 00000003h
loc_005B09E4:   lea edx, var_A0
loc_005B09EA:   lea ecx, var_50
loc_005B09ED:   mov var_B0, eax
  loc_005B09F3: mov var_C8, 00000005h
loc_005B09FD:   mov var_D0, eax
  loc_005B0A03: mov var_98, 00435A2Ch ; "##,##0"
  loc_005B0A0D: mov var_A0, 00000008h
  loc_005B0A17: call [00401238h] ; __vbaVarDup
  loc_005B0A1D: push 00000001h
loc_005B0A1F:   lea eax, var_50
  loc_005B0A22: push 00000001h
loc_005B0A24:   push eax
loc_005B0A25:   lea eax, [esi+00000070h]
loc_005B0A28:   lea ecx, var_60
loc_005B0A2B:   push eax
loc_005B0A2C:   push ecx
  loc_005B0A2D: call [00401078h] ; rtcVarFromFormatVar
loc_005B0A33:   lea edx, var_60
loc_005B0A36:   push edx
  loc_005B0A37: call [00401034h] ; __vbaStrVarMove
loc_005B0A3D:   mov edx, var_B0
  loc_005B0A43: sub esp, 00000010h
loc_005B0A46:   mov var_68, eax
  loc_005B0A49: mov var_70, 00000008h
loc_005B0A50:   mov ecx, esp
loc_005B0A52:   mov [ecx], edx
loc_005B0A54:   mov edx, var_AC
  loc_005B0A5A: sub esp, 00000010h
loc_005B0A5D:   mov [ecx+00000004h], edx
loc_005B0A60:   mov edx, var_A8
loc_005B0A66:   mov [ecx+00000008h], edx
loc_005B0A69:   mov edx, var_A4
loc_005B0A6F:   mov [ecx+0000000Ch], edx
loc_005B0A72:   mov edx, var_D0
loc_005B0A78:   mov ecx, esp
  loc_005B0A7A: sub esp, 00000010h
loc_005B0A7D:   mov [ecx], edx
loc_005B0A7F:   mov edx, var_CC
loc_005B0A85:   mov [ecx+00000004h], edx
loc_005B0A88:   mov edx, var_C8
loc_005B0A8E:   mov [ecx+00000008h], edx
loc_005B0A91:   mov edx, var_C4
loc_005B0A97:   mov [ecx+0000000Ch], edx
loc_005B0A9A:   mov edx, var_70
loc_005B0A9D:   mov ecx, esp
  loc_005B0A9F: push 00000002h
  loc_005B0AA1: push 00000041h
loc_005B0AA3:   push esi
loc_005B0AA4:   mov [ecx], edx
loc_005B0AA6:   mov edx, var_6C
loc_005B0AA9:   mov [ecx+00000004h], edx
loc_005B0AAC:   mov [ecx+00000008h], eax
loc_005B0AAF:   mov eax, var_64
loc_005B0AB2:   mov [ecx+0000000Ch], eax
loc_005B0AB5:   mov ecx, [esi]
loc_005B0AB7:   Call [ecx+00000390h]
loc_005B0ABD:   lea edx, var_30
loc_005B0AC0:   push eax
loc_005B0AC1:   push edx
loc_005B0AC2:   Call edi
loc_005B0AC4:   push eax
loc_005B0AC5:   Call ebx
  loc_005B0AC7: add esp, 0000003Ch
loc_005B0ACA:   lea ecx, var_30
  loc_005B0ACD: call [004012A0h] ; __vbaFreeObj
loc_005B0AD3:   lea eax, var_70
loc_005B0AD6:   lea ecx, var_60
loc_005B0AD9:   push eax
loc_005B0ADA:   lea edx, var_50
loc_005B0ADD:   push ecx
loc_005B0ADE:   push edx
  loc_005B0ADF: push 00000003h
  loc_005B0AE1: call [0040103Ch] ; __vbaFreeVarList
loc_005B0AE7:   movsx eax, [esi+00000034h]
loc_005B0AEB:   mov var_A8, eax
  loc_005B0AF1: mov eax, 00000003h
  loc_005B0AF6: add esp, 00000010h
loc_005B0AF9:   lea edx, var_A0
loc_005B0AFF:   lea ecx, var_50
loc_005B0B02:   mov var_B0, eax
  loc_005B0B08: mov var_C8, 00000006h
loc_005B0B12:   mov var_D0, eax
  loc_005B0B18: mov var_98, 00435A2Ch ; "##,##0"
  loc_005B0B22: mov var_A0, 00000008h
  loc_005B0B2C: call [00401238h] ; __vbaVarDup
  loc_005B0B32: push 00000001h
loc_005B0B34:   lea ecx, var_50
  loc_005B0B37: push 00000001h
loc_005B0B39:   lea eax, [esi+000000A0h]
loc_005B0B3F:   push ecx
loc_005B0B40:   lea edx, var_60
loc_005B0B43:   push eax
loc_005B0B44:   push edx
  loc_005B0B45: call [00401078h] ; rtcVarFromFormatVar
loc_005B0B4B:   lea eax, var_60
loc_005B0B4E:   push eax
  loc_005B0B4F: call [00401034h] ; __vbaStrVarMove
loc_005B0B55:   mov edx, var_B0
  loc_005B0B5B: sub esp, 00000010h
loc_005B0B5E:   mov ecx, esp
loc_005B0B60:   mov var_68, eax
  loc_005B0B63: mov var_70, 00000008h
loc_005B0B6A:   mov [ecx], edx
loc_005B0B6C:   mov edx, var_AC
loc_005B0B72:   mov [ecx+00000004h], edx
loc_005B0B75:   mov edx, var_A8
loc_005B0B7B:   mov [ecx+00000008h], edx
loc_005B0B7E:   mov edx, var_A4
  loc_005B0B84: sub esp, 00000010h
loc_005B0B87:   mov [ecx+0000000Ch], edx
loc_005B0B8A:   mov edx, var_D0
loc_005B0B90:   mov ecx, esp
  loc_005B0B92: sub esp, 00000010h
loc_005B0B95:   mov [ecx], edx
loc_005B0B97:   mov edx, var_CC
loc_005B0B9D:   mov [ecx+00000004h], edx
loc_005B0BA0:   mov edx, var_C8
loc_005B0BA6:   mov [ecx+00000008h], edx
loc_005B0BA9:   mov edx, var_C4
loc_005B0BAF:   mov [ecx+0000000Ch], edx
loc_005B0BB2:   mov edx, var_70
loc_005B0BB5:   mov ecx, esp
  loc_005B0BB7: push 00000002h
  loc_005B0BB9: push 00000041h
loc_005B0BBB:   push esi
loc_005B0BBC:   mov [ecx], edx
loc_005B0BBE:   mov edx, var_6C
loc_005B0BC1:   mov [ecx+00000004h], edx
loc_005B0BC4:   mov [ecx+00000008h], eax
loc_005B0BC7:   mov eax, var_64
loc_005B0BCA:   mov [ecx+0000000Ch], eax
loc_005B0BCD:   mov ecx, [esi]
loc_005B0BCF:   Call [ecx+00000390h]
loc_005B0BD5:   lea edx, var_30
loc_005B0BD8:   push eax
loc_005B0BD9:   push edx
loc_005B0BDA:   Call edi
loc_005B0BDC:   push eax
loc_005B0BDD:   Call ebx
  loc_005B0BDF: add esp, 0000003Ch
loc_005B0BE2:   lea ecx, var_30
  loc_005B0BE5: call [004012A0h] ; __vbaFreeObj
loc_005B0BEB:   lea eax, var_70
loc_005B0BEE:   lea ecx, var_60
loc_005B0BF1:   push eax
loc_005B0BF2:   lea edx, var_50
loc_005B0BF5:   push ecx
loc_005B0BF6:   push edx
  loc_005B0BF7: push 00000003h
  loc_005B0BF9: call [0040103Ch] ; __vbaFreeVarList
loc_005B0BFF:   movsx eax, [esi+00000034h]
loc_005B0C03:   mov var_A8, eax
  loc_005B0C09: mov eax, 00000003h
  loc_005B0C0E: add esp, 00000010h
loc_005B0C11:   lea edx, var_A0
loc_005B0C17:   lea ecx, var_50
loc_005B0C1A:   mov var_B0, eax
  loc_005B0C20: mov var_C8, 00000007h
loc_005B0C2A:   mov var_D0, eax
  loc_005B0C30: mov var_98, 00435A2Ch ; "##,##0"
  loc_005B0C3A: mov var_A0, 00000008h
  loc_005B0C44: call [00401238h] ; __vbaVarDup
  loc_005B0C4A: push 00000001h
loc_005B0C4C:   lea ecx, var_50
  loc_005B0C4F: push 00000001h
loc_005B0C51:   lea eax, [esi+000000B0h]
loc_005B0C57:   push ecx
loc_005B0C58:   lea edx, var_60
loc_005B0C5B:   push eax
loc_005B0C5C:   push edx
  loc_005B0C5D: call [00401078h] ; rtcVarFromFormatVar
loc_005B0C63:   lea eax, var_60
loc_005B0C66:   push eax
  loc_005B0C67: call [00401034h] ; __vbaStrVarMove
loc_005B0C6D:   mov edx, var_B0
  loc_005B0C73: sub esp, 00000010h
loc_005B0C76:   mov ecx, esp
loc_005B0C78:   mov var_68, eax
  loc_005B0C7B: mov var_70, 00000008h
  loc_005B0C82: sub esp, 00000010h
loc_005B0C85:   mov [ecx], edx
loc_005B0C87:   mov edx, var_AC
loc_005B0C8D:   mov [ecx+00000004h], edx
loc_005B0C90:   mov edx, var_A8
loc_005B0C96:   mov [ecx+00000008h], edx
loc_005B0C99:   mov edx, var_A4
loc_005B0C9F:   mov [ecx+0000000Ch], edx
loc_005B0CA2:   mov edx, var_D0
loc_005B0CA8:   mov ecx, esp
  loc_005B0CAA: sub esp, 00000010h
loc_005B0CAD:   mov [ecx], edx
loc_005B0CAF:   mov edx, var_CC
loc_005B0CB5:   mov [ecx+00000004h], edx
loc_005B0CB8:   mov edx, var_C8
loc_005B0CBE:   mov [ecx+00000008h], edx
loc_005B0CC1:   mov edx, var_C4
loc_005B0CC7:   mov [ecx+0000000Ch], edx
loc_005B0CCA:   mov edx, var_70
loc_005B0CCD:   mov ecx, esp
  loc_005B0CCF: push 00000002h
  loc_005B0CD1: push 00000041h
loc_005B0CD3:   push esi
loc_005B0CD4:   mov [ecx], edx
loc_005B0CD6:   mov edx, var_6C
loc_005B0CD9:   mov [ecx+00000004h], edx
loc_005B0CDC:   mov [ecx+00000008h], eax
loc_005B0CDF:   mov eax, var_64
loc_005B0CE2:   mov [ecx+0000000Ch], eax
loc_005B0CE5:   mov ecx, [esi]
loc_005B0CE7:   Call [ecx+00000390h]
loc_005B0CED:   lea edx, var_30
loc_005B0CF0:   push eax
loc_005B0CF1:   push edx
loc_005B0CF2:   Call edi
loc_005B0CF4:   push eax
loc_005B0CF5:   Call ebx
  loc_005B0CF7: add esp, 0000003Ch
loc_005B0CFA:   lea ecx, var_30
  loc_005B0CFD: call [004012A0h] ; __vbaFreeObj
loc_005B0D03:   lea eax, var_70
loc_005B0D06:   lea ecx, var_60
loc_005B0D09:   push eax
loc_005B0D0A:   lea edx, var_50
loc_005B0D0D:   push ecx
loc_005B0D0E:   push edx
  loc_005B0D0F: push 00000003h
  loc_005B0D11: call [0040103Ch] ; __vbaFreeVarList
loc_005B0D17:   mov ax, [esi+00000034h]
loc_005B0D1B:   mov ecx, [esi]
  loc_005B0D1D: add esp, 00000010h
  loc_005B0D20: add ax, 0001h
  loc_005B0D24: jo 005B0EECh
loc_005B0D2A:   push esi
loc_005B0D2B:   mov [esi+00000034h], ax
loc_005B0D2F:   Call [ecx+000006F8h]
loc_005B0D35:   test eax, eax
  loc_005B0D37: jge 005B0D4Bh
  loc_005B0D39: push 000006F8h
  loc_005B0D3E: push 00434B78h
loc_005B0D43:   push esi
loc_005B0D44:   push eax
  loc_005B0D45: call [00401080h] ; __vbaHresultCheckObj
loc_005B0D4B:   mov edx, [esi]
loc_005B0D4D:   push esi
loc_005B0D4E:   Call [edx+00000348h]
loc_005B0D54:   push eax
loc_005B0D55:   lea eax, var_30
loc_005B0D58:   push eax
loc_005B0D59:   Call edi
loc_005B0D5B:   mov ebx, eax
  loc_005B0D5D: push 0042DFECh
loc_005B0D62:   push ebx
loc_005B0D63:   mov ecx, [ebx]
loc_005B0D65:   Call [ecx+000000A4h]
loc_005B0D6B:   test eax, eax
loc_005B0D6D:   fnclex
  loc_005B0D6F: jge 005B0D83h
  loc_005B0D71: push 000000A4h
  loc_005B0D76: push 0042DFCCh
loc_005B0D7B:   push ebx
loc_005B0D7C:   push eax
  loc_005B0D7D: call [00401080h] ; __vbaHresultCheckObj
loc_005B0D83:   lea ecx, var_30
  loc_005B0D86: call [004012A0h] ; __vbaFreeObj
loc_005B0D8C:   mov edx, [esi]
loc_005B0D8E:   push esi
loc_005B0D8F:   Call [edx+0000033Ch]
loc_005B0D95:   push eax
loc_005B0D96:   lea eax, var_30
loc_005B0D99:   push eax
loc_005B0D9A:   Call edi
loc_005B0D9C:   mov ebx, eax
  loc_005B0D9E: push 00434A94h
loc_005B0DA3:   push ebx
loc_005B0DA4:   mov ecx, [ebx]
loc_005B0DA6:   Call [ecx+000000A4h]
loc_005B0DAC:   test eax, eax
loc_005B0DAE:   fnclex
  loc_005B0DB0: jge 005B0DC4h
  loc_005B0DB2: push 000000A4h
  loc_005B0DB7: push 0042DFCCh
loc_005B0DBC:   push ebx
loc_005B0DBD:   push eax
  loc_005B0DBE: call [00401080h] ; __vbaHresultCheckObj
loc_005B0DC4:   lea ecx, var_30
  loc_005B0DC7: call [004012A0h] ; __vbaFreeObj
loc_005B0DCD:   mov edx, [esi]
loc_005B0DCF:   push esi
loc_005B0DD0:   Call [edx+00000344h]
loc_005B0DD6:   push eax
loc_005B0DD7:   lea eax, var_30
loc_005B0DDA:   push eax
loc_005B0DDB:   Call edi
loc_005B0DDD:   mov ebx, eax
  loc_005B0DDF: push 0042DFECh
loc_005B0DE4:   push ebx
loc_005B0DE5:   mov ecx, [ebx]
loc_005B0DE7:   Call [ecx+000000A4h]
loc_005B0DED:   test eax, eax
loc_005B0DEF:   fnclex
  loc_005B0DF1: jge 005B0E05h
  loc_005B0DF3: push 000000A4h
  loc_005B0DF8: push 0042DFCCh
loc_005B0DFD:   push ebx
loc_005B0DFE:   push eax
  loc_005B0DFF: call [00401080h] ; __vbaHresultCheckObj
  loc_005B0E05: mov ebx, [004012A0h] ; __vbaFreeObj
loc_005B0E0B:   lea ecx, var_30
loc_005B0E0E:   Call ebx
loc_005B0E10:   mov edx, [esi]
loc_005B0E12:   push esi
loc_005B0E13:   Call [edx+00000340h]
loc_005B0E19:   push eax
loc_005B0E1A:   lea eax, var_30
loc_005B0E1D:   push eax
loc_005B0E1E:   Call edi
loc_005B0E20:   mov esi, eax
  loc_005B0E22: push 0042DFECh
loc_005B0E27:   push esi
loc_005B0E28:   mov ecx, [esi]
loc_005B0E2A:   Call [ecx+000000A4h]
loc_005B0E30:   test eax, eax
loc_005B0E32:   fnclex
  loc_005B0E34: jge 005B0E48h
  loc_005B0E36: push 000000A4h
  loc_005B0E3B: push 0042DFCCh
loc_005B0E40:   push esi
loc_005B0E41:   push eax
  loc_005B0E42: call [00401080h] ; __vbaHresultCheckObj
loc_005B0E48:   lea ecx, var_30
loc_005B0E4B:   Call ebx
  loc_005B0E4D: mov var_4, 00000000h
loc_005B0E54:   fwait
  loc_005B0E55: push 005B0EC8h
  loc_005B0E5A: jmp 005B0EBBh
loc_005B0E5C:   lea edx, var_2C
loc_005B0E5F:   lea eax, var_28
loc_005B0E62:   push edx
loc_005B0E63:   lea ecx, var_24
loc_005B0E66:   push eax
loc_005B0E67:   lea edx, var_20
loc_005B0E6A:   push ecx
loc_005B0E6B:   lea eax, var_1C
loc_005B0E6E:   push edx
loc_005B0E6F:   lea ecx, var_18
loc_005B0E72:   push eax
loc_005B0E73:   push ecx
  loc_005B0E74: push 00000006h
  loc_005B0E76: call [00401204h] ; __vbaFreeStrList
loc_005B0E7C:   lea edx, var_40
loc_005B0E7F:   lea eax, var_3C
loc_005B0E82:   push edx
loc_005B0E83:   lea ecx, var_38
loc_005B0E86:   push eax
loc_005B0E87:   lea edx, var_34
loc_005B0E8A:   push ecx
loc_005B0E8B:   lea eax, var_30
loc_005B0E8E:   push edx
loc_005B0E8F:   push eax
  loc_005B0E90: push 00000005h
  loc_005B0E92: call [0040104Ch] ; __vbaFreeObjList
loc_005B0E98:   lea ecx, var_90
loc_005B0E9E:   lea edx, var_80
loc_005B0EA1:   push ecx
loc_005B0EA2:   lea eax, var_70
loc_005B0EA5:   push edx
loc_005B0EA6:   lea ecx, var_60
loc_005B0EA9:   push eax
loc_005B0EAA:   lea edx, var_50
loc_005B0EAD:   push ecx
loc_005B0EAE:   push edx
  loc_005B0EAF: push 00000005h
  loc_005B0EB1: call [0040103Ch] ; __vbaFreeVarList
  loc_005B0EB7: add esp, 0000004Ch
loc_005B0EBA:   ret
loc_005B0EBB:   lea ecx, var_148
  loc_005B0EC1: call [004012A0h] ; __vbaFreeObj
loc_005B0EC7:   ret
loc_005B0EC8:   mov eax, Me
loc_005B0ECB:   push eax
loc_005B0ECC:   mov ecx, [eax]
loc_005B0ECE:   Call [ecx+00000008h]
loc_005B0ED1:   mov eax, var_4
loc_005B0ED4:   mov ecx, var_14
loc_005B0ED7:   pop edi
loc_005B0ED8:   pop esi
loc_005B0ED9:   mov fs: [00000000h] , ecx
loc_005B0EE0:   pop ebx
loc_005B0EE1:   mov esp, ebp
loc_005B0EE3:   pop ebp
  loc_005B0EE4: retn 0004h
End Sub

Private Sub txtKode_KeyPress(KeyAscii As Integer) '5B0F00
loc_005B0F00:   push ebp
loc_005B0F01:   mov ebp, esp
  loc_005B0F03: sub esp, 0000000Ch
  loc_005B0F06: push 00405696h ; __vbaExceptHandler
loc_005B0F0B:   mov eax, fs: [00000000h]
loc_005B0F11:   push eax
loc_005B0F12:   mov fs: [00000000h] , esp
  loc_005B0F19: sub esp, 0000003Ch
loc_005B0F1C:   push ebx
loc_005B0F1D:   push esi
loc_005B0F1E:   push edi
loc_005B0F1F:   mov var_C, esp
  loc_005B0F22: mov var_8, 00402F88h
loc_005B0F29:   mov edi, Me
loc_005B0F2C:   mov eax, edi
  loc_005B0F2E: and eax, 00000001h
loc_005B0F31:   mov var_4, eax
  loc_005B0F34: and edi, FFFFFFFEh
loc_005B0F37:   push edi
loc_005B0F38:   mov Me, edi
loc_005B0F3B:   mov ecx, [edi]
loc_005B0F3D:   Call [ecx+00000004h]
loc_005B0F40:   mov esi, KeyAscii
loc_005B0F43:   lea eax, var_2C
  loc_005B0F46: xor ebx, ebx
loc_005B0F48:   movsx edx, [esi]
loc_005B0F4B:   push edx
loc_005B0F4C:   push eax
loc_005B0F4D:   mov var_18, ebx
loc_005B0F50:   mov var_1C, ebx
loc_005B0F53:   mov var_2C, ebx
loc_005B0F56:   mov var_3C, ebx
  loc_005B0F59: call [004011B0h] ; rtcVarBstrFromAnsi
loc_005B0F5F:   lea ecx, var_2C
loc_005B0F62:   lea edx, var_3C
loc_005B0F65:   push ecx
loc_005B0F66:   push edx
  loc_005B0F67: call [00401124h] ; rtcUpperCaseVar
loc_005B0F6D:   lea eax, var_3C
loc_005B0F70:   lea ecx, var_18
loc_005B0F73:   push eax
loc_005B0F74:   push ecx
  loc_005B0F75: call [004011C0h] ; __vbaStrVarVal
loc_005B0F7B:   push eax
  loc_005B0F7C: call [00401050h] ; rtcAnsiValueBstr
loc_005B0F82:   lea ecx, var_18
loc_005B0F85:   mov [esi], ax
  loc_005B0F88: call [0040129Ch] ; __vbaFreeStr
loc_005B0F8E:   lea edx, var_3C
loc_005B0F91:   lea eax, var_2C
loc_005B0F94:   push edx
loc_005B0F95:   push eax
  loc_005B0F96: push 00000002h
  loc_005B0F98: call [0040103Ch] ; __vbaFreeVarList
  loc_005B0F9E: add esp, 0000000Ch
  loc_005B0FA1: cmp [esi], 000Dh
  loc_005B0FA5: jnz 005B1043h
loc_005B0FAB:   mov ecx, [edi]
loc_005B0FAD:   push edi
loc_005B0FAE:   Call [ecx+00000348h]
loc_005B0FB4:   lea edx, var_1C
loc_005B0FB7:   push eax
loc_005B0FB8:   push edx
  loc_005B0FB9: call [004010B8h] ; __vbaObjSet
loc_005B0FBF:   mov esi, eax
loc_005B0FC1:   lea ecx, var_18
loc_005B0FC4:   push ecx
loc_005B0FC5:   push esi
loc_005B0FC6:   mov eax, [esi]
loc_005B0FC8:   Call [eax+000000A0h]
loc_005B0FCE:   cmp eax, ebx
loc_005B0FD0:   fnclex
  loc_005B0FD2: jge 005B0FE6h
  loc_005B0FD4: push 000000A0h
  loc_005B0FD9: push 0042DFCCh
loc_005B0FDE:   push esi
loc_005B0FDF:   push eax
  loc_005B0FE0: call [00401080h] ; __vbaHresultCheckObj
loc_005B0FE6:   mov edx, var_18
loc_005B0FE9:   push edx
  loc_005B0FEA: push 0042DFECh
  loc_005B0FEF: call [0040112Ch] ; __vbaStrCmp
loc_005B0FF5:   mov esi, eax
loc_005B0FF7:   lea ecx, var_18
loc_005B0FFA:   neg esi
loc_005B0FFC:   sbb esi, esi
loc_005B0FFE:   inc esi
loc_005B0FFF:   neg esi
  loc_005B1001: call [0040129Ch] ; __vbaFreeStr
  loc_005B1007: mov ebx, [004012A0h] ; __vbaFreeObj
loc_005B100D:   lea ecx, var_1C
loc_005B1010:   Call ebx
loc_005B1012:   test si, si
  loc_005B1015: jz 005B1041h
loc_005B1017:   mov eax, [edi]
  loc_005B1019: push 00000000h
  loc_005B101B: push 80011000h
loc_005B1020:   push edi
loc_005B1021:   Call [eax+00000394h]
loc_005B1027:   lea ecx, var_1C
loc_005B102A:   push eax
loc_005B102B:   push ecx
  loc_005B102C: call [004010B8h] ; __vbaObjSet
loc_005B1032:   push eax
  loc_005B1033: call [0040102Ch] ; __vbaLateIdCall
  loc_005B1039: add esp, 0000000Ch
loc_005B103C:   lea ecx, var_1C
loc_005B103F:   Call ebx
  loc_005B1041: xor ebx, ebx
loc_005B1043:   mov var_4, ebx
  loc_005B1046: push 005B1074h
  loc_005B104B: jmp 005B1073h
loc_005B104D:   lea ecx, var_18
  loc_005B1050: call [0040129Ch] ; __vbaFreeStr
loc_005B1056:   lea ecx, var_1C
  loc_005B1059: call [004012A0h] ; __vbaFreeObj
loc_005B105F:   lea edx, var_3C
loc_005B1062:   lea eax, var_2C
loc_005B1065:   push edx
loc_005B1066:   push eax
  loc_005B1067: push 00000002h
  loc_005B1069: call [0040103Ch] ; __vbaFreeVarList
  loc_005B106F: add esp, 0000000Ch
loc_005B1072:   ret
loc_005B1073:   ret
loc_005B1074:   mov eax, Me
loc_005B1077:   push eax
loc_005B1078:   mov ecx, [eax]
loc_005B107A:   Call [ecx+00000008h]
loc_005B107D:   mov eax, var_4
loc_005B1080:   mov ecx, var_14
loc_005B1083:   pop edi
loc_005B1084:   pop esi
loc_005B1085:   mov fs: [00000000h] , ecx
loc_005B108C:   pop ebx
loc_005B108D:   mov esp, ebp
loc_005B108F:   pop ebp
  loc_005B1090: retn 0008h
End Sub

Private Sub cmbJenis_UnknownEvent_C() '5A6440
loc_005A6440:   push ebp
loc_005A6441:   mov ebp, esp
  loc_005A6443: sub esp, 0000000Ch
  loc_005A6446: push 00405696h ; __vbaExceptHandler
loc_005A644B:   mov eax, fs: [00000000h]
loc_005A6451:   push eax
loc_005A6452:   mov fs: [00000000h] , esp
  loc_005A6459: sub esp, 00000048h
loc_005A645C:   push ebx
loc_005A645D:   push esi
loc_005A645E:   push edi
loc_005A645F:   mov var_C, esp
  loc_005A6462: mov var_8, 00402D30h
loc_005A6469:   mov esi, Me
loc_005A646C:   mov eax, esi
  loc_005A646E: and eax, 00000001h
loc_005A6471:   mov var_4, eax
  loc_005A6474: and esi, FFFFFFFEh
loc_005A6477:   push esi
loc_005A6478:   mov Me, esi
loc_005A647B:   mov ecx, [esi]
loc_005A647D:   Call [ecx+00000004h]
loc_005A6480:   mov edx, [esi]
  loc_005A6482: xor eax, eax
loc_005A6484:   push eax
loc_005A6485:   push FFFFFDFBh
loc_005A648A:   push esi
loc_005A648B:   mov var_18, eax
loc_005A648E:   mov var_1C, eax
loc_005A6491:   mov var_2C, eax
loc_005A6494:   Call [edx+00000388h]
  loc_005A649A: mov ebx, [004010B8h] ; __vbaObjSet
loc_005A64A0:   push eax
loc_005A64A1:   lea eax, var_1C
loc_005A64A4:   push eax
loc_005A64A5:   Call ebx
loc_005A64A7:   lea ecx, var_2C
loc_005A64AA:   push eax
loc_005A64AB:   push ecx
  loc_005A64AC: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005A64B2: add esp, 00000010h
loc_005A64B5:   push eax
  loc_005A64B6: call [00401034h] ; __vbaStrVarMove
loc_005A64BC:   mov edx, eax
loc_005A64BE:   lea ecx, var_18
  loc_005A64C1: call [0040126Ch] ; __vbaStrMove
loc_005A64C7:   push eax
  loc_005A64C8: push 00433F74h ; "Tunai"
  loc_005A64CD: call [0040112Ch] ; __vbaStrCmp
loc_005A64D3:   mov edi, eax
loc_005A64D5:   lea ecx, var_18
loc_005A64D8:   neg edi
loc_005A64DA:   sbb edi, edi
loc_005A64DC:   inc edi
loc_005A64DD:   neg edi
  loc_005A64DF: call [0040129Ch] ; __vbaFreeStr
loc_005A64E5:   lea ecx, var_1C
  loc_005A64E8: call [004012A0h] ; __vbaFreeObj
loc_005A64EE:   lea ecx, var_2C
  loc_005A64F1: call [00401020h] ; __vbaFreeVar
loc_005A64F7:   test di, di
  loc_005A64FA: jz 005A6540h
loc_005A64FC:   mov edx, [esi]
loc_005A64FE:   push esi
loc_005A64FF:   Call [edx+00000320h]
loc_005A6505:   push eax
loc_005A6506:   lea eax, var_1C
loc_005A6509:   push eax
loc_005A650A:   Call ebx
loc_005A650C:   mov edi, eax
  loc_005A650E: push 00000000h
loc_005A6510:   push edi
loc_005A6511:   mov ecx, [edi]
loc_005A6513:   Call [ecx+0000009Ch]
loc_005A6519:   test eax, eax
loc_005A651B:   fnclex
  loc_005A651D: jge 005A6531h
  loc_005A651F: push 0000009Ch
  loc_005A6524: push 00433F80h
loc_005A6529:   push edi
loc_005A652A:   push eax
  loc_005A652B: call [00401080h] ; __vbaHresultCheckObj
  loc_005A6531: mov edi, [004012A0h] ; __vbaFreeObj
loc_005A6537:   lea ecx, var_1C
loc_005A653A:   Call edi
  loc_005A653C: xor eax, eax
  loc_005A653E: jmp 005A6583h
loc_005A6540:   mov eax, [esi]
loc_005A6542:   push esi
loc_005A6543:   Call [eax+00000320h]
loc_005A6549:   lea ecx, var_1C
loc_005A654C:   push eax
loc_005A654D:   push ecx
loc_005A654E:   Call ebx
loc_005A6550:   mov edi, eax
loc_005A6552:   push FFFFFFFFh
loc_005A6554:   push edi
loc_005A6555:   mov edx, [edi]
loc_005A6557:   Call [edx+0000009Ch]
loc_005A655D:   test eax, eax
loc_005A655F:   fnclex
  loc_005A6561: jge 005A6575h
  loc_005A6563: push 0000009Ch
  loc_005A6568: push 00433F80h
loc_005A656D:   push edi
loc_005A656E:   push eax
  loc_005A656F: call [00401080h] ; __vbaHresultCheckObj
  loc_005A6575: mov edi, [004012A0h] ; __vbaFreeObj
loc_005A657B:   lea ecx, var_1C
loc_005A657E:   Call edi
  loc_005A6580: or eax, FFFFFFFFh
  loc_005A6583: sub esp, 00000010h
  loc_005A6586: mov ecx, 0000000Bh
loc_005A658B:   mov edx, esp
  loc_005A658D: push 80010007h
loc_005A6592:   push esi
loc_005A6593:   mov [edx], ecx
loc_005A6595:   mov ecx, var_38
loc_005A6598:   mov [edx+00000004h], ecx
loc_005A659B:   mov ecx, [esi]
loc_005A659D:   mov [edx+00000008h], eax
loc_005A65A0:   mov eax, var_30
loc_005A65A3:   mov [edx+0000000Ch], eax
loc_005A65A6:   Call [ecx+00000384h]
loc_005A65AC:   lea edx, var_1C
loc_005A65AF:   push eax
loc_005A65B0:   push edx
loc_005A65B1:   Call ebx
loc_005A65B3:   push eax
  loc_005A65B4: call [00401280h] ; __vbaLateIdSt
loc_005A65BA:   lea ecx, var_1C
loc_005A65BD:   Call edi
  loc_005A65BF: mov var_4, 00000000h
  loc_005A65C6: push 005A65EAh
  loc_005A65CB: jmp 005A65E9h
loc_005A65CD:   lea ecx, var_18
  loc_005A65D0: call [0040129Ch] ; __vbaFreeStr
loc_005A65D6:   lea ecx, var_1C
  loc_005A65D9: call [004012A0h] ; __vbaFreeObj
loc_005A65DF:   lea ecx, var_2C
  loc_005A65E2: call [00401020h] ; __vbaFreeVar
loc_005A65E8:   ret
loc_005A65E9:   ret
loc_005A65EA:   mov eax, Me
loc_005A65ED:   push eax
loc_005A65EE:   mov ecx, [eax]
loc_005A65F0:   Call [ecx+00000008h]
loc_005A65F3:   mov eax, var_4
loc_005A65F6:   mov ecx, var_14
loc_005A65F9:   pop edi
loc_005A65FA:   pop esi
loc_005A65FB:   mov fs: [00000000h] , ecx
loc_005A6602:   pop ebx
loc_005A6603:   mov esp, ebp
loc_005A6605:   pop ebp
  loc_005A6606: retn 0004h
End Sub

Private Sub CmdSimpan_UnknownEvent_10() '5A98C0
loc_005A98C0:   push ebp
loc_005A98C1:   mov ebp, esp
  loc_005A98C3: sub esp, 00000018h
  loc_005A98C6: push 00405696h ; __vbaExceptHandler
loc_005A98CB:   mov eax, fs: [00000000h]
loc_005A98D1:   push eax
loc_005A98D2:   mov fs: [00000000h] , esp
  loc_005A98D9: mov eax, 00000590h
  loc_005A98DE: call 00405690h ; __vbaChkstk
loc_005A98E3:   push ebx
loc_005A98E4:   push esi
loc_005A98E5:   push edi
loc_005A98E6:   mov var_18, esp
  loc_005A98E9: mov var_14, 00402E08h ; "'"
loc_005A98F0:   mov eax, Me
  loc_005A98F3: and eax, 00000001h
loc_005A98F6:   mov var_10, eax
loc_005A98F9:   mov ecx, Me
  loc_005A98FC: and ecx, FFFFFFFEh
loc_005A98FF:   mov Me, ecx
  loc_005A9902: mov var_C, 00000000h
loc_005A9909:   mov edx, Me
loc_005A990C:   mov eax, [edx]
loc_005A990E:   mov ecx, Me
loc_005A9911:   push ecx
loc_005A9912:   Call [eax+00000004h]
  loc_005A9915: mov var_4, 00000001h
  loc_005A991C: mov var_4, 00000002h
loc_005A9923:   mov edx, Me
loc_005A9926:   mov eax, [edx]
loc_005A9928:   mov ecx, Me
loc_005A992B:   push ecx
loc_005A992C:   Call [eax+0000030Ch]
loc_005A9932:   push eax
loc_005A9933:   lea edx, var_64
loc_005A9936:   push edx
  loc_005A9937: call [004010B8h] ; __vbaObjSet
loc_005A993D:   mov var_434, eax
loc_005A9943:   lea eax, var_38
loc_005A9946:   push eax
loc_005A9947:   mov ecx, var_434
loc_005A994D:   mov edx, [ecx]
loc_005A994F:   mov eax, var_434
loc_005A9955:   push eax
loc_005A9956:   Call [edx+000000A0h]
loc_005A995C:   fnclex
loc_005A995E:   mov var_438, eax
  loc_005A9964: cmp var_438, 00000000h
  loc_005A996B: jge 005A9993h
  loc_005A996D: push 000000A0h
  loc_005A9972: push 0042DFCCh
loc_005A9977:   mov ecx, var_434
loc_005A997D:   push ecx
loc_005A997E:   mov edx, var_438
loc_005A9984:   push edx
  loc_005A9985: call [00401080h] ; __vbaHresultCheckObj
loc_005A998B:   mov var_4B4, eax
  loc_005A9991: jmp 005A999Dh
  loc_005A9993: mov var_4B4, 00000000h
loc_005A999D:   mov eax, var_38
loc_005A99A0:   push eax
  loc_005A99A1: push 0042DFECh
  loc_005A99A6: call [0040112Ch] ; __vbaStrCmp
loc_005A99AC:   neg eax
loc_005A99AE:   sbb eax, eax
loc_005A99B0:   inc eax
loc_005A99B1:   neg eax
loc_005A99B3:   mov var_43C, ax
loc_005A99BA:   lea ecx, var_38
  loc_005A99BD: call [0040129Ch] ; __vbaFreeStr
loc_005A99C3:   lea ecx, var_64
  loc_005A99C6: call [004012A0h] ; __vbaFreeObj
loc_005A99CC:   movsx ecx, var_43C
loc_005A99D3:   test ecx, ecx
  loc_005A99D5: jz 005A9B2Ch
  loc_005A99DB: mov var_4, 00000003h
  loc_005A99E2: mov var_B0, 80020004h
  loc_005A99EC: mov var_B8, 0000000Ah
  loc_005A99F6: mov var_A0, 80020004h
  loc_005A9A00: mov var_A8, 0000000Ah
  loc_005A9A0A: mov var_260, 0042E624h ; "Error"
  loc_005A9A14: mov var_268, 00000008h
loc_005A9A1E:   lea edx, var_268
loc_005A9A24:   lea ecx, var_98
  loc_005A9A2A: call [00401238h] ; __vbaVarDup
  loc_005A9A30: mov var_250, 00435C70h ; "Nota transaksi kosong !"
  loc_005A9A3A: mov var_258, 00000008h
loc_005A9A44:   lea edx, var_258
loc_005A9A4A:   lea ecx, var_88
  loc_005A9A50: call [00401238h] ; __vbaVarDup
loc_005A9A56:   lea edx, var_B8
loc_005A9A5C:   push edx
loc_005A9A5D:   lea eax, var_A8
loc_005A9A63:   push eax
loc_005A9A64:   lea ecx, var_98
loc_005A9A6A:   push ecx
  loc_005A9A6B: push 00000010h
loc_005A9A6D:   lea edx, var_88
loc_005A9A73:   push edx
  loc_005A9A74: call [004010B4h] ; rtcMsgBox
loc_005A9A7A:   lea eax, var_B8
loc_005A9A80:   push eax
loc_005A9A81:   lea ecx, var_A8
loc_005A9A87:   push ecx
loc_005A9A88:   lea edx, var_98
loc_005A9A8E:   push edx
loc_005A9A8F:   lea eax, var_88
loc_005A9A95:   push eax
  loc_005A9A96: push 00000004h
  loc_005A9A98: call [0040103Ch] ; __vbaFreeVarList
  loc_005A9A9E: add esp, 00000014h
  loc_005A9AA1: mov var_4, 00000004h
loc_005A9AA8:   mov ecx, Me
loc_005A9AAB:   mov edx, [ecx]
loc_005A9AAD:   mov eax, Me
loc_005A9AB0:   push eax
loc_005A9AB1:   Call [edx+0000030Ch]
loc_005A9AB7:   push eax
loc_005A9AB8:   lea ecx, var_64
loc_005A9ABB:   push ecx
  loc_005A9ABC: call [004010B8h] ; __vbaObjSet
loc_005A9AC2:   mov var_434, eax
loc_005A9AC8:   mov edx, var_434
loc_005A9ACE:   mov eax, [edx]
loc_005A9AD0:   mov ecx, var_434
loc_005A9AD6:   push ecx
loc_005A9AD7:   Call [eax+00000204h]
loc_005A9ADD:   fnclex
loc_005A9ADF:   mov var_438, eax
  loc_005A9AE5: cmp var_438, 00000000h
  loc_005A9AEC: jge 005A9B14h
  loc_005A9AEE: push 00000204h
  loc_005A9AF3: push 0042DFCCh
loc_005A9AF8:   mov edx, var_434
loc_005A9AFE:   push edx
loc_005A9AFF:   mov eax, var_438
loc_005A9B05:   push eax
  loc_005A9B06: call [00401080h] ; __vbaHresultCheckObj
loc_005A9B0C:   mov var_4B8, eax
  loc_005A9B12: jmp 005A9B1Eh
  loc_005A9B14: mov var_4B8, 00000000h
loc_005A9B1E:   lea ecx, var_64
  loc_005A9B21: call [004012A0h] ; __vbaFreeObj
  loc_005A9B27: jmp 005AEA44h
  loc_005A9B2C: mov var_4, 00000005h
loc_005A9B33:   mov ecx, Me
  loc_005A9B36: cmp [ecx+00000034h], 0001h
  loc_005A9B3B: jnz 005A9C47h
  loc_005A9B41: mov var_4, 00000006h
  loc_005A9B48: mov var_B0, 80020004h
  loc_005A9B52: mov var_B8, 0000000Ah
  loc_005A9B5C: mov var_A0, 80020004h
  loc_005A9B66: mov var_A8, 0000000Ah
  loc_005A9B70: mov var_260, 0042E624h ; "Error"
  loc_005A9B7A: mov var_268, 00000008h
loc_005A9B84:   lea edx, var_268
loc_005A9B8A:   lea ecx, var_98
  loc_005A9B90: call [00401238h] ; __vbaVarDup
  loc_005A9B96: mov var_250, 00435CA4h ; "Grid barang kosong"
  loc_005A9BA0: mov var_258, 00000008h
loc_005A9BAA:   lea edx, var_258
loc_005A9BB0:   lea ecx, var_88
  loc_005A9BB6: call [00401238h] ; __vbaVarDup
loc_005A9BBC:   lea edx, var_B8
loc_005A9BC2:   push edx
loc_005A9BC3:   lea eax, var_A8
loc_005A9BC9:   push eax
loc_005A9BCA:   lea ecx, var_98
loc_005A9BD0:   push ecx
  loc_005A9BD1: push 00000010h
loc_005A9BD3:   lea edx, var_88
loc_005A9BD9:   push edx
  loc_005A9BDA: call [004010B4h] ; rtcMsgBox
loc_005A9BE0:   lea eax, var_B8
loc_005A9BE6:   push eax
loc_005A9BE7:   lea ecx, var_A8
loc_005A9BED:   push ecx
loc_005A9BEE:   lea edx, var_98
loc_005A9BF4:   push edx
loc_005A9BF5:   lea eax, var_88
loc_005A9BFB:   push eax
  loc_005A9BFC: push 00000004h
  loc_005A9BFE: call [0040103Ch] ; __vbaFreeVarList
  loc_005A9C04: add esp, 00000014h
  loc_005A9C07: mov var_4, 00000007h
  loc_005A9C0E: push 00000000h
  loc_005A9C10: push 80011000h
loc_005A9C15:   mov ecx, Me
loc_005A9C18:   mov edx, [ecx]
loc_005A9C1A:   mov eax, Me
loc_005A9C1D:   push eax
loc_005A9C1E:   Call [edx+00000380h]
loc_005A9C24:   push eax
loc_005A9C25:   lea ecx, var_64
loc_005A9C28:   push ecx
  loc_005A9C29: call [004010B8h] ; __vbaObjSet
loc_005A9C2F:   push eax
  loc_005A9C30: call [0040102Ch] ; __vbaLateIdCall
  loc_005A9C36: add esp, 0000000Ch
loc_005A9C39:   lea ecx, var_64
  loc_005A9C3C: call [004012A0h] ; __vbaFreeObj
  loc_005A9C42: jmp 005AEA44h
  loc_005A9C47: mov var_4, 00000009h
  loc_005A9C4E: call 0055B320h
  loc_005A9C53: mov var_4, 0000000Ah
loc_005A9C5A:   mov edx, Me
loc_005A9C5D:   mov eax, [edx]
loc_005A9C5F:   mov ecx, Me
loc_005A9C62:   push ecx
loc_005A9C63:   Call [eax+00000710h]
loc_005A9C69:   mov var_434, eax
  loc_005A9C6F: cmp var_434, 00000000h
  loc_005A9C76: jge 005A9C9Bh
  loc_005A9C78: push 00000710h
  loc_005A9C7D: push 00434B78h
loc_005A9C82:   mov edx, Me
loc_005A9C85:   push edx
loc_005A9C86:   mov eax, var_434
loc_005A9C8C:   push eax
  loc_005A9C8D: call [00401080h] ; __vbaHresultCheckObj
loc_005A9C93:   mov var_4BC, eax
  loc_005A9C99: jmp 005A9CA5h
  loc_005A9C9B: mov var_4BC, 00000000h
  loc_005A9CA5: mov var_4, 0000000Bh
loc_005A9CAC:   mov ecx, Me
loc_005A9CAF:   mov edx, [ecx]
loc_005A9CB1:   mov eax, Me
loc_005A9CB4:   push eax
loc_005A9CB5:   Call [edx+00000338h]
loc_005A9CBB:   push eax
loc_005A9CBC:   lea ecx, var_64
loc_005A9CBF:   push ecx
  loc_005A9CC0: call [004010B8h] ; __vbaObjSet
loc_005A9CC6:   mov var_434, eax
loc_005A9CCC:   lea edx, var_38
loc_005A9CCF:   push edx
loc_005A9CD0:   mov eax, var_434
loc_005A9CD6:   mov ecx, [eax]
loc_005A9CD8:   mov edx, var_434
loc_005A9CDE:   push edx
loc_005A9CDF:   Call [ecx+00000050h]
loc_005A9CE2:   fnclex
loc_005A9CE4:   mov var_438, eax
  loc_005A9CEA: cmp var_438, 00000000h
  loc_005A9CF1: jge 005A9D16h
  loc_005A9CF3: push 00000050h
  loc_005A9CF5: push 00433F80h
loc_005A9CFA:   mov eax, var_434
loc_005A9D00:   push eax
loc_005A9D01:   mov ecx, var_438
loc_005A9D07:   push ecx
  loc_005A9D08: call [00401080h] ; __vbaHresultCheckObj
loc_005A9D0E:   mov var_4C0, eax
  loc_005A9D14: jmp 005A9D20h
  loc_005A9D16: mov var_4C0, 00000000h
loc_005A9D20:   mov edx, Me
loc_005A9D23:   mov eax, [edx]
loc_005A9D25:   mov ecx, Me
loc_005A9D28:   push ecx
loc_005A9D29:   Call [eax+00000300h]
loc_005A9D2F:   push eax
loc_005A9D30:   lea edx, var_68
loc_005A9D33:   push edx
  loc_005A9D34: call [004010B8h] ; __vbaObjSet
loc_005A9D3A:   mov var_43C, eax
loc_005A9D40:   lea eax, var_40
loc_005A9D43:   push eax
loc_005A9D44:   mov ecx, var_43C
loc_005A9D4A:   mov edx, [ecx]
loc_005A9D4C:   mov eax, var_43C
loc_005A9D52:   push eax
loc_005A9D53:   Call [edx+00000050h]
loc_005A9D56:   fnclex
loc_005A9D58:   mov var_440, eax
  loc_005A9D5E: cmp var_440, 00000000h
  loc_005A9D65: jge 005A9D8Ah
  loc_005A9D67: push 00000050h
  loc_005A9D69: push 00433F80h
loc_005A9D6E:   mov ecx, var_43C
loc_005A9D74:   push ecx
loc_005A9D75:   mov edx, var_440
loc_005A9D7B:   push edx
  loc_005A9D7C: call [00401080h] ; __vbaHresultCheckObj
loc_005A9D82:   mov var_4C4, eax
  loc_005A9D88: jmp 005A9D94h
  loc_005A9D8A: mov var_4C4, 00000000h
  loc_005A9D94: mov var_260, 00435A20h ; "00"
  loc_005A9D9E: mov var_268, 00000008h
loc_005A9DA8:   lea edx, var_268
loc_005A9DAE:   lea ecx, var_C8
  loc_005A9DB4: call [00401238h] ; __vbaVarDup
loc_005A9DBA:   mov eax, var_40
loc_005A9DBD:   mov var_478, eax
  loc_005A9DC3: mov var_40, 00000000h
loc_005A9DCA:   mov ecx, var_478
loc_005A9DD0:   mov var_B0, ecx
  loc_005A9DD6: mov var_B8, 00000008h
  loc_005A9DE0: push 00000001h
  loc_005A9DE2: push 00000001h
loc_005A9DE4:   lea edx, var_C8
loc_005A9DEA:   push edx
loc_005A9DEB:   lea eax, var_B8
loc_005A9DF1:   push eax
loc_005A9DF2:   lea ecx, var_D8
loc_005A9DF8:   push ecx
  loc_005A9DF9: call [00401078h] ; rtcVarFromFormatVar
loc_005A9DFF:   lea edx, var_D8
loc_005A9E05:   push edx
loc_005A9E06:   lea eax, var_44
loc_005A9E09:   push eax
  loc_005A9E0A: call [004011C0h] ; __vbaStrVarVal
loc_005A9E10:   push eax
  loc_005A9E11: call [004012A4h] ; rtcR8ValFromBstr
  loc_005A9E17: fstp real8 ptr var_430
  loc_005A9E1D: mov var_250, 00435A20h ; "00"
  loc_005A9E27: mov var_258, 00000008h
loc_005A9E31:   lea edx, var_258
loc_005A9E37:   lea ecx, var_98
  loc_005A9E3D: call [00401238h] ; __vbaVarDup
loc_005A9E43:   mov ecx, var_38
loc_005A9E46:   mov var_47C, ecx
  loc_005A9E4C: mov var_38, 00000000h
loc_005A9E53:   mov edx, var_47C
loc_005A9E59:   mov var_80, edx
  loc_005A9E5C: mov var_88, 00000008h
  loc_005A9E66: push 00000001h
  loc_005A9E68: push 00000001h
loc_005A9E6A:   lea eax, var_98
loc_005A9E70:   push eax
loc_005A9E71:   lea ecx, var_88
loc_005A9E77:   push ecx
loc_005A9E78:   lea edx, var_A8
loc_005A9E7E:   push edx
  loc_005A9E7F: call [00401078h] ; rtcVarFromFormatVar
loc_005A9E85:   lea eax, var_A8
loc_005A9E8B:   push eax
loc_005A9E8C:   lea ecx, var_3C
loc_005A9E8F:   push ecx
  loc_005A9E90: call [004011C0h] ; __vbaStrVarVal
loc_005A9E96:   push eax
  loc_005A9E97: call [004012A4h] ; rtcR8ValFromBstr
  loc_005A9E9D: fsub st0, real8 ptr var_430
loc_005A9EA3:   fnstsw ax
  loc_005A9EA5: test al, 0Dh
  loc_005A9EA7: jnz 005AEBCCh
  loc_005A9EAD: call [00401224h] ; __vbaFpCy
loc_005A9EB3:   mov var_34, eax
loc_005A9EB6:   mov var_30, edx
loc_005A9EB9:   lea edx, var_44
loc_005A9EBC:   push edx
loc_005A9EBD:   lea eax, var_3C
loc_005A9EC0:   push eax
  loc_005A9EC1: push 00000002h
  loc_005A9EC3: call [00401204h] ; __vbaFreeStrList
  loc_005A9EC9: add esp, 0000000Ch
loc_005A9ECC:   lea ecx, var_68
loc_005A9ECF:   push ecx
loc_005A9ED0:   lea edx, var_64
loc_005A9ED3:   push edx
  loc_005A9ED4: push 00000002h
  loc_005A9ED6: call [0040104Ch] ; __vbaFreeObjList
  loc_005A9EDC: add esp, 0000000Ch
loc_005A9EDF:   lea eax, var_D8
loc_005A9EE5:   push eax
loc_005A9EE6:   lea ecx, var_C8
loc_005A9EEC:   push ecx
loc_005A9EED:   lea edx, var_B8
loc_005A9EF3:   push edx
loc_005A9EF4:   lea eax, var_A8
loc_005A9EFA:   push eax
loc_005A9EFB:   lea ecx, var_98
loc_005A9F01:   push ecx
loc_005A9F02:   lea edx, var_88
loc_005A9F08:   push edx
  loc_005A9F09: push 00000006h
  loc_005A9F0B: call [0040103Ch] ; __vbaFreeVarList
  loc_005A9F11: add esp, 0000001Ch
  loc_005A9F14: mov var_4, 0000000Ch
  loc_005A9F1B: push 00000000h
loc_005A9F1D:   push FFFFFDFBh
loc_005A9F22:   mov eax, Me
loc_005A9F25:   mov ecx, [eax]
loc_005A9F27:   mov edx, Me
loc_005A9F2A:   push edx
loc_005A9F2B:   Call [ecx+00000388h]
loc_005A9F31:   push eax
loc_005A9F32:   lea eax, var_64
loc_005A9F35:   push eax
  loc_005A9F36: call [004010B8h] ; __vbaObjSet
loc_005A9F3C:   push eax
loc_005A9F3D:   lea ecx, var_88
loc_005A9F43:   push ecx
  loc_005A9F44: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005A9F4A: add esp, 00000010h
loc_005A9F4D:   push eax
  loc_005A9F4E: call [00401034h] ; __vbaStrVarMove
loc_005A9F54:   mov edx, eax
loc_005A9F56:   lea ecx, var_38
  loc_005A9F59: call [0040126Ch] ; __vbaStrMove
loc_005A9F5F:   push eax
  loc_005A9F60: push 00433F74h ; "Tunai"
  loc_005A9F65: call [0040112Ch] ; __vbaStrCmp
loc_005A9F6B:   neg eax
loc_005A9F6D:   sbb eax, eax
loc_005A9F6F:   inc eax
loc_005A9F70:   neg eax
loc_005A9F72:   mov var_434, ax
loc_005A9F79:   lea ecx, var_38
  loc_005A9F7C: call [0040129Ch] ; __vbaFreeStr
loc_005A9F82:   lea ecx, var_64
  loc_005A9F85: call [004012A0h] ; __vbaFreeObj
loc_005A9F8B:   lea ecx, var_88
  loc_005A9F91: call [00401020h] ; __vbaFreeVar
loc_005A9F97:   movsx edx, var_434
loc_005A9F9E:   test edx, edx
  loc_005A9FA0: jz 005AC00Ah
  loc_005A9FA6: mov var_4, 0000000Dh
  loc_005A9FAD: mov edx, 00433F74h ; "Tunai"
  loc_005A9FB2: mov ecx, 006680D4h
  loc_005A9FB7: call [004011FCh] ; __vbaStrCopy
  loc_005A9FBD: mov var_4, 0000000Eh
loc_005A9FC4:   lea eax, var_88
loc_005A9FCA:   push eax
  loc_005A9FCB: call [00401220h] ; rtcGetDateVar
loc_005A9FD1:   mov ecx, Me
loc_005A9FD4:   mov edx, [ecx]
loc_005A9FD6:   mov eax, Me
loc_005A9FD9:   push eax
loc_005A9FDA:   Call [edx+00000338h]
loc_005A9FE0:   push eax
loc_005A9FE1:   lea ecx, var_74
loc_005A9FE4:   push ecx
  loc_005A9FE5: call [004010B8h] ; __vbaObjSet
loc_005A9FEB:   mov var_434, eax
loc_005A9FF1:   lea edx, var_4C
loc_005A9FF4:   push edx
loc_005A9FF5:   mov eax, var_434
loc_005A9FFB:   mov ecx, [eax]
loc_005A9FFD:   mov edx, var_434
loc_005AA003:   push edx
loc_005AA004:   Call [ecx+00000050h]
loc_005AA007:   fnclex
loc_005AA009:   mov var_438, eax
  loc_005AA00F: cmp var_438, 00000000h
  loc_005AA016: jge 005AA03Bh
  loc_005AA018: push 00000050h
  loc_005AA01A: push 00433F80h
loc_005AA01F:   mov eax, var_434
loc_005AA025:   push eax
loc_005AA026:   mov ecx, var_438
loc_005AA02C:   push ecx
  loc_005AA02D: call [00401080h] ; __vbaHresultCheckObj
loc_005AA033:   mov var_4C8, eax
  loc_005AA039: jmp 005AA045h
  loc_005AA03B: mov var_4C8, 00000000h
  loc_005AA045: push 00435CD0h ; "INSERT INTO Penjualan"
  loc_005AA04A: push 00435D00h ; "(no_nota,tgl_nota,kode_pelanggan,nama_pelanggan,jenis,userid,catatan,tot_jual,Kode_Perusahaan)"
  loc_005AA04F: call [00401070h] ; __vbaStrCat
loc_005AA055:   mov edx, eax
loc_005AA057:   lea ecx, var_38
  loc_005AA05A: call [0040126Ch] ; __vbaStrMove
loc_005AA060:   push eax
  loc_005AA061: push 00434520h ; "VALUES ('"
  loc_005AA066: call [00401070h] ; __vbaStrCat
loc_005AA06C:   mov edx, eax
loc_005AA06E:   lea ecx, var_3C
  loc_005AA071: call [0040126Ch] ; __vbaStrMove
loc_005AA077:   push eax
loc_005AA078:   mov edx, Me
loc_005AA07B:   mov eax, [edx+0000006Ch]
loc_005AA07E:   push eax
  loc_005AA07F: call [00401070h] ; __vbaStrCat
loc_005AA085:   mov edx, eax
loc_005AA087:   lea ecx, var_40
  loc_005AA08A: call [0040126Ch] ; __vbaStrMove
loc_005AA090:   push eax
  loc_005AA091: push 0042EC30h ; "','"
  loc_005AA096: call [00401070h] ; __vbaStrCat
loc_005AA09C:   mov var_B0, eax
  loc_005AA0A2: mov var_B8, 00000008h
  loc_005AA0AC: mov var_250, 00434538h ; "yyyy/MM/dd"
  loc_005AA0B6: mov var_258, 00000008h
loc_005AA0C0:   lea edx, var_258
loc_005AA0C6:   lea ecx, var_98
  loc_005AA0CC: call [00401238h] ; __vbaVarDup
  loc_005AA0D2: push 00000001h
  loc_005AA0D4: push 00000001h
loc_005AA0D6:   lea ecx, var_98
loc_005AA0DC:   push ecx
loc_005AA0DD:   lea edx, var_88
loc_005AA0E3:   push edx
loc_005AA0E4:   lea eax, var_A8
loc_005AA0EA:   push eax
  loc_005AA0EB: call [00401078h] ; rtcVarFromFormatVar
  loc_005AA0F1: mov var_260, 0042EC30h ; "','"
  loc_005AA0FB: mov var_268, 00000008h
  loc_005AA105: push 00000000h
loc_005AA107:   push FFFFFDFBh
loc_005AA10C:   mov ecx, Me
loc_005AA10F:   mov edx, [ecx]
loc_005AA111:   mov eax, Me
loc_005AA114:   push eax
loc_005AA115:   Call [edx+0000038Ch]
loc_005AA11B:   push eax
loc_005AA11C:   lea ecx, var_64
loc_005AA11F:   push ecx
  loc_005AA120: call [004010B8h] ; __vbaObjSet
loc_005AA126:   push eax
loc_005AA127:   lea edx, var_E8
loc_005AA12D:   push edx
  loc_005AA12E: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AA134: add esp, 00000010h
loc_005AA137:   push eax
  loc_005AA138: call [00401034h] ; __vbaStrVarMove
loc_005AA13E:   mov var_F0, eax
  loc_005AA144: mov var_F8, 00000008h
  loc_005AA14E: mov var_270, 0042EC30h ; "','"
  loc_005AA158: mov var_278, 00000008h
loc_005AA162:   mov eax, Me
loc_005AA165:   mov ecx, [eax]
loc_005AA167:   mov edx, Me
loc_005AA16A:   push edx
loc_005AA16B:   Call [ecx+00000310h]
loc_005AA171:   push eax
loc_005AA172:   lea eax, var_68
loc_005AA175:   push eax
  loc_005AA176: call [004010B8h] ; __vbaObjSet
loc_005AA17C:   mov var_43C, eax
loc_005AA182:   lea ecx, var_44
loc_005AA185:   push ecx
loc_005AA186:   mov edx, var_43C
loc_005AA18C:   mov eax, [edx]
loc_005AA18E:   mov ecx, var_43C
loc_005AA194:   push ecx
loc_005AA195:   Call [eax+000000A0h]
loc_005AA19B:   fnclex
loc_005AA19D:   mov var_440, eax
  loc_005AA1A3: cmp var_440, 00000000h
  loc_005AA1AA: jge 005AA1D2h
  loc_005AA1AC: push 000000A0h
  loc_005AA1B1: push 0042DFCCh
loc_005AA1B6:   mov edx, var_43C
loc_005AA1BC:   push edx
loc_005AA1BD:   mov eax, var_440
loc_005AA1C3:   push eax
  loc_005AA1C4: call [00401080h] ; __vbaHresultCheckObj
loc_005AA1CA:   mov var_4CC, eax
  loc_005AA1D0: jmp 005AA1DCh
  loc_005AA1D2: mov var_4CC, 00000000h
loc_005AA1DC:   mov ecx, var_44
loc_005AA1DF:   mov var_480, ecx
  loc_005AA1E5: mov var_44, 00000000h
loc_005AA1EC:   mov edx, var_480
loc_005AA1F2:   mov var_120, edx
  loc_005AA1F8: mov var_128, 00000008h
  loc_005AA202: mov var_280, 0042EC30h ; "','"
  loc_005AA20C: mov var_288, 00000008h
  loc_005AA216: push 00000000h
loc_005AA218:   push FFFFFDFBh
loc_005AA21D:   mov eax, Me
loc_005AA220:   mov ecx, [eax]
loc_005AA222:   mov edx, Me
loc_005AA225:   push edx
loc_005AA226:   Call [ecx+00000388h]
loc_005AA22C:   push eax
loc_005AA22D:   lea eax, var_6C
loc_005AA230:   push eax
  loc_005AA231: call [004010B8h] ; __vbaObjSet
loc_005AA237:   push eax
loc_005AA238:   lea ecx, var_158
loc_005AA23E:   push ecx
  loc_005AA23F: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AA245: add esp, 00000010h
loc_005AA248:   push eax
  loc_005AA249: call [00401034h] ; __vbaStrVarMove
loc_005AA24F:   mov var_160, eax
  loc_005AA255: mov var_168, 00000008h
  loc_005AA25F: mov var_290, 0042EC30h ; "','"
  loc_005AA269: mov var_298, 00000008h
  loc_005AA273: mov var_2A0, 0042EC30h ; "','"
  loc_005AA27D: mov var_2A8, 00000008h
loc_005AA287:   mov edx, Me
loc_005AA28A:   mov eax, [edx]
loc_005AA28C:   mov ecx, Me
loc_005AA28F:   push ecx
loc_005AA290:   Call [eax+00000314h]
loc_005AA296:   push eax
loc_005AA297:   lea edx, var_70
loc_005AA29A:   push edx
  loc_005AA29B: call [004010B8h] ; __vbaObjSet
loc_005AA2A1:   mov var_444, eax
loc_005AA2A7:   lea eax, var_48
loc_005AA2AA:   push eax
loc_005AA2AB:   mov ecx, var_444
loc_005AA2B1:   mov edx, [ecx]
loc_005AA2B3:   mov eax, var_444
loc_005AA2B9:   push eax
loc_005AA2BA:   Call [edx+000000A0h]
loc_005AA2C0:   fnclex
loc_005AA2C2:   mov var_448, eax
  loc_005AA2C8: cmp var_448, 00000000h
  loc_005AA2CF: jge 005AA2F7h
  loc_005AA2D1: push 000000A0h
  loc_005AA2D6: push 0042DFCCh
loc_005AA2DB:   mov ecx, var_444
loc_005AA2E1:   push ecx
loc_005AA2E2:   mov edx, var_448
loc_005AA2E8:   push edx
  loc_005AA2E9: call [00401080h] ; __vbaHresultCheckObj
loc_005AA2EF:   mov var_4D0, eax
  loc_005AA2F5: jmp 005AA301h
  loc_005AA2F7: mov var_4D0, 00000000h
loc_005AA301:   mov eax, var_48
loc_005AA304:   mov var_484, eax
  loc_005AA30A: mov var_48, 00000000h
loc_005AA311:   mov ecx, var_484
loc_005AA317:   mov var_1B0, ecx
  loc_005AA31D: mov var_1B8, 00000008h
  loc_005AA327: mov var_2B0, 0042EC30h ; "','"
  loc_005AA331: mov var_2B8, 00000008h
  loc_005AA33B: mov var_2C0, 00435A20h ; "00"
  loc_005AA345: mov var_2C8, 00000008h
loc_005AA34F:   lea edx, var_2C8
loc_005AA355:   lea ecx, var_1F8
  loc_005AA35B: call [00401238h] ; __vbaVarDup
loc_005AA361:   mov edx, var_4C
loc_005AA364:   mov var_488, edx
  loc_005AA36A: mov var_4C, 00000000h
loc_005AA371:   mov eax, var_488
loc_005AA377:   mov var_1E0, eax
  loc_005AA37D: mov var_1E8, 00000008h
  loc_005AA387: push 00000001h
  loc_005AA389: push 00000001h
loc_005AA38B:   lea ecx, var_1F8
loc_005AA391:   push ecx
loc_005AA392:   lea edx, var_1E8
loc_005AA398:   push edx
loc_005AA399:   lea eax, var_208
loc_005AA39F:   push eax
  loc_005AA3A0: call [00401078h] ; rtcVarFromFormatVar
  loc_005AA3A6: mov var_2D0, 0042EC30h ; "','"
  loc_005AA3B0: mov var_2D8, 00000008h
  loc_005AA3BA: mov var_2E0, 00435DC4h ; "P1"
  loc_005AA3C4: mov var_2E8, 00000008h
  loc_005AA3CE: mov var_2F0, 0042EC3Ch ; "')"
  loc_005AA3D8: mov var_2F8, 00000008h
loc_005AA3E2:   lea ecx, var_B8
loc_005AA3E8:   push ecx
loc_005AA3E9:   lea edx, var_A8
loc_005AA3EF:   push edx
loc_005AA3F0:   lea eax, var_C8
loc_005AA3F6:   push eax
  loc_005AA3F7: call [004011C4h] ; __vbaVarCat
loc_005AA3FD:   push eax
loc_005AA3FE:   lea ecx, var_268
loc_005AA404:   push ecx
loc_005AA405:   lea edx, var_D8
loc_005AA40B:   push edx
  loc_005AA40C: call [004011C4h] ; __vbaVarCat
loc_005AA412:   push eax
loc_005AA413:   lea eax, var_F8
loc_005AA419:   push eax
loc_005AA41A:   lea ecx, var_108
loc_005AA420:   push ecx
  loc_005AA421: call [004011C4h] ; __vbaVarCat
loc_005AA427:   push eax
loc_005AA428:   lea edx, var_278
loc_005AA42E:   push edx
loc_005AA42F:   lea eax, var_118
loc_005AA435:   push eax
  loc_005AA436: call [004011C4h] ; __vbaVarCat
loc_005AA43C:   push eax
loc_005AA43D:   lea ecx, var_128
loc_005AA443:   push ecx
loc_005AA444:   lea edx, var_138
loc_005AA44A:   push edx
  loc_005AA44B: call [004011C4h] ; __vbaVarCat
loc_005AA451:   push eax
loc_005AA452:   lea eax, var_288
loc_005AA458:   push eax
loc_005AA459:   lea ecx, var_148
loc_005AA45F:   push ecx
  loc_005AA460: call [004011C4h] ; __vbaVarCat
loc_005AA466:   push eax
loc_005AA467:   lea edx, var_168
loc_005AA46D:   push edx
loc_005AA46E:   lea eax, var_178
loc_005AA474:   push eax
  loc_005AA475: call [004011C4h] ; __vbaVarCat
loc_005AA47B:   push eax
loc_005AA47C:   lea ecx, var_298
loc_005AA482:   push ecx
loc_005AA483:   lea edx, var_188
loc_005AA489:   push edx
  loc_005AA48A: call [004011C4h] ; __vbaVarCat
loc_005AA490:   push eax
  loc_005AA491: push 006680B4h
loc_005AA496:   lea eax, var_198
loc_005AA49C:   push eax
  loc_005AA49D: call [004011C4h] ; __vbaVarCat
loc_005AA4A3:   push eax
loc_005AA4A4:   lea ecx, var_2A8
loc_005AA4AA:   push ecx
loc_005AA4AB:   lea edx, var_1A8
loc_005AA4B1:   push edx
  loc_005AA4B2: call [004011C4h] ; __vbaVarCat
loc_005AA4B8:   push eax
loc_005AA4B9:   lea eax, var_1B8
loc_005AA4BF:   push eax
loc_005AA4C0:   lea ecx, var_1C8
loc_005AA4C6:   push ecx
  loc_005AA4C7: call [004011C4h] ; __vbaVarCat
loc_005AA4CD:   push eax
loc_005AA4CE:   lea edx, var_2B8
loc_005AA4D4:   push edx
loc_005AA4D5:   lea eax, var_1D8
loc_005AA4DB:   push eax
  loc_005AA4DC: call [004011C4h] ; __vbaVarCat
loc_005AA4E2:   push eax
loc_005AA4E3:   lea ecx, var_208
loc_005AA4E9:   push ecx
loc_005AA4EA:   lea edx, var_218
loc_005AA4F0:   push edx
  loc_005AA4F1: call [004011C4h] ; __vbaVarCat
loc_005AA4F7:   push eax
loc_005AA4F8:   lea eax, var_2D8
loc_005AA4FE:   push eax
loc_005AA4FF:   lea ecx, var_228
loc_005AA505:   push ecx
  loc_005AA506: call [004011C4h] ; __vbaVarCat
loc_005AA50C:   push eax
loc_005AA50D:   lea edx, var_2E8
loc_005AA513:   push edx
loc_005AA514:   lea eax, var_238
loc_005AA51A:   push eax
  loc_005AA51B: call [004011C4h] ; __vbaVarCat
loc_005AA521:   push eax
loc_005AA522:   lea ecx, var_2F8
loc_005AA528:   push ecx
loc_005AA529:   lea edx, var_248
loc_005AA52F:   push edx
  loc_005AA530: call [004011C4h] ; __vbaVarCat
loc_005AA536:   push eax
  loc_005AA537: call [00401034h] ; __vbaStrVarMove
loc_005AA53D:   mov edx, eax
  loc_005AA53F: mov ecx, 0066809Ch
  loc_005AA544: call [0040126Ch] ; __vbaStrMove
loc_005AA54A:   lea eax, var_40
loc_005AA54D:   push eax
loc_005AA54E:   lea ecx, var_3C
loc_005AA551:   push ecx
loc_005AA552:   lea edx, var_38
loc_005AA555:   push edx
  loc_005AA556: push 00000003h
  loc_005AA558: call [00401204h] ; __vbaFreeStrList
  loc_005AA55E: add esp, 00000010h
loc_005AA561:   lea eax, var_74
loc_005AA564:   push eax
loc_005AA565:   lea ecx, var_70
loc_005AA568:   push ecx
loc_005AA569:   lea edx, var_6C
loc_005AA56C:   push edx
loc_005AA56D:   lea eax, var_68
loc_005AA570:   push eax
loc_005AA571:   lea ecx, var_64
loc_005AA574:   push ecx
  loc_005AA575: push 00000005h
  loc_005AA577: call [0040104Ch] ; __vbaFreeObjList
  loc_005AA57D: add esp, 00000018h
loc_005AA580:   lea edx, var_248
loc_005AA586:   push edx
loc_005AA587:   lea eax, var_238
loc_005AA58D:   push eax
loc_005AA58E:   lea ecx, var_228
loc_005AA594:   push ecx
loc_005AA595:   lea edx, var_218
loc_005AA59B:   push edx
loc_005AA59C:   lea eax, var_208
loc_005AA5A2:   push eax
loc_005AA5A3:   lea ecx, var_1D8
loc_005AA5A9:   push ecx
loc_005AA5AA:   lea edx, var_1F8
loc_005AA5B0:   push edx
loc_005AA5B1:   lea eax, var_1E8
loc_005AA5B7:   push eax
loc_005AA5B8:   lea ecx, var_1C8
loc_005AA5BE:   push ecx
loc_005AA5BF:   lea edx, var_1B8
loc_005AA5C5:   push edx
loc_005AA5C6:   lea eax, var_1A8
loc_005AA5CC:   push eax
loc_005AA5CD:   lea ecx, var_198
loc_005AA5D3:   push ecx
loc_005AA5D4:   lea edx, var_188
loc_005AA5DA:   push edx
loc_005AA5DB:   lea eax, var_178
loc_005AA5E1:   push eax
loc_005AA5E2:   lea ecx, var_168
loc_005AA5E8:   push ecx
loc_005AA5E9:   lea edx, var_148
loc_005AA5EF:   push edx
loc_005AA5F0:   lea eax, var_158
loc_005AA5F6:   push eax
loc_005AA5F7:   lea ecx, var_138
loc_005AA5FD:   push ecx
loc_005AA5FE:   lea edx, var_128
loc_005AA604:   push edx
loc_005AA605:   lea eax, var_118
loc_005AA60B:   push eax
loc_005AA60C:   lea ecx, var_108
loc_005AA612:   push ecx
loc_005AA613:   lea edx, var_F8
loc_005AA619:   push edx
loc_005AA61A:   lea eax, var_D8
loc_005AA620:   push eax
loc_005AA621:   lea ecx, var_E8
loc_005AA627:   push ecx
loc_005AA628:   lea edx, var_C8
loc_005AA62E:   push edx
loc_005AA62F:   lea eax, var_A8
loc_005AA635:   push eax
loc_005AA636:   lea ecx, var_B8
loc_005AA63C:   push ecx
loc_005AA63D:   lea edx, var_98
loc_005AA643:   push edx
loc_005AA644:   lea eax, var_88
loc_005AA64A:   push eax
  loc_005AA64B: push 0000001Dh
  loc_005AA64D: call [0040103Ch] ; __vbaFreeVarList
  loc_005AA653: add esp, 00000078h
  loc_005AA656: mov var_4, 0000000Fh
  loc_005AA65D: cmp [0066803Ch], 00000000h
  loc_005AA664: jnz 005AA682h
  loc_005AA666: push 0066803Ch
  loc_005AA66B: push 0042DEFCh
  loc_005AA670: call [004011E8h] ; __vbaNew2
  loc_005AA676: mov var_4D4, 0066803Ch
  loc_005AA680: jmp 005AA68Ch
  loc_005AA682: mov var_4D4, 0066803Ch
loc_005AA68C:   mov ecx, var_4D4
loc_005AA692:   mov edx, [ecx]
loc_005AA694:   mov var_434, edx
  loc_005AA69A: mov var_80, 80020004h
  loc_005AA6A1: mov var_88, 0000000Ah
loc_005AA6AB:   lea ecx, var_88
  loc_005AA6B1: call [0040123Ch] ; __vbaFreeVarg
loc_005AA6B7:   lea eax, var_64
loc_005AA6BA:   push eax
  loc_005AA6BB: push 00000001h
loc_005AA6BD:   lea ecx, var_88
loc_005AA6C3:   push ecx
loc_005AA6C4:   mov edx, [0066809Ch]
loc_005AA6CA:   push edx
loc_005AA6CB:   mov eax, var_434
loc_005AA6D1:   mov ecx, [eax]
loc_005AA6D3:   mov edx, var_434
loc_005AA6D9:   push edx
loc_005AA6DA:   Call [ecx+00000040h]
loc_005AA6DD:   fnclex
loc_005AA6DF:   mov var_438, eax
  loc_005AA6E5: cmp var_438, 00000000h
  loc_005AA6EC: jge 005AA711h
  loc_005AA6EE: push 00000040h
  loc_005AA6F0: push 0042E1B0h
loc_005AA6F5:   mov eax, var_434
loc_005AA6FB:   push eax
loc_005AA6FC:   mov ecx, var_438
loc_005AA702:   push ecx
  loc_005AA703: call [00401080h] ; __vbaHresultCheckObj
loc_005AA709:   mov var_4D8, eax
  loc_005AA70F: jmp 005AA71Bh
  loc_005AA711: mov var_4D8, 00000000h
loc_005AA71B:   lea ecx, var_64
  loc_005AA71E: call [004012A0h] ; __vbaFreeObj
loc_005AA724:   lea ecx, var_88
  loc_005AA72A: call [00401020h] ; __vbaFreeVar
  loc_005AA730: mov var_4, 00000010h
loc_005AA737:   lea edx, var_88
loc_005AA73D:   push edx
  loc_005AA73E: call [00401220h] ; rtcGetDateVar
loc_005AA744:   mov eax, Me
loc_005AA747:   mov ecx, [eax]
loc_005AA749:   mov edx, Me
loc_005AA74C:   push edx
loc_005AA74D:   Call [ecx+00000338h]
loc_005AA753:   push eax
loc_005AA754:   lea eax, var_64
loc_005AA757:   push eax
  loc_005AA758: call [004010B8h] ; __vbaObjSet
loc_005AA75E:   mov var_434, eax
loc_005AA764:   lea ecx, var_44
loc_005AA767:   push ecx
loc_005AA768:   mov edx, var_434
loc_005AA76E:   mov eax, [edx]
loc_005AA770:   mov ecx, var_434
loc_005AA776:   push ecx
loc_005AA777:   Call [eax+00000050h]
loc_005AA77A:   fnclex
loc_005AA77C:   mov var_438, eax
  loc_005AA782: cmp var_438, 00000000h
  loc_005AA789: jge 005AA7AEh
  loc_005AA78B: push 00000050h
  loc_005AA78D: push 00433F80h
loc_005AA792:   mov edx, var_434
loc_005AA798:   push edx
loc_005AA799:   mov eax, var_438
loc_005AA79F:   push eax
  loc_005AA7A0: call [00401080h] ; __vbaHresultCheckObj
loc_005AA7A6:   mov var_4DC, eax
  loc_005AA7AC: jmp 005AA7B8h
  loc_005AA7AE: mov var_4DC, 00000000h
loc_005AA7B8:   mov ecx, Me
loc_005AA7BB:   mov edx, [ecx]
loc_005AA7BD:   mov eax, Me
loc_005AA7C0:   push eax
loc_005AA7C1:   Call [edx+00000300h]
loc_005AA7C7:   push eax
loc_005AA7C8:   lea ecx, var_68
loc_005AA7CB:   push ecx
  loc_005AA7CC: call [004010B8h] ; __vbaObjSet
loc_005AA7D2:   mov var_43C, eax
loc_005AA7D8:   lea edx, var_48
loc_005AA7DB:   push edx
loc_005AA7DC:   mov eax, var_43C
loc_005AA7E2:   mov ecx, [eax]
loc_005AA7E4:   mov edx, var_43C
loc_005AA7EA:   push edx
loc_005AA7EB:   Call [ecx+00000050h]
loc_005AA7EE:   fnclex
loc_005AA7F0:   mov var_440, eax
  loc_005AA7F6: cmp var_440, 00000000h
  loc_005AA7FD: jge 005AA822h
  loc_005AA7FF: push 00000050h
  loc_005AA801: push 00433F80h
loc_005AA806:   mov eax, var_43C
loc_005AA80C:   push eax
loc_005AA80D:   mov ecx, var_440
loc_005AA813:   push ecx
  loc_005AA814: call [00401080h] ; __vbaHresultCheckObj
loc_005AA81A:   mov var_4E0, eax
  loc_005AA820: jmp 005AA82Ch
  loc_005AA822: mov var_4E0, 00000000h
  loc_005AA82C: push 00435DD0h ; "INSERT INTO Laba"
  loc_005AA831: push 00435E04h ; "(no_nota,tgl_nota,tot_jual,tot_beli,laba)"
  loc_005AA836: call [00401070h] ; __vbaStrCat
loc_005AA83C:   mov edx, eax
loc_005AA83E:   lea ecx, var_38
  loc_005AA841: call [0040126Ch] ; __vbaStrMove
loc_005AA847:   push eax
  loc_005AA848: push 00434520h ; "VALUES ('"
  loc_005AA84D: call [00401070h] ; __vbaStrCat
loc_005AA853:   mov edx, eax
loc_005AA855:   lea ecx, var_3C
  loc_005AA858: call [0040126Ch] ; __vbaStrMove
loc_005AA85E:   push eax
loc_005AA85F:   mov edx, Me
loc_005AA862:   mov eax, [edx+0000006Ch]
loc_005AA865:   push eax
  loc_005AA866: call [00401070h] ; __vbaStrCat
loc_005AA86C:   mov edx, eax
loc_005AA86E:   lea ecx, var_40
  loc_005AA871: call [0040126Ch] ; __vbaStrMove
loc_005AA877:   push eax
  loc_005AA878: push 0042EC30h ; "','"
  loc_005AA87D: call [00401070h] ; __vbaStrCat
loc_005AA883:   mov var_B0, eax
  loc_005AA889: mov var_B8, 00000008h
  loc_005AA893: mov var_250, 00434538h ; "yyyy/MM/dd"
  loc_005AA89D: mov var_258, 00000008h
loc_005AA8A7:   lea edx, var_258
loc_005AA8AD:   lea ecx, var_98
  loc_005AA8B3: call [00401238h] ; __vbaVarDup
  loc_005AA8B9: push 00000001h
  loc_005AA8BB: push 00000001h
loc_005AA8BD:   lea ecx, var_98
loc_005AA8C3:   push ecx
loc_005AA8C4:   lea edx, var_88
loc_005AA8CA:   push edx
loc_005AA8CB:   lea eax, var_A8
loc_005AA8D1:   push eax
  loc_005AA8D2: call [00401078h] ; rtcVarFromFormatVar
  loc_005AA8D8: mov var_260, 0042EC30h ; "','"
  loc_005AA8E2: mov var_268, 00000008h
  loc_005AA8EC: mov var_270, 00435A20h ; "00"
  loc_005AA8F6: mov var_278, 00000008h
loc_005AA900:   lea edx, var_278
loc_005AA906:   lea ecx, var_F8
  loc_005AA90C: call [00401238h] ; __vbaVarDup
loc_005AA912:   mov ecx, var_44
loc_005AA915:   mov var_48C, ecx
  loc_005AA91B: mov var_44, 00000000h
loc_005AA922:   mov edx, var_48C
loc_005AA928:   mov var_E0, edx
  loc_005AA92E: mov var_E8, 00000008h
  loc_005AA938: push 00000001h
  loc_005AA93A: push 00000001h
loc_005AA93C:   lea eax, var_F8
loc_005AA942:   push eax
loc_005AA943:   lea ecx, var_E8
loc_005AA949:   push ecx
loc_005AA94A:   lea edx, var_108
loc_005AA950:   push edx
  loc_005AA951: call [00401078h] ; rtcVarFromFormatVar
  loc_005AA957: mov var_280, 0042EC30h ; "','"
  loc_005AA961: mov var_288, 00000008h
  loc_005AA96B: mov var_290, 00435A20h ; "00"
  loc_005AA975: mov var_298, 00000008h
loc_005AA97F:   lea edx, var_298
loc_005AA985:   lea ecx, var_148
  loc_005AA98B: call [00401238h] ; __vbaVarDup
loc_005AA991:   mov eax, var_48
loc_005AA994:   mov var_490, eax
  loc_005AA99A: mov var_48, 00000000h
loc_005AA9A1:   mov ecx, var_490
loc_005AA9A7:   mov var_130, ecx
  loc_005AA9AD: mov var_138, 00000008h
  loc_005AA9B7: push 00000001h
  loc_005AA9B9: push 00000001h
loc_005AA9BB:   lea edx, var_148
loc_005AA9C1:   push edx
loc_005AA9C2:   lea eax, var_138
loc_005AA9C8:   push eax
loc_005AA9C9:   lea ecx, var_158
loc_005AA9CF:   push ecx
  loc_005AA9D0: call [00401078h] ; rtcVarFromFormatVar
  loc_005AA9D6: mov var_2A0, 0042EC30h ; "','"
  loc_005AA9E0: mov var_2A8, 00000008h
loc_005AA9EA:   mov edx, var_34
loc_005AA9ED:   mov var_2B0, edx
loc_005AA9F3:   mov eax, var_30
loc_005AA9F6:   mov var_2AC, eax
  loc_005AA9FC: mov var_2B8, 00000006h
  loc_005AAA06: mov var_2C0, 0042EC3Ch ; "')"
  loc_005AAA10: mov var_2C8, 00000008h
loc_005AAA1A:   lea ecx, var_B8
loc_005AAA20:   push ecx
loc_005AAA21:   lea edx, var_A8
loc_005AAA27:   push edx
loc_005AAA28:   lea eax, var_C8
loc_005AAA2E:   push eax
  loc_005AAA2F: call [004011C4h] ; __vbaVarCat
loc_005AAA35:   push eax
loc_005AAA36:   lea ecx, var_268
loc_005AAA3C:   push ecx
loc_005AAA3D:   lea edx, var_D8
loc_005AAA43:   push edx
  loc_005AAA44: call [004011C4h] ; __vbaVarCat
loc_005AAA4A:   push eax
loc_005AAA4B:   lea eax, var_108
loc_005AAA51:   push eax
loc_005AAA52:   lea ecx, var_118
loc_005AAA58:   push ecx
  loc_005AAA59: call [004011C4h] ; __vbaVarCat
loc_005AAA5F:   push eax
loc_005AAA60:   lea edx, var_288
loc_005AAA66:   push edx
loc_005AAA67:   lea eax, var_128
loc_005AAA6D:   push eax
  loc_005AAA6E: call [004011C4h] ; __vbaVarCat
loc_005AAA74:   push eax
loc_005AAA75:   lea ecx, var_158
loc_005AAA7B:   push ecx
loc_005AAA7C:   lea edx, var_168
loc_005AAA82:   push edx
  loc_005AAA83: call [004011C4h] ; __vbaVarCat
loc_005AAA89:   push eax
loc_005AAA8A:   lea eax, var_2A8
loc_005AAA90:   push eax
loc_005AAA91:   lea ecx, var_178
loc_005AAA97:   push ecx
  loc_005AAA98: call [004011C4h] ; __vbaVarCat
loc_005AAA9E:   push eax
loc_005AAA9F:   lea edx, var_2B8
loc_005AAAA5:   push edx
loc_005AAAA6:   lea eax, var_188
loc_005AAAAC:   push eax
  loc_005AAAAD: call [004011C4h] ; __vbaVarCat
loc_005AAAB3:   push eax
loc_005AAAB4:   lea ecx, var_2C8
loc_005AAABA:   push ecx
loc_005AAABB:   lea edx, var_198
loc_005AAAC1:   push edx
  loc_005AAAC2: call [004011C4h] ; __vbaVarCat
loc_005AAAC8:   push eax
  loc_005AAAC9: call [00401034h] ; __vbaStrVarMove
loc_005AAACF:   mov edx, eax
  loc_005AAAD1: mov ecx, 0066809Ch
  loc_005AAAD6: call [0040126Ch] ; __vbaStrMove
loc_005AAADC:   lea eax, var_40
loc_005AAADF:   push eax
loc_005AAAE0:   lea ecx, var_3C
loc_005AAAE3:   push ecx
loc_005AAAE4:   lea edx, var_38
loc_005AAAE7:   push edx
  loc_005AAAE8: push 00000003h
  loc_005AAAEA: call [00401204h] ; __vbaFreeStrList
  loc_005AAAF0: add esp, 00000010h
loc_005AAAF3:   lea eax, var_68
loc_005AAAF6:   push eax
loc_005AAAF7:   lea ecx, var_64
loc_005AAAFA:   push ecx
  loc_005AAAFB: push 00000002h
  loc_005AAAFD: call [0040104Ch] ; __vbaFreeObjList
  loc_005AAB03: add esp, 0000000Ch
loc_005AAB06:   lea edx, var_198
loc_005AAB0C:   push edx
loc_005AAB0D:   lea eax, var_188
loc_005AAB13:   push eax
loc_005AAB14:   lea ecx, var_178
loc_005AAB1A:   push ecx
loc_005AAB1B:   lea edx, var_168
loc_005AAB21:   push edx
loc_005AAB22:   lea eax, var_158
loc_005AAB28:   push eax
loc_005AAB29:   lea ecx, var_128
loc_005AAB2F:   push ecx
loc_005AAB30:   lea edx, var_148
loc_005AAB36:   push edx
loc_005AAB37:   lea eax, var_138
loc_005AAB3D:   push eax
loc_005AAB3E:   lea ecx, var_118
loc_005AAB44:   push ecx
loc_005AAB45:   lea edx, var_108
loc_005AAB4B:   push edx
loc_005AAB4C:   lea eax, var_D8
loc_005AAB52:   push eax
loc_005AAB53:   lea ecx, var_F8
loc_005AAB59:   push ecx
loc_005AAB5A:   lea edx, var_E8
loc_005AAB60:   push edx
loc_005AAB61:   lea eax, var_C8
loc_005AAB67:   push eax
loc_005AAB68:   lea ecx, var_A8
loc_005AAB6E:   push ecx
loc_005AAB6F:   lea edx, var_B8
loc_005AAB75:   push edx
loc_005AAB76:   lea eax, var_98
loc_005AAB7C:   push eax
loc_005AAB7D:   lea ecx, var_88
loc_005AAB83:   push ecx
  loc_005AAB84: push 00000012h
  loc_005AAB86: call [0040103Ch] ; __vbaFreeVarList
  loc_005AAB8C: add esp, 0000004Ch
  loc_005AAB8F: mov var_4, 00000011h
  loc_005AAB96: cmp [0066803Ch], 00000000h
  loc_005AAB9D: jnz 005AABBBh
  loc_005AAB9F: push 0066803Ch
  loc_005AABA4: push 0042DEFCh
  loc_005AABA9: call [004011E8h] ; __vbaNew2
  loc_005AABAF: mov var_4E4, 0066803Ch
  loc_005AABB9: jmp 005AABC5h
  loc_005AABBB: mov var_4E4, 0066803Ch
loc_005AABC5:   mov edx, var_4E4
loc_005AABCB:   mov eax, [edx]
loc_005AABCD:   mov var_434, eax
  loc_005AABD3: mov var_80, 80020004h
  loc_005AABDA: mov var_88, 0000000Ah
loc_005AABE4:   lea ecx, var_88
  loc_005AABEA: call [0040123Ch] ; __vbaFreeVarg
loc_005AABF0:   lea ecx, var_64
loc_005AABF3:   push ecx
  loc_005AABF4: push 00000001h
loc_005AABF6:   lea edx, var_88
loc_005AABFC:   push edx
loc_005AABFD:   mov eax, [0066809Ch]
loc_005AAC02:   push eax
loc_005AAC03:   mov ecx, var_434
loc_005AAC09:   mov edx, [ecx]
loc_005AAC0B:   mov eax, var_434
loc_005AAC11:   push eax
loc_005AAC12:   Call [edx+00000040h]
loc_005AAC15:   fnclex
loc_005AAC17:   mov var_438, eax
  loc_005AAC1D: cmp var_438, 00000000h
  loc_005AAC24: jge 005AAC49h
  loc_005AAC26: push 00000040h
  loc_005AAC28: push 0042E1B0h
loc_005AAC2D:   mov ecx, var_434
loc_005AAC33:   push ecx
loc_005AAC34:   mov edx, var_438
loc_005AAC3A:   push edx
  loc_005AAC3B: call [00401080h] ; __vbaHresultCheckObj
loc_005AAC41:   mov var_4E8, eax
  loc_005AAC47: jmp 005AAC53h
  loc_005AAC49: mov var_4E8, 00000000h
loc_005AAC53:   lea ecx, var_64
  loc_005AAC56: call [004012A0h] ; __vbaFreeObj
loc_005AAC5C:   lea ecx, var_88
  loc_005AAC62: call [00401020h] ; __vbaFreeVar
  loc_005AAC68: mov var_4, 00000012h
loc_005AAC6F:   mov eax, Me
loc_005AAC72:   mov cx, [eax+00000034h]
  loc_005AAC76: sub cx, 0001h
  loc_005AAC7A: jo 005AEBD1h
loc_005AAC80:   mov var_458, cx
  loc_005AAC87: mov var_454, 0001h
  loc_005AAC90: mov var_24, 0001h
  loc_005AAC96: jmp 005AACADh
loc_005AAC98:   mov dx, var_24
loc_005AAC9C:   Add dx, var_454
  loc_005AACA3: jo 005AEBD1h
loc_005AACA9:   mov var_24, dx
loc_005AACAD:   mov ax, var_24
loc_005AACB1:   cmp ax, var_458
  loc_005AACB8: jg 005AB933h
  loc_005AACBE: mov var_4, 00000013h
loc_005AACC5:   movsx ecx, var_24
loc_005AACC9:   mov var_2D0, ecx
  loc_005AACCF: mov var_2D8, 00000003h
  loc_005AACD9: mov var_2F0, 00000002h
  loc_005AACE3: mov var_2F8, 00000003h
  loc_005AACED: mov eax, 00000010h
  loc_005AACF2: call 00405690h ; __vbaChkstk
loc_005AACF7:   mov edx, esp
loc_005AACF9:   mov eax, var_2D8
loc_005AACFF:   mov [edx], eax
loc_005AAD01:   mov ecx, var_2D4
loc_005AAD07:   mov [edx+00000004h], ecx
loc_005AAD0A:   mov eax, var_2D0
loc_005AAD10:   mov [edx+00000008h], eax
loc_005AAD13:   mov ecx, var_2CC
loc_005AAD19:   mov [edx+0000000Ch], ecx
  loc_005AAD1C: mov eax, 00000010h
  loc_005AAD21: call 00405690h ; __vbaChkstk
loc_005AAD26:   mov edx, esp
loc_005AAD28:   mov eax, var_2F8
loc_005AAD2E:   mov [edx], eax
loc_005AAD30:   mov ecx, var_2F4
loc_005AAD36:   mov [edx+00000004h], ecx
loc_005AAD39:   mov eax, var_2F0
loc_005AAD3F:   mov [edx+00000008h], eax
loc_005AAD42:   mov ecx, var_2EC
loc_005AAD48:   mov [edx+0000000Ch], ecx
  loc_005AAD4B: push 00000002h
  loc_005AAD4D: push 00000041h
loc_005AAD4F:   mov edx, Me
loc_005AAD52:   mov eax, [edx]
loc_005AAD54:   mov ecx, Me
loc_005AAD57:   push ecx
loc_005AAD58:   Call [eax+00000390h]
loc_005AAD5E:   push eax
loc_005AAD5F:   lea edx, var_6C
loc_005AAD62:   push edx
  loc_005AAD63: call [004010B8h] ; __vbaObjSet
loc_005AAD69:   push eax
loc_005AAD6A:   lea eax, var_A8
loc_005AAD70:   push eax
  loc_005AAD71: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AAD77: add esp, 00000030h
loc_005AAD7A:   movsx ecx, var_24
loc_005AAD7E:   mov var_3D0, ecx
  loc_005AAD84: mov var_3D8, 00000003h
  loc_005AAD8E: mov var_3F0, 00000005h
  loc_005AAD98: mov var_3F8, 00000003h
  loc_005AADA2: mov eax, 00000010h
  loc_005AADA7: call 00405690h ; __vbaChkstk
loc_005AADAC:   mov edx, esp
loc_005AADAE:   mov eax, var_3D8
loc_005AADB4:   mov [edx], eax
loc_005AADB6:   mov ecx, var_3D4
loc_005AADBC:   mov [edx+00000004h], ecx
loc_005AADBF:   mov eax, var_3D0
loc_005AADC5:   mov [edx+00000008h], eax
loc_005AADC8:   mov ecx, var_3CC
loc_005AADCE:   mov [edx+0000000Ch], ecx
  loc_005AADD1: mov eax, 00000010h
  loc_005AADD6: call 00405690h ; __vbaChkstk
loc_005AADDB:   mov edx, esp
loc_005AADDD:   mov eax, var_3F8
loc_005AADE3:   mov [edx], eax
loc_005AADE5:   mov ecx, var_3F4
loc_005AADEB:   mov [edx+00000004h], ecx
loc_005AADEE:   mov eax, var_3F0
loc_005AADF4:   mov [edx+00000008h], eax
loc_005AADF7:   mov ecx, var_3EC
loc_005AADFD:   mov [edx+0000000Ch], ecx
  loc_005AAE00: push 00000002h
  loc_005AAE02: push 00000041h
loc_005AAE04:   mov edx, Me
loc_005AAE07:   mov eax, [edx]
loc_005AAE09:   mov ecx, Me
loc_005AAE0C:   push ecx
loc_005AAE0D:   Call [eax+00000390h]
loc_005AAE13:   push eax
loc_005AAE14:   lea edx, var_78
loc_005AAE17:   push edx
  loc_005AAE18: call [004010B8h] ; __vbaObjSet
loc_005AAE1E:   push eax
loc_005AAE1F:   lea eax, var_198
loc_005AAE25:   push eax
  loc_005AAE26: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AAE2C: add esp, 00000030h
loc_005AAE2F:   movsx ecx, var_24
loc_005AAE33:   mov var_250, ecx
  loc_005AAE39: mov var_258, 00000003h
  loc_005AAE43: mov var_270, 00000000h
  loc_005AAE4D: mov var_278, 00000003h
loc_005AAE57:   movsx edx, var_24
loc_005AAE5B:   mov var_290, edx
  loc_005AAE61: mov var_298, 00000003h
  loc_005AAE6B: mov var_2B0, 00000001h
  loc_005AAE75: mov var_2B8, 00000003h
  loc_005AAE7F: push 00435E5Ch ; "INSERT INTO penjualan_detail"
  loc_005AAE84: push 00435E9Ch ; "(no_nota,kode_barang,nama_barang,harga_jual,diskon,jumlah,sub_total)"
  loc_005AAE89: call [00401070h] ; __vbaStrCat
loc_005AAE8F:   mov edx, eax
loc_005AAE91:   lea ecx, var_38
  loc_005AAE94: call [0040126Ch] ; __vbaStrMove
loc_005AAE9A:   push eax
  loc_005AAE9B: push 00434620h ; " VALUES ('"
  loc_005AAEA0: call [00401070h] ; __vbaStrCat
loc_005AAEA6:   mov edx, eax
loc_005AAEA8:   lea ecx, var_3C
  loc_005AAEAB: call [0040126Ch] ; __vbaStrMove
loc_005AAEB1:   push eax
loc_005AAEB2:   mov eax, Me
loc_005AAEB5:   mov ecx, [eax+0000006Ch]
loc_005AAEB8:   push ecx
  loc_005AAEB9: call [00401070h] ; __vbaStrCat
loc_005AAEBF:   mov edx, eax
loc_005AAEC1:   lea ecx, var_40
  loc_005AAEC4: call [0040126Ch] ; __vbaStrMove
loc_005AAECA:   push eax
  loc_005AAECB: push 0042EC30h ; "','"
  loc_005AAED0: call [00401070h] ; __vbaStrCat
loc_005AAED6:   mov edx, eax
loc_005AAED8:   lea ecx, var_44
  loc_005AAEDB: call [0040126Ch] ; __vbaStrMove
loc_005AAEE1:   push eax
  loc_005AAEE2: mov eax, 00000010h
  loc_005AAEE7: call 00405690h ; __vbaChkstk
loc_005AAEEC:   mov edx, esp
loc_005AAEEE:   mov eax, var_258
loc_005AAEF4:   mov [edx], eax
loc_005AAEF6:   mov ecx, var_254
loc_005AAEFC:   mov [edx+00000004h], ecx
loc_005AAEFF:   mov eax, var_250
loc_005AAF05:   mov [edx+00000008h], eax
loc_005AAF08:   mov ecx, var_24C
loc_005AAF0E:   mov [edx+0000000Ch], ecx
  loc_005AAF11: mov eax, 00000010h
  loc_005AAF16: call 00405690h ; __vbaChkstk
loc_005AAF1B:   mov edx, esp
loc_005AAF1D:   mov eax, var_278
loc_005AAF23:   mov [edx], eax
loc_005AAF25:   mov ecx, var_274
loc_005AAF2B:   mov [edx+00000004h], ecx
loc_005AAF2E:   mov eax, var_270
loc_005AAF34:   mov [edx+00000008h], eax
loc_005AAF37:   mov ecx, var_26C
loc_005AAF3D:   mov [edx+0000000Ch], ecx
  loc_005AAF40: push 00000002h
  loc_005AAF42: push 00000041h
loc_005AAF44:   mov edx, Me
loc_005AAF47:   mov eax, [edx]
loc_005AAF49:   mov ecx, Me
loc_005AAF4C:   push ecx
loc_005AAF4D:   Call [eax+00000390h]
loc_005AAF53:   push eax
loc_005AAF54:   lea edx, var_64
loc_005AAF57:   push edx
  loc_005AAF58: call [004010B8h] ; __vbaObjSet
loc_005AAF5E:   push eax
loc_005AAF5F:   lea eax, var_88
loc_005AAF65:   push eax
  loc_005AAF66: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AAF6C: add esp, 00000030h
loc_005AAF6F:   push eax
  loc_005AAF70: call [00401034h] ; __vbaStrVarMove
loc_005AAF76:   mov edx, eax
loc_005AAF78:   lea ecx, var_48
  loc_005AAF7B: call [0040126Ch] ; __vbaStrMove
loc_005AAF81:   push eax
  loc_005AAF82: call [00401070h] ; __vbaStrCat
loc_005AAF88:   mov edx, eax
loc_005AAF8A:   lea ecx, var_4C
  loc_005AAF8D: call [0040126Ch] ; __vbaStrMove
loc_005AAF93:   push eax
  loc_005AAF94: push 0042EC30h ; "','"
  loc_005AAF99: call [00401070h] ; __vbaStrCat
loc_005AAF9F:   mov edx, eax
loc_005AAFA1:   lea ecx, var_50
  loc_005AAFA4: call [0040126Ch] ; __vbaStrMove
loc_005AAFAA:   push eax
  loc_005AAFAB: mov eax, 00000010h
  loc_005AAFB0: call 00405690h ; __vbaChkstk
loc_005AAFB5:   mov ecx, esp
loc_005AAFB7:   mov edx, var_298
loc_005AAFBD:   mov [ecx], edx
loc_005AAFBF:   mov eax, var_294
loc_005AAFC5:   mov [ecx+00000004h], eax
loc_005AAFC8:   mov edx, var_290
loc_005AAFCE:   mov [ecx+00000008h], edx
loc_005AAFD1:   mov eax, var_28C
loc_005AAFD7:   mov [ecx+0000000Ch], eax
  loc_005AAFDA: mov eax, 00000010h
  loc_005AAFDF: call 00405690h ; __vbaChkstk
loc_005AAFE4:   mov ecx, esp
loc_005AAFE6:   mov edx, var_2B8
loc_005AAFEC:   mov [ecx], edx
loc_005AAFEE:   mov eax, var_2B4
loc_005AAFF4:   mov [ecx+00000004h], eax
loc_005AAFF7:   mov edx, var_2B0
loc_005AAFFD:   mov [ecx+00000008h], edx
loc_005AB000:   mov eax, var_2AC
loc_005AB006:   mov [ecx+0000000Ch], eax
  loc_005AB009: push 00000002h
  loc_005AB00B: push 00000041h
loc_005AB00D:   mov ecx, Me
loc_005AB010:   mov edx, [ecx]
loc_005AB012:   mov eax, Me
loc_005AB015:   push eax
loc_005AB016:   Call [edx+00000390h]
loc_005AB01C:   push eax
loc_005AB01D:   lea ecx, var_68
loc_005AB020:   push ecx
  loc_005AB021: call [004010B8h] ; __vbaObjSet
loc_005AB027:   push eax
loc_005AB028:   lea edx, var_98
loc_005AB02E:   push edx
  loc_005AB02F: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AB035: add esp, 00000030h
loc_005AB038:   push eax
  loc_005AB039: call [00401034h] ; __vbaStrVarMove
loc_005AB03F:   mov edx, eax
loc_005AB041:   lea ecx, var_54
  loc_005AB044: call [0040126Ch] ; __vbaStrMove
loc_005AB04A:   push eax
  loc_005AB04B: call [00401070h] ; __vbaStrCat
loc_005AB051:   mov edx, eax
loc_005AB053:   lea ecx, var_58
  loc_005AB056: call [0040126Ch] ; __vbaStrMove
loc_005AB05C:   push eax
  loc_005AB05D: push 0042EC30h ; "','"
  loc_005AB062: call [00401070h] ; __vbaStrCat
loc_005AB068:   mov var_E0, eax
  loc_005AB06E: mov var_E8, 00000008h
  loc_005AB078: mov var_310, 00435A20h ; "00"
  loc_005AB082: mov var_318, 00000008h
loc_005AB08C:   lea edx, var_318
loc_005AB092:   lea ecx, var_C8
  loc_005AB098: call [00401238h] ; __vbaVarDup
loc_005AB09E:   lea eax, var_A8
loc_005AB0A4:   push eax
  loc_005AB0A5: call [00401034h] ; __vbaStrVarMove
loc_005AB0AB:   mov var_B0, eax
  loc_005AB0B1: mov var_B8, 00000008h
  loc_005AB0BB: push 00000001h
  loc_005AB0BD: push 00000001h
loc_005AB0BF:   lea ecx, var_C8
loc_005AB0C5:   push ecx
loc_005AB0C6:   lea edx, var_B8
loc_005AB0CC:   push edx
loc_005AB0CD:   lea eax, var_D8
loc_005AB0D3:   push eax
  loc_005AB0D4: call [00401078h] ; rtcVarFromFormatVar
  loc_005AB0DA: mov var_320, 0042EC30h ; "','"
  loc_005AB0E4: mov var_328, 00000008h
loc_005AB0EE:   movsx ecx, var_24
loc_005AB0F2:   mov var_330, ecx
  loc_005AB0F8: mov var_338, 00000003h
  loc_005AB102: mov var_350, 00000003h
  loc_005AB10C: mov var_358, 00000003h
  loc_005AB116: mov eax, 00000010h
  loc_005AB11B: call 00405690h ; __vbaChkstk
loc_005AB120:   mov edx, esp
loc_005AB122:   mov eax, var_338
loc_005AB128:   mov [edx], eax
loc_005AB12A:   mov ecx, var_334
loc_005AB130:   mov [edx+00000004h], ecx
loc_005AB133:   mov eax, var_330
loc_005AB139:   mov [edx+00000008h], eax
loc_005AB13C:   mov ecx, var_32C
loc_005AB142:   mov [edx+0000000Ch], ecx
  loc_005AB145: mov eax, 00000010h
  loc_005AB14A: call 00405690h ; __vbaChkstk
loc_005AB14F:   mov edx, esp
loc_005AB151:   mov eax, var_358
loc_005AB157:   mov [edx], eax
loc_005AB159:   mov ecx, var_354
loc_005AB15F:   mov [edx+00000004h], ecx
loc_005AB162:   mov eax, var_350
loc_005AB168:   mov [edx+00000008h], eax
loc_005AB16B:   mov ecx, var_34C
loc_005AB171:   mov [edx+0000000Ch], ecx
  loc_005AB174: push 00000002h
  loc_005AB176: push 00000041h
loc_005AB178:   mov edx, Me
loc_005AB17B:   mov eax, [edx]
loc_005AB17D:   mov ecx, Me
loc_005AB180:   push ecx
loc_005AB181:   Call [eax+00000390h]
loc_005AB187:   push eax
loc_005AB188:   lea edx, var_70
loc_005AB18B:   push edx
  loc_005AB18C: call [004010B8h] ; __vbaObjSet
loc_005AB192:   push eax
loc_005AB193:   lea eax, var_118
loc_005AB199:   push eax
  loc_005AB19A: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AB1A0: add esp, 00000030h
loc_005AB1A3:   push eax
  loc_005AB1A4: call [00401034h] ; __vbaStrVarMove
loc_005AB1AA:   mov var_120, eax
  loc_005AB1B0: mov var_128, 00000008h
  loc_005AB1BA: mov var_370, 0042EC30h ; "','"
  loc_005AB1C4: mov var_378, 00000008h
loc_005AB1CE:   movsx ecx, var_24
loc_005AB1D2:   mov var_380, ecx
  loc_005AB1D8: mov var_388, 00000003h
  loc_005AB1E2: mov var_3A0, 00000004h
  loc_005AB1EC: mov var_3A8, 00000003h
  loc_005AB1F6: mov eax, 00000010h
  loc_005AB1FB: call 00405690h ; __vbaChkstk
loc_005AB200:   mov edx, esp
loc_005AB202:   mov eax, var_388
loc_005AB208:   mov [edx], eax
loc_005AB20A:   mov ecx, var_384
loc_005AB210:   mov [edx+00000004h], ecx
loc_005AB213:   mov eax, var_380
loc_005AB219:   mov [edx+00000008h], eax
loc_005AB21C:   mov ecx, var_37C
loc_005AB222:   mov [edx+0000000Ch], ecx
  loc_005AB225: mov eax, 00000010h
  loc_005AB22A: call 00405690h ; __vbaChkstk
loc_005AB22F:   mov edx, esp
loc_005AB231:   mov eax, var_3A8
loc_005AB237:   mov [edx], eax
loc_005AB239:   mov ecx, var_3A4
loc_005AB23F:   mov [edx+00000004h], ecx
loc_005AB242:   mov eax, var_3A0
loc_005AB248:   mov [edx+00000008h], eax
loc_005AB24B:   mov ecx, var_39C
loc_005AB251:   mov [edx+0000000Ch], ecx
  loc_005AB254: push 00000002h
  loc_005AB256: push 00000041h
loc_005AB258:   mov edx, Me
loc_005AB25B:   mov eax, [edx]
loc_005AB25D:   mov ecx, Me
loc_005AB260:   push ecx
loc_005AB261:   Call [eax+00000390h]
loc_005AB267:   push eax
loc_005AB268:   lea edx, var_74
loc_005AB26B:   push edx
  loc_005AB26C: call [004010B8h] ; __vbaObjSet
loc_005AB272:   push eax
loc_005AB273:   lea eax, var_158
loc_005AB279:   push eax
  loc_005AB27A: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AB280: add esp, 00000030h
loc_005AB283:   push eax
  loc_005AB284: call [00401034h] ; __vbaStrVarMove
loc_005AB28A:   mov var_160, eax
  loc_005AB290: mov var_168, 00000008h
  loc_005AB29A: mov var_3C0, 0042EC30h ; "','"
  loc_005AB2A4: mov var_3C8, 00000008h
  loc_005AB2AE: mov var_410, 00435A20h ; "00"
  loc_005AB2B8: mov var_418, 00000008h
loc_005AB2C2:   lea edx, var_418
loc_005AB2C8:   lea ecx, var_1B8
  loc_005AB2CE: call [00401238h] ; __vbaVarDup
loc_005AB2D4:   lea ecx, var_198
loc_005AB2DA:   push ecx
  loc_005AB2DB: call [00401034h] ; __vbaStrVarMove
loc_005AB2E1:   mov var_1A0, eax
  loc_005AB2E7: mov var_1A8, 00000008h
  loc_005AB2F1: push 00000001h
  loc_005AB2F3: push 00000001h
loc_005AB2F5:   lea edx, var_1B8
loc_005AB2FB:   push edx
loc_005AB2FC:   lea eax, var_1A8
loc_005AB302:   push eax
loc_005AB303:   lea ecx, var_1C8
loc_005AB309:   push ecx
  loc_005AB30A: call [00401078h] ; rtcVarFromFormatVar
  loc_005AB310: mov var_420, 0042EC3Ch ; "')"
  loc_005AB31A: mov var_428, 00000008h
loc_005AB324:   lea edx, var_E8
loc_005AB32A:   push edx
loc_005AB32B:   lea eax, var_D8
loc_005AB331:   push eax
loc_005AB332:   lea ecx, var_F8
loc_005AB338:   push ecx
  loc_005AB339: call [004011C4h] ; __vbaVarCat
loc_005AB33F:   push eax
loc_005AB340:   lea edx, var_328
loc_005AB346:   push edx
loc_005AB347:   lea eax, var_108
loc_005AB34D:   push eax
  loc_005AB34E: call [004011C4h] ; __vbaVarCat
loc_005AB354:   push eax
loc_005AB355:   lea ecx, var_128
loc_005AB35B:   push ecx
loc_005AB35C:   lea edx, var_138
loc_005AB362:   push edx
  loc_005AB363: call [004011C4h] ; __vbaVarCat
loc_005AB369:   push eax
loc_005AB36A:   lea eax, var_378
loc_005AB370:   push eax
loc_005AB371:   lea ecx, var_148
loc_005AB377:   push ecx
  loc_005AB378: call [004011C4h] ; __vbaVarCat
loc_005AB37E:   push eax
loc_005AB37F:   lea edx, var_168
loc_005AB385:   push edx
loc_005AB386:   lea eax, var_178
loc_005AB38C:   push eax
  loc_005AB38D: call [004011C4h] ; __vbaVarCat
loc_005AB393:   push eax
loc_005AB394:   lea ecx, var_3C8
loc_005AB39A:   push ecx
loc_005AB39B:   lea edx, var_188
loc_005AB3A1:   push edx
  loc_005AB3A2: call [004011C4h] ; __vbaVarCat
loc_005AB3A8:   push eax
loc_005AB3A9:   lea eax, var_1C8
loc_005AB3AF:   push eax
loc_005AB3B0:   lea ecx, var_1D8
loc_005AB3B6:   push ecx
  loc_005AB3B7: call [004011C4h] ; __vbaVarCat
loc_005AB3BD:   push eax
loc_005AB3BE:   lea edx, var_428
loc_005AB3C4:   push edx
loc_005AB3C5:   lea eax, var_1E8
loc_005AB3CB:   push eax
  loc_005AB3CC: call [004011C4h] ; __vbaVarCat
loc_005AB3D2:   push eax
  loc_005AB3D3: call [00401034h] ; __vbaStrVarMove
loc_005AB3D9:   mov edx, eax
  loc_005AB3DB: mov ecx, 0066809Ch
  loc_005AB3E0: call [0040126Ch] ; __vbaStrMove
loc_005AB3E6:   lea ecx, var_58
loc_005AB3E9:   push ecx
loc_005AB3EA:   lea edx, var_54
loc_005AB3ED:   push edx
loc_005AB3EE:   lea eax, var_50
loc_005AB3F1:   push eax
loc_005AB3F2:   lea ecx, var_4C
loc_005AB3F5:   push ecx
loc_005AB3F6:   lea edx, var_48
loc_005AB3F9:   push edx
loc_005AB3FA:   lea eax, var_44
loc_005AB3FD:   push eax
loc_005AB3FE:   lea ecx, var_40
loc_005AB401:   push ecx
loc_005AB402:   lea edx, var_3C
loc_005AB405:   push edx
loc_005AB406:   lea eax, var_38
loc_005AB409:   push eax
  loc_005AB40A: push 00000009h
  loc_005AB40C: call [00401204h] ; __vbaFreeStrList
  loc_005AB412: add esp, 00000028h
loc_005AB415:   lea ecx, var_78
loc_005AB418:   push ecx
loc_005AB419:   lea edx, var_74
loc_005AB41C:   push edx
loc_005AB41D:   lea eax, var_70
loc_005AB420:   push eax
loc_005AB421:   lea ecx, var_6C
loc_005AB424:   push ecx
loc_005AB425:   lea edx, var_68
loc_005AB428:   push edx
loc_005AB429:   lea eax, var_64
loc_005AB42C:   push eax
  loc_005AB42D: push 00000006h
  loc_005AB42F: call [0040104Ch] ; __vbaFreeObjList
  loc_005AB435: add esp, 0000001Ch
loc_005AB438:   lea ecx, var_1E8
loc_005AB43E:   push ecx
loc_005AB43F:   lea edx, var_1D8
loc_005AB445:   push edx
loc_005AB446:   lea eax, var_1C8
loc_005AB44C:   push eax
loc_005AB44D:   lea ecx, var_188
loc_005AB453:   push ecx
loc_005AB454:   lea edx, var_1B8
loc_005AB45A:   push edx
loc_005AB45B:   lea eax, var_1A8
loc_005AB461:   push eax
loc_005AB462:   lea ecx, var_198
loc_005AB468:   push ecx
loc_005AB469:   lea edx, var_178
loc_005AB46F:   push edx
loc_005AB470:   lea eax, var_168
loc_005AB476:   push eax
loc_005AB477:   lea ecx, var_148
loc_005AB47D:   push ecx
loc_005AB47E:   lea edx, var_158
loc_005AB484:   push edx
loc_005AB485:   lea eax, var_138
loc_005AB48B:   push eax
loc_005AB48C:   lea ecx, var_128
loc_005AB492:   push ecx
loc_005AB493:   lea edx, var_108
loc_005AB499:   push edx
loc_005AB49A:   lea eax, var_118
loc_005AB4A0:   push eax
loc_005AB4A1:   lea ecx, var_F8
loc_005AB4A7:   push ecx
loc_005AB4A8:   lea edx, var_D8
loc_005AB4AE:   push edx
loc_005AB4AF:   lea eax, var_E8
loc_005AB4B5:   push eax
loc_005AB4B6:   lea ecx, var_C8
loc_005AB4BC:   push ecx
loc_005AB4BD:   lea edx, var_B8
loc_005AB4C3:   push edx
loc_005AB4C4:   lea eax, var_A8
loc_005AB4CA:   push eax
loc_005AB4CB:   lea ecx, var_98
loc_005AB4D1:   push ecx
loc_005AB4D2:   lea edx, var_88
loc_005AB4D8:   push edx
  loc_005AB4D9: push 00000017h
  loc_005AB4DB: call [0040103Ch] ; __vbaFreeVarList
  loc_005AB4E1: add esp, 00000060h
  loc_005AB4E4: mov var_4, 00000014h
  loc_005AB4EB: cmp [0066803Ch], 00000000h
  loc_005AB4F2: jnz 005AB510h
  loc_005AB4F4: push 0066803Ch
  loc_005AB4F9: push 0042DEFCh
  loc_005AB4FE: call [004011E8h] ; __vbaNew2
  loc_005AB504: mov var_4EC, 0066803Ch
  loc_005AB50E: jmp 005AB51Ah
  loc_005AB510: mov var_4EC, 0066803Ch
loc_005AB51A:   mov eax, var_4EC
loc_005AB520:   mov ecx, [eax]
loc_005AB522:   mov var_434, ecx
  loc_005AB528: mov var_80, 80020004h
  loc_005AB52F: mov var_88, 0000000Ah
loc_005AB539:   lea ecx, var_88
  loc_005AB53F: call [0040123Ch] ; __vbaFreeVarg
loc_005AB545:   lea edx, var_64
loc_005AB548:   push edx
  loc_005AB549: push 00000001h
loc_005AB54B:   lea eax, var_88
loc_005AB551:   push eax
loc_005AB552:   mov ecx, [0066809Ch]
loc_005AB558:   push ecx
loc_005AB559:   mov edx, var_434
loc_005AB55F:   mov eax, [edx]
loc_005AB561:   mov ecx, var_434
loc_005AB567:   push ecx
loc_005AB568:   Call [eax+00000040h]
loc_005AB56B:   fnclex
loc_005AB56D:   mov var_438, eax
  loc_005AB573: cmp var_438, 00000000h
  loc_005AB57A: jge 005AB59Fh
  loc_005AB57C: push 00000040h
  loc_005AB57E: push 0042E1B0h
loc_005AB583:   mov edx, var_434
loc_005AB589:   push edx
loc_005AB58A:   mov eax, var_438
loc_005AB590:   push eax
  loc_005AB591: call [00401080h] ; __vbaHresultCheckObj
loc_005AB597:   mov var_4F0, eax
  loc_005AB59D: jmp 005AB5A9h
  loc_005AB59F: mov var_4F0, 00000000h
loc_005AB5A9:   lea ecx, var_64
  loc_005AB5AC: call [004012A0h] ; __vbaFreeObj
loc_005AB5B2:   lea ecx, var_88
  loc_005AB5B8: call [00401020h] ; __vbaFreeVar
  loc_005AB5BE: mov var_4, 00000015h
loc_005AB5C5:   movsx ecx, var_24
loc_005AB5C9:   mov var_250, ecx
  loc_005AB5CF: mov var_258, 00000003h
  loc_005AB5D9: mov var_270, 00000004h
  loc_005AB5E3: mov var_278, 00000003h
  loc_005AB5ED: mov eax, 00000010h
  loc_005AB5F2: call 00405690h ; __vbaChkstk
loc_005AB5F7:   mov edx, esp
loc_005AB5F9:   mov eax, var_258
loc_005AB5FF:   mov [edx], eax
loc_005AB601:   mov ecx, var_254
loc_005AB607:   mov [edx+00000004h], ecx
loc_005AB60A:   mov eax, var_250
loc_005AB610:   mov [edx+00000008h], eax
loc_005AB613:   mov ecx, var_24C
loc_005AB619:   mov [edx+0000000Ch], ecx
  loc_005AB61C: mov eax, 00000010h
  loc_005AB621: call 00405690h ; __vbaChkstk
loc_005AB626:   mov edx, esp
loc_005AB628:   mov eax, var_278
loc_005AB62E:   mov [edx], eax
loc_005AB630:   mov ecx, var_274
loc_005AB636:   mov [edx+00000004h], ecx
loc_005AB639:   mov eax, var_270
loc_005AB63F:   mov [edx+00000008h], eax
loc_005AB642:   mov ecx, var_26C
loc_005AB648:   mov [edx+0000000Ch], ecx
  loc_005AB64B: push 00000002h
  loc_005AB64D: push 00000041h
loc_005AB64F:   mov edx, Me
loc_005AB652:   mov eax, [edx]
loc_005AB654:   mov ecx, Me
loc_005AB657:   push ecx
loc_005AB658:   Call [eax+00000390h]
loc_005AB65E:   push eax
loc_005AB65F:   lea edx, var_64
loc_005AB662:   push edx
  loc_005AB663: call [004010B8h] ; __vbaObjSet
loc_005AB669:   push eax
loc_005AB66A:   lea eax, var_88
loc_005AB670:   push eax
  loc_005AB671: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AB677: add esp, 00000030h
loc_005AB67A:   push eax
  loc_005AB67B: call [00401034h] ; __vbaStrVarMove
loc_005AB681:   mov edx, eax
loc_005AB683:   lea ecx, var_38
  loc_005AB686: call [0040126Ch] ; __vbaStrMove
loc_005AB68C:   push eax
  loc_005AB68D: call [004012A4h] ; rtcR8ValFromBstr
  loc_005AB693: call [00401244h] ; __vbaFpI2
loc_005AB699:   mov var_28, ax
loc_005AB69D:   lea ecx, var_38
  loc_005AB6A0: call [0040129Ch] ; __vbaFreeStr
loc_005AB6A6:   lea ecx, var_64
  loc_005AB6A9: call [004012A0h] ; __vbaFreeObj
loc_005AB6AF:   lea ecx, var_88
  loc_005AB6B5: call [00401020h] ; __vbaFreeVar
  loc_005AB6BB: mov var_4, 00000016h
loc_005AB6C2:   movsx ecx, var_24
loc_005AB6C6:   mov var_250, ecx
  loc_005AB6CC: mov var_258, 00000003h
  loc_005AB6D6: mov var_270, 00000000h
  loc_005AB6E0: mov var_278, 00000003h
  loc_005AB6EA: mov eax, 00000010h
  loc_005AB6EF: call 00405690h ; __vbaChkstk
loc_005AB6F4:   mov edx, esp
loc_005AB6F6:   mov eax, var_258
loc_005AB6FC:   mov [edx], eax
loc_005AB6FE:   mov ecx, var_254
loc_005AB704:   mov [edx+00000004h], ecx
loc_005AB707:   mov eax, var_250
loc_005AB70D:   mov [edx+00000008h], eax
loc_005AB710:   mov ecx, var_24C
loc_005AB716:   mov [edx+0000000Ch], ecx
  loc_005AB719: mov eax, 00000010h
  loc_005AB71E: call 00405690h ; __vbaChkstk
loc_005AB723:   mov edx, esp
loc_005AB725:   mov eax, var_278
loc_005AB72B:   mov [edx], eax
loc_005AB72D:   mov ecx, var_274
loc_005AB733:   mov [edx+00000004h], ecx
loc_005AB736:   mov eax, var_270
loc_005AB73C:   mov [edx+00000008h], eax
loc_005AB73F:   mov ecx, var_26C
loc_005AB745:   mov [edx+0000000Ch], ecx
  loc_005AB748: push 00000002h
  loc_005AB74A: push 00000041h
loc_005AB74C:   mov edx, Me
loc_005AB74F:   mov eax, [edx]
loc_005AB751:   mov ecx, Me
loc_005AB754:   push ecx
loc_005AB755:   Call [eax+00000390h]
loc_005AB75B:   push eax
loc_005AB75C:   lea edx, var_64
loc_005AB75F:   push edx
  loc_005AB760: call [004010B8h] ; __vbaObjSet
loc_005AB766:   push eax
loc_005AB767:   lea eax, var_88
loc_005AB76D:   push eax
  loc_005AB76E: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AB774: add esp, 00000030h
loc_005AB777:   push eax
  loc_005AB778: call [00401034h] ; __vbaStrVarMove
loc_005AB77E:   mov edx, eax
loc_005AB780:   lea ecx, var_2C
  loc_005AB783: call [0040126Ch] ; __vbaStrMove
loc_005AB789:   lea ecx, var_64
  loc_005AB78C: call [004012A0h] ; __vbaFreeObj
loc_005AB792:   lea ecx, var_88
  loc_005AB798: call [00401020h] ; __vbaFreeVar
  loc_005AB79E: mov var_4, 00000017h
  loc_005AB7A5: push 00435F2Ch ; "UPDATE brg_barang SET stok=stok - "
loc_005AB7AA:   mov cx, var_28
loc_005AB7AE:   push ecx
  loc_005AB7AF: call [0040100Ch] ; __vbaStrI2
loc_005AB7B5:   mov edx, eax
loc_005AB7B7:   lea ecx, var_38
  loc_005AB7BA: call [0040126Ch] ; __vbaStrMove
loc_005AB7C0:   push eax
  loc_005AB7C1: call [00401070h] ; __vbaStrCat
loc_005AB7C7:   mov edx, eax
loc_005AB7C9:   lea ecx, var_3C
  loc_005AB7CC: call [0040126Ch] ; __vbaStrMove
loc_005AB7D2:   push eax
  loc_005AB7D3: push 0042DFECh
  loc_005AB7D8: call [00401070h] ; __vbaStrCat
loc_005AB7DE:   mov edx, eax
loc_005AB7E0:   lea ecx, var_40
  loc_005AB7E3: call [0040126Ch] ; __vbaStrMove
loc_005AB7E9:   push eax
  loc_005AB7EA: push 00433824h ; " WHERE kode_barang='"
  loc_005AB7EF: call [00401070h] ; __vbaStrCat
loc_005AB7F5:   mov edx, eax
loc_005AB7F7:   lea ecx, var_44
  loc_005AB7FA: call [0040126Ch] ; __vbaStrMove
loc_005AB800:   push eax
loc_005AB801:   mov edx, var_2C
loc_005AB804:   push edx
  loc_005AB805: call [00401070h] ; __vbaStrCat
loc_005AB80B:   mov edx, eax
loc_005AB80D:   lea ecx, var_48
  loc_005AB810: call [0040126Ch] ; __vbaStrMove
loc_005AB816:   push eax
  loc_005AB817: push 0042EBA8h ; "'"
  loc_005AB81C: call [00401070h] ; __vbaStrCat
loc_005AB822:   mov edx, eax
  loc_005AB824: mov ecx, 0066809Ch
  loc_005AB829: call [0040126Ch] ; __vbaStrMove
loc_005AB82F:   lea eax, var_48
loc_005AB832:   push eax
loc_005AB833:   lea ecx, var_44
loc_005AB836:   push ecx
loc_005AB837:   lea edx, var_40
loc_005AB83A:   push edx
loc_005AB83B:   lea eax, var_3C
loc_005AB83E:   push eax
loc_005AB83F:   lea ecx, var_38
loc_005AB842:   push ecx
  loc_005AB843: push 00000005h
  loc_005AB845: call [00401204h] ; __vbaFreeStrList
  loc_005AB84B: add esp, 00000018h
  loc_005AB84E: mov var_4, 00000018h
  loc_005AB855: cmp [0066803Ch], 00000000h
  loc_005AB85C: jnz 005AB87Ah
  loc_005AB85E: push 0066803Ch
  loc_005AB863: push 0042DEFCh
  loc_005AB868: call [004011E8h] ; __vbaNew2
  loc_005AB86E: mov var_4F4, 0066803Ch
  loc_005AB878: jmp 005AB884h
  loc_005AB87A: mov var_4F4, 0066803Ch
loc_005AB884:   mov edx, var_4F4
loc_005AB88A:   mov eax, [edx]
loc_005AB88C:   mov var_434, eax
  loc_005AB892: mov var_80, 80020004h
  loc_005AB899: mov var_88, 0000000Ah
loc_005AB8A3:   lea ecx, var_88
  loc_005AB8A9: call [0040123Ch] ; __vbaFreeVarg
loc_005AB8AF:   lea ecx, var_64
loc_005AB8B2:   push ecx
  loc_005AB8B3: push 00000001h
loc_005AB8B5:   lea edx, var_88
loc_005AB8BB:   push edx
loc_005AB8BC:   mov eax, [0066809Ch]
loc_005AB8C1:   push eax
loc_005AB8C2:   mov ecx, var_434
loc_005AB8C8:   mov edx, [ecx]
loc_005AB8CA:   mov eax, var_434
loc_005AB8D0:   push eax
loc_005AB8D1:   Call [edx+00000040h]
loc_005AB8D4:   fnclex
loc_005AB8D6:   mov var_438, eax
  loc_005AB8DC: cmp var_438, 00000000h
  loc_005AB8E3: jge 005AB908h
  loc_005AB8E5: push 00000040h
  loc_005AB8E7: push 0042E1B0h
loc_005AB8EC:   mov ecx, var_434
loc_005AB8F2:   push ecx
loc_005AB8F3:   mov edx, var_438
loc_005AB8F9:   push edx
  loc_005AB8FA: call [00401080h] ; __vbaHresultCheckObj
loc_005AB900:   mov var_4F8, eax
  loc_005AB906: jmp 005AB912h
  loc_005AB908: mov var_4F8, 00000000h
loc_005AB912:   lea ecx, var_64
  loc_005AB915: call [004012A0h] ; __vbaFreeObj
loc_005AB91B:   lea ecx, var_88
  loc_005AB921: call [00401020h] ; __vbaFreeVar
  loc_005AB927: mov var_4, 00000019h
  loc_005AB92E: jmp 005AAC98h
  loc_005AB933: mov var_4, 0000001Ah
loc_005AB93A:   push FFFFFFFFh
  loc_005AB93C: call [004010BCh] ; __vbaOnError
  loc_005AB942: mov var_4, 0000001Bh
  loc_005AB949: cmp [00668454h], 00000000h
  loc_005AB950: jnz 005AB96Eh
  loc_005AB952: push 00668454h
  loc_005AB957: push 0040F860h
  loc_005AB95C: call [004011E8h] ; __vbaNew2
  loc_005AB962: mov var_4FC, 00668454h
  loc_005AB96C: jmp 005AB978h
  loc_005AB96E: mov var_4FC, 00668454h
loc_005AB978:   mov eax, var_4FC
loc_005AB97E:   mov ecx, [eax]
loc_005AB980:   push ecx
loc_005AB981:   lea edx, var_44C
loc_005AB987:   push edx
  loc_005AB988: call [004010C8h] ; __vbaObjSetAddref
  loc_005AB98E: mov var_4, 0000001Ch
loc_005AB995:   mov eax, var_44C
loc_005AB99B:   mov ecx, [eax]
loc_005AB99D:   mov edx, var_44C
loc_005AB9A3:   push edx
loc_005AB9A4:   Call [ecx+0000031Ch]
loc_005AB9AA:   push eax
loc_005AB9AB:   lea eax, var_64
loc_005AB9AE:   push eax
  loc_005AB9AF: call [004010B8h] ; __vbaObjSet
loc_005AB9B5:   mov var_434, eax
  loc_005AB9BB: push 00435FDCh ; "Uang Bayar"
loc_005AB9C0:   mov ecx, var_434
loc_005AB9C6:   mov edx, [ecx]
loc_005AB9C8:   mov eax, var_434
loc_005AB9CE:   push eax
loc_005AB9CF:   Call [edx+00000054h]
loc_005AB9D2:   fnclex
loc_005AB9D4:   mov var_438, eax
  loc_005AB9DA: cmp var_438, 00000000h
  loc_005AB9E1: jge 005ABA06h
  loc_005AB9E3: push 00000054h
  loc_005AB9E5: push 00433F80h
loc_005AB9EA:   mov ecx, var_434
loc_005AB9F0:   push ecx
loc_005AB9F1:   mov edx, var_438
loc_005AB9F7:   push edx
  loc_005AB9F8: call [00401080h] ; __vbaHresultCheckObj
loc_005AB9FE:   mov var_500, eax
  loc_005ABA04: jmp 005ABA10h
  loc_005ABA06: mov var_500, 00000000h
loc_005ABA10:   lea ecx, var_64
  loc_005ABA13: call [004012A0h] ; __vbaFreeObj
  loc_005ABA19: mov var_4, 0000001Dh
loc_005ABA20:   mov eax, var_44C
loc_005ABA26:   mov ecx, [eax]
loc_005ABA28:   mov edx, var_44C
loc_005ABA2E:   push edx
loc_005ABA2F:   Call [ecx+00000318h]
loc_005ABA35:   push eax
loc_005ABA36:   lea eax, var_64
loc_005ABA39:   push eax
  loc_005ABA3A: call [004010B8h] ; __vbaObjSet
loc_005ABA40:   mov var_434, eax
  loc_005ABA46: push 00435FF8h ; "Uang Kembali"
loc_005ABA4B:   mov ecx, var_434
loc_005ABA51:   mov edx, [ecx]
loc_005ABA53:   mov eax, var_434
loc_005ABA59:   push eax
loc_005ABA5A:   Call [edx+00000054h]
loc_005ABA5D:   fnclex
loc_005ABA5F:   mov var_438, eax
  loc_005ABA65: cmp var_438, 00000000h
  loc_005ABA6C: jge 005ABA91h
  loc_005ABA6E: push 00000054h
  loc_005ABA70: push 00433F80h
loc_005ABA75:   mov ecx, var_434
loc_005ABA7B:   push ecx
loc_005ABA7C:   mov edx, var_438
loc_005ABA82:   push edx
  loc_005ABA83: call [00401080h] ; __vbaHresultCheckObj
loc_005ABA89:   mov var_504, eax
  loc_005ABA8F: jmp 005ABA9Bh
  loc_005ABA91: mov var_504, 00000000h
loc_005ABA9B:   lea ecx, var_64
  loc_005ABA9E: call [004012A0h] ; __vbaFreeObj
  loc_005ABAA4: mov var_4, 0000001Eh
loc_005ABAAB:   mov eax, var_44C
loc_005ABAB1:   mov ecx, [eax]
loc_005ABAB3:   mov edx, var_44C
loc_005ABAB9:   push edx
loc_005ABABA:   Call [ecx+00000308h]
loc_005ABAC0:   push eax
loc_005ABAC1:   lea eax, var_68
loc_005ABAC4:   push eax
  loc_005ABAC5: call [004010B8h] ; __vbaObjSet
loc_005ABACB:   mov var_43C, eax
loc_005ABAD1:   mov ecx, Me
loc_005ABAD4:   mov edx, [ecx]
loc_005ABAD6:   mov eax, Me
loc_005ABAD9:   push eax
loc_005ABADA:   Call [edx+0000030Ch]
loc_005ABAE0:   push eax
loc_005ABAE1:   lea ecx, var_64
loc_005ABAE4:   push ecx
  loc_005ABAE5: call [004010B8h] ; __vbaObjSet
loc_005ABAEB:   mov var_434, eax
loc_005ABAF1:   lea edx, var_38
loc_005ABAF4:   push edx
loc_005ABAF5:   mov eax, var_434
loc_005ABAFB:   mov ecx, [eax]
loc_005ABAFD:   mov edx, var_434
loc_005ABB03:   push edx
loc_005ABB04:   Call [ecx+000000A0h]
loc_005ABB0A:   fnclex
loc_005ABB0C:   mov var_438, eax
  loc_005ABB12: cmp var_438, 00000000h
  loc_005ABB19: jge 005ABB41h
  loc_005ABB1B: push 000000A0h
  loc_005ABB20: push 0042DFCCh
loc_005ABB25:   mov eax, var_434
loc_005ABB2B:   push eax
loc_005ABB2C:   mov ecx, var_438
loc_005ABB32:   push ecx
  loc_005ABB33: call [00401080h] ; __vbaHresultCheckObj
loc_005ABB39:   mov var_508, eax
  loc_005ABB3F: jmp 005ABB4Bh
  loc_005ABB41: mov var_508, 00000000h
loc_005ABB4B:   mov edx, var_38
loc_005ABB4E:   push edx
loc_005ABB4F:   mov eax, var_43C
loc_005ABB55:   mov ecx, [eax]
loc_005ABB57:   mov edx, var_43C
loc_005ABB5D:   push edx
loc_005ABB5E:   Call [ecx+000000A4h]
loc_005ABB64:   fnclex
loc_005ABB66:   mov var_440, eax
  loc_005ABB6C: cmp var_440, 00000000h
  loc_005ABB73: jge 005ABB9Bh
  loc_005ABB75: push 000000A4h
  loc_005ABB7A: push 0042DFCCh
loc_005ABB7F:   mov eax, var_43C
loc_005ABB85:   push eax
loc_005ABB86:   mov ecx, var_440
loc_005ABB8C:   push ecx
  loc_005ABB8D: call [00401080h] ; __vbaHresultCheckObj
loc_005ABB93:   mov var_50C, eax
  loc_005ABB99: jmp 005ABBA5h
  loc_005ABB9B: mov var_50C, 00000000h
loc_005ABBA5:   lea ecx, var_38
  loc_005ABBA8: call [0040129Ch] ; __vbaFreeStr
loc_005ABBAE:   lea edx, var_68
loc_005ABBB1:   push edx
loc_005ABBB2:   lea eax, var_64
loc_005ABBB5:   push eax
  loc_005ABBB6: push 00000002h
  loc_005ABBB8: call [0040104Ch] ; __vbaFreeObjList
  loc_005ABBBE: add esp, 0000000Ch
  loc_005ABBC1: mov var_4, 0000001Fh
loc_005ABBC8:   mov ecx, var_44C
loc_005ABBCE:   mov edx, [ecx]
loc_005ABBD0:   mov eax, var_44C
loc_005ABBD6:   push eax
loc_005ABBD7:   Call [edx+00000304h]
loc_005ABBDD:   push eax
loc_005ABBDE:   lea ecx, var_68
loc_005ABBE1:   push ecx
  loc_005ABBE2: call [004010B8h] ; __vbaObjSet
loc_005ABBE8:   mov var_43C, eax
loc_005ABBEE:   mov edx, Me
loc_005ABBF1:   mov eax, [edx]
loc_005ABBF3:   mov ecx, Me
loc_005ABBF6:   push ecx
loc_005ABBF7:   Call [eax+00000338h]
loc_005ABBFD:   push eax
loc_005ABBFE:   lea edx, var_64
loc_005ABC01:   push edx
  loc_005ABC02: call [004010B8h] ; __vbaObjSet
loc_005ABC08:   mov var_434, eax
loc_005ABC0E:   lea eax, var_38
loc_005ABC11:   push eax
loc_005ABC12:   mov ecx, var_434
loc_005ABC18:   mov edx, [ecx]
loc_005ABC1A:   mov eax, var_434
loc_005ABC20:   push eax
loc_005ABC21:   Call [edx+00000050h]
loc_005ABC24:   fnclex
loc_005ABC26:   mov var_438, eax
  loc_005ABC2C: cmp var_438, 00000000h
  loc_005ABC33: jge 005ABC58h
  loc_005ABC35: push 00000050h
  loc_005ABC37: push 00433F80h
loc_005ABC3C:   mov ecx, var_434
loc_005ABC42:   push ecx
loc_005ABC43:   mov edx, var_438
loc_005ABC49:   push edx
  loc_005ABC4A: call [00401080h] ; __vbaHresultCheckObj
loc_005ABC50:   mov var_510, eax
  loc_005ABC56: jmp 005ABC62h
  loc_005ABC58: mov var_510, 00000000h
loc_005ABC62:   mov eax, var_38
loc_005ABC65:   push eax
loc_005ABC66:   mov ecx, var_43C
loc_005ABC6C:   mov edx, [ecx]
loc_005ABC6E:   mov eax, var_43C
loc_005ABC74:   push eax
loc_005ABC75:   Call [edx+000000A4h]
loc_005ABC7B:   fnclex
loc_005ABC7D:   mov var_440, eax
  loc_005ABC83: cmp var_440, 00000000h
  loc_005ABC8A: jge 005ABCB2h
  loc_005ABC8C: push 000000A4h
  loc_005ABC91: push 0042DFCCh
loc_005ABC96:   mov ecx, var_43C
loc_005ABC9C:   push ecx
loc_005ABC9D:   mov edx, var_440
loc_005ABCA3:   push edx
  loc_005ABCA4: call [00401080h] ; __vbaHresultCheckObj
loc_005ABCAA:   mov var_514, eax
  loc_005ABCB0: jmp 005ABCBCh
  loc_005ABCB2: mov var_514, 00000000h
loc_005ABCBC:   lea ecx, var_38
  loc_005ABCBF: call [0040129Ch] ; __vbaFreeStr
loc_005ABCC5:   lea eax, var_68
loc_005ABCC8:   push eax
loc_005ABCC9:   lea ecx, var_64
loc_005ABCCC:   push ecx
  loc_005ABCCD: push 00000002h
  loc_005ABCCF: call [0040104Ch] ; __vbaFreeObjList
  loc_005ABCD5: add esp, 0000000Ch
  loc_005ABCD8: mov var_4, 00000020h
loc_005ABCDF:   mov edx, var_44C
loc_005ABCE5:   mov eax, [edx]
loc_005ABCE7:   mov ecx, var_44C
loc_005ABCED:   push ecx
loc_005ABCEE:   Call [eax+00000300h]
loc_005ABCF4:   push eax
loc_005ABCF5:   lea edx, var_64
loc_005ABCF8:   push edx
  loc_005ABCF9: call [004010B8h] ; __vbaObjSet
loc_005ABCFF:   mov var_434, eax
  loc_005ABD05: push 0042DFECh
loc_005ABD0A:   mov eax, var_434
loc_005ABD10:   mov ecx, [eax]
loc_005ABD12:   mov edx, var_434
loc_005ABD18:   push edx
loc_005ABD19:   Call [ecx+000000A4h]
loc_005ABD1F:   fnclex
loc_005ABD21:   mov var_438, eax
  loc_005ABD27: cmp var_438, 00000000h
  loc_005ABD2E: jge 005ABD56h
  loc_005ABD30: push 000000A4h
  loc_005ABD35: push 0042DFCCh
loc_005ABD3A:   mov eax, var_434
loc_005ABD40:   push eax
loc_005ABD41:   mov ecx, var_438
loc_005ABD47:   push ecx
  loc_005ABD48: call [00401080h] ; __vbaHresultCheckObj
loc_005ABD4E:   mov var_518, eax
  loc_005ABD54: jmp 005ABD60h
  loc_005ABD56: mov var_518, 00000000h
loc_005ABD60:   lea ecx, var_64
  loc_005ABD63: call [004012A0h] ; __vbaFreeObj
  loc_005ABD69: mov var_4, 00000021h
loc_005ABD70:   mov edx, var_44C
loc_005ABD76:   mov eax, [edx]
loc_005ABD78:   mov ecx, var_44C
loc_005ABD7E:   push ecx
loc_005ABD7F:   Call [eax+000002FCh]
loc_005ABD85:   push eax
loc_005ABD86:   lea edx, var_68
loc_005ABD89:   push edx
  loc_005ABD8A: call [004010B8h] ; __vbaObjSet
loc_005ABD90:   mov var_43C, eax
loc_005ABD96:   mov eax, Me
loc_005ABD99:   mov ecx, [eax]
loc_005ABD9B:   mov edx, Me
loc_005ABD9E:   push edx
loc_005ABD9F:   Call [ecx+00000338h]
loc_005ABDA5:   push eax
loc_005ABDA6:   lea eax, var_64
loc_005ABDA9:   push eax
  loc_005ABDAA: call [004010B8h] ; __vbaObjSet
loc_005ABDB0:   mov var_434, eax
loc_005ABDB6:   lea ecx, var_38
loc_005ABDB9:   push ecx
loc_005ABDBA:   mov edx, var_434
loc_005ABDC0:   mov eax, [edx]
loc_005ABDC2:   mov ecx, var_434
loc_005ABDC8:   push ecx
loc_005ABDC9:   Call [eax+00000050h]
loc_005ABDCC:   fnclex
loc_005ABDCE:   mov var_438, eax
  loc_005ABDD4: cmp var_438, 00000000h
  loc_005ABDDB: jge 005ABE00h
  loc_005ABDDD: push 00000050h
  loc_005ABDDF: push 00433F80h
loc_005ABDE4:   mov edx, var_434
loc_005ABDEA:   push edx
loc_005ABDEB:   mov eax, var_438
loc_005ABDF1:   push eax
  loc_005ABDF2: call [00401080h] ; __vbaHresultCheckObj
loc_005ABDF8:   mov var_51C, eax
  loc_005ABDFE: jmp 005ABE0Ah
  loc_005ABE00: mov var_51C, 00000000h
loc_005ABE0A:   mov ecx, var_38
loc_005ABE0D:   push ecx
loc_005ABE0E:   mov edx, var_43C
loc_005ABE14:   mov eax, [edx]
loc_005ABE16:   mov ecx, var_43C
loc_005ABE1C:   push ecx
loc_005ABE1D:   Call [eax+000000A4h]
loc_005ABE23:   fnclex
loc_005ABE25:   mov var_440, eax
  loc_005ABE2B: cmp var_440, 00000000h
  loc_005ABE32: jge 005ABE5Ah
  loc_005ABE34: push 000000A4h
  loc_005ABE39: push 0042DFCCh
loc_005ABE3E:   mov edx, var_43C
loc_005ABE44:   push edx
loc_005ABE45:   mov eax, var_440
loc_005ABE4B:   push eax
  loc_005ABE4C: call [00401080h] ; __vbaHresultCheckObj
loc_005ABE52:   mov var_520, eax
  loc_005ABE58: jmp 005ABE64h
  loc_005ABE5A: mov var_520, 00000000h
loc_005ABE64:   lea ecx, var_38
  loc_005ABE67: call [0040129Ch] ; __vbaFreeStr
loc_005ABE6D:   lea ecx, var_68
loc_005ABE70:   push ecx
loc_005ABE71:   lea edx, var_64
loc_005ABE74:   push edx
  loc_005ABE75: push 00000002h
  loc_005ABE77: call [0040104Ch] ; __vbaFreeObjList
  loc_005ABE7D: add esp, 0000000Ch
  loc_005ABE80: mov var_4, 00000022h
loc_005ABE87:   mov eax, var_44C
loc_005ABE8D:   mov ecx, [eax]
loc_005ABE8F:   mov edx, var_44C
loc_005ABE95:   push edx
loc_005ABE96:   Call [ecx+00000300h]
loc_005ABE9C:   push eax
loc_005ABE9D:   lea eax, var_64
loc_005ABEA0:   push eax
  loc_005ABEA1: call [004010B8h] ; __vbaObjSet
loc_005ABEA7:   mov var_434, eax
loc_005ABEAD:   mov ecx, var_434
loc_005ABEB3:   mov edx, [ecx]
loc_005ABEB5:   mov eax, var_434
loc_005ABEBB:   push eax
loc_005ABEBC:   Call [edx+00000204h]
loc_005ABEC2:   fnclex
loc_005ABEC4:   mov var_438, eax
  loc_005ABECA: cmp var_438, 00000000h
  loc_005ABED1: jge 005ABEF9h
  loc_005ABED3: push 00000204h
  loc_005ABED8: push 0042DFCCh
loc_005ABEDD:   mov ecx, var_434
loc_005ABEE3:   push ecx
loc_005ABEE4:   mov edx, var_438
loc_005ABEEA:   push edx
  loc_005ABEEB: call [00401080h] ; __vbaHresultCheckObj
loc_005ABEF1:   mov var_524, eax
  loc_005ABEF7: jmp 005ABF03h
  loc_005ABEF9: mov var_524, 00000000h
loc_005ABF03:   lea ecx, var_64
  loc_005ABF06: call [004012A0h] ; __vbaFreeObj
  loc_005ABF0C: mov var_4, 00000023h
  loc_005ABF13: mov var_260, 80020004h
  loc_005ABF1D: mov var_268, 0000000Ah
  loc_005ABF27: mov var_250, 00000001h
  loc_005ABF31: mov var_258, 00000002h
  loc_005ABF3B: mov eax, 00000010h
  loc_005ABF40: call 00405690h ; __vbaChkstk
loc_005ABF45:   mov eax, esp
loc_005ABF47:   mov ecx, var_268
loc_005ABF4D:   mov [eax], ecx
loc_005ABF4F:   mov edx, var_264
loc_005ABF55:   mov [eax+00000004h], edx
loc_005ABF58:   mov ecx, var_260
loc_005ABF5E:   mov [eax+00000008h], ecx
loc_005ABF61:   mov edx, var_25C
loc_005ABF67:   mov [eax+0000000Ch], edx
  loc_005ABF6A: mov eax, 00000010h
  loc_005ABF6F: call 00405690h ; __vbaChkstk
loc_005ABF74:   mov eax, esp
loc_005ABF76:   mov ecx, var_258
loc_005ABF7C:   mov [eax], ecx
loc_005ABF7E:   mov edx, var_254
loc_005ABF84:   mov [eax+00000004h], edx
loc_005ABF87:   mov ecx, var_250
loc_005ABF8D:   mov [eax+00000008h], ecx
loc_005ABF90:   mov edx, var_24C
loc_005ABF96:   mov [eax+0000000Ch], edx
loc_005ABF99:   mov eax, var_44C
loc_005ABF9F:   mov ecx, [eax]
loc_005ABFA1:   mov edx, var_44C
loc_005ABFA7:   push edx
loc_005ABFA8:   Call [ecx+000002B0h]
loc_005ABFAE:   fnclex
loc_005ABFB0:   mov var_434, eax
  loc_005ABFB6: cmp var_434, 00000000h
  loc_005ABFBD: jge 005ABFE5h
  loc_005ABFBF: push 000002B0h
  loc_005ABFC4: push 00435F74h
loc_005ABFC9:   mov eax, var_44C
loc_005ABFCF:   push eax
loc_005ABFD0:   mov ecx, var_434
loc_005ABFD6:   push ecx
  loc_005ABFD7: call [00401080h] ; __vbaHresultCheckObj
loc_005ABFDD:   mov var_528, eax
  loc_005ABFE3: jmp 005ABFEFh
  loc_005ABFE5: mov var_528, 00000000h
  loc_005ABFEF: mov var_4, 00000024h
  loc_005ABFF6: push 00000000h
loc_005ABFF8:   lea edx, var_44C
loc_005ABFFE:   push edx
  loc_005ABFFF: call [004010C8h] ; __vbaObjSetAddref
  loc_005AC005: jmp 005AE853h
  loc_005AC00A: mov var_4, 00000025h
  loc_005AC011: push 00000000h
loc_005AC013:   push FFFFFDFBh
loc_005AC018:   mov eax, Me
loc_005AC01B:   mov ecx, [eax]
loc_005AC01D:   mov edx, Me
loc_005AC020:   push edx
loc_005AC021:   Call [ecx+00000388h]
loc_005AC027:   push eax
loc_005AC028:   lea eax, var_64
loc_005AC02B:   push eax
  loc_005AC02C: call [004010B8h] ; __vbaObjSet
loc_005AC032:   push eax
loc_005AC033:   lea ecx, var_88
loc_005AC039:   push ecx
  loc_005AC03A: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AC040: add esp, 00000010h
loc_005AC043:   push eax
  loc_005AC044: call [00401034h] ; __vbaStrVarMove
loc_005AC04A:   mov edx, eax
loc_005AC04C:   lea ecx, var_38
  loc_005AC04F: call [0040126Ch] ; __vbaStrMove
loc_005AC055:   push eax
  loc_005AC056: push 00433F94h ; "Kredit"
  loc_005AC05B: call [0040112Ch] ; __vbaStrCmp
loc_005AC061:   neg eax
loc_005AC063:   sbb eax, eax
loc_005AC065:   inc eax
loc_005AC066:   neg eax
loc_005AC068:   mov var_434, ax
loc_005AC06F:   lea ecx, var_38
  loc_005AC072: call [0040129Ch] ; __vbaFreeStr
loc_005AC078:   lea ecx, var_64
  loc_005AC07B: call [004012A0h] ; __vbaFreeObj
loc_005AC081:   lea ecx, var_88
  loc_005AC087: call [00401020h] ; __vbaFreeVar
loc_005AC08D:   movsx edx, var_434
loc_005AC094:   test edx, edx
  loc_005AC096: jz 005AE853h
  loc_005AC09C: mov var_4, 00000026h
  loc_005AC0A3: mov edx, 00433F94h ; "Kredit"
  loc_005AC0A8: mov ecx, 006680D4h
  loc_005AC0AD: call [004011FCh] ; __vbaStrCopy
  loc_005AC0B3: mov var_4, 00000027h
loc_005AC0BA:   lea eax, var_88
loc_005AC0C0:   push eax
  loc_005AC0C1: call [00401220h] ; rtcGetDateVar
loc_005AC0C7:   mov ecx, Me
loc_005AC0CA:   mov edx, [ecx]
loc_005AC0CC:   mov eax, Me
loc_005AC0CF:   push eax
loc_005AC0D0:   Call [edx+00000338h]
loc_005AC0D6:   push eax
loc_005AC0D7:   lea ecx, var_74
loc_005AC0DA:   push ecx
  loc_005AC0DB: call [004010B8h] ; __vbaObjSet
loc_005AC0E1:   mov var_434, eax
loc_005AC0E7:   lea edx, var_4C
loc_005AC0EA:   push edx
loc_005AC0EB:   mov eax, var_434
loc_005AC0F1:   mov ecx, [eax]
loc_005AC0F3:   mov edx, var_434
loc_005AC0F9:   push edx
loc_005AC0FA:   Call [ecx+00000050h]
loc_005AC0FD:   fnclex
loc_005AC0FF:   mov var_438, eax
  loc_005AC105: cmp var_438, 00000000h
  loc_005AC10C: jge 005AC131h
  loc_005AC10E: push 00000050h
  loc_005AC110: push 00433F80h
loc_005AC115:   mov eax, var_434
loc_005AC11B:   push eax
loc_005AC11C:   mov ecx, var_438
loc_005AC122:   push ecx
  loc_005AC123: call [00401080h] ; __vbaHresultCheckObj
loc_005AC129:   mov var_52C, eax
  loc_005AC12F: jmp 005AC13Bh
  loc_005AC131: mov var_52C, 00000000h
  loc_005AC13B: push 00435CD0h ; "INSERT INTO Penjualan"
  loc_005AC140: push 00435D00h ; "(no_nota,tgl_nota,kode_pelanggan,nama_pelanggan,jenis,userid,catatan,tot_jual,Kode_Perusahaan)"
  loc_005AC145: call [00401070h] ; __vbaStrCat
loc_005AC14B:   mov edx, eax
loc_005AC14D:   lea ecx, var_38
  loc_005AC150: call [0040126Ch] ; __vbaStrMove
loc_005AC156:   push eax
  loc_005AC157: push 00434520h ; "VALUES ('"
  loc_005AC15C: call [00401070h] ; __vbaStrCat
loc_005AC162:   mov edx, eax
loc_005AC164:   lea ecx, var_3C
  loc_005AC167: call [0040126Ch] ; __vbaStrMove
loc_005AC16D:   push eax
loc_005AC16E:   mov edx, Me
loc_005AC171:   mov eax, [edx+0000006Ch]
loc_005AC174:   push eax
  loc_005AC175: call [00401070h] ; __vbaStrCat
loc_005AC17B:   mov edx, eax
loc_005AC17D:   lea ecx, var_40
  loc_005AC180: call [0040126Ch] ; __vbaStrMove
loc_005AC186:   push eax
  loc_005AC187: push 0042EC30h ; "','"
  loc_005AC18C: call [00401070h] ; __vbaStrCat
loc_005AC192:   mov var_B0, eax
  loc_005AC198: mov var_B8, 00000008h
  loc_005AC1A2: mov var_250, 00434538h ; "yyyy/MM/dd"
  loc_005AC1AC: mov var_258, 00000008h
loc_005AC1B6:   lea edx, var_258
loc_005AC1BC:   lea ecx, var_98
  loc_005AC1C2: call [00401238h] ; __vbaVarDup
  loc_005AC1C8: push 00000001h
  loc_005AC1CA: push 00000001h
loc_005AC1CC:   lea ecx, var_98
loc_005AC1D2:   push ecx
loc_005AC1D3:   lea edx, var_88
loc_005AC1D9:   push edx
loc_005AC1DA:   lea eax, var_A8
loc_005AC1E0:   push eax
  loc_005AC1E1: call [00401078h] ; rtcVarFromFormatVar
  loc_005AC1E7: mov var_260, 0042EC30h ; "','"
  loc_005AC1F1: mov var_268, 00000008h
  loc_005AC1FB: push 00000000h
loc_005AC1FD:   push FFFFFDFBh
loc_005AC202:   mov ecx, Me
loc_005AC205:   mov edx, [ecx]
loc_005AC207:   mov eax, Me
loc_005AC20A:   push eax
loc_005AC20B:   Call [edx+0000038Ch]
loc_005AC211:   push eax
loc_005AC212:   lea ecx, var_64
loc_005AC215:   push ecx
  loc_005AC216: call [004010B8h] ; __vbaObjSet
loc_005AC21C:   push eax
loc_005AC21D:   lea edx, var_E8
loc_005AC223:   push edx
  loc_005AC224: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AC22A: add esp, 00000010h
loc_005AC22D:   push eax
  loc_005AC22E: call [00401034h] ; __vbaStrVarMove
loc_005AC234:   mov var_F0, eax
  loc_005AC23A: mov var_F8, 00000008h
  loc_005AC244: mov var_270, 0042EC30h ; "','"
  loc_005AC24E: mov var_278, 00000008h
loc_005AC258:   mov eax, Me
loc_005AC25B:   mov ecx, [eax]
loc_005AC25D:   mov edx, Me
loc_005AC260:   push edx
loc_005AC261:   Call [ecx+00000310h]
loc_005AC267:   push eax
loc_005AC268:   lea eax, var_68
loc_005AC26B:   push eax
  loc_005AC26C: call [004010B8h] ; __vbaObjSet
loc_005AC272:   mov var_43C, eax
loc_005AC278:   lea ecx, var_44
loc_005AC27B:   push ecx
loc_005AC27C:   mov edx, var_43C
loc_005AC282:   mov eax, [edx]
loc_005AC284:   mov ecx, var_43C
loc_005AC28A:   push ecx
loc_005AC28B:   Call [eax+000000A0h]
loc_005AC291:   fnclex
loc_005AC293:   mov var_440, eax
  loc_005AC299: cmp var_440, 00000000h
  loc_005AC2A0: jge 005AC2C8h
  loc_005AC2A2: push 000000A0h
  loc_005AC2A7: push 0042DFCCh
loc_005AC2AC:   mov edx, var_43C
loc_005AC2B2:   push edx
loc_005AC2B3:   mov eax, var_440
loc_005AC2B9:   push eax
  loc_005AC2BA: call [00401080h] ; __vbaHresultCheckObj
loc_005AC2C0:   mov var_530, eax
  loc_005AC2C6: jmp 005AC2D2h
  loc_005AC2C8: mov var_530, 00000000h
loc_005AC2D2:   mov ecx, var_44
loc_005AC2D5:   mov var_494, ecx
  loc_005AC2DB: mov var_44, 00000000h
loc_005AC2E2:   mov edx, var_494
loc_005AC2E8:   mov var_120, edx
  loc_005AC2EE: mov var_128, 00000008h
  loc_005AC2F8: mov var_280, 0042EC30h ; "','"
  loc_005AC302: mov var_288, 00000008h
  loc_005AC30C: push 00000000h
loc_005AC30E:   push FFFFFDFBh
loc_005AC313:   mov eax, Me
loc_005AC316:   mov ecx, [eax]
loc_005AC318:   mov edx, Me
loc_005AC31B:   push edx
loc_005AC31C:   Call [ecx+00000388h]
loc_005AC322:   push eax
loc_005AC323:   lea eax, var_6C
loc_005AC326:   push eax
  loc_005AC327: call [004010B8h] ; __vbaObjSet
loc_005AC32D:   push eax
loc_005AC32E:   lea ecx, var_158
loc_005AC334:   push ecx
  loc_005AC335: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AC33B: add esp, 00000010h
loc_005AC33E:   push eax
  loc_005AC33F: call [00401034h] ; __vbaStrVarMove
loc_005AC345:   mov var_160, eax
  loc_005AC34B: mov var_168, 00000008h
  loc_005AC355: mov var_290, 0042EC30h ; "','"
  loc_005AC35F: mov var_298, 00000008h
  loc_005AC369: mov var_2A0, 0042EC30h ; "','"
  loc_005AC373: mov var_2A8, 00000008h
loc_005AC37D:   mov edx, Me
loc_005AC380:   mov eax, [edx]
loc_005AC382:   mov ecx, Me
loc_005AC385:   push ecx
loc_005AC386:   Call [eax+00000314h]
loc_005AC38C:   push eax
loc_005AC38D:   lea edx, var_70
loc_005AC390:   push edx
  loc_005AC391: call [004010B8h] ; __vbaObjSet
loc_005AC397:   mov var_444, eax
loc_005AC39D:   lea eax, var_48
loc_005AC3A0:   push eax
loc_005AC3A1:   mov ecx, var_444
loc_005AC3A7:   mov edx, [ecx]
loc_005AC3A9:   mov eax, var_444
loc_005AC3AF:   push eax
loc_005AC3B0:   Call [edx+000000A0h]
loc_005AC3B6:   fnclex
loc_005AC3B8:   mov var_448, eax
  loc_005AC3BE: cmp var_448, 00000000h
  loc_005AC3C5: jge 005AC3EDh
  loc_005AC3C7: push 000000A0h
  loc_005AC3CC: push 0042DFCCh
loc_005AC3D1:   mov ecx, var_444
loc_005AC3D7:   push ecx
loc_005AC3D8:   mov edx, var_448
loc_005AC3DE:   push edx
  loc_005AC3DF: call [00401080h] ; __vbaHresultCheckObj
loc_005AC3E5:   mov var_534, eax
  loc_005AC3EB: jmp 005AC3F7h
  loc_005AC3ED: mov var_534, 00000000h
loc_005AC3F7:   mov eax, var_48
loc_005AC3FA:   mov var_498, eax
  loc_005AC400: mov var_48, 00000000h
loc_005AC407:   mov ecx, var_498
loc_005AC40D:   mov var_1B0, ecx
  loc_005AC413: mov var_1B8, 00000008h
  loc_005AC41D: mov var_2B0, 0042EC30h ; "','"
  loc_005AC427: mov var_2B8, 00000008h
  loc_005AC431: mov var_2C0, 00435A20h ; "00"
  loc_005AC43B: mov var_2C8, 00000008h
loc_005AC445:   lea edx, var_2C8
loc_005AC44B:   lea ecx, var_1F8
  loc_005AC451: call [00401238h] ; __vbaVarDup
loc_005AC457:   mov edx, var_4C
loc_005AC45A:   mov var_49C, edx
  loc_005AC460: mov var_4C, 00000000h
loc_005AC467:   mov eax, var_49C
loc_005AC46D:   mov var_1E0, eax
  loc_005AC473: mov var_1E8, 00000008h
  loc_005AC47D: push 00000001h
  loc_005AC47F: push 00000001h
loc_005AC481:   lea ecx, var_1F8
loc_005AC487:   push ecx
loc_005AC488:   lea edx, var_1E8
loc_005AC48E:   push edx
loc_005AC48F:   lea eax, var_208
loc_005AC495:   push eax
  loc_005AC496: call [00401078h] ; rtcVarFromFormatVar
  loc_005AC49C: mov var_2D0, 0042EC30h ; "','"
  loc_005AC4A6: mov var_2D8, 00000008h
  loc_005AC4B0: mov var_2E0, 00435DC4h ; "P1"
  loc_005AC4BA: mov var_2E8, 00000008h
  loc_005AC4C4: mov var_2F0, 0042EC3Ch ; "')"
  loc_005AC4CE: mov var_2F8, 00000008h
loc_005AC4D8:   lea ecx, var_B8
loc_005AC4DE:   push ecx
loc_005AC4DF:   lea edx, var_A8
loc_005AC4E5:   push edx
loc_005AC4E6:   lea eax, var_C8
loc_005AC4EC:   push eax
  loc_005AC4ED: call [004011C4h] ; __vbaVarCat
loc_005AC4F3:   push eax
loc_005AC4F4:   lea ecx, var_268
loc_005AC4FA:   push ecx
loc_005AC4FB:   lea edx, var_D8
loc_005AC501:   push edx
  loc_005AC502: call [004011C4h] ; __vbaVarCat
loc_005AC508:   push eax
loc_005AC509:   lea eax, var_F8
loc_005AC50F:   push eax
loc_005AC510:   lea ecx, var_108
loc_005AC516:   push ecx
  loc_005AC517: call [004011C4h] ; __vbaVarCat
loc_005AC51D:   push eax
loc_005AC51E:   lea edx, var_278
loc_005AC524:   push edx
loc_005AC525:   lea eax, var_118
loc_005AC52B:   push eax
  loc_005AC52C: call [004011C4h] ; __vbaVarCat
loc_005AC532:   push eax
loc_005AC533:   lea ecx, var_128
loc_005AC539:   push ecx
loc_005AC53A:   lea edx, var_138
loc_005AC540:   push edx
  loc_005AC541: call [004011C4h] ; __vbaVarCat
loc_005AC547:   push eax
loc_005AC548:   lea eax, var_288
loc_005AC54E:   push eax
loc_005AC54F:   lea ecx, var_148
loc_005AC555:   push ecx
  loc_005AC556: call [004011C4h] ; __vbaVarCat
loc_005AC55C:   push eax
loc_005AC55D:   lea edx, var_168
loc_005AC563:   push edx
loc_005AC564:   lea eax, var_178
loc_005AC56A:   push eax
  loc_005AC56B: call [004011C4h] ; __vbaVarCat
loc_005AC571:   push eax
loc_005AC572:   lea ecx, var_298
loc_005AC578:   push ecx
loc_005AC579:   lea edx, var_188
loc_005AC57F:   push edx
  loc_005AC580: call [004011C4h] ; __vbaVarCat
loc_005AC586:   push eax
  loc_005AC587: push 006680B4h
loc_005AC58C:   lea eax, var_198
loc_005AC592:   push eax
  loc_005AC593: call [004011C4h] ; __vbaVarCat
loc_005AC599:   push eax
loc_005AC59A:   lea ecx, var_2A8
loc_005AC5A0:   push ecx
loc_005AC5A1:   lea edx, var_1A8
loc_005AC5A7:   push edx
  loc_005AC5A8: call [004011C4h] ; __vbaVarCat
loc_005AC5AE:   push eax
loc_005AC5AF:   lea eax, var_1B8
loc_005AC5B5:   push eax
loc_005AC5B6:   lea ecx, var_1C8
loc_005AC5BC:   push ecx
  loc_005AC5BD: call [004011C4h] ; __vbaVarCat
loc_005AC5C3:   push eax
loc_005AC5C4:   lea edx, var_2B8
loc_005AC5CA:   push edx
loc_005AC5CB:   lea eax, var_1D8
loc_005AC5D1:   push eax
  loc_005AC5D2: call [004011C4h] ; __vbaVarCat
loc_005AC5D8:   push eax
loc_005AC5D9:   lea ecx, var_208
loc_005AC5DF:   push ecx
loc_005AC5E0:   lea edx, var_218
loc_005AC5E6:   push edx
  loc_005AC5E7: call [004011C4h] ; __vbaVarCat
loc_005AC5ED:   push eax
loc_005AC5EE:   lea eax, var_2D8
loc_005AC5F4:   push eax
loc_005AC5F5:   lea ecx, var_228
loc_005AC5FB:   push ecx
  loc_005AC5FC: call [004011C4h] ; __vbaVarCat
loc_005AC602:   push eax
loc_005AC603:   lea edx, var_2E8
loc_005AC609:   push edx
loc_005AC60A:   lea eax, var_238
loc_005AC610:   push eax
  loc_005AC611: call [004011C4h] ; __vbaVarCat
loc_005AC617:   push eax
loc_005AC618:   lea ecx, var_2F8
loc_005AC61E:   push ecx
loc_005AC61F:   lea edx, var_248
loc_005AC625:   push edx
  loc_005AC626: call [004011C4h] ; __vbaVarCat
loc_005AC62C:   push eax
  loc_005AC62D: call [00401034h] ; __vbaStrVarMove
loc_005AC633:   mov edx, eax
  loc_005AC635: mov ecx, 0066809Ch
  loc_005AC63A: call [0040126Ch] ; __vbaStrMove
loc_005AC640:   lea eax, var_40
loc_005AC643:   push eax
loc_005AC644:   lea ecx, var_3C
loc_005AC647:   push ecx
loc_005AC648:   lea edx, var_38
loc_005AC64B:   push edx
  loc_005AC64C: push 00000003h
  loc_005AC64E: call [00401204h] ; __vbaFreeStrList
  loc_005AC654: add esp, 00000010h
loc_005AC657:   lea eax, var_74
loc_005AC65A:   push eax
loc_005AC65B:   lea ecx, var_70
loc_005AC65E:   push ecx
loc_005AC65F:   lea edx, var_6C
loc_005AC662:   push edx
loc_005AC663:   lea eax, var_68
loc_005AC666:   push eax
loc_005AC667:   lea ecx, var_64
loc_005AC66A:   push ecx
  loc_005AC66B: push 00000005h
  loc_005AC66D: call [0040104Ch] ; __vbaFreeObjList
  loc_005AC673: add esp, 00000018h
loc_005AC676:   lea edx, var_248
loc_005AC67C:   push edx
loc_005AC67D:   lea eax, var_238
loc_005AC683:   push eax
loc_005AC684:   lea ecx, var_228
loc_005AC68A:   push ecx
loc_005AC68B:   lea edx, var_218
loc_005AC691:   push edx
loc_005AC692:   lea eax, var_208
loc_005AC698:   push eax
loc_005AC699:   lea ecx, var_1D8
loc_005AC69F:   push ecx
loc_005AC6A0:   lea edx, var_1F8
loc_005AC6A6:   push edx
loc_005AC6A7:   lea eax, var_1E8
loc_005AC6AD:   push eax
loc_005AC6AE:   lea ecx, var_1C8
loc_005AC6B4:   push ecx
loc_005AC6B5:   lea edx, var_1B8
loc_005AC6BB:   push edx
loc_005AC6BC:   lea eax, var_1A8
loc_005AC6C2:   push eax
loc_005AC6C3:   lea ecx, var_198
loc_005AC6C9:   push ecx
loc_005AC6CA:   lea edx, var_188
loc_005AC6D0:   push edx
loc_005AC6D1:   lea eax, var_178
loc_005AC6D7:   push eax
loc_005AC6D8:   lea ecx, var_168
loc_005AC6DE:   push ecx
loc_005AC6DF:   lea edx, var_148
loc_005AC6E5:   push edx
loc_005AC6E6:   lea eax, var_158
loc_005AC6EC:   push eax
loc_005AC6ED:   lea ecx, var_138
loc_005AC6F3:   push ecx
loc_005AC6F4:   lea edx, var_128
loc_005AC6FA:   push edx
loc_005AC6FB:   lea eax, var_118
loc_005AC701:   push eax
loc_005AC702:   lea ecx, var_108
loc_005AC708:   push ecx
loc_005AC709:   lea edx, var_F8
loc_005AC70F:   push edx
loc_005AC710:   lea eax, var_D8
loc_005AC716:   push eax
loc_005AC717:   lea ecx, var_E8
loc_005AC71D:   push ecx
loc_005AC71E:   lea edx, var_C8
loc_005AC724:   push edx
loc_005AC725:   lea eax, var_A8
loc_005AC72B:   push eax
loc_005AC72C:   lea ecx, var_B8
loc_005AC732:   push ecx
loc_005AC733:   lea edx, var_98
loc_005AC739:   push edx
loc_005AC73A:   lea eax, var_88
loc_005AC740:   push eax
  loc_005AC741: push 0000001Dh
  loc_005AC743: call [0040103Ch] ; __vbaFreeVarList
  loc_005AC749: add esp, 00000078h
  loc_005AC74C: mov var_4, 00000028h
  loc_005AC753: cmp [0066803Ch], 00000000h
  loc_005AC75A: jnz 005AC778h
  loc_005AC75C: push 0066803Ch
  loc_005AC761: push 0042DEFCh
  loc_005AC766: call [004011E8h] ; __vbaNew2
  loc_005AC76C: mov var_538, 0066803Ch
  loc_005AC776: jmp 005AC782h
  loc_005AC778: mov var_538, 0066803Ch
loc_005AC782:   mov ecx, var_538
loc_005AC788:   mov edx, [ecx]
loc_005AC78A:   mov var_434, edx
  loc_005AC790: mov var_80, 80020004h
  loc_005AC797: mov var_88, 0000000Ah
loc_005AC7A1:   lea ecx, var_88
  loc_005AC7A7: call [0040123Ch] ; __vbaFreeVarg
loc_005AC7AD:   lea eax, var_64
loc_005AC7B0:   push eax
  loc_005AC7B1: push 00000001h
loc_005AC7B3:   lea ecx, var_88
loc_005AC7B9:   push ecx
loc_005AC7BA:   mov edx, [0066809Ch]
loc_005AC7C0:   push edx
loc_005AC7C1:   mov eax, var_434
loc_005AC7C7:   mov ecx, [eax]
loc_005AC7C9:   mov edx, var_434
loc_005AC7CF:   push edx
loc_005AC7D0:   Call [ecx+00000040h]
loc_005AC7D3:   fnclex
loc_005AC7D5:   mov var_438, eax
  loc_005AC7DB: cmp var_438, 00000000h
  loc_005AC7E2: jge 005AC807h
  loc_005AC7E4: push 00000040h
  loc_005AC7E6: push 0042E1B0h
loc_005AC7EB:   mov eax, var_434
loc_005AC7F1:   push eax
loc_005AC7F2:   mov ecx, var_438
loc_005AC7F8:   push ecx
  loc_005AC7F9: call [00401080h] ; __vbaHresultCheckObj
loc_005AC7FF:   mov var_53C, eax
  loc_005AC805: jmp 005AC811h
  loc_005AC807: mov var_53C, 00000000h
loc_005AC811:   lea ecx, var_64
  loc_005AC814: call [004012A0h] ; __vbaFreeObj
loc_005AC81A:   lea ecx, var_88
  loc_005AC820: call [00401020h] ; __vbaFreeVar
  loc_005AC826: mov var_4, 00000029h
loc_005AC82D:   lea edx, var_88
loc_005AC833:   push edx
  loc_005AC834: call [00401220h] ; rtcGetDateVar
loc_005AC83A:   mov eax, Me
loc_005AC83D:   mov ecx, [eax]
loc_005AC83F:   mov edx, Me
loc_005AC842:   push edx
loc_005AC843:   Call [ecx+00000338h]
loc_005AC849:   push eax
loc_005AC84A:   lea eax, var_64
loc_005AC84D:   push eax
  loc_005AC84E: call [004010B8h] ; __vbaObjSet
loc_005AC854:   mov var_434, eax
loc_005AC85A:   lea ecx, var_44
loc_005AC85D:   push ecx
loc_005AC85E:   mov edx, var_434
loc_005AC864:   mov eax, [edx]
loc_005AC866:   mov ecx, var_434
loc_005AC86C:   push ecx
loc_005AC86D:   Call [eax+00000050h]
loc_005AC870:   fnclex
loc_005AC872:   mov var_438, eax
  loc_005AC878: cmp var_438, 00000000h
  loc_005AC87F: jge 005AC8A4h
  loc_005AC881: push 00000050h
  loc_005AC883: push 00433F80h
loc_005AC888:   mov edx, var_434
loc_005AC88E:   push edx
loc_005AC88F:   mov eax, var_438
loc_005AC895:   push eax
  loc_005AC896: call [00401080h] ; __vbaHresultCheckObj
loc_005AC89C:   mov var_540, eax
  loc_005AC8A2: jmp 005AC8AEh
  loc_005AC8A4: mov var_540, 00000000h
loc_005AC8AE:   mov ecx, Me
loc_005AC8B1:   mov edx, [ecx]
loc_005AC8B3:   mov eax, Me
loc_005AC8B6:   push eax
loc_005AC8B7:   Call [edx+00000300h]
loc_005AC8BD:   push eax
loc_005AC8BE:   lea ecx, var_68
loc_005AC8C1:   push ecx
  loc_005AC8C2: call [004010B8h] ; __vbaObjSet
loc_005AC8C8:   mov var_43C, eax
loc_005AC8CE:   lea edx, var_48
loc_005AC8D1:   push edx
loc_005AC8D2:   mov eax, var_43C
loc_005AC8D8:   mov ecx, [eax]
loc_005AC8DA:   mov edx, var_43C
loc_005AC8E0:   push edx
loc_005AC8E1:   Call [ecx+00000050h]
loc_005AC8E4:   fnclex
loc_005AC8E6:   mov var_440, eax
  loc_005AC8EC: cmp var_440, 00000000h
  loc_005AC8F3: jge 005AC918h
  loc_005AC8F5: push 00000050h
  loc_005AC8F7: push 00433F80h
loc_005AC8FC:   mov eax, var_43C
loc_005AC902:   push eax
loc_005AC903:   mov ecx, var_440
loc_005AC909:   push ecx
  loc_005AC90A: call [00401080h] ; __vbaHresultCheckObj
loc_005AC910:   mov var_544, eax
  loc_005AC916: jmp 005AC922h
  loc_005AC918: mov var_544, 00000000h
  loc_005AC922: push 00435DD0h ; "INSERT INTO Laba"
  loc_005AC927: push 00435E04h ; "(no_nota,tgl_nota,tot_jual,tot_beli,laba)"
  loc_005AC92C: call [00401070h] ; __vbaStrCat
loc_005AC932:   mov edx, eax
loc_005AC934:   lea ecx, var_38
  loc_005AC937: call [0040126Ch] ; __vbaStrMove
loc_005AC93D:   push eax
  loc_005AC93E: push 00434520h ; "VALUES ('"
  loc_005AC943: call [00401070h] ; __vbaStrCat
loc_005AC949:   mov edx, eax
loc_005AC94B:   lea ecx, var_3C
  loc_005AC94E: call [0040126Ch] ; __vbaStrMove
loc_005AC954:   push eax
loc_005AC955:   mov edx, Me
loc_005AC958:   mov eax, [edx+0000006Ch]
loc_005AC95B:   push eax
  loc_005AC95C: call [00401070h] ; __vbaStrCat
loc_005AC962:   mov edx, eax
loc_005AC964:   lea ecx, var_40
  loc_005AC967: call [0040126Ch] ; __vbaStrMove
loc_005AC96D:   push eax
  loc_005AC96E: push 0042EC30h ; "','"
  loc_005AC973: call [00401070h] ; __vbaStrCat
loc_005AC979:   mov var_B0, eax
  loc_005AC97F: mov var_B8, 00000008h
  loc_005AC989: mov var_250, 00434538h ; "yyyy/MM/dd"
  loc_005AC993: mov var_258, 00000008h
loc_005AC99D:   lea edx, var_258
loc_005AC9A3:   lea ecx, var_98
  loc_005AC9A9: call [00401238h] ; __vbaVarDup
  loc_005AC9AF: push 00000001h
  loc_005AC9B1: push 00000001h
loc_005AC9B3:   lea ecx, var_98
loc_005AC9B9:   push ecx
loc_005AC9BA:   lea edx, var_88
loc_005AC9C0:   push edx
loc_005AC9C1:   lea eax, var_A8
loc_005AC9C7:   push eax
  loc_005AC9C8: call [00401078h] ; rtcVarFromFormatVar
  loc_005AC9CE: mov var_260, 0042EC30h ; "','"
  loc_005AC9D8: mov var_268, 00000008h
  loc_005AC9E2: mov var_270, 00435A20h ; "00"
  loc_005AC9EC: mov var_278, 00000008h
loc_005AC9F6:   lea edx, var_278
loc_005AC9FC:   lea ecx, var_F8
  loc_005ACA02: call [00401238h] ; __vbaVarDup
loc_005ACA08:   mov ecx, var_44
loc_005ACA0B:   mov var_4A0, ecx
  loc_005ACA11: mov var_44, 00000000h
loc_005ACA18:   mov edx, var_4A0
loc_005ACA1E:   mov var_E0, edx
  loc_005ACA24: mov var_E8, 00000008h
  loc_005ACA2E: push 00000001h
  loc_005ACA30: push 00000001h
loc_005ACA32:   lea eax, var_F8
loc_005ACA38:   push eax
loc_005ACA39:   lea ecx, var_E8
loc_005ACA3F:   push ecx
loc_005ACA40:   lea edx, var_108
loc_005ACA46:   push edx
  loc_005ACA47: call [00401078h] ; rtcVarFromFormatVar
  loc_005ACA4D: mov var_280, 0042EC30h ; "','"
  loc_005ACA57: mov var_288, 00000008h
  loc_005ACA61: mov var_290, 00435A20h ; "00"
  loc_005ACA6B: mov var_298, 00000008h
loc_005ACA75:   lea edx, var_298
loc_005ACA7B:   lea ecx, var_148
  loc_005ACA81: call [00401238h] ; __vbaVarDup
loc_005ACA87:   mov eax, var_48
loc_005ACA8A:   mov var_4A4, eax
  loc_005ACA90: mov var_48, 00000000h
loc_005ACA97:   mov ecx, var_4A4
loc_005ACA9D:   mov var_130, ecx
  loc_005ACAA3: mov var_138, 00000008h
  loc_005ACAAD: push 00000001h
  loc_005ACAAF: push 00000001h
loc_005ACAB1:   lea edx, var_148
loc_005ACAB7:   push edx
loc_005ACAB8:   lea eax, var_138
loc_005ACABE:   push eax
loc_005ACABF:   lea ecx, var_158
loc_005ACAC5:   push ecx
  loc_005ACAC6: call [00401078h] ; rtcVarFromFormatVar
  loc_005ACACC: mov var_2A0, 0042EC30h ; "','"
  loc_005ACAD6: mov var_2A8, 00000008h
loc_005ACAE0:   mov edx, var_34
loc_005ACAE3:   mov var_2B0, edx
loc_005ACAE9:   mov eax, var_30
loc_005ACAEC:   mov var_2AC, eax
  loc_005ACAF2: mov var_2B8, 00000006h
  loc_005ACAFC: mov var_2C0, 0042EC3Ch ; "')"
  loc_005ACB06: mov var_2C8, 00000008h
loc_005ACB10:   lea ecx, var_B8
loc_005ACB16:   push ecx
loc_005ACB17:   lea edx, var_A8
loc_005ACB1D:   push edx
loc_005ACB1E:   lea eax, var_C8
loc_005ACB24:   push eax
  loc_005ACB25: call [004011C4h] ; __vbaVarCat
loc_005ACB2B:   push eax
loc_005ACB2C:   lea ecx, var_268
loc_005ACB32:   push ecx
loc_005ACB33:   lea edx, var_D8
loc_005ACB39:   push edx
  loc_005ACB3A: call [004011C4h] ; __vbaVarCat
loc_005ACB40:   push eax
loc_005ACB41:   lea eax, var_108
loc_005ACB47:   push eax
loc_005ACB48:   lea ecx, var_118
loc_005ACB4E:   push ecx
  loc_005ACB4F: call [004011C4h] ; __vbaVarCat
loc_005ACB55:   push eax
loc_005ACB56:   lea edx, var_288
loc_005ACB5C:   push edx
loc_005ACB5D:   lea eax, var_128
loc_005ACB63:   push eax
  loc_005ACB64: call [004011C4h] ; __vbaVarCat
loc_005ACB6A:   push eax
loc_005ACB6B:   lea ecx, var_158
loc_005ACB71:   push ecx
loc_005ACB72:   lea edx, var_168
loc_005ACB78:   push edx
  loc_005ACB79: call [004011C4h] ; __vbaVarCat
loc_005ACB7F:   push eax
loc_005ACB80:   lea eax, var_2A8
loc_005ACB86:   push eax
loc_005ACB87:   lea ecx, var_178
loc_005ACB8D:   push ecx
  loc_005ACB8E: call [004011C4h] ; __vbaVarCat
loc_005ACB94:   push eax
loc_005ACB95:   lea edx, var_2B8
loc_005ACB9B:   push edx
loc_005ACB9C:   lea eax, var_188
loc_005ACBA2:   push eax
  loc_005ACBA3: call [004011C4h] ; __vbaVarCat
loc_005ACBA9:   push eax
loc_005ACBAA:   lea ecx, var_2C8
loc_005ACBB0:   push ecx
loc_005ACBB1:   lea edx, var_198
loc_005ACBB7:   push edx
  loc_005ACBB8: call [004011C4h] ; __vbaVarCat
loc_005ACBBE:   push eax
  loc_005ACBBF: call [00401034h] ; __vbaStrVarMove
loc_005ACBC5:   mov edx, eax
  loc_005ACBC7: mov ecx, 0066809Ch
  loc_005ACBCC: call [0040126Ch] ; __vbaStrMove
loc_005ACBD2:   lea eax, var_40
loc_005ACBD5:   push eax
loc_005ACBD6:   lea ecx, var_3C
loc_005ACBD9:   push ecx
loc_005ACBDA:   lea edx, var_38
loc_005ACBDD:   push edx
  loc_005ACBDE: push 00000003h
  loc_005ACBE0: call [00401204h] ; __vbaFreeStrList
  loc_005ACBE6: add esp, 00000010h
loc_005ACBE9:   lea eax, var_68
loc_005ACBEC:   push eax
loc_005ACBED:   lea ecx, var_64
loc_005ACBF0:   push ecx
  loc_005ACBF1: push 00000002h
  loc_005ACBF3: call [0040104Ch] ; __vbaFreeObjList
  loc_005ACBF9: add esp, 0000000Ch
loc_005ACBFC:   lea edx, var_198
loc_005ACC02:   push edx
loc_005ACC03:   lea eax, var_188
loc_005ACC09:   push eax
loc_005ACC0A:   lea ecx, var_178
loc_005ACC10:   push ecx
loc_005ACC11:   lea edx, var_168
loc_005ACC17:   push edx
loc_005ACC18:   lea eax, var_158
loc_005ACC1E:   push eax
loc_005ACC1F:   lea ecx, var_128
loc_005ACC25:   push ecx
loc_005ACC26:   lea edx, var_148
loc_005ACC2C:   push edx
loc_005ACC2D:   lea eax, var_138
loc_005ACC33:   push eax
loc_005ACC34:   lea ecx, var_118
loc_005ACC3A:   push ecx
loc_005ACC3B:   lea edx, var_108
loc_005ACC41:   push edx
loc_005ACC42:   lea eax, var_D8
loc_005ACC48:   push eax
loc_005ACC49:   lea ecx, var_F8
loc_005ACC4F:   push ecx
loc_005ACC50:   lea edx, var_E8
loc_005ACC56:   push edx
loc_005ACC57:   lea eax, var_C8
loc_005ACC5D:   push eax
loc_005ACC5E:   lea ecx, var_A8
loc_005ACC64:   push ecx
loc_005ACC65:   lea edx, var_B8
loc_005ACC6B:   push edx
loc_005ACC6C:   lea eax, var_98
loc_005ACC72:   push eax
loc_005ACC73:   lea ecx, var_88
loc_005ACC79:   push ecx
  loc_005ACC7A: push 00000012h
  loc_005ACC7C: call [0040103Ch] ; __vbaFreeVarList
  loc_005ACC82: add esp, 0000004Ch
  loc_005ACC85: mov var_4, 0000002Ah
  loc_005ACC8C: cmp [0066803Ch], 00000000h
  loc_005ACC93: jnz 005ACCB1h
  loc_005ACC95: push 0066803Ch
  loc_005ACC9A: push 0042DEFCh
  loc_005ACC9F: call [004011E8h] ; __vbaNew2
  loc_005ACCA5: mov var_548, 0066803Ch
  loc_005ACCAF: jmp 005ACCBBh
  loc_005ACCB1: mov var_548, 0066803Ch
loc_005ACCBB:   mov edx, var_548
loc_005ACCC1:   mov eax, [edx]
loc_005ACCC3:   mov var_434, eax
  loc_005ACCC9: mov var_80, 80020004h
  loc_005ACCD0: mov var_88, 0000000Ah
loc_005ACCDA:   lea ecx, var_88
  loc_005ACCE0: call [0040123Ch] ; __vbaFreeVarg
loc_005ACCE6:   lea ecx, var_64
loc_005ACCE9:   push ecx
  loc_005ACCEA: push 00000001h
loc_005ACCEC:   lea edx, var_88
loc_005ACCF2:   push edx
loc_005ACCF3:   mov eax, [0066809Ch]
loc_005ACCF8:   push eax
loc_005ACCF9:   mov ecx, var_434
loc_005ACCFF:   mov edx, [ecx]
loc_005ACD01:   mov eax, var_434
loc_005ACD07:   push eax
loc_005ACD08:   Call [edx+00000040h]
loc_005ACD0B:   fnclex
loc_005ACD0D:   mov var_438, eax
  loc_005ACD13: cmp var_438, 00000000h
  loc_005ACD1A: jge 005ACD3Fh
  loc_005ACD1C: push 00000040h
  loc_005ACD1E: push 0042E1B0h
loc_005ACD23:   mov ecx, var_434
loc_005ACD29:   push ecx
loc_005ACD2A:   mov edx, var_438
loc_005ACD30:   push edx
  loc_005ACD31: call [00401080h] ; __vbaHresultCheckObj
loc_005ACD37:   mov var_54C, eax
  loc_005ACD3D: jmp 005ACD49h
  loc_005ACD3F: mov var_54C, 00000000h
loc_005ACD49:   lea ecx, var_64
  loc_005ACD4C: call [004012A0h] ; __vbaFreeObj
loc_005ACD52:   lea ecx, var_88
  loc_005ACD58: call [00401020h] ; __vbaFreeVar
  loc_005ACD5E: mov var_4, 0000002Bh
loc_005ACD65:   lea eax, var_98
loc_005ACD6B:   push eax
  loc_005ACD6C: call [00401220h] ; rtcGetDateVar
loc_005ACD72:   mov ecx, Me
loc_005ACD75:   mov edx, [ecx]
loc_005ACD77:   mov eax, Me
loc_005ACD7A:   push eax
loc_005ACD7B:   Call [edx+00000384h]
loc_005ACD81:   push eax
loc_005ACD82:   lea ecx, var_74
loc_005ACD85:   push ecx
  loc_005ACD86: call [004010B8h] ; __vbaObjSet
loc_005ACD8C:   mov edx, Me
loc_005ACD8F:   mov eax, [edx]
loc_005ACD91:   mov ecx, Me
loc_005ACD94:   push ecx
loc_005ACD95:   Call [eax+00000338h]
loc_005ACD9B:   push eax
loc_005ACD9C:   lea edx, var_6C
loc_005ACD9F:   push edx
  loc_005ACDA0: call [004010B8h] ; __vbaObjSet
loc_005ACDA6:   mov var_434, eax
loc_005ACDAC:   lea eax, var_5C
loc_005ACDAF:   push eax
loc_005ACDB0:   mov ecx, var_434
loc_005ACDB6:   mov edx, [ecx]
loc_005ACDB8:   mov eax, var_434
loc_005ACDBE:   push eax
loc_005ACDBF:   Call [edx+00000050h]
loc_005ACDC2:   fnclex
loc_005ACDC4:   mov var_438, eax
  loc_005ACDCA: cmp var_438, 00000000h
  loc_005ACDD1: jge 005ACDF6h
  loc_005ACDD3: push 00000050h
  loc_005ACDD5: push 00433F80h
loc_005ACDDA:   mov ecx, var_434
loc_005ACDE0:   push ecx
loc_005ACDE1:   mov edx, var_438
loc_005ACDE7:   push edx
  loc_005ACDE8: call [00401080h] ; __vbaHresultCheckObj
loc_005ACDEE:   mov var_550, eax
  loc_005ACDF4: jmp 005ACE00h
  loc_005ACDF6: mov var_550, 00000000h
loc_005ACE00:   mov eax, Me
loc_005ACE03:   mov ecx, [eax]
loc_005ACE05:   mov edx, Me
loc_005ACE08:   push edx
loc_005ACE09:   Call [ecx+00000338h]
loc_005ACE0F:   push eax
loc_005ACE10:   lea eax, var_70
loc_005ACE13:   push eax
  loc_005ACE14: call [004010B8h] ; __vbaObjSet
loc_005ACE1A:   mov var_43C, eax
loc_005ACE20:   lea ecx, var_60
loc_005ACE23:   push ecx
loc_005ACE24:   mov edx, var_43C
loc_005ACE2A:   mov eax, [edx]
loc_005ACE2C:   mov ecx, var_43C
loc_005ACE32:   push ecx
loc_005ACE33:   Call [eax+00000050h]
loc_005ACE36:   fnclex
loc_005ACE38:   mov var_440, eax
  loc_005ACE3E: cmp var_440, 00000000h
  loc_005ACE45: jge 005ACE6Ah
  loc_005ACE47: push 00000050h
  loc_005ACE49: push 00433F80h
loc_005ACE4E:   mov edx, var_43C
loc_005ACE54:   push edx
loc_005ACE55:   mov eax, var_440
loc_005ACE5B:   push eax
  loc_005ACE5C: call [00401080h] ; __vbaHresultCheckObj
loc_005ACE62:   mov var_554, eax
  loc_005ACE68: jmp 005ACE74h
  loc_005ACE6A: mov var_554, 00000000h
loc_005ACE74:   mov ecx, Me
loc_005ACE77:   mov edx, [ecx]
loc_005ACE79:   mov eax, Me
loc_005ACE7C:   push eax
loc_005ACE7D:   Call [edx+00000310h]
loc_005ACE83:   push eax
loc_005ACE84:   lea ecx, var_68
loc_005ACE87:   push ecx
  loc_005ACE88: call [004010B8h] ; __vbaObjSet
loc_005ACE8E:   mov var_444, eax
loc_005ACE94:   lea edx, var_50
loc_005ACE97:   push edx
loc_005ACE98:   mov eax, var_444
loc_005ACE9E:   mov ecx, [eax]
loc_005ACEA0:   mov edx, var_444
loc_005ACEA6:   push edx
loc_005ACEA7:   Call [ecx+000000A0h]
loc_005ACEAD:   fnclex
loc_005ACEAF:   mov var_448, eax
  loc_005ACEB5: cmp var_448, 00000000h
  loc_005ACEBC: jge 005ACEE4h
  loc_005ACEBE: push 000000A0h
  loc_005ACEC3: push 0042DFCCh
loc_005ACEC8:   mov eax, var_444
loc_005ACECE:   push eax
loc_005ACECF:   mov ecx, var_448
loc_005ACED5:   push ecx
  loc_005ACED6: call [00401080h] ; __vbaHresultCheckObj
loc_005ACEDC:   mov var_558, eax
  loc_005ACEE2: jmp 005ACEEEh
  loc_005ACEE4: mov var_558, 00000000h
  loc_005ACEEE: push 00436018h ; "INSERT INTO Piutang"
  loc_005ACEF3: push 00436044h ; "(no_nota,kode_pelanggan,nama_pelanggan,tgl_nota,tgl_jatuh_tempo,tot_jual,sisa_piutang,Kode_Perusahaan)"
  loc_005ACEF8: call [00401070h] ; __vbaStrCat
loc_005ACEFE:   mov edx, eax
loc_005ACF00:   lea ecx, var_38
  loc_005ACF03: call [0040126Ch] ; __vbaStrMove
loc_005ACF09:   push eax
  loc_005ACF0A: push 00434520h ; "VALUES ('"
  loc_005ACF0F: call [00401070h] ; __vbaStrCat
loc_005ACF15:   mov edx, eax
loc_005ACF17:   lea ecx, var_3C
  loc_005ACF1A: call [0040126Ch] ; __vbaStrMove
loc_005ACF20:   push eax
loc_005ACF21:   mov edx, Me
loc_005ACF24:   mov eax, [edx+0000006Ch]
loc_005ACF27:   push eax
  loc_005ACF28: call [00401070h] ; __vbaStrCat
loc_005ACF2E:   mov edx, eax
loc_005ACF30:   lea ecx, var_40
  loc_005ACF33: call [0040126Ch] ; __vbaStrMove
loc_005ACF39:   push eax
  loc_005ACF3A: push 0042EC30h ; "','"
  loc_005ACF3F: call [00401070h] ; __vbaStrCat
loc_005ACF45:   mov edx, eax
loc_005ACF47:   lea ecx, var_44
  loc_005ACF4A: call [0040126Ch] ; __vbaStrMove
loc_005ACF50:   push eax
  loc_005ACF51: push 00000000h
loc_005ACF53:   push FFFFFDFBh
loc_005ACF58:   mov ecx, Me
loc_005ACF5B:   mov edx, [ecx]
loc_005ACF5D:   mov eax, Me
loc_005ACF60:   push eax
loc_005ACF61:   Call [edx+0000038Ch]
loc_005ACF67:   push eax
loc_005ACF68:   lea ecx, var_64
loc_005ACF6B:   push ecx
  loc_005ACF6C: call [004010B8h] ; __vbaObjSet
loc_005ACF72:   push eax
loc_005ACF73:   lea edx, var_88
loc_005ACF79:   push edx
  loc_005ACF7A: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005ACF80: add esp, 00000010h
loc_005ACF83:   push eax
  loc_005ACF84: call [00401034h] ; __vbaStrVarMove
loc_005ACF8A:   mov edx, eax
loc_005ACF8C:   lea ecx, var_48
  loc_005ACF8F: call [0040126Ch] ; __vbaStrMove
loc_005ACF95:   push eax
  loc_005ACF96: call [00401070h] ; __vbaStrCat
loc_005ACF9C:   mov edx, eax
loc_005ACF9E:   lea ecx, var_4C
  loc_005ACFA1: call [0040126Ch] ; __vbaStrMove
loc_005ACFA7:   push eax
  loc_005ACFA8: push 0042EC30h ; "','"
  loc_005ACFAD: call [00401070h] ; __vbaStrCat
loc_005ACFB3:   mov edx, eax
loc_005ACFB5:   lea ecx, var_54
  loc_005ACFB8: call [0040126Ch] ; __vbaStrMove
loc_005ACFBE:   push eax
loc_005ACFBF:   mov eax, var_50
loc_005ACFC2:   push eax
  loc_005ACFC3: call [00401070h] ; __vbaStrCat
loc_005ACFC9:   mov edx, eax
loc_005ACFCB:   lea ecx, var_58
  loc_005ACFCE: call [0040126Ch] ; __vbaStrMove
loc_005ACFD4:   push eax
  loc_005ACFD5: push 0042EC30h ; "','"
  loc_005ACFDA: call [00401070h] ; __vbaStrCat
loc_005ACFE0:   mov var_C0, eax
  loc_005ACFE6: mov var_C8, 00000008h
  loc_005ACFF0: mov var_250, 00434538h ; "yyyy/MM/dd"
  loc_005ACFFA: mov var_258, 00000008h
loc_005AD004:   lea edx, var_258
loc_005AD00A:   lea ecx, var_A8
  loc_005AD010: call [00401238h] ; __vbaVarDup
  loc_005AD016: push 00000001h
  loc_005AD018: push 00000001h
loc_005AD01A:   lea ecx, var_A8
loc_005AD020:   push ecx
loc_005AD021:   lea edx, var_98
loc_005AD027:   push edx
loc_005AD028:   lea eax, var_B8
loc_005AD02E:   push eax
  loc_005AD02F: call [00401078h] ; rtcVarFromFormatVar
  loc_005AD035: mov var_260, 0042EC30h ; "','"
  loc_005AD03F: mov var_268, 00000008h
  loc_005AD049: mov var_270, 00434538h ; "yyyy/MM/dd"
  loc_005AD053: mov var_278, 00000008h
loc_005AD05D:   lea edx, var_278
loc_005AD063:   lea ecx, var_108
  loc_005AD069: call [00401238h] ; __vbaVarDup
loc_005AD06F:   mov ecx, var_74
loc_005AD072:   mov var_4A8, ecx
  loc_005AD078: mov var_74, 00000000h
loc_005AD07F:   mov edx, var_4A8
loc_005AD085:   mov var_F0, edx
  loc_005AD08B: mov var_F8, 00000009h
  loc_005AD095: push 00000001h
  loc_005AD097: push 00000001h
loc_005AD099:   lea eax, var_108
loc_005AD09F:   push eax
loc_005AD0A0:   lea ecx, var_F8
loc_005AD0A6:   push ecx
loc_005AD0A7:   lea edx, var_118
loc_005AD0AD:   push edx
  loc_005AD0AE: call [00401078h] ; rtcVarFromFormatVar
  loc_005AD0B4: mov var_280, 0042EC30h ; "','"
  loc_005AD0BE: mov var_288, 00000008h
  loc_005AD0C8: mov var_290, 00435A20h ; "00"
  loc_005AD0D2: mov var_298, 00000008h
loc_005AD0DC:   lea edx, var_298
loc_005AD0E2:   lea ecx, var_158
  loc_005AD0E8: call [00401238h] ; __vbaVarDup
loc_005AD0EE:   mov eax, var_5C
loc_005AD0F1:   mov var_4AC, eax
  loc_005AD0F7: mov var_5C, 00000000h
loc_005AD0FE:   mov ecx, var_4AC
loc_005AD104:   mov var_140, ecx
  loc_005AD10A: mov var_148, 00000008h
  loc_005AD114: push 00000001h
  loc_005AD116: push 00000001h
loc_005AD118:   lea edx, var_158
loc_005AD11E:   push edx
loc_005AD11F:   lea eax, var_148
loc_005AD125:   push eax
loc_005AD126:   lea ecx, var_168
loc_005AD12C:   push ecx
  loc_005AD12D: call [00401078h] ; rtcVarFromFormatVar
  loc_005AD133: mov var_2A0, 0042EC30h ; "','"
  loc_005AD13D: mov var_2A8, 00000008h
  loc_005AD147: mov var_2B0, 00435A20h ; "00"
  loc_005AD151: mov var_2B8, 00000008h
loc_005AD15B:   lea edx, var_2B8
loc_005AD161:   lea ecx, var_1A8
  loc_005AD167: call [00401238h] ; __vbaVarDup
loc_005AD16D:   mov edx, var_60
loc_005AD170:   mov var_4B0, edx
  loc_005AD176: mov var_60, 00000000h
loc_005AD17D:   mov eax, var_4B0
loc_005AD183:   mov var_190, eax
  loc_005AD189: mov var_198, 00000008h
  loc_005AD193: push 00000001h
  loc_005AD195: push 00000001h
loc_005AD197:   lea ecx, var_1A8
loc_005AD19D:   push ecx
loc_005AD19E:   lea edx, var_198
loc_005AD1A4:   push edx
loc_005AD1A5:   lea eax, var_1B8
loc_005AD1AB:   push eax
  loc_005AD1AC: call [00401078h] ; rtcVarFromFormatVar
  loc_005AD1B2: mov var_2C0, 0042EC30h ; "','"
  loc_005AD1BC: mov var_2C8, 00000008h
  loc_005AD1C6: mov var_2D0, 00435DC4h ; "P1"
  loc_005AD1D0: mov var_2D8, 00000008h
  loc_005AD1DA: mov var_2E0, 0042EC3Ch ; "')"
  loc_005AD1E4: mov var_2E8, 00000008h
loc_005AD1EE:   lea ecx, var_C8
loc_005AD1F4:   push ecx
loc_005AD1F5:   lea edx, var_B8
loc_005AD1FB:   push edx
loc_005AD1FC:   lea eax, var_D8
loc_005AD202:   push eax
  loc_005AD203: call [004011C4h] ; __vbaVarCat
loc_005AD209:   push eax
loc_005AD20A:   lea ecx, var_268
loc_005AD210:   push ecx
loc_005AD211:   lea edx, var_E8
loc_005AD217:   push edx
  loc_005AD218: call [004011C4h] ; __vbaVarCat
loc_005AD21E:   push eax
loc_005AD21F:   lea eax, var_118
loc_005AD225:   push eax
loc_005AD226:   lea ecx, var_128
loc_005AD22C:   push ecx
  loc_005AD22D: call [004011C4h] ; __vbaVarCat
loc_005AD233:   push eax
loc_005AD234:   lea edx, var_288
loc_005AD23A:   push edx
loc_005AD23B:   lea eax, var_138
loc_005AD241:   push eax
  loc_005AD242: call [004011C4h] ; __vbaVarCat
loc_005AD248:   push eax
loc_005AD249:   lea ecx, var_168
loc_005AD24F:   push ecx
loc_005AD250:   lea edx, var_178
loc_005AD256:   push edx
  loc_005AD257: call [004011C4h] ; __vbaVarCat
loc_005AD25D:   push eax
loc_005AD25E:   lea eax, var_2A8
loc_005AD264:   push eax
loc_005AD265:   lea ecx, var_188
loc_005AD26B:   push ecx
  loc_005AD26C: call [004011C4h] ; __vbaVarCat
loc_005AD272:   push eax
loc_005AD273:   lea edx, var_1B8
loc_005AD279:   push edx
loc_005AD27A:   lea eax, var_1C8
loc_005AD280:   push eax
  loc_005AD281: call [004011C4h] ; __vbaVarCat
loc_005AD287:   push eax
loc_005AD288:   lea ecx, var_2C8
loc_005AD28E:   push ecx
loc_005AD28F:   lea edx, var_1D8
loc_005AD295:   push edx
  loc_005AD296: call [004011C4h] ; __vbaVarCat
loc_005AD29C:   push eax
loc_005AD29D:   lea eax, var_2D8
loc_005AD2A3:   push eax
loc_005AD2A4:   lea ecx, var_1E8
loc_005AD2AA:   push ecx
  loc_005AD2AB: call [004011C4h] ; __vbaVarCat
loc_005AD2B1:   push eax
loc_005AD2B2:   lea edx, var_2E8
loc_005AD2B8:   push edx
loc_005AD2B9:   lea eax, var_1F8
loc_005AD2BF:   push eax
  loc_005AD2C0: call [004011C4h] ; __vbaVarCat
loc_005AD2C6:   push eax
  loc_005AD2C7: call [00401034h] ; __vbaStrVarMove
loc_005AD2CD:   mov edx, eax
  loc_005AD2CF: mov ecx, 0066809Ch
  loc_005AD2D4: call [0040126Ch] ; __vbaStrMove
loc_005AD2DA:   lea ecx, var_58
loc_005AD2DD:   push ecx
loc_005AD2DE:   lea edx, var_50
loc_005AD2E1:   push edx
loc_005AD2E2:   lea eax, var_54
loc_005AD2E5:   push eax
loc_005AD2E6:   lea ecx, var_4C
loc_005AD2E9:   push ecx
loc_005AD2EA:   lea edx, var_48
loc_005AD2ED:   push edx
loc_005AD2EE:   lea eax, var_44
loc_005AD2F1:   push eax
loc_005AD2F2:   lea ecx, var_40
loc_005AD2F5:   push ecx
loc_005AD2F6:   lea edx, var_3C
loc_005AD2F9:   push edx
loc_005AD2FA:   lea eax, var_38
loc_005AD2FD:   push eax
  loc_005AD2FE: push 00000009h
  loc_005AD300: call [00401204h] ; __vbaFreeStrList
  loc_005AD306: add esp, 00000028h
loc_005AD309:   lea ecx, var_74
loc_005AD30C:   push ecx
loc_005AD30D:   lea edx, var_70
loc_005AD310:   push edx
loc_005AD311:   lea eax, var_6C
loc_005AD314:   push eax
loc_005AD315:   lea ecx, var_68
loc_005AD318:   push ecx
loc_005AD319:   lea edx, var_64
loc_005AD31C:   push edx
  loc_005AD31D: push 00000005h
  loc_005AD31F: call [0040104Ch] ; __vbaFreeObjList
  loc_005AD325: add esp, 00000018h
loc_005AD328:   lea eax, var_1F8
loc_005AD32E:   push eax
loc_005AD32F:   lea ecx, var_1E8
loc_005AD335:   push ecx
loc_005AD336:   lea edx, var_1D8
loc_005AD33C:   push edx
loc_005AD33D:   lea eax, var_1C8
loc_005AD343:   push eax
loc_005AD344:   lea ecx, var_1B8
loc_005AD34A:   push ecx
loc_005AD34B:   lea edx, var_188
loc_005AD351:   push edx
loc_005AD352:   lea eax, var_1A8
loc_005AD358:   push eax
loc_005AD359:   lea ecx, var_198
loc_005AD35F:   push ecx
loc_005AD360:   lea edx, var_178
loc_005AD366:   push edx
loc_005AD367:   lea eax, var_168
loc_005AD36D:   push eax
loc_005AD36E:   lea ecx, var_138
loc_005AD374:   push ecx
loc_005AD375:   lea edx, var_158
loc_005AD37B:   push edx
loc_005AD37C:   lea eax, var_148
loc_005AD382:   push eax
loc_005AD383:   lea ecx, var_128
loc_005AD389:   push ecx
loc_005AD38A:   lea edx, var_118
loc_005AD390:   push edx
loc_005AD391:   lea eax, var_E8
loc_005AD397:   push eax
loc_005AD398:   lea ecx, var_108
loc_005AD39E:   push ecx
loc_005AD39F:   lea edx, var_F8
loc_005AD3A5:   push edx
loc_005AD3A6:   lea eax, var_D8
loc_005AD3AC:   push eax
loc_005AD3AD:   lea ecx, var_B8
loc_005AD3B3:   push ecx
loc_005AD3B4:   lea edx, var_C8
loc_005AD3BA:   push edx
loc_005AD3BB:   lea eax, var_A8
loc_005AD3C1:   push eax
loc_005AD3C2:   lea ecx, var_98
loc_005AD3C8:   push ecx
loc_005AD3C9:   lea edx, var_88
loc_005AD3CF:   push edx
  loc_005AD3D0: push 00000018h
  loc_005AD3D2: call [0040103Ch] ; __vbaFreeVarList
  loc_005AD3D8: add esp, 00000064h
  loc_005AD3DB: mov var_4, 0000002Ch
  loc_005AD3E2: cmp [0066803Ch], 00000000h
  loc_005AD3E9: jnz 005AD407h
  loc_005AD3EB: push 0066803Ch
  loc_005AD3F0: push 0042DEFCh
  loc_005AD3F5: call [004011E8h] ; __vbaNew2
  loc_005AD3FB: mov var_55C, 0066803Ch
  loc_005AD405: jmp 005AD411h
  loc_005AD407: mov var_55C, 0066803Ch
loc_005AD411:   mov eax, var_55C
loc_005AD417:   mov ecx, [eax]
loc_005AD419:   mov var_434, ecx
  loc_005AD41F: mov var_80, 80020004h
  loc_005AD426: mov var_88, 0000000Ah
loc_005AD430:   lea ecx, var_88
  loc_005AD436: call [0040123Ch] ; __vbaFreeVarg
loc_005AD43C:   lea edx, var_64
loc_005AD43F:   push edx
  loc_005AD440: push 00000001h
loc_005AD442:   lea eax, var_88
loc_005AD448:   push eax
loc_005AD449:   mov ecx, [0066809Ch]
loc_005AD44F:   push ecx
loc_005AD450:   mov edx, var_434
loc_005AD456:   mov eax, [edx]
loc_005AD458:   mov ecx, var_434
loc_005AD45E:   push ecx
loc_005AD45F:   Call [eax+00000040h]
loc_005AD462:   fnclex
loc_005AD464:   mov var_438, eax
  loc_005AD46A: cmp var_438, 00000000h
  loc_005AD471: jge 005AD496h
  loc_005AD473: push 00000040h
  loc_005AD475: push 0042E1B0h
loc_005AD47A:   mov edx, var_434
loc_005AD480:   push edx
loc_005AD481:   mov eax, var_438
loc_005AD487:   push eax
  loc_005AD488: call [00401080h] ; __vbaHresultCheckObj
loc_005AD48E:   mov var_560, eax
  loc_005AD494: jmp 005AD4A0h
  loc_005AD496: mov var_560, 00000000h
loc_005AD4A0:   lea ecx, var_64
  loc_005AD4A3: call [004012A0h] ; __vbaFreeObj
loc_005AD4A9:   lea ecx, var_88
  loc_005AD4AF: call [00401020h] ; __vbaFreeVar
  loc_005AD4B5: mov var_4, 0000002Dh
loc_005AD4BC:   mov ecx, Me
loc_005AD4BF:   mov dx, [ecx+00000034h]
  loc_005AD4C3: sub dx, 0001h
  loc_005AD4C7: jo 005AEBD1h
loc_005AD4CD:   mov var_460, dx
  loc_005AD4D4: mov var_45C, 0001h
  loc_005AD4DD: mov var_24, 0001h
  loc_005AD4E3: jmp 005AD4FAh
loc_005AD4E5:   mov ax, var_24
loc_005AD4E9:   Add ax, var_45C
  loc_005AD4F0: jo 005AEBD1h
loc_005AD4F6:   mov var_24, ax
loc_005AD4FA:   mov cx, var_24
loc_005AD4FE:   cmp cx, var_460
  loc_005AD505: jg 005AE181h
  loc_005AD50B: mov var_4, 0000002Eh
loc_005AD512:   movsx edx, var_24
loc_005AD516:   mov var_2D0, edx
  loc_005AD51C: mov var_2D8, 00000003h
  loc_005AD526: mov var_2F0, 00000002h
  loc_005AD530: mov var_2F8, 00000003h
  loc_005AD53A: mov eax, 00000010h
  loc_005AD53F: call 00405690h ; __vbaChkstk
loc_005AD544:   mov eax, esp
loc_005AD546:   mov ecx, var_2D8
loc_005AD54C:   mov [eax], ecx
loc_005AD54E:   mov edx, var_2D4
loc_005AD554:   mov [eax+00000004h], edx
loc_005AD557:   mov ecx, var_2D0
loc_005AD55D:   mov [eax+00000008h], ecx
loc_005AD560:   mov edx, var_2CC
loc_005AD566:   mov [eax+0000000Ch], edx
  loc_005AD569: mov eax, 00000010h
  loc_005AD56E: call 00405690h ; __vbaChkstk
loc_005AD573:   mov eax, esp
loc_005AD575:   mov ecx, var_2F8
loc_005AD57B:   mov [eax], ecx
loc_005AD57D:   mov edx, var_2F4
loc_005AD583:   mov [eax+00000004h], edx
loc_005AD586:   mov ecx, var_2F0
loc_005AD58C:   mov [eax+00000008h], ecx
loc_005AD58F:   mov edx, var_2EC
loc_005AD595:   mov [eax+0000000Ch], edx
  loc_005AD598: push 00000002h
  loc_005AD59A: push 00000041h
loc_005AD59C:   mov eax, Me
loc_005AD59F:   mov ecx, [eax]
loc_005AD5A1:   mov edx, Me
loc_005AD5A4:   push edx
loc_005AD5A5:   Call [ecx+00000390h]
loc_005AD5AB:   push eax
loc_005AD5AC:   lea eax, var_6C
loc_005AD5AF:   push eax
  loc_005AD5B0: call [004010B8h] ; __vbaObjSet
loc_005AD5B6:   push eax
loc_005AD5B7:   lea ecx, var_A8
loc_005AD5BD:   push ecx
  loc_005AD5BE: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AD5C4: add esp, 00000030h
loc_005AD5C7:   movsx edx, var_24
loc_005AD5CB:   mov var_3D0, edx
  loc_005AD5D1: mov var_3D8, 00000003h
  loc_005AD5DB: mov var_3F0, 00000005h
  loc_005AD5E5: mov var_3F8, 00000003h
  loc_005AD5EF: mov eax, 00000010h
  loc_005AD5F4: call 00405690h ; __vbaChkstk
loc_005AD5F9:   mov eax, esp
loc_005AD5FB:   mov ecx, var_3D8
loc_005AD601:   mov [eax], ecx
loc_005AD603:   mov edx, var_3D4
loc_005AD609:   mov [eax+00000004h], edx
loc_005AD60C:   mov ecx, var_3D0
loc_005AD612:   mov [eax+00000008h], ecx
loc_005AD615:   mov edx, var_3CC
loc_005AD61B:   mov [eax+0000000Ch], edx
  loc_005AD61E: mov eax, 00000010h
  loc_005AD623: call 00405690h ; __vbaChkstk
loc_005AD628:   mov eax, esp
loc_005AD62A:   mov ecx, var_3F8
loc_005AD630:   mov [eax], ecx
loc_005AD632:   mov edx, var_3F4
loc_005AD638:   mov [eax+00000004h], edx
loc_005AD63B:   mov ecx, var_3F0
loc_005AD641:   mov [eax+00000008h], ecx
loc_005AD644:   mov edx, var_3EC
loc_005AD64A:   mov [eax+0000000Ch], edx
  loc_005AD64D: push 00000002h
  loc_005AD64F: push 00000041h
loc_005AD651:   mov eax, Me
loc_005AD654:   mov ecx, [eax]
loc_005AD656:   mov edx, Me
loc_005AD659:   push edx
loc_005AD65A:   Call [ecx+00000390h]
loc_005AD660:   push eax
loc_005AD661:   lea eax, var_78
loc_005AD664:   push eax
  loc_005AD665: call [004010B8h] ; __vbaObjSet
loc_005AD66B:   push eax
loc_005AD66C:   lea ecx, var_198
loc_005AD672:   push ecx
  loc_005AD673: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AD679: add esp, 00000030h
loc_005AD67C:   movsx edx, var_24
loc_005AD680:   mov var_250, edx
  loc_005AD686: mov var_258, 00000003h
  loc_005AD690: mov var_270, 00000000h
  loc_005AD69A: mov var_278, 00000003h
loc_005AD6A4:   movsx eax, var_24
loc_005AD6A8:   mov var_290, eax
  loc_005AD6AE: mov var_298, 00000003h
  loc_005AD6B8: mov var_2B0, 00000001h
  loc_005AD6C2: mov var_2B8, 00000003h
  loc_005AD6CC: push 00435E5Ch ; "INSERT INTO penjualan_detail"
  loc_005AD6D1: push 00435E9Ch ; "(no_nota,kode_barang,nama_barang,harga_jual,diskon,jumlah,sub_total)"
  loc_005AD6D6: call [00401070h] ; __vbaStrCat
loc_005AD6DC:   mov edx, eax
loc_005AD6DE:   lea ecx, var_38
  loc_005AD6E1: call [0040126Ch] ; __vbaStrMove
loc_005AD6E7:   push eax
  loc_005AD6E8: push 00434620h ; " VALUES ('"
  loc_005AD6ED: call [00401070h] ; __vbaStrCat
loc_005AD6F3:   mov edx, eax
loc_005AD6F5:   lea ecx, var_3C
  loc_005AD6F8: call [0040126Ch] ; __vbaStrMove
loc_005AD6FE:   push eax
loc_005AD6FF:   mov ecx, Me
loc_005AD702:   mov edx, [ecx+0000006Ch]
loc_005AD705:   push edx
  loc_005AD706: call [00401070h] ; __vbaStrCat
loc_005AD70C:   mov edx, eax
loc_005AD70E:   lea ecx, var_40
  loc_005AD711: call [0040126Ch] ; __vbaStrMove
loc_005AD717:   push eax
  loc_005AD718: push 0042EC30h ; "','"
  loc_005AD71D: call [00401070h] ; __vbaStrCat
loc_005AD723:   mov edx, eax
loc_005AD725:   lea ecx, var_44
  loc_005AD728: call [0040126Ch] ; __vbaStrMove
loc_005AD72E:   push eax
  loc_005AD72F: mov eax, 00000010h
  loc_005AD734: call 00405690h ; __vbaChkstk
loc_005AD739:   mov eax, esp
loc_005AD73B:   mov ecx, var_258
loc_005AD741:   mov [eax], ecx
loc_005AD743:   mov edx, var_254
loc_005AD749:   mov [eax+00000004h], edx
loc_005AD74C:   mov ecx, var_250
loc_005AD752:   mov [eax+00000008h], ecx
loc_005AD755:   mov edx, var_24C
loc_005AD75B:   mov [eax+0000000Ch], edx
  loc_005AD75E: mov eax, 00000010h
  loc_005AD763: call 00405690h ; __vbaChkstk
loc_005AD768:   mov eax, esp
loc_005AD76A:   mov ecx, var_278
loc_005AD770:   mov [eax], ecx
loc_005AD772:   mov edx, var_274
loc_005AD778:   mov [eax+00000004h], edx
loc_005AD77B:   mov ecx, var_270
loc_005AD781:   mov [eax+00000008h], ecx
loc_005AD784:   mov edx, var_26C
loc_005AD78A:   mov [eax+0000000Ch], edx
  loc_005AD78D: push 00000002h
  loc_005AD78F: push 00000041h
loc_005AD791:   mov eax, Me
loc_005AD794:   mov ecx, [eax]
loc_005AD796:   mov edx, Me
loc_005AD799:   push edx
loc_005AD79A:   Call [ecx+00000390h]
loc_005AD7A0:   push eax
loc_005AD7A1:   lea eax, var_64
loc_005AD7A4:   push eax
  loc_005AD7A5: call [004010B8h] ; __vbaObjSet
loc_005AD7AB:   push eax
loc_005AD7AC:   lea ecx, var_88
loc_005AD7B2:   push ecx
  loc_005AD7B3: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AD7B9: add esp, 00000030h
loc_005AD7BC:   push eax
  loc_005AD7BD: call [00401034h] ; __vbaStrVarMove
loc_005AD7C3:   mov edx, eax
loc_005AD7C5:   lea ecx, var_48
  loc_005AD7C8: call [0040126Ch] ; __vbaStrMove
loc_005AD7CE:   push eax
  loc_005AD7CF: call [00401070h] ; __vbaStrCat
loc_005AD7D5:   mov edx, eax
loc_005AD7D7:   lea ecx, var_4C
  loc_005AD7DA: call [0040126Ch] ; __vbaStrMove
loc_005AD7E0:   push eax
  loc_005AD7E1: push 0042EC30h ; "','"
  loc_005AD7E6: call [00401070h] ; __vbaStrCat
loc_005AD7EC:   mov edx, eax
loc_005AD7EE:   lea ecx, var_50
  loc_005AD7F1: call [0040126Ch] ; __vbaStrMove
loc_005AD7F7:   push eax
  loc_005AD7F8: mov eax, 00000010h
  loc_005AD7FD: call 00405690h ; __vbaChkstk
loc_005AD802:   mov edx, esp
loc_005AD804:   mov eax, var_298
loc_005AD80A:   mov [edx], eax
loc_005AD80C:   mov ecx, var_294
loc_005AD812:   mov [edx+00000004h], ecx
loc_005AD815:   mov eax, var_290
loc_005AD81B:   mov [edx+00000008h], eax
loc_005AD81E:   mov ecx, var_28C
loc_005AD824:   mov [edx+0000000Ch], ecx
  loc_005AD827: mov eax, 00000010h
  loc_005AD82C: call 00405690h ; __vbaChkstk
loc_005AD831:   mov edx, esp
loc_005AD833:   mov eax, var_2B8
loc_005AD839:   mov [edx], eax
loc_005AD83B:   mov ecx, var_2B4
loc_005AD841:   mov [edx+00000004h], ecx
loc_005AD844:   mov eax, var_2B0
loc_005AD84A:   mov [edx+00000008h], eax
loc_005AD84D:   mov ecx, var_2AC
loc_005AD853:   mov [edx+0000000Ch], ecx
  loc_005AD856: push 00000002h
  loc_005AD858: push 00000041h
loc_005AD85A:   mov edx, Me
loc_005AD85D:   mov eax, [edx]
loc_005AD85F:   mov ecx, Me
loc_005AD862:   push ecx
loc_005AD863:   Call [eax+00000390h]
loc_005AD869:   push eax
loc_005AD86A:   lea edx, var_68
loc_005AD86D:   push edx
  loc_005AD86E: call [004010B8h] ; __vbaObjSet
loc_005AD874:   push eax
loc_005AD875:   lea eax, var_98
loc_005AD87B:   push eax
  loc_005AD87C: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AD882: add esp, 00000030h
loc_005AD885:   push eax
  loc_005AD886: call [00401034h] ; __vbaStrVarMove
loc_005AD88C:   mov edx, eax
loc_005AD88E:   lea ecx, var_54
  loc_005AD891: call [0040126Ch] ; __vbaStrMove
loc_005AD897:   push eax
  loc_005AD898: call [00401070h] ; __vbaStrCat
loc_005AD89E:   mov edx, eax
loc_005AD8A0:   lea ecx, var_58
  loc_005AD8A3: call [0040126Ch] ; __vbaStrMove
loc_005AD8A9:   push eax
  loc_005AD8AA: push 0042EC30h ; "','"
  loc_005AD8AF: call [00401070h] ; __vbaStrCat
loc_005AD8B5:   mov var_E0, eax
  loc_005AD8BB: mov var_E8, 00000008h
  loc_005AD8C5: mov var_310, 00435A20h ; "00"
  loc_005AD8CF: mov var_318, 00000008h
loc_005AD8D9:   lea edx, var_318
loc_005AD8DF:   lea ecx, var_C8
  loc_005AD8E5: call [00401238h] ; __vbaVarDup
loc_005AD8EB:   lea ecx, var_A8
loc_005AD8F1:   push ecx
  loc_005AD8F2: call [00401034h] ; __vbaStrVarMove
loc_005AD8F8:   mov var_B0, eax
  loc_005AD8FE: mov var_B8, 00000008h
  loc_005AD908: push 00000001h
  loc_005AD90A: push 00000001h
loc_005AD90C:   lea edx, var_C8
loc_005AD912:   push edx
loc_005AD913:   lea eax, var_B8
loc_005AD919:   push eax
loc_005AD91A:   lea ecx, var_D8
loc_005AD920:   push ecx
  loc_005AD921: call [00401078h] ; rtcVarFromFormatVar
  loc_005AD927: mov var_320, 0042EC30h ; "','"
  loc_005AD931: mov var_328, 00000008h
loc_005AD93B:   movsx edx, var_24
loc_005AD93F:   mov var_330, edx
  loc_005AD945: mov var_338, 00000003h
  loc_005AD94F: mov var_350, 00000003h
  loc_005AD959: mov var_358, 00000003h
  loc_005AD963: mov eax, 00000010h
  loc_005AD968: call 00405690h ; __vbaChkstk
loc_005AD96D:   mov eax, esp
loc_005AD96F:   mov ecx, var_338
loc_005AD975:   mov [eax], ecx
loc_005AD977:   mov edx, var_334
loc_005AD97D:   mov [eax+00000004h], edx
loc_005AD980:   mov ecx, var_330
loc_005AD986:   mov [eax+00000008h], ecx
loc_005AD989:   mov edx, var_32C
loc_005AD98F:   mov [eax+0000000Ch], edx
  loc_005AD992: mov eax, 00000010h
  loc_005AD997: call 00405690h ; __vbaChkstk
loc_005AD99C:   mov eax, esp
loc_005AD99E:   mov ecx, var_358
loc_005AD9A4:   mov [eax], ecx
loc_005AD9A6:   mov edx, var_354
loc_005AD9AC:   mov [eax+00000004h], edx
loc_005AD9AF:   mov ecx, var_350
loc_005AD9B5:   mov [eax+00000008h], ecx
loc_005AD9B8:   mov edx, var_34C
loc_005AD9BE:   mov [eax+0000000Ch], edx
  loc_005AD9C1: push 00000002h
  loc_005AD9C3: push 00000041h
loc_005AD9C5:   mov eax, Me
loc_005AD9C8:   mov ecx, [eax]
loc_005AD9CA:   mov edx, Me
loc_005AD9CD:   push edx
loc_005AD9CE:   Call [ecx+00000390h]
loc_005AD9D4:   push eax
loc_005AD9D5:   lea eax, var_70
loc_005AD9D8:   push eax
  loc_005AD9D9: call [004010B8h] ; __vbaObjSet
loc_005AD9DF:   push eax
loc_005AD9E0:   lea ecx, var_118
loc_005AD9E6:   push ecx
  loc_005AD9E7: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AD9ED: add esp, 00000030h
loc_005AD9F0:   push eax
  loc_005AD9F1: call [00401034h] ; __vbaStrVarMove
loc_005AD9F7:   mov var_120, eax
  loc_005AD9FD: mov var_128, 00000008h
  loc_005ADA07: mov var_370, 0042EC30h ; "','"
  loc_005ADA11: mov var_378, 00000008h
loc_005ADA1B:   movsx edx, var_24
loc_005ADA1F:   mov var_380, edx
  loc_005ADA25: mov var_388, 00000003h
  loc_005ADA2F: mov var_3A0, 00000004h
  loc_005ADA39: mov var_3A8, 00000003h
  loc_005ADA43: mov eax, 00000010h
  loc_005ADA48: call 00405690h ; __vbaChkstk
loc_005ADA4D:   mov eax, esp
loc_005ADA4F:   mov ecx, var_388
loc_005ADA55:   mov [eax], ecx
loc_005ADA57:   mov edx, var_384
loc_005ADA5D:   mov [eax+00000004h], edx
loc_005ADA60:   mov ecx, var_380
loc_005ADA66:   mov [eax+00000008h], ecx
loc_005ADA69:   mov edx, var_37C
loc_005ADA6F:   mov [eax+0000000Ch], edx
  loc_005ADA72: mov eax, 00000010h
  loc_005ADA77: call 00405690h ; __vbaChkstk
loc_005ADA7C:   mov eax, esp
loc_005ADA7E:   mov ecx, var_3A8
loc_005ADA84:   mov [eax], ecx
loc_005ADA86:   mov edx, var_3A4
loc_005ADA8C:   mov [eax+00000004h], edx
loc_005ADA8F:   mov ecx, var_3A0
loc_005ADA95:   mov [eax+00000008h], ecx
loc_005ADA98:   mov edx, var_39C
loc_005ADA9E:   mov [eax+0000000Ch], edx
  loc_005ADAA1: push 00000002h
  loc_005ADAA3: push 00000041h
loc_005ADAA5:   mov eax, Me
loc_005ADAA8:   mov ecx, [eax]
loc_005ADAAA:   mov edx, Me
loc_005ADAAD:   push edx
loc_005ADAAE:   Call [ecx+00000390h]
loc_005ADAB4:   push eax
loc_005ADAB5:   lea eax, var_74
loc_005ADAB8:   push eax
  loc_005ADAB9: call [004010B8h] ; __vbaObjSet
loc_005ADABF:   push eax
loc_005ADAC0:   lea ecx, var_158
loc_005ADAC6:   push ecx
  loc_005ADAC7: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005ADACD: add esp, 00000030h
loc_005ADAD0:   push eax
  loc_005ADAD1: call [00401034h] ; __vbaStrVarMove
loc_005ADAD7:   mov var_160, eax
  loc_005ADADD: mov var_168, 00000008h
  loc_005ADAE7: mov var_3C0, 0042EC30h ; "','"
  loc_005ADAF1: mov var_3C8, 00000008h
  loc_005ADAFB: mov var_410, 00435A20h ; "00"
  loc_005ADB05: mov var_418, 00000008h
loc_005ADB0F:   lea edx, var_418
loc_005ADB15:   lea ecx, var_1B8
  loc_005ADB1B: call [00401238h] ; __vbaVarDup
loc_005ADB21:   lea edx, var_198
loc_005ADB27:   push edx
  loc_005ADB28: call [00401034h] ; __vbaStrVarMove
loc_005ADB2E:   mov var_1A0, eax
  loc_005ADB34: mov var_1A8, 00000008h
  loc_005ADB3E: push 00000001h
  loc_005ADB40: push 00000001h
loc_005ADB42:   lea eax, var_1B8
loc_005ADB48:   push eax
loc_005ADB49:   lea ecx, var_1A8
loc_005ADB4F:   push ecx
loc_005ADB50:   lea edx, var_1C8
loc_005ADB56:   push edx
  loc_005ADB57: call [00401078h] ; rtcVarFromFormatVar
  loc_005ADB5D: mov var_420, 0042EC3Ch ; "')"
  loc_005ADB67: mov var_428, 00000008h
loc_005ADB71:   lea eax, var_E8
loc_005ADB77:   push eax
loc_005ADB78:   lea ecx, var_D8
loc_005ADB7E:   push ecx
loc_005ADB7F:   lea edx, var_F8
loc_005ADB85:   push edx
  loc_005ADB86: call [004011C4h] ; __vbaVarCat
loc_005ADB8C:   push eax
loc_005ADB8D:   lea eax, var_328
loc_005ADB93:   push eax
loc_005ADB94:   lea ecx, var_108
loc_005ADB9A:   push ecx
  loc_005ADB9B: call [004011C4h] ; __vbaVarCat
loc_005ADBA1:   push eax
loc_005ADBA2:   lea edx, var_128
loc_005ADBA8:   push edx
loc_005ADBA9:   lea eax, var_138
loc_005ADBAF:   push eax
  loc_005ADBB0: call [004011C4h] ; __vbaVarCat
loc_005ADBB6:   push eax
loc_005ADBB7:   lea ecx, var_378
loc_005ADBBD:   push ecx
loc_005ADBBE:   lea edx, var_148
loc_005ADBC4:   push edx
  loc_005ADBC5: call [004011C4h] ; __vbaVarCat
loc_005ADBCB:   push eax
loc_005ADBCC:   lea eax, var_168
loc_005ADBD2:   push eax
loc_005ADBD3:   lea ecx, var_178
loc_005ADBD9:   push ecx
  loc_005ADBDA: call [004011C4h] ; __vbaVarCat
loc_005ADBE0:   push eax
loc_005ADBE1:   lea edx, var_3C8
loc_005ADBE7:   push edx
loc_005ADBE8:   lea eax, var_188
loc_005ADBEE:   push eax
  loc_005ADBEF: call [004011C4h] ; __vbaVarCat
loc_005ADBF5:   push eax
loc_005ADBF6:   lea ecx, var_1C8
loc_005ADBFC:   push ecx
loc_005ADBFD:   lea edx, var_1D8
loc_005ADC03:   push edx
  loc_005ADC04: call [004011C4h] ; __vbaVarCat
loc_005ADC0A:   push eax
loc_005ADC0B:   lea eax, var_428
loc_005ADC11:   push eax
loc_005ADC12:   lea ecx, var_1E8
loc_005ADC18:   push ecx
  loc_005ADC19: call [004011C4h] ; __vbaVarCat
loc_005ADC1F:   push eax
  loc_005ADC20: call [00401034h] ; __vbaStrVarMove
loc_005ADC26:   mov edx, eax
  loc_005ADC28: mov ecx, 0066809Ch
  loc_005ADC2D: call [0040126Ch] ; __vbaStrMove
loc_005ADC33:   lea edx, var_58
loc_005ADC36:   push edx
loc_005ADC37:   lea eax, var_54
loc_005ADC3A:   push eax
loc_005ADC3B:   lea ecx, var_50
loc_005ADC3E:   push ecx
loc_005ADC3F:   lea edx, var_4C
loc_005ADC42:   push edx
loc_005ADC43:   lea eax, var_48
loc_005ADC46:   push eax
loc_005ADC47:   lea ecx, var_44
loc_005ADC4A:   push ecx
loc_005ADC4B:   lea edx, var_40
loc_005ADC4E:   push edx
loc_005ADC4F:   lea eax, var_3C
loc_005ADC52:   push eax
loc_005ADC53:   lea ecx, var_38
loc_005ADC56:   push ecx
  loc_005ADC57: push 00000009h
  loc_005ADC59: call [00401204h] ; __vbaFreeStrList
  loc_005ADC5F: add esp, 00000028h
loc_005ADC62:   lea edx, var_78
loc_005ADC65:   push edx
loc_005ADC66:   lea eax, var_74
loc_005ADC69:   push eax
loc_005ADC6A:   lea ecx, var_70
loc_005ADC6D:   push ecx
loc_005ADC6E:   lea edx, var_6C
loc_005ADC71:   push edx
loc_005ADC72:   lea eax, var_68
loc_005ADC75:   push eax
loc_005ADC76:   lea ecx, var_64
loc_005ADC79:   push ecx
  loc_005ADC7A: push 00000006h
  loc_005ADC7C: call [0040104Ch] ; __vbaFreeObjList
  loc_005ADC82: add esp, 0000001Ch
loc_005ADC85:   lea edx, var_1E8
loc_005ADC8B:   push edx
loc_005ADC8C:   lea eax, var_1D8
loc_005ADC92:   push eax
loc_005ADC93:   lea ecx, var_1C8
loc_005ADC99:   push ecx
loc_005ADC9A:   lea edx, var_188
loc_005ADCA0:   push edx
loc_005ADCA1:   lea eax, var_1B8
loc_005ADCA7:   push eax
loc_005ADCA8:   lea ecx, var_1A8
loc_005ADCAE:   push ecx
loc_005ADCAF:   lea edx, var_198
loc_005ADCB5:   push edx
loc_005ADCB6:   lea eax, var_178
loc_005ADCBC:   push eax
loc_005ADCBD:   lea ecx, var_168
loc_005ADCC3:   push ecx
loc_005ADCC4:   lea edx, var_148
loc_005ADCCA:   push edx
loc_005ADCCB:   lea eax, var_158
loc_005ADCD1:   push eax
loc_005ADCD2:   lea ecx, var_138
loc_005ADCD8:   push ecx
loc_005ADCD9:   lea edx, var_128
loc_005ADCDF:   push edx
loc_005ADCE0:   lea eax, var_108
loc_005ADCE6:   push eax
loc_005ADCE7:   lea ecx, var_118
loc_005ADCED:   push ecx
loc_005ADCEE:   lea edx, var_F8
loc_005ADCF4:   push edx
loc_005ADCF5:   lea eax, var_D8
loc_005ADCFB:   push eax
loc_005ADCFC:   lea ecx, var_E8
loc_005ADD02:   push ecx
loc_005ADD03:   lea edx, var_C8
loc_005ADD09:   push edx
loc_005ADD0A:   lea eax, var_B8
loc_005ADD10:   push eax
loc_005ADD11:   lea ecx, var_A8
loc_005ADD17:   push ecx
loc_005ADD18:   lea edx, var_98
loc_005ADD1E:   push edx
loc_005ADD1F:   lea eax, var_88
loc_005ADD25:   push eax
  loc_005ADD26: push 00000017h
  loc_005ADD28: call [0040103Ch] ; __vbaFreeVarList
  loc_005ADD2E: add esp, 00000060h
  loc_005ADD31: mov var_4, 0000002Fh
  loc_005ADD38: cmp [0066803Ch], 00000000h
  loc_005ADD3F: jnz 005ADD5Dh
  loc_005ADD41: push 0066803Ch
  loc_005ADD46: push 0042DEFCh
  loc_005ADD4B: call [004011E8h] ; __vbaNew2
  loc_005ADD51: mov var_564, 0066803Ch
  loc_005ADD5B: jmp 005ADD67h
  loc_005ADD5D: mov var_564, 0066803Ch
loc_005ADD67:   mov ecx, var_564
loc_005ADD6D:   mov edx, [ecx]
loc_005ADD6F:   mov var_434, edx
  loc_005ADD75: mov var_80, 80020004h
  loc_005ADD7C: mov var_88, 0000000Ah
loc_005ADD86:   lea ecx, var_88
  loc_005ADD8C: call [0040123Ch] ; __vbaFreeVarg
loc_005ADD92:   lea eax, var_64
loc_005ADD95:   push eax
  loc_005ADD96: push 00000001h
loc_005ADD98:   lea ecx, var_88
loc_005ADD9E:   push ecx
loc_005ADD9F:   mov edx, [0066809Ch]
loc_005ADDA5:   push edx
loc_005ADDA6:   mov eax, var_434
loc_005ADDAC:   mov ecx, [eax]
loc_005ADDAE:   mov edx, var_434
loc_005ADDB4:   push edx
loc_005ADDB5:   Call [ecx+00000040h]
loc_005ADDB8:   fnclex
loc_005ADDBA:   mov var_438, eax
  loc_005ADDC0: cmp var_438, 00000000h
  loc_005ADDC7: jge 005ADDECh
  loc_005ADDC9: push 00000040h
  loc_005ADDCB: push 0042E1B0h
loc_005ADDD0:   mov eax, var_434
loc_005ADDD6:   push eax
loc_005ADDD7:   mov ecx, var_438
loc_005ADDDD:   push ecx
  loc_005ADDDE: call [00401080h] ; __vbaHresultCheckObj
loc_005ADDE4:   mov var_568, eax
  loc_005ADDEA: jmp 005ADDF6h
  loc_005ADDEC: mov var_568, 00000000h
loc_005ADDF6:   lea ecx, var_64
  loc_005ADDF9: call [004012A0h] ; __vbaFreeObj
loc_005ADDFF:   lea ecx, var_88
  loc_005ADE05: call [00401020h] ; __vbaFreeVar
  loc_005ADE0B: mov var_4, 00000030h
loc_005ADE12:   movsx edx, var_24
loc_005ADE16:   mov var_250, edx
  loc_005ADE1C: mov var_258, 00000003h
  loc_005ADE26: mov var_270, 00000004h
  loc_005ADE30: mov var_278, 00000003h
  loc_005ADE3A: mov eax, 00000010h
  loc_005ADE3F: call 00405690h ; __vbaChkstk
loc_005ADE44:   mov eax, esp
loc_005ADE46:   mov ecx, var_258
loc_005ADE4C:   mov [eax], ecx
loc_005ADE4E:   mov edx, var_254
loc_005ADE54:   mov [eax+00000004h], edx
loc_005ADE57:   mov ecx, var_250
loc_005ADE5D:   mov [eax+00000008h], ecx
loc_005ADE60:   mov edx, var_24C
loc_005ADE66:   mov [eax+0000000Ch], edx
  loc_005ADE69: mov eax, 00000010h
  loc_005ADE6E: call 00405690h ; __vbaChkstk
loc_005ADE73:   mov eax, esp
loc_005ADE75:   mov ecx, var_278
loc_005ADE7B:   mov [eax], ecx
loc_005ADE7D:   mov edx, var_274
loc_005ADE83:   mov [eax+00000004h], edx
loc_005ADE86:   mov ecx, var_270
loc_005ADE8C:   mov [eax+00000008h], ecx
loc_005ADE8F:   mov edx, var_26C
loc_005ADE95:   mov [eax+0000000Ch], edx
  loc_005ADE98: push 00000002h
  loc_005ADE9A: push 00000041h
loc_005ADE9C:   mov eax, Me
loc_005ADE9F:   mov ecx, [eax]
loc_005ADEA1:   mov edx, Me
loc_005ADEA4:   push edx
loc_005ADEA5:   Call [ecx+00000390h]
loc_005ADEAB:   push eax
loc_005ADEAC:   lea eax, var_64
loc_005ADEAF:   push eax
  loc_005ADEB0: call [004010B8h] ; __vbaObjSet
loc_005ADEB6:   push eax
loc_005ADEB7:   lea ecx, var_88
loc_005ADEBD:   push ecx
  loc_005ADEBE: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005ADEC4: add esp, 00000030h
loc_005ADEC7:   push eax
  loc_005ADEC8: call [00401034h] ; __vbaStrVarMove
loc_005ADECE:   mov edx, eax
loc_005ADED0:   lea ecx, var_38
  loc_005ADED3: call [0040126Ch] ; __vbaStrMove
loc_005ADED9:   push eax
  loc_005ADEDA: call [004012A4h] ; rtcR8ValFromBstr
  loc_005ADEE0: call [00401244h] ; __vbaFpI2
loc_005ADEE6:   mov var_28, ax
loc_005ADEEA:   lea ecx, var_38
  loc_005ADEED: call [0040129Ch] ; __vbaFreeStr
loc_005ADEF3:   lea ecx, var_64
  loc_005ADEF6: call [004012A0h] ; __vbaFreeObj
loc_005ADEFC:   lea ecx, var_88
  loc_005ADF02: call [00401020h] ; __vbaFreeVar
  loc_005ADF08: mov var_4, 00000031h
loc_005ADF0F:   movsx edx, var_24
loc_005ADF13:   mov var_250, edx
  loc_005ADF19: mov var_258, 00000003h
  loc_005ADF23: mov var_270, 00000000h
  loc_005ADF2D: mov var_278, 00000003h
  loc_005ADF37: mov eax, 00000010h
  loc_005ADF3C: call 00405690h ; __vbaChkstk
loc_005ADF41:   mov eax, esp
loc_005ADF43:   mov ecx, var_258
loc_005ADF49:   mov [eax], ecx
loc_005ADF4B:   mov edx, var_254
loc_005ADF51:   mov [eax+00000004h], edx
loc_005ADF54:   mov ecx, var_250
loc_005ADF5A:   mov [eax+00000008h], ecx
loc_005ADF5D:   mov edx, var_24C
loc_005ADF63:   mov [eax+0000000Ch], edx
  loc_005ADF66: mov eax, 00000010h
  loc_005ADF6B: call 00405690h ; __vbaChkstk
loc_005ADF70:   mov eax, esp
loc_005ADF72:   mov ecx, var_278
loc_005ADF78:   mov [eax], ecx
loc_005ADF7A:   mov edx, var_274
loc_005ADF80:   mov [eax+00000004h], edx
loc_005ADF83:   mov ecx, var_270
loc_005ADF89:   mov [eax+00000008h], ecx
loc_005ADF8C:   mov edx, var_26C
loc_005ADF92:   mov [eax+0000000Ch], edx
  loc_005ADF95: push 00000002h
  loc_005ADF97: push 00000041h
loc_005ADF99:   mov eax, Me
loc_005ADF9C:   mov ecx, [eax]
loc_005ADF9E:   mov edx, Me
loc_005ADFA1:   push edx
loc_005ADFA2:   Call [ecx+00000390h]
loc_005ADFA8:   push eax
loc_005ADFA9:   lea eax, var_64
loc_005ADFAC:   push eax
  loc_005ADFAD: call [004010B8h] ; __vbaObjSet
loc_005ADFB3:   push eax
loc_005ADFB4:   lea ecx, var_88
loc_005ADFBA:   push ecx
  loc_005ADFBB: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005ADFC1: add esp, 00000030h
loc_005ADFC4:   push eax
  loc_005ADFC5: call [00401034h] ; __vbaStrVarMove
loc_005ADFCB:   mov edx, eax
loc_005ADFCD:   lea ecx, var_2C
  loc_005ADFD0: call [0040126Ch] ; __vbaStrMove
loc_005ADFD6:   lea ecx, var_64
  loc_005ADFD9: call [004012A0h] ; __vbaFreeObj
loc_005ADFDF:   lea ecx, var_88
  loc_005ADFE5: call [00401020h] ; __vbaFreeVar
  loc_005ADFEB: mov var_4, 00000032h
  loc_005ADFF2: push 00435F2Ch ; "UPDATE brg_barang SET stok=stok - "
loc_005ADFF7:   mov dx, var_28
loc_005ADFFB:   push edx
  loc_005ADFFC: call [0040100Ch] ; __vbaStrI2
loc_005AE002:   mov edx, eax
loc_005AE004:   lea ecx, var_38
  loc_005AE007: call [0040126Ch] ; __vbaStrMove
loc_005AE00D:   push eax
  loc_005AE00E: call [00401070h] ; __vbaStrCat
loc_005AE014:   mov edx, eax
loc_005AE016:   lea ecx, var_3C
  loc_005AE019: call [0040126Ch] ; __vbaStrMove
loc_005AE01F:   push eax
  loc_005AE020: push 0042DFECh
  loc_005AE025: call [00401070h] ; __vbaStrCat
loc_005AE02B:   mov edx, eax
loc_005AE02D:   lea ecx, var_40
  loc_005AE030: call [0040126Ch] ; __vbaStrMove
loc_005AE036:   push eax
  loc_005AE037: push 00433824h ; " WHERE kode_barang='"
  loc_005AE03C: call [00401070h] ; __vbaStrCat
loc_005AE042:   mov edx, eax
loc_005AE044:   lea ecx, var_44
  loc_005AE047: call [0040126Ch] ; __vbaStrMove
loc_005AE04D:   push eax
loc_005AE04E:   mov eax, var_2C
loc_005AE051:   push eax
  loc_005AE052: call [00401070h] ; __vbaStrCat
loc_005AE058:   mov edx, eax
loc_005AE05A:   lea ecx, var_48
  loc_005AE05D: call [0040126Ch] ; __vbaStrMove
loc_005AE063:   push eax
  loc_005AE064: push 0042EBA8h ; "'"
  loc_005AE069: call [00401070h] ; __vbaStrCat
loc_005AE06F:   mov edx, eax
  loc_005AE071: mov ecx, 0066809Ch
  loc_005AE076: call [0040126Ch] ; __vbaStrMove
loc_005AE07C:   lea ecx, var_48
loc_005AE07F:   push ecx
loc_005AE080:   lea edx, var_44
loc_005AE083:   push edx
loc_005AE084:   lea eax, var_40
loc_005AE087:   push eax
loc_005AE088:   lea ecx, var_3C
loc_005AE08B:   push ecx
loc_005AE08C:   lea edx, var_38
loc_005AE08F:   push edx
  loc_005AE090: push 00000005h
  loc_005AE092: call [00401204h] ; __vbaFreeStrList
  loc_005AE098: add esp, 00000018h
  loc_005AE09B: mov var_4, 00000033h
  loc_005AE0A2: cmp [0066803Ch], 00000000h
  loc_005AE0A9: jnz 005AE0C7h
  loc_005AE0AB: push 0066803Ch
  loc_005AE0B0: push 0042DEFCh
  loc_005AE0B5: call [004011E8h] ; __vbaNew2
  loc_005AE0BB: mov var_56C, 0066803Ch
  loc_005AE0C5: jmp 005AE0D1h
  loc_005AE0C7: mov var_56C, 0066803Ch
loc_005AE0D1:   mov eax, var_56C
loc_005AE0D7:   mov ecx, [eax]
loc_005AE0D9:   mov var_434, ecx
  loc_005AE0DF: mov var_80, 80020004h
  loc_005AE0E6: mov var_88, 0000000Ah
loc_005AE0F0:   lea ecx, var_88
  loc_005AE0F6: call [0040123Ch] ; __vbaFreeVarg
loc_005AE0FC:   lea edx, var_64
loc_005AE0FF:   push edx
  loc_005AE100: push 00000001h
loc_005AE102:   lea eax, var_88
loc_005AE108:   push eax
loc_005AE109:   mov ecx, [0066809Ch]
loc_005AE10F:   push ecx
loc_005AE110:   mov edx, var_434
loc_005AE116:   mov eax, [edx]
loc_005AE118:   mov ecx, var_434
loc_005AE11E:   push ecx
loc_005AE11F:   Call [eax+00000040h]
loc_005AE122:   fnclex
loc_005AE124:   mov var_438, eax
  loc_005AE12A: cmp var_438, 00000000h
  loc_005AE131: jge 005AE156h
  loc_005AE133: push 00000040h
  loc_005AE135: push 0042E1B0h
loc_005AE13A:   mov edx, var_434
loc_005AE140:   push edx
loc_005AE141:   mov eax, var_438
loc_005AE147:   push eax
  loc_005AE148: call [00401080h] ; __vbaHresultCheckObj
loc_005AE14E:   mov var_570, eax
  loc_005AE154: jmp 005AE160h
  loc_005AE156: mov var_570, 00000000h
loc_005AE160:   lea ecx, var_64
  loc_005AE163: call [004012A0h] ; __vbaFreeObj
loc_005AE169:   lea ecx, var_88
  loc_005AE16F: call [00401020h] ; __vbaFreeVar
  loc_005AE175: mov var_4, 00000034h
  loc_005AE17C: jmp 005AD4E5h
  loc_005AE181: mov var_4, 00000035h
loc_005AE188:   push FFFFFFFFh
  loc_005AE18A: call [004010BCh] ; __vbaOnError
  loc_005AE190: mov var_4, 00000036h
  loc_005AE197: cmp [00668454h], 00000000h
  loc_005AE19E: jnz 005AE1BCh
  loc_005AE1A0: push 00668454h
  loc_005AE1A5: push 0040F860h
  loc_005AE1AA: call [004011E8h] ; __vbaNew2
  loc_005AE1B0: mov var_574, 00668454h
  loc_005AE1BA: jmp 005AE1C6h
  loc_005AE1BC: mov var_574, 00668454h
loc_005AE1C6:   mov ecx, var_574
loc_005AE1CC:   mov edx, [ecx]
loc_005AE1CE:   push edx
loc_005AE1CF:   lea eax, var_450
loc_005AE1D5:   push eax
  loc_005AE1D6: call [004010C8h] ; __vbaObjSetAddref
  loc_005AE1DC: mov var_4, 00000037h
loc_005AE1E3:   mov ecx, var_450
loc_005AE1E9:   mov edx, [ecx]
loc_005AE1EB:   mov eax, var_450
loc_005AE1F1:   push eax
loc_005AE1F2:   Call [edx+0000031Ch]
loc_005AE1F8:   push eax
loc_005AE1F9:   lea ecx, var_64
loc_005AE1FC:   push ecx
  loc_005AE1FD: call [004010B8h] ; __vbaObjSet
loc_005AE203:   mov var_434, eax
  loc_005AE209: push 00436118h ; "DP"
loc_005AE20E:   mov edx, var_434
loc_005AE214:   mov eax, [edx]
loc_005AE216:   mov ecx, var_434
loc_005AE21C:   push ecx
loc_005AE21D:   Call [eax+00000054h]
loc_005AE220:   fnclex
loc_005AE222:   mov var_438, eax
  loc_005AE228: cmp var_438, 00000000h
  loc_005AE22F: jge 005AE254h
  loc_005AE231: push 00000054h
  loc_005AE233: push 00433F80h
loc_005AE238:   mov edx, var_434
loc_005AE23E:   push edx
loc_005AE23F:   mov eax, var_438
loc_005AE245:   push eax
  loc_005AE246: call [00401080h] ; __vbaHresultCheckObj
loc_005AE24C:   mov var_578, eax
  loc_005AE252: jmp 005AE25Eh
  loc_005AE254: mov var_578, 00000000h
loc_005AE25E:   lea ecx, var_64
  loc_005AE261: call [004012A0h] ; __vbaFreeObj
  loc_005AE267: mov var_4, 00000038h
loc_005AE26E:   mov ecx, var_450
loc_005AE274:   mov edx, [ecx]
loc_005AE276:   mov eax, var_450
loc_005AE27C:   push eax
loc_005AE27D:   Call [edx+00000318h]
loc_005AE283:   push eax
loc_005AE284:   lea ecx, var_64
loc_005AE287:   push ecx
  loc_005AE288: call [004010B8h] ; __vbaObjSet
loc_005AE28E:   mov var_434, eax
  loc_005AE294: push 00436124h ; "Sisa Piutang"
loc_005AE299:   mov edx, var_434
loc_005AE29F:   mov eax, [edx]
loc_005AE2A1:   mov ecx, var_434
loc_005AE2A7:   push ecx
loc_005AE2A8:   Call [eax+00000054h]
loc_005AE2AB:   fnclex
loc_005AE2AD:   mov var_438, eax
  loc_005AE2B3: cmp var_438, 00000000h
  loc_005AE2BA: jge 005AE2DFh
  loc_005AE2BC: push 00000054h
  loc_005AE2BE: push 00433F80h
loc_005AE2C3:   mov edx, var_434
loc_005AE2C9:   push edx
loc_005AE2CA:   mov eax, var_438
loc_005AE2D0:   push eax
  loc_005AE2D1: call [00401080h] ; __vbaHresultCheckObj
loc_005AE2D7:   mov var_57C, eax
  loc_005AE2DD: jmp 005AE2E9h
  loc_005AE2DF: mov var_57C, 00000000h
loc_005AE2E9:   lea ecx, var_64
  loc_005AE2EC: call [004012A0h] ; __vbaFreeObj
  loc_005AE2F2: mov var_4, 00000039h
loc_005AE2F9:   mov ecx, var_450
loc_005AE2FF:   mov edx, [ecx]
loc_005AE301:   mov eax, var_450
loc_005AE307:   push eax
loc_005AE308:   Call [edx+00000308h]
loc_005AE30E:   push eax
loc_005AE30F:   lea ecx, var_68
loc_005AE312:   push ecx
  loc_005AE313: call [004010B8h] ; __vbaObjSet
loc_005AE319:   mov var_43C, eax
loc_005AE31F:   mov edx, Me
loc_005AE322:   mov eax, [edx]
loc_005AE324:   mov ecx, Me
loc_005AE327:   push ecx
loc_005AE328:   Call [eax+0000030Ch]
loc_005AE32E:   push eax
loc_005AE32F:   lea edx, var_64
loc_005AE332:   push edx
  loc_005AE333: call [004010B8h] ; __vbaObjSet
loc_005AE339:   mov var_434, eax
loc_005AE33F:   lea eax, var_38
loc_005AE342:   push eax
loc_005AE343:   mov ecx, var_434
loc_005AE349:   mov edx, [ecx]
loc_005AE34B:   mov eax, var_434
loc_005AE351:   push eax
loc_005AE352:   Call [edx+000000A0h]
loc_005AE358:   fnclex
loc_005AE35A:   mov var_438, eax
  loc_005AE360: cmp var_438, 00000000h
  loc_005AE367: jge 005AE38Fh
  loc_005AE369: push 000000A0h
  loc_005AE36E: push 0042DFCCh
loc_005AE373:   mov ecx, var_434
loc_005AE379:   push ecx
loc_005AE37A:   mov edx, var_438
loc_005AE380:   push edx
  loc_005AE381: call [00401080h] ; __vbaHresultCheckObj
loc_005AE387:   mov var_580, eax
  loc_005AE38D: jmp 005AE399h
  loc_005AE38F: mov var_580, 00000000h
loc_005AE399:   mov eax, var_38
loc_005AE39C:   push eax
loc_005AE39D:   mov ecx, var_43C
loc_005AE3A3:   mov edx, [ecx]
loc_005AE3A5:   mov eax, var_43C
loc_005AE3AB:   push eax
loc_005AE3AC:   Call [edx+000000A4h]
loc_005AE3B2:   fnclex
loc_005AE3B4:   mov var_440, eax
  loc_005AE3BA: cmp var_440, 00000000h
  loc_005AE3C1: jge 005AE3E9h
  loc_005AE3C3: push 000000A4h
  loc_005AE3C8: push 0042DFCCh
loc_005AE3CD:   mov ecx, var_43C
loc_005AE3D3:   push ecx
loc_005AE3D4:   mov edx, var_440
loc_005AE3DA:   push edx
  loc_005AE3DB: call [00401080h] ; __vbaHresultCheckObj
loc_005AE3E1:   mov var_584, eax
  loc_005AE3E7: jmp 005AE3F3h
  loc_005AE3E9: mov var_584, 00000000h
loc_005AE3F3:   lea ecx, var_38
  loc_005AE3F6: call [0040129Ch] ; __vbaFreeStr
loc_005AE3FC:   lea eax, var_68
loc_005AE3FF:   push eax
loc_005AE400:   lea ecx, var_64
loc_005AE403:   push ecx
  loc_005AE404: push 00000002h
  loc_005AE406: call [0040104Ch] ; __vbaFreeObjList
  loc_005AE40C: add esp, 0000000Ch
  loc_005AE40F: mov var_4, 0000003Ah
loc_005AE416:   mov edx, var_450
loc_005AE41C:   mov eax, [edx]
loc_005AE41E:   mov ecx, var_450
loc_005AE424:   push ecx
loc_005AE425:   Call [eax+00000304h]
loc_005AE42B:   push eax
loc_005AE42C:   lea edx, var_68
loc_005AE42F:   push edx
  loc_005AE430: call [004010B8h] ; __vbaObjSet
loc_005AE436:   mov var_43C, eax
loc_005AE43C:   mov eax, Me
loc_005AE43F:   mov ecx, [eax]
loc_005AE441:   mov edx, Me
loc_005AE444:   push edx
loc_005AE445:   Call [ecx+00000338h]
loc_005AE44B:   push eax
loc_005AE44C:   lea eax, var_64
loc_005AE44F:   push eax
  loc_005AE450: call [004010B8h] ; __vbaObjSet
loc_005AE456:   mov var_434, eax
loc_005AE45C:   lea ecx, var_38
loc_005AE45F:   push ecx
loc_005AE460:   mov edx, var_434
loc_005AE466:   mov eax, [edx]
loc_005AE468:   mov ecx, var_434
loc_005AE46E:   push ecx
loc_005AE46F:   Call [eax+00000050h]
loc_005AE472:   fnclex
loc_005AE474:   mov var_438, eax
  loc_005AE47A: cmp var_438, 00000000h
  loc_005AE481: jge 005AE4A6h
  loc_005AE483: push 00000050h
  loc_005AE485: push 00433F80h
loc_005AE48A:   mov edx, var_434
loc_005AE490:   push edx
loc_005AE491:   mov eax, var_438
loc_005AE497:   push eax
  loc_005AE498: call [00401080h] ; __vbaHresultCheckObj
loc_005AE49E:   mov var_588, eax
  loc_005AE4A4: jmp 005AE4B0h
  loc_005AE4A6: mov var_588, 00000000h
loc_005AE4B0:   mov ecx, var_38
loc_005AE4B3:   push ecx
loc_005AE4B4:   mov edx, var_43C
loc_005AE4BA:   mov eax, [edx]
loc_005AE4BC:   mov ecx, var_43C
loc_005AE4C2:   push ecx
loc_005AE4C3:   Call [eax+000000A4h]
loc_005AE4C9:   fnclex
loc_005AE4CB:   mov var_440, eax
  loc_005AE4D1: cmp var_440, 00000000h
  loc_005AE4D8: jge 005AE500h
  loc_005AE4DA: push 000000A4h
  loc_005AE4DF: push 0042DFCCh
loc_005AE4E4:   mov edx, var_43C
loc_005AE4EA:   push edx
loc_005AE4EB:   mov eax, var_440
loc_005AE4F1:   push eax
  loc_005AE4F2: call [00401080h] ; __vbaHresultCheckObj
loc_005AE4F8:   mov var_58C, eax
  loc_005AE4FE: jmp 005AE50Ah
  loc_005AE500: mov var_58C, 00000000h
loc_005AE50A:   lea ecx, var_38
  loc_005AE50D: call [0040129Ch] ; __vbaFreeStr
loc_005AE513:   lea ecx, var_68
loc_005AE516:   push ecx
loc_005AE517:   lea edx, var_64
loc_005AE51A:   push edx
  loc_005AE51B: push 00000002h
  loc_005AE51D: call [0040104Ch] ; __vbaFreeObjList
  loc_005AE523: add esp, 0000000Ch
  loc_005AE526: mov var_4, 0000003Bh
loc_005AE52D:   mov eax, var_450
loc_005AE533:   mov ecx, [eax]
loc_005AE535:   mov edx, var_450
loc_005AE53B:   push edx
loc_005AE53C:   Call [ecx+00000300h]
loc_005AE542:   push eax
loc_005AE543:   lea eax, var_64
loc_005AE546:   push eax
  loc_005AE547: call [004010B8h] ; __vbaObjSet
loc_005AE54D:   mov var_434, eax
  loc_005AE553: push 0042DFECh
loc_005AE558:   mov ecx, var_434
loc_005AE55E:   mov edx, [ecx]
loc_005AE560:   mov eax, var_434
loc_005AE566:   push eax
loc_005AE567:   Call [edx+000000A4h]
loc_005AE56D:   fnclex
loc_005AE56F:   mov var_438, eax
  loc_005AE575: cmp var_438, 00000000h
  loc_005AE57C: jge 005AE5A4h
  loc_005AE57E: push 000000A4h
  loc_005AE583: push 0042DFCCh
loc_005AE588:   mov ecx, var_434
loc_005AE58E:   push ecx
loc_005AE58F:   mov edx, var_438
loc_005AE595:   push edx
  loc_005AE596: call [00401080h] ; __vbaHresultCheckObj
loc_005AE59C:   mov var_590, eax
  loc_005AE5A2: jmp 005AE5AEh
  loc_005AE5A4: mov var_590, 00000000h
loc_005AE5AE:   lea ecx, var_64
  loc_005AE5B1: call [004012A0h] ; __vbaFreeObj
  loc_005AE5B7: mov var_4, 0000003Ch
loc_005AE5BE:   mov eax, var_450
loc_005AE5C4:   mov ecx, [eax]
loc_005AE5C6:   mov edx, var_450
loc_005AE5CC:   push edx
loc_005AE5CD:   Call [ecx+000002FCh]
loc_005AE5D3:   push eax
loc_005AE5D4:   lea eax, var_68
loc_005AE5D7:   push eax
  loc_005AE5D8: call [004010B8h] ; __vbaObjSet
loc_005AE5DE:   mov var_43C, eax
loc_005AE5E4:   mov ecx, Me
loc_005AE5E7:   mov edx, [ecx]
loc_005AE5E9:   mov eax, Me
loc_005AE5EC:   push eax
loc_005AE5ED:   Call [edx+00000338h]
loc_005AE5F3:   push eax
loc_005AE5F4:   lea ecx, var_64
loc_005AE5F7:   push ecx
  loc_005AE5F8: call [004010B8h] ; __vbaObjSet
loc_005AE5FE:   mov var_434, eax
loc_005AE604:   lea edx, var_38
loc_005AE607:   push edx
loc_005AE608:   mov eax, var_434
loc_005AE60E:   mov ecx, [eax]
loc_005AE610:   mov edx, var_434
loc_005AE616:   push edx
loc_005AE617:   Call [ecx+00000050h]
loc_005AE61A:   fnclex
loc_005AE61C:   mov var_438, eax
  loc_005AE622: cmp var_438, 00000000h
  loc_005AE629: jge 005AE64Eh
  loc_005AE62B: push 00000050h
  loc_005AE62D: push 00433F80h
loc_005AE632:   mov eax, var_434
loc_005AE638:   push eax
loc_005AE639:   mov ecx, var_438
loc_005AE63F:   push ecx
  loc_005AE640: call [00401080h] ; __vbaHresultCheckObj
loc_005AE646:   mov var_594, eax
  loc_005AE64C: jmp 005AE658h
  loc_005AE64E: mov var_594, 00000000h
loc_005AE658:   mov edx, var_38
loc_005AE65B:   push edx
loc_005AE65C:   mov eax, var_43C
loc_005AE662:   mov ecx, [eax]
loc_005AE664:   mov edx, var_43C
loc_005AE66A:   push edx
loc_005AE66B:   Call [ecx+000000A4h]
loc_005AE671:   fnclex
loc_005AE673:   mov var_440, eax
  loc_005AE679: cmp var_440, 00000000h
  loc_005AE680: jge 005AE6A8h
  loc_005AE682: push 000000A4h
  loc_005AE687: push 0042DFCCh
loc_005AE68C:   mov eax, var_43C
loc_005AE692:   push eax
loc_005AE693:   mov ecx, var_440
loc_005AE699:   push ecx
  loc_005AE69A: call [00401080h] ; __vbaHresultCheckObj
loc_005AE6A0:   mov var_598, eax
  loc_005AE6A6: jmp 005AE6B2h
  loc_005AE6A8: mov var_598, 00000000h
loc_005AE6B2:   lea ecx, var_38
  loc_005AE6B5: call [0040129Ch] ; __vbaFreeStr
loc_005AE6BB:   lea edx, var_68
loc_005AE6BE:   push edx
loc_005AE6BF:   lea eax, var_64
loc_005AE6C2:   push eax
  loc_005AE6C3: push 00000002h
  loc_005AE6C5: call [0040104Ch] ; __vbaFreeObjList
  loc_005AE6CB: add esp, 0000000Ch
  loc_005AE6CE: mov var_4, 0000003Dh
loc_005AE6D5:   mov ecx, var_450
loc_005AE6DB:   mov edx, [ecx]
loc_005AE6DD:   mov eax, var_450
loc_005AE6E3:   push eax
loc_005AE6E4:   Call [edx+00000300h]
loc_005AE6EA:   push eax
loc_005AE6EB:   lea ecx, var_64
loc_005AE6EE:   push ecx
  loc_005AE6EF: call [004010B8h] ; __vbaObjSet
loc_005AE6F5:   mov var_434, eax
loc_005AE6FB:   mov edx, var_434
loc_005AE701:   mov eax, [edx]
loc_005AE703:   mov ecx, var_434
loc_005AE709:   push ecx
loc_005AE70A:   Call [eax+00000204h]
loc_005AE710:   fnclex
loc_005AE712:   mov var_438, eax
  loc_005AE718: cmp var_438, 00000000h
  loc_005AE71F: jge 005AE747h
  loc_005AE721: push 00000204h
  loc_005AE726: push 0042DFCCh
loc_005AE72B:   mov edx, var_434
loc_005AE731:   push edx
loc_005AE732:   mov eax, var_438
loc_005AE738:   push eax
  loc_005AE739: call [00401080h] ; __vbaHresultCheckObj
loc_005AE73F:   mov var_59C, eax
  loc_005AE745: jmp 005AE751h
  loc_005AE747: mov var_59C, 00000000h
loc_005AE751:   lea ecx, var_64
  loc_005AE754: call [004012A0h] ; __vbaFreeObj
  loc_005AE75A: mov var_4, 0000003Eh
  loc_005AE761: mov var_260, 80020004h
  loc_005AE76B: mov var_268, 0000000Ah
  loc_005AE775: mov var_250, 00000001h
  loc_005AE77F: mov var_258, 00000002h
  loc_005AE789: mov eax, 00000010h
  loc_005AE78E: call 00405690h ; __vbaChkstk
loc_005AE793:   mov ecx, esp
loc_005AE795:   mov edx, var_268
loc_005AE79B:   mov [ecx], edx
loc_005AE79D:   mov eax, var_264
loc_005AE7A3:   mov [ecx+00000004h], eax
loc_005AE7A6:   mov edx, var_260
loc_005AE7AC:   mov [ecx+00000008h], edx
loc_005AE7AF:   mov eax, var_25C
loc_005AE7B5:   mov [ecx+0000000Ch], eax
  loc_005AE7B8: mov eax, 00000010h
  loc_005AE7BD: call 00405690h ; __vbaChkstk
loc_005AE7C2:   mov ecx, esp
loc_005AE7C4:   mov edx, var_258
loc_005AE7CA:   mov [ecx], edx
loc_005AE7CC:   mov eax, var_254
loc_005AE7D2:   mov [ecx+00000004h], eax
loc_005AE7D5:   mov edx, var_250
loc_005AE7DB:   mov [ecx+00000008h], edx
loc_005AE7DE:   mov eax, var_24C
loc_005AE7E4:   mov [ecx+0000000Ch], eax
loc_005AE7E7:   mov ecx, var_450
loc_005AE7ED:   mov edx, [ecx]
loc_005AE7EF:   mov eax, var_450
loc_005AE7F5:   push eax
loc_005AE7F6:   Call [edx+000002B0h]
loc_005AE7FC:   fnclex
loc_005AE7FE:   mov var_434, eax
  loc_005AE804: cmp var_434, 00000000h
  loc_005AE80B: jge 005AE833h
  loc_005AE80D: push 000002B0h
  loc_005AE812: push 00435F74h
loc_005AE817:   mov ecx, var_450
loc_005AE81D:   push ecx
loc_005AE81E:   mov edx, var_434
loc_005AE824:   push edx
  loc_005AE825: call [00401080h] ; __vbaHresultCheckObj
loc_005AE82B:   mov var_5A0, eax
  loc_005AE831: jmp 005AE83Dh
  loc_005AE833: mov var_5A0, 00000000h
  loc_005AE83D: mov var_4, 0000003Fh
  loc_005AE844: push 00000000h
loc_005AE846:   lea eax, var_450
loc_005AE84C:   push eax
  loc_005AE84D: call [004010C8h] ; __vbaObjSetAddref
  loc_005AE853: mov var_4, 00000041h
loc_005AE85A:   mov ecx, Me
loc_005AE85D:   mov edx, [ecx]
loc_005AE85F:   mov eax, Me
loc_005AE862:   push eax
loc_005AE863:   Call [edx+00000710h]
loc_005AE869:   mov var_434, eax
  loc_005AE86F: cmp var_434, 00000000h
  loc_005AE876: jge 005AE89Bh
  loc_005AE878: push 00000710h
  loc_005AE87D: push 00434B78h
loc_005AE882:   mov ecx, Me
loc_005AE885:   push ecx
loc_005AE886:   mov edx, var_434
loc_005AE88C:   push edx
  loc_005AE88D: call [00401080h] ; __vbaHresultCheckObj
loc_005AE893:   mov var_5A4, eax
  loc_005AE899: jmp 005AE8A5h
  loc_005AE89B: mov var_5A4, 00000000h
  loc_005AE8A5: mov var_4, 00000042h
loc_005AE8AC:   mov eax, Me
loc_005AE8AF:   mov ecx, [eax]
loc_005AE8B1:   mov edx, Me
loc_005AE8B4:   push edx
loc_005AE8B5:   Call [ecx+0000030Ch]
loc_005AE8BB:   push eax
loc_005AE8BC:   lea eax, var_64
loc_005AE8BF:   push eax
  loc_005AE8C0: call [004010B8h] ; __vbaObjSet
loc_005AE8C6:   mov var_434, eax
loc_005AE8CC:   mov ecx, Me
loc_005AE8CF:   mov edx, [ecx+0000006Ch]
loc_005AE8D2:   push edx
loc_005AE8D3:   mov eax, var_434
loc_005AE8D9:   mov ecx, [eax]
loc_005AE8DB:   mov edx, var_434
loc_005AE8E1:   push edx
loc_005AE8E2:   Call [ecx+000000A4h]
loc_005AE8E8:   fnclex
loc_005AE8EA:   mov var_438, eax
  loc_005AE8F0: cmp var_438, 00000000h
  loc_005AE8F7: jge 005AE91Fh
  loc_005AE8F9: push 000000A4h
  loc_005AE8FE: push 0042DFCCh
loc_005AE903:   mov eax, var_434
loc_005AE909:   push eax
loc_005AE90A:   mov ecx, var_438
loc_005AE910:   push ecx
  loc_005AE911: call [00401080h] ; __vbaHresultCheckObj
loc_005AE917:   mov var_5A8, eax
  loc_005AE91D: jmp 005AE929h
  loc_005AE91F: mov var_5A8, 00000000h
loc_005AE929:   lea ecx, var_64
  loc_005AE92C: call [004012A0h] ; __vbaFreeObj
  loc_005AE932: mov var_4, 00000043h
loc_005AE939:   mov edx, Me
loc_005AE93C:   mov eax, [edx]
loc_005AE93E:   mov ecx, Me
loc_005AE941:   push ecx
loc_005AE942:   Call [eax+0000070Ch]
loc_005AE948:   mov var_434, eax
  loc_005AE94E: cmp var_434, 00000000h
  loc_005AE955: jge 005AE97Ah
  loc_005AE957: push 0000070Ch
  loc_005AE95C: push 00434B78h
loc_005AE961:   mov edx, Me
loc_005AE964:   push edx
loc_005AE965:   mov eax, var_434
loc_005AE96B:   push eax
  loc_005AE96C: call [00401080h] ; __vbaHresultCheckObj
loc_005AE972:   mov var_5AC, eax
  loc_005AE978: jmp 005AE984h
  loc_005AE97A: mov var_5AC, 00000000h
  loc_005AE984: mov var_4, 00000044h
loc_005AE98B:   mov ecx, Me
loc_005AE98E:   mov edx, [ecx]
loc_005AE990:   mov eax, Me
loc_005AE993:   push eax
loc_005AE994:   Call [edx+00000338h]
loc_005AE99A:   push eax
loc_005AE99B:   lea ecx, var_64
loc_005AE99E:   push ecx
  loc_005AE99F: call [004010B8h] ; __vbaObjSet
loc_005AE9A5:   mov var_434, eax
  loc_005AE9AB: push 0042E3A4h
loc_005AE9B0:   mov edx, var_434
loc_005AE9B6:   mov eax, [edx]
loc_005AE9B8:   mov ecx, var_434
loc_005AE9BE:   push ecx
loc_005AE9BF:   Call [eax+00000054h]
loc_005AE9C2:   fnclex
loc_005AE9C4:   mov var_438, eax
  loc_005AE9CA: cmp var_438, 00000000h
  loc_005AE9D1: jge 005AE9F6h
  loc_005AE9D3: push 00000054h
  loc_005AE9D5: push 00433F80h
loc_005AE9DA:   mov edx, var_434
loc_005AE9E0:   push edx
loc_005AE9E1:   mov eax, var_438
loc_005AE9E7:   push eax
  loc_005AE9E8: call [00401080h] ; __vbaHresultCheckObj
loc_005AE9EE:   mov var_5B0, eax
  loc_005AE9F4: jmp 005AEA00h
  loc_005AE9F6: mov var_5B0, 00000000h
loc_005AEA00:   lea ecx, var_64
  loc_005AEA03: call [004012A0h] ; __vbaFreeObj
  loc_005AEA09: mov var_4, 00000045h
  loc_005AEA10: push 00000000h
  loc_005AEA12: push 80011000h
loc_005AEA17:   mov ecx, Me
loc_005AEA1A:   mov edx, [ecx]
loc_005AEA1C:   mov eax, Me
loc_005AEA1F:   push eax
loc_005AEA20:   Call [edx+0000037Ch]
loc_005AEA26:   push eax
loc_005AEA27:   lea ecx, var_64
loc_005AEA2A:   push ecx
  loc_005AEA2B: call [004010B8h] ; __vbaObjSet
loc_005AEA31:   push eax
  loc_005AEA32: call [0040102Ch] ; __vbaLateIdCall
  loc_005AEA38: add esp, 0000000Ch
loc_005AEA3B:   lea ecx, var_64
  loc_005AEA3E: call [004012A0h] ; __vbaFreeObj
  loc_005AEA44: mov var_10, 00000000h
loc_005AEA4B:   fwait
  loc_005AEA4C: push 005AEBAAh
  loc_005AEA51: jmp 005AEB87h
loc_005AEA56:   lea edx, var_60
loc_005AEA59:   push edx
loc_005AEA5A:   lea eax, var_5C
loc_005AEA5D:   push eax
loc_005AEA5E:   lea ecx, var_58
loc_005AEA61:   push ecx
loc_005AEA62:   lea edx, var_54
loc_005AEA65:   push edx
loc_005AEA66:   lea eax, var_50
loc_005AEA69:   push eax
loc_005AEA6A:   lea ecx, var_4C
loc_005AEA6D:   push ecx
loc_005AEA6E:   lea edx, var_48
loc_005AEA71:   push edx
loc_005AEA72:   lea eax, var_44
loc_005AEA75:   push eax
loc_005AEA76:   lea ecx, var_40
loc_005AEA79:   push ecx
loc_005AEA7A:   lea edx, var_3C
loc_005AEA7D:   push edx
loc_005AEA7E:   lea eax, var_38
loc_005AEA81:   push eax
  loc_005AEA82: push 0000000Bh
  loc_005AEA84: call [00401204h] ; __vbaFreeStrList
  loc_005AEA8A: add esp, 00000030h
loc_005AEA8D:   lea ecx, var_78
loc_005AEA90:   push ecx
loc_005AEA91:   lea edx, var_74
loc_005AEA94:   push edx
loc_005AEA95:   lea eax, var_70
loc_005AEA98:   push eax
loc_005AEA99:   lea ecx, var_6C
loc_005AEA9C:   push ecx
loc_005AEA9D:   lea edx, var_68
loc_005AEAA0:   push edx
loc_005AEAA1:   lea eax, var_64
loc_005AEAA4:   push eax
  loc_005AEAA5: push 00000006h
  loc_005AEAA7: call [0040104Ch] ; __vbaFreeObjList
  loc_005AEAAD: add esp, 0000001Ch
loc_005AEAB0:   lea ecx, var_248
loc_005AEAB6:   push ecx
loc_005AEAB7:   lea edx, var_238
loc_005AEABD:   push edx
loc_005AEABE:   lea eax, var_228
loc_005AEAC4:   push eax
loc_005AEAC5:   lea ecx, var_218
loc_005AEACB:   push ecx
loc_005AEACC:   lea edx, var_208
loc_005AEAD2:   push edx
loc_005AEAD3:   lea eax, var_1F8
loc_005AEAD9:   push eax
loc_005AEADA:   lea ecx, var_1E8
loc_005AEAE0:   push ecx
loc_005AEAE1:   lea edx, var_1D8
loc_005AEAE7:   push edx
loc_005AEAE8:   lea eax, var_1C8
loc_005AEAEE:   push eax
loc_005AEAEF:   lea ecx, var_1B8
loc_005AEAF5:   push ecx
loc_005AEAF6:   lea edx, var_1A8
loc_005AEAFC:   push edx
loc_005AEAFD:   lea eax, var_198
loc_005AEB03:   push eax
loc_005AEB04:   lea ecx, var_188
loc_005AEB0A:   push ecx
loc_005AEB0B:   lea edx, var_178
loc_005AEB11:   push edx
loc_005AEB12:   lea eax, var_168
loc_005AEB18:   push eax
loc_005AEB19:   lea ecx, var_158
loc_005AEB1F:   push ecx
loc_005AEB20:   lea edx, var_148
loc_005AEB26:   push edx
loc_005AEB27:   lea eax, var_138
loc_005AEB2D:   push eax
loc_005AEB2E:   lea ecx, var_128
loc_005AEB34:   push ecx
loc_005AEB35:   lea edx, var_118
loc_005AEB3B:   push edx
loc_005AEB3C:   lea eax, var_108
loc_005AEB42:   push eax
loc_005AEB43:   lea ecx, var_F8
loc_005AEB49:   push ecx
loc_005AEB4A:   lea edx, var_E8
loc_005AEB50:   push edx
loc_005AEB51:   lea eax, var_D8
loc_005AEB57:   push eax
loc_005AEB58:   lea ecx, var_C8
loc_005AEB5E:   push ecx
loc_005AEB5F:   lea edx, var_B8
loc_005AEB65:   push edx
loc_005AEB66:   lea eax, var_A8
loc_005AEB6C:   push eax
loc_005AEB6D:   lea ecx, var_98
loc_005AEB73:   push ecx
loc_005AEB74:   lea edx, var_88
loc_005AEB7A:   push edx
  loc_005AEB7B: push 0000001Dh
  loc_005AEB7D: call [0040103Ch] ; __vbaFreeVarList
  loc_005AEB83: add esp, 00000078h
loc_005AEB86:   ret
loc_005AEB87:   lea eax, var_450
loc_005AEB8D:   push eax
loc_005AEB8E:   lea ecx, var_44C
loc_005AEB94:   push ecx
  loc_005AEB95: push 00000002h
  loc_005AEB97: call [0040104Ch] ; __vbaFreeObjList
  loc_005AEB9D: add esp, 0000000Ch
loc_005AEBA0:   lea ecx, var_2C
  loc_005AEBA3: call [0040129Ch] ; __vbaFreeStr
loc_005AEBA9:   ret
loc_005AEBAA:   mov edx, Me
loc_005AEBAD:   mov eax, [edx]
loc_005AEBAF:   mov ecx, Me
loc_005AEBB2:   push ecx
loc_005AEBB3:   Call [eax+00000008h]
loc_005AEBB6:   mov eax, var_10
loc_005AEBB9:   mov ecx, var_20
loc_005AEBBC:   mov fs: [00000000h] , ecx
loc_005AEBC3:   pop edi
loc_005AEBC4:   pop esi
loc_005AEBC5:   pop ebx
loc_005AEBC6:   mov esp, ebp
loc_005AEBC8:   pop ebp
  loc_005AEBC9: retn 0004h
End Sub

Private Sub CmdBaru_UnknownEvent_10() '5A9560
loc_005A9560:   push ebp
loc_005A9561:   mov ebp, esp
  loc_005A9563: sub esp, 0000000Ch
  loc_005A9566: push 00405696h ; __vbaExceptHandler
loc_005A956B:   mov eax, fs: [00000000h]
loc_005A9571:   push eax
loc_005A9572:   mov fs: [00000000h] , esp
  loc_005A9579: sub esp, 0000006Ch
loc_005A957C:   push ebx
loc_005A957D:   push esi
loc_005A957E:   push edi
loc_005A957F:   mov var_C, esp
  loc_005A9582: mov var_8, 00402DF8h
loc_005A9589:   mov esi, Me
loc_005A958C:   mov eax, esi
  loc_005A958E: and eax, 00000001h
loc_005A9591:   mov var_4, eax
  loc_005A9594: and esi, FFFFFFFEh
loc_005A9597:   push esi
loc_005A9598:   mov Me, esi
loc_005A959B:   mov ecx, [esi]
loc_005A959D:   Call [ecx+00000004h]
loc_005A95A0:   lea edx, var_2C
  loc_005A95A3: xor eax, eax
loc_005A95A5:   push edx
loc_005A95A6:   mov var_18, eax
loc_005A95A9:   mov var_1C, eax
loc_005A95AC:   mov var_2C, eax
loc_005A95AF:   mov var_3C, eax
loc_005A95B2:   mov var_4C, eax
loc_005A95B5:   mov var_5C, eax
  loc_005A95B8: call [00401220h] ; rtcGetDateVar
loc_005A95BE:   mov eax, [esi]
loc_005A95C0:   push esi
loc_005A95C1:   Call [eax+00000360h]
  loc_005A95C7: mov edi, [004010B8h] ; __vbaObjSet
loc_005A95CD:   lea ecx, var_1C
loc_005A95D0:   push eax
loc_005A95D1:   push ecx
loc_005A95D2:   Call edi
loc_005A95D4:   mov ebx, eax
loc_005A95D6:   lea edx, var_5C
loc_005A95D9:   lea ecx, var_3C
loc_005A95DC:   mov var_70, ebx
  loc_005A95DF: mov var_54, 00433FA8h ; "dd/MM/yyyy"
  loc_005A95E6: mov var_5C, 00000008h
  loc_005A95ED: call [00401238h] ; __vbaVarDup
  loc_005A95F3: push 00000001h
loc_005A95F5:   lea edx, var_3C
  loc_005A95F8: push 00000001h
loc_005A95FA:   lea eax, var_2C
loc_005A95FD:   push edx
loc_005A95FE:   lea ecx, var_4C
loc_005A9601:   push eax
loc_005A9602:   push ecx
  loc_005A9603: call [00401078h] ; rtcVarFromFormatVar
loc_005A9609:   mov ebx, [ebx]
loc_005A960B:   lea edx, var_4C
loc_005A960E:   lea eax, var_18
loc_005A9611:   push edx
loc_005A9612:   push eax
  loc_005A9613: call [004011C0h] ; __vbaStrVarVal
loc_005A9619:   mov ecx, ebx
loc_005A961B:   mov ebx, var_70
loc_005A961E:   push eax
loc_005A961F:   push ebx
loc_005A9620:   Call [ecx+00000054h]
loc_005A9623:   test eax, eax
loc_005A9625:   fnclex
  loc_005A9627: jge 005A9638h
  loc_005A9629: push 00000054h
  loc_005A962B: push 00433F80h
loc_005A9630:   push ebx
loc_005A9631:   push eax
  loc_005A9632: call [00401080h] ; __vbaHresultCheckObj
loc_005A9638:   lea ecx, var_18
  loc_005A963B: call [0040129Ch] ; __vbaFreeStr
loc_005A9641:   lea ecx, var_1C
  loc_005A9644: call [004012A0h] ; __vbaFreeObj
loc_005A964A:   lea edx, var_4C
loc_005A964D:   lea eax, var_3C
loc_005A9650:   push edx
loc_005A9651:   lea ecx, var_2C
loc_005A9654:   push eax
loc_005A9655:   push ecx
  loc_005A9656: push 00000003h
  loc_005A9658: call [0040103Ch] ; __vbaFreeVarList
loc_005A965E:   mov edx, [esi]
  loc_005A9660: add esp, 00000010h
loc_005A9663:   push esi
loc_005A9664:   Call [edx+00000710h]
loc_005A966A:   test eax, eax
  loc_005A966C: jge 005A9680h
  loc_005A966E: push 00000710h
  loc_005A9673: push 00434B78h
loc_005A9678:   push esi
loc_005A9679:   push eax
  loc_005A967A: call [00401080h] ; __vbaHresultCheckObj
loc_005A9680:   mov eax, [esi]
loc_005A9682:   push esi
loc_005A9683:   Call [eax+0000030Ch]
loc_005A9689:   lea ecx, var_1C
loc_005A968C:   push eax
loc_005A968D:   push ecx
loc_005A968E:   Call edi
loc_005A9690:   mov ebx, eax
loc_005A9692:   mov eax, [esi+0000006Ch]
loc_005A9695:   push eax
loc_005A9696:   push ebx
loc_005A9697:   mov edx, [ebx]
loc_005A9699:   Call [edx+000000A4h]
loc_005A969F:   test eax, eax
loc_005A96A1:   fnclex
  loc_005A96A3: jge 005A96BBh
  loc_005A96A5: push 000000A4h
  loc_005A96AA: push 0042DFCCh
loc_005A96AF:   push ebx
  loc_005A96B0: mov ebx, [00401080h] ; __vbaHresultCheckObj
loc_005A96B6:   push eax
loc_005A96B7:   Call ebx
  loc_005A96B9: jmp 005A96C1h
  loc_005A96BB: mov ebx, [00401080h] ; __vbaHresultCheckObj
loc_005A96C1:   lea ecx, var_1C
  loc_005A96C4: call [004012A0h] ; __vbaFreeObj
loc_005A96CA:   mov ecx, [esi]
loc_005A96CC:   push esi
loc_005A96CD:   Call [ecx+00000708h]
loc_005A96D3:   test eax, eax
  loc_005A96D5: jge 005A96E5h
  loc_005A96D7: push 00000708h
  loc_005A96DC: push 00434B78h
loc_005A96E1:   push esi
loc_005A96E2:   push eax
loc_005A96E3:   Call ebx
loc_005A96E5:   mov edx, [esi]
loc_005A96E7:   push esi
loc_005A96E8:   Call [edx+00000714h]
loc_005A96EE:   test eax, eax
  loc_005A96F0: jge 005A9700h
  loc_005A96F2: push 00000714h
  loc_005A96F7: push 00434B78h
loc_005A96FC:   push esi
loc_005A96FD:   push eax
loc_005A96FE:   Call ebx
  loc_005A9700: sub esp, 00000010h
  loc_005A9703: mov ecx, 0000000Bh
loc_005A9708:   mov edx, esp
loc_005A970A:   mov var_5C, ecx
  loc_005A970D: xor eax, eax
  loc_005A970F: push 8001000Dh
loc_005A9714:   mov [edx], ecx
loc_005A9716:   mov ecx, var_58
loc_005A9719:   mov var_54, eax
loc_005A971C:   push esi
loc_005A971D:   mov [edx+00000004h], ecx
loc_005A9720:   mov ecx, [esi]
loc_005A9722:   mov [edx+00000008h], eax
loc_005A9725:   mov eax, var_50
loc_005A9728:   mov [edx+0000000Ch], eax
loc_005A972B:   Call [ecx+0000037Ch]
loc_005A9731:   lea edx, var_1C
loc_005A9734:   push eax
loc_005A9735:   push edx
loc_005A9736:   Call edi
  loc_005A9738: mov ebx, [00401280h] ; __vbaLateIdSt
loc_005A973E:   push eax
loc_005A973F:   Call ebx
loc_005A9741:   lea ecx, var_1C
  loc_005A9744: call [004012A0h] ; __vbaFreeObj
  loc_005A974A: sub esp, 00000010h
  loc_005A974D: mov ecx, 0000000Bh
loc_005A9752:   mov edx, esp
loc_005A9754:   mov var_5C, ecx
  loc_005A9757: or eax, FFFFFFFFh
  loc_005A975A: push 8001000Dh
loc_005A975F:   mov [edx], ecx
loc_005A9761:   mov ecx, var_58
loc_005A9764:   mov var_54, eax
loc_005A9767:   push esi
loc_005A9768:   mov [edx+00000004h], ecx
loc_005A976B:   mov ecx, [esi]
loc_005A976D:   mov [edx+00000008h], eax
loc_005A9770:   mov eax, var_50
loc_005A9773:   mov [edx+0000000Ch], eax
loc_005A9776:   Call [ecx+00000380h]
loc_005A977C:   lea edx, var_1C
loc_005A977F:   push eax
loc_005A9780:   push edx
loc_005A9781:   Call edi
loc_005A9783:   push eax
loc_005A9784:   Call ebx
loc_005A9786:   lea ecx, var_1C
  loc_005A9789: call [004012A0h] ; __vbaFreeObj
  loc_005A978F: sub esp, 00000010h
  loc_005A9792: mov ecx, 00000008h
loc_005A9797:   mov edx, esp
loc_005A9799:   mov var_5C, ecx
  loc_005A979C: mov eax, 004341FCh ; "&Batal"
loc_005A97A1:   push FFFFFDFAh
loc_005A97A6:   mov [edx], ecx
loc_005A97A8:   mov ecx, var_58
loc_005A97AB:   mov var_54, eax
loc_005A97AE:   push esi
loc_005A97AF:   mov [edx+00000004h], ecx
loc_005A97B2:   mov ecx, [esi]
loc_005A97B4:   mov [edx+00000008h], eax
loc_005A97B7:   mov eax, var_50
loc_005A97BA:   mov [edx+0000000Ch], eax
loc_005A97BD:   Call [ecx+00000398h]
loc_005A97C3:   lea edx, var_1C
loc_005A97C6:   push eax
loc_005A97C7:   push edx
loc_005A97C8:   Call edi
loc_005A97CA:   push eax
loc_005A97CB:   Call ebx
loc_005A97CD:   lea ecx, var_1C
  loc_005A97D0: call [004012A0h] ; __vbaFreeObj
  loc_005A97D6: or eax, FFFFFFFFh
  loc_005A97D9: sub esp, 00000010h
loc_005A97DC:   mov edx, esp
  loc_005A97DE: mov ecx, 0000000Bh
loc_005A97E3:   mov var_5C, ecx
loc_005A97E6:   mov var_54, eax
loc_005A97E9:   mov [edx], ecx
loc_005A97EB:   mov ecx, var_58
loc_005A97EE:   mov [edx+00000004h], ecx
loc_005A97F1:   mov ecx, [esi]
  loc_005A97F3: push 8001000Dh
loc_005A97F8:   push esi
loc_005A97F9:   mov [edx+00000008h], eax
loc_005A97FC:   mov eax, var_50
loc_005A97FF:   mov [edx+0000000Ch], eax
loc_005A9802:   Call [ecx+00000394h]
loc_005A9808:   lea edx, var_1C
loc_005A980B:   push eax
loc_005A980C:   push edx
loc_005A980D:   Call edi
loc_005A980F:   push eax
loc_005A9810:   Call ebx
  loc_005A9812: mov ebx, [004012A0h] ; __vbaFreeObj
loc_005A9818:   lea ecx, var_1C
loc_005A981B:   Call ebx
loc_005A981D:   mov eax, [esi]
loc_005A981F:   push esi
loc_005A9820:   Call [eax+00000348h]
loc_005A9826:   lea ecx, var_1C
loc_005A9829:   push eax
loc_005A982A:   push ecx
loc_005A982B:   Call edi
loc_005A982D:   mov edi, eax
loc_005A982F:   push edi
loc_005A9830:   mov edx, [edi]
loc_005A9832:   Call [edx+00000204h]
loc_005A9838:   test eax, eax
loc_005A983A:   fnclex
  loc_005A983C: jge 005A9850h
  loc_005A983E: push 00000204h
  loc_005A9843: push 0042DFCCh
loc_005A9848:   push edi
loc_005A9849:   push eax
  loc_005A984A: call [00401080h] ; __vbaHresultCheckObj
loc_005A9850:   lea ecx, var_1C
loc_005A9853:   Call ebx
  loc_005A9855: mov [esi+00000034h], 0001h
  loc_005A985B: mov var_4, 00000000h
  loc_005A9862: push 005A9894h
  loc_005A9867: jmp 005A9893h
loc_005A9869:   lea ecx, var_18
  loc_005A986C: call [0040129Ch] ; __vbaFreeStr
loc_005A9872:   lea ecx, var_1C
  loc_005A9875: call [004012A0h] ; __vbaFreeObj
loc_005A987B:   lea eax, var_4C
loc_005A987E:   lea ecx, var_3C
loc_005A9881:   push eax
loc_005A9882:   lea edx, var_2C
loc_005A9885:   push ecx
loc_005A9886:   push edx
  loc_005A9887: push 00000003h
  loc_005A9889: call [0040103Ch] ; __vbaFreeVarList
  loc_005A988F: add esp, 00000010h
loc_005A9892:   ret
loc_005A9893:   ret
loc_005A9894:   mov eax, Me
loc_005A9897:   push eax
loc_005A9898:   mov ecx, [eax]
loc_005A989A:   Call [ecx+00000008h]
loc_005A989D:   mov eax, var_4
loc_005A98A0:   mov ecx, var_14
loc_005A98A3:   pop edi
loc_005A98A4:   pop esi
loc_005A98A5:   mov fs: [00000000h] , ecx
loc_005A98AC:   pop ebx
loc_005A98AD:   mov esp, ebp
loc_005A98AF:   pop ebp
  loc_005A98B0: retn 0004h
End Sub

Private Sub cmdKeluar_UnknownEvent_10() '5AEBE0
loc_005AEBE0:   push ebp
loc_005AEBE1:   mov ebp, esp
  loc_005AEBE3: sub esp, 0000000Ch
  loc_005AEBE6: push 00405696h ; __vbaExceptHandler
loc_005AEBEB:   mov eax, fs: [00000000h]
loc_005AEBF1:   push eax
loc_005AEBF2:   mov fs: [00000000h] , esp
  loc_005AEBF9: sub esp, 0000002Ch
loc_005AEBFC:   push ebx
loc_005AEBFD:   push esi
loc_005AEBFE:   push edi
loc_005AEBFF:   mov var_C, esp
  loc_005AEC02: mov var_8, 00402F48h
loc_005AEC09:   mov edi, Me
loc_005AEC0C:   mov eax, edi
  loc_005AEC0E: and eax, 00000001h
loc_005AEC11:   mov var_4, eax
  loc_005AEC14: and edi, FFFFFFFEh
loc_005AEC17:   push edi
loc_005AEC18:   mov Me, edi
loc_005AEC1B:   mov ecx, [edi]
loc_005AEC1D:   Call [ecx+00000004h]
loc_005AEC20:   mov edx, [edi]
  loc_005AEC22: xor ebx, ebx
loc_005AEC24:   push ebx
loc_005AEC25:   push FFFFFDFAh
loc_005AEC2A:   push edi
loc_005AEC2B:   mov var_18, ebx
loc_005AEC2E:   mov var_1C, ebx
loc_005AEC31:   mov var_2C, ebx
loc_005AEC34:   Call [edx+00000398h]
loc_005AEC3A:   push eax
loc_005AEC3B:   lea eax, var_1C
loc_005AEC3E:   push eax
  loc_005AEC3F: call [004010B8h] ; __vbaObjSet
loc_005AEC45:   lea ecx, var_2C
loc_005AEC48:   push eax
loc_005AEC49:   push ecx
  loc_005AEC4A: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005AEC50: add esp, 00000010h
loc_005AEC53:   push eax
  loc_005AEC54: call [00401034h] ; __vbaStrVarMove
loc_005AEC5A:   mov edx, eax
loc_005AEC5C:   lea ecx, var_18
  loc_005AEC5F: call [0040126Ch] ; __vbaStrMove
loc_005AEC65:   push eax
  loc_005AEC66: push 0042E930h ; "&Keluar"
  loc_005AEC6B: call [0040112Ch] ; __vbaStrCmp
loc_005AEC71:   mov esi, eax
loc_005AEC73:   lea ecx, var_18
loc_005AEC76:   neg esi
loc_005AEC78:   sbb esi, esi
loc_005AEC7A:   inc esi
loc_005AEC7B:   neg esi
  loc_005AEC7D: call [0040129Ch] ; __vbaFreeStr
loc_005AEC83:   lea ecx, var_1C
  loc_005AEC86: call [004012A0h] ; __vbaFreeObj
loc_005AEC8C:   lea ecx, var_2C
  loc_005AEC8F: call [00401020h] ; __vbaFreeVar
loc_005AEC95:   cmp si, bx
  loc_005AEC98: jz 005AED35h
loc_005AEC9E:   cmp [00668228h], ebx
  loc_005AECA4: jnz 005AECB6h
  loc_005AECA6: push 00668228h
  loc_005AECAB: push 004296BCh
  loc_005AECB0: call [004011E8h] ; __vbaNew2
loc_005AECB6:   mov esi, [00668228h]
loc_005AECBC:   push FFFFFFFFh
loc_005AECBE:   push esi
loc_005AECBF:   mov edx, [esi]
loc_005AECC1:   Call [edx+00000094h]
loc_005AECC7:   cmp eax, ebx
loc_005AECC9:   fnclex
  loc_005AECCB: jge 005AECDFh
  loc_005AECCD: push 00000094h
  loc_005AECD2: push 00433880h
loc_005AECD7:   push esi
loc_005AECD8:   push eax
  loc_005AECD9: call [00401080h] ; __vbaHresultCheckObj
loc_005AECDF:   cmp [0066A318h], ebx
  loc_005AECE5: jnz 005AECF7h
  loc_005AECE7: push 0066A318h
  loc_005AECEC: push 0042E390h
  loc_005AECF1: call [004011E8h] ; __vbaNew2
loc_005AECF7:   mov esi, [0066A318h]
loc_005AECFD:   lea eax, var_1C
loc_005AED00:   push edi
loc_005AED01:   push eax
loc_005AED02:   mov edx, [esi]
loc_005AED04:   mov var_40, edx
  loc_005AED07: call [004010C8h] ; __vbaObjSetAddref
loc_005AED0D:   mov ecx, var_40
loc_005AED10:   push eax
loc_005AED11:   push esi
loc_005AED12:   Call [ecx+00000010h]
loc_005AED15:   cmp eax, ebx
loc_005AED17:   fnclex
  loc_005AED19: jge 005AED2Ah
  loc_005AED1B: push 00000010h
  loc_005AED1D: push 0042E380h
loc_005AED22:   push esi
loc_005AED23:   push eax
  loc_005AED24: call [00401080h] ; __vbaHresultCheckObj
loc_005AED2A:   lea ecx, var_1C
  loc_005AED2D: call [004012A0h] ; __vbaFreeObj
  loc_005AED33: jmp 005AED54h
loc_005AED35:   mov edx, [edi]
loc_005AED37:   push edi
loc_005AED38:   Call [edx+0000070Ch]
loc_005AED3E:   cmp eax, ebx
  loc_005AED40: jge 005AED54h
  loc_005AED42: push 0000070Ch
  loc_005AED47: push 00434B78h
loc_005AED4C:   push edi
loc_005AED4D:   push eax
  loc_005AED4E: call [00401080h] ; __vbaHresultCheckObj
loc_005AED54:   mov var_4, ebx
  loc_005AED57: push 005AED7Bh
  loc_005AED5C: jmp 005AED7Ah
loc_005AED5E:   lea ecx, var_18
  loc_005AED61: call [0040129Ch] ; __vbaFreeStr
loc_005AED67:   lea ecx, var_1C
  loc_005AED6A: call [004012A0h] ; __vbaFreeObj
loc_005AED70:   lea ecx, var_2C
  loc_005AED73: call [00401020h] ; __vbaFreeVar
loc_005AED79:   ret
loc_005AED7A:   ret
loc_005AED7B:   mov eax, Me
loc_005AED7E:   push eax
loc_005AED7F:   mov ecx, [eax]
loc_005AED81:   Call [ecx+00000008h]
loc_005AED84:   mov eax, var_4
loc_005AED87:   mov ecx, var_14
loc_005AED8A:   pop edi
loc_005AED8B:   pop esi
loc_005AED8C:   mov fs: [00000000h] , ecx
loc_005AED93:   pop ebx
loc_005AED94:   mov esp, ebp
loc_005AED96:   pop ebp
  loc_005AED97: retn 0004h
End Sub

Private Sub Form_Load() '5A6610
loc_005A6610:   push ebp
loc_005A6611:   mov ebp, esp
  loc_005A6613: sub esp, 0000000Ch
  loc_005A6616: push 00405696h ; __vbaExceptHandler
loc_005A661B:   mov eax, fs: [00000000h]
loc_005A6621:   push eax
loc_005A6622:   mov fs: [00000000h] , esp
  loc_005A6629: sub esp, 0000006Ch
loc_005A662C:   push ebx
loc_005A662D:   push esi
loc_005A662E:   push edi
loc_005A662F:   mov var_C, esp
  loc_005A6632: mov var_8, 00402D40h
loc_005A6639:   mov esi, Me
loc_005A663C:   mov eax, esi
  loc_005A663E: and eax, 00000001h
loc_005A6641:   mov var_4, eax
  loc_005A6644: and esi, FFFFFFFEh
loc_005A6647:   push esi
loc_005A6648:   mov Me, esi
loc_005A664B:   mov ecx, [esi]
loc_005A664D:   Call [ecx+00000004h]
  loc_005A6650: xor eax, eax
loc_005A6652:   mov var_18, eax
loc_005A6655:   mov var_1C, eax
loc_005A6658:   mov var_2C, eax
loc_005A665B:   mov var_3C, eax
loc_005A665E:   mov var_4C, eax
loc_005A6661:   mov var_5C, eax
  loc_005A6664: call 0055B320h
loc_005A6669:   lea edx, var_2C
loc_005A666C:   push edx
  loc_005A666D: call [00401220h] ; rtcGetDateVar
loc_005A6673:   mov eax, [esi]
loc_005A6675:   push esi
loc_005A6676:   Call [eax+00000360h]
loc_005A667C:   lea ecx, var_1C
loc_005A667F:   push eax
loc_005A6680:   push ecx
  loc_005A6681: call [004010B8h] ; __vbaObjSet
loc_005A6687:   lea edx, var_5C
loc_005A668A:   lea ecx, var_3C
loc_005A668D:   mov edi, eax
  loc_005A668F: mov var_54, 00433FA8h ; "dd/MM/yyyy"
  loc_005A6696: mov var_5C, 00000008h
  loc_005A669D: call [00401238h] ; __vbaVarDup
  loc_005A66A3: push 00000001h
loc_005A66A5:   lea edx, var_3C
  loc_005A66A8: push 00000001h
loc_005A66AA:   lea eax, var_2C
loc_005A66AD:   push edx
loc_005A66AE:   lea ecx, var_4C
loc_005A66B1:   push eax
loc_005A66B2:   push ecx
  loc_005A66B3: call [00401078h] ; rtcVarFromFormatVar
loc_005A66B9:   mov ebx, [edi]
loc_005A66BB:   lea edx, var_4C
loc_005A66BE:   lea eax, var_18
loc_005A66C1:   push edx
loc_005A66C2:   push eax
  loc_005A66C3: call [004011C0h] ; __vbaStrVarVal
loc_005A66C9:   push eax
loc_005A66CA:   push edi
loc_005A66CB:   Call [ebx+00000054h]
loc_005A66CE:   test eax, eax
loc_005A66D0:   fnclex
  loc_005A66D2: jge 005A66E3h
  loc_005A66D4: push 00000054h
  loc_005A66D6: push 00433F80h
loc_005A66DB:   push edi
loc_005A66DC:   push eax
  loc_005A66DD: call [00401080h] ; __vbaHresultCheckObj
loc_005A66E3:   lea ecx, var_18
  loc_005A66E6: call [0040129Ch] ; __vbaFreeStr
  loc_005A66EC: mov edi, [004012A0h] ; __vbaFreeObj
loc_005A66F2:   lea ecx, var_1C
loc_005A66F5:   Call edi
loc_005A66F7:   lea ecx, var_4C
loc_005A66FA:   lea edx, var_3C
loc_005A66FD:   push ecx
loc_005A66FE:   lea eax, var_2C
loc_005A6701:   push edx
loc_005A6702:   push eax
  loc_005A6703: push 00000003h
  loc_005A6705: call [0040103Ch] ; __vbaFreeVarList
loc_005A670B:   mov ecx, [esi]
  loc_005A670D: add esp, 00000010h
loc_005A6710:   push esi
loc_005A6711:   Call [ecx+0000035Ch]
loc_005A6717:   lea edx, var_1C
loc_005A671A:   push eax
loc_005A671B:   push edx
  loc_005A671C: call [004010B8h] ; __vbaObjSet
loc_005A6722:   mov ebx, [eax]
loc_005A6724:   mov var_70, eax
loc_005A6727:   lea eax, var_18
  loc_005A672A: push 006680B4h
loc_005A672F:   push eax
  loc_005A6730: call [004011C0h] ; __vbaStrVarVal
loc_005A6736:   mov ecx, ebx
loc_005A6738:   mov ebx, var_70
loc_005A673B:   push eax
loc_005A673C:   push ebx
loc_005A673D:   Call [ecx+00000054h]
loc_005A6740:   test eax, eax
loc_005A6742:   fnclex
  loc_005A6744: jge 005A6755h
  loc_005A6746: push 00000054h
  loc_005A6748: push 00433F80h
loc_005A674D:   push ebx
loc_005A674E:   push eax
  loc_005A674F: call [00401080h] ; __vbaHresultCheckObj
loc_005A6755:   lea ecx, var_18
  loc_005A6758: call [0040129Ch] ; __vbaFreeStr
loc_005A675E:   lea ecx, var_1C
loc_005A6761:   Call edi
  loc_005A6763: sub esp, 00000010h
  loc_005A6766: mov ecx, 00000008h
loc_005A676B:   mov edx, esp
loc_005A676D:   mov var_5C, ecx
  loc_005A6770: mov eax, 00433F74h ; "Tunai"
  loc_005A6775: push 00000001h
loc_005A6777:   mov [edx], ecx
loc_005A6779:   mov ecx, var_58
loc_005A677C:   mov var_54, eax
loc_005A677F:   push FFFFFDD7h
loc_005A6784:   mov [edx+00000004h], ecx
loc_005A6787:   mov ecx, [esi]
loc_005A6789:   push esi
loc_005A678A:   mov [edx+00000008h], eax
loc_005A678D:   mov eax, var_50
loc_005A6790:   mov [edx+0000000Ch], eax
loc_005A6793:   Call [ecx+00000388h]
  loc_005A6799: mov ebx, [004010B8h] ; __vbaObjSet
loc_005A679F:   lea edx, var_1C
loc_005A67A2:   push eax
loc_005A67A3:   push edx
loc_005A67A4:   Call ebx
loc_005A67A6:   push eax
  loc_005A67A7: call [0040102Ch] ; __vbaLateIdCall
  loc_005A67AD: add esp, 0000001Ch
loc_005A67B0:   lea ecx, var_1C
loc_005A67B3:   Call edi
  loc_005A67B5: sub esp, 00000010h
  loc_005A67B8: mov ecx, 00000008h
loc_005A67BD:   mov edx, esp
loc_005A67BF:   mov var_5C, ecx
  loc_005A67C2: mov eax, 00433F94h ; "Kredit"
  loc_005A67C7: push 00000001h
loc_005A67C9:   mov [edx], ecx
loc_005A67CB:   mov ecx, var_58
loc_005A67CE:   mov var_54, eax
loc_005A67D1:   push FFFFFDD7h
loc_005A67D6:   mov [edx+00000004h], ecx
loc_005A67D9:   mov ecx, [esi]
loc_005A67DB:   push esi
loc_005A67DC:   mov [edx+00000008h], eax
loc_005A67DF:   mov eax, var_50
loc_005A67E2:   mov [edx+0000000Ch], eax
loc_005A67E5:   Call [ecx+00000388h]
loc_005A67EB:   lea edx, var_1C
loc_005A67EE:   push eax
loc_005A67EF:   push edx
loc_005A67F0:   Call ebx
loc_005A67F2:   push eax
  loc_005A67F3: call [0040102Ch] ; __vbaLateIdCall
  loc_005A67F9: add esp, 0000001Ch
loc_005A67FC:   lea ecx, var_1C
loc_005A67FF:   Call edi
loc_005A6801:   mov eax, [esi]
loc_005A6803:   push esi
loc_005A6804:   Call [eax+00000718h]
loc_005A680A:   test eax, eax
  loc_005A680C: jge 005A6820h
  loc_005A680E: push 00000718h
  loc_005A6813: push 00434B78h
loc_005A6818:   push esi
loc_005A6819:   push eax
  loc_005A681A: call [00401080h] ; __vbaHresultCheckObj
loc_005A6820:   mov ecx, [esi]
loc_005A6822:   push esi
loc_005A6823:   Call [ecx+00000710h]
loc_005A6829:   test eax, eax
  loc_005A682B: jge 005A683Fh
  loc_005A682D: push 00000710h
  loc_005A6832: push 00434B78h
loc_005A6837:   push esi
loc_005A6838:   push eax
  loc_005A6839: call [00401080h] ; __vbaHresultCheckObj
loc_005A683F:   mov edx, [esi]
loc_005A6841:   push esi
loc_005A6842:   Call [edx+0000030Ch]
loc_005A6848:   push eax
loc_005A6849:   lea eax, var_1C
loc_005A684C:   push eax
loc_005A684D:   Call ebx
loc_005A684F:   mov edx, [esi+0000006Ch]
loc_005A6852:   mov ecx, [eax]
loc_005A6854:   push edx
loc_005A6855:   push eax
loc_005A6856:   mov var_70, eax
loc_005A6859:   Call [ecx+000000A4h]
loc_005A685F:   test eax, eax
loc_005A6861:   fnclex
  loc_005A6863: jge 005A687Ah
loc_005A6865:   mov ecx, var_70
  loc_005A6868: push 000000A4h
  loc_005A686D: push 0042DFCCh
loc_005A6872:   push ecx
loc_005A6873:   push eax
  loc_005A6874: call [00401080h] ; __vbaHresultCheckObj
loc_005A687A:   lea ecx, var_1C
loc_005A687D:   Call edi
loc_005A687F:   mov edx, [esi]
loc_005A6881:   push esi
loc_005A6882:   Call [edx+000006FCh]
loc_005A6888:   test eax, eax
  loc_005A688A: jge 005A689Eh
  loc_005A688C: push 000006FCh
  loc_005A6891: push 00434B78h
loc_005A6896:   push esi
loc_005A6897:   push eax
  loc_005A6898: call [00401080h] ; __vbaHresultCheckObj
  loc_005A689E: sub esp, 00000010h
  loc_005A68A1: mov ecx, 0000000Bh
loc_005A68A6:   mov edx, esp
loc_005A68A8:   mov var_5C, ecx
  loc_005A68AB: xor eax, eax
  loc_005A68AD: push 8001000Dh
loc_005A68B2:   mov [edx], ecx
loc_005A68B4:   mov ecx, var_58
loc_005A68B7:   mov var_54, eax
loc_005A68BA:   push esi
loc_005A68BB:   mov [edx+00000004h], ecx
loc_005A68BE:   mov ecx, [esi]
loc_005A68C0:   mov [edx+00000008h], eax
loc_005A68C3:   mov eax, var_50
loc_005A68C6:   mov [edx+0000000Ch], eax
loc_005A68C9:   Call [ecx+0000037Ch]
loc_005A68CF:   lea edx, var_1C
loc_005A68D2:   push eax
loc_005A68D3:   push edx
loc_005A68D4:   Call ebx
loc_005A68D6:   push eax
  loc_005A68D7: call [00401280h] ; __vbaLateIdSt
loc_005A68DD:   lea ecx, var_1C
loc_005A68E0:   Call edi
  loc_005A68E2: sub esp, 00000010h
  loc_005A68E5: mov ecx, 0000000Bh
loc_005A68EA:   mov edx, esp
loc_005A68EC:   mov var_5C, ecx
  loc_005A68EF: or eax, FFFFFFFFh
  loc_005A68F2: push 8001000Dh
loc_005A68F7:   mov [edx], ecx
loc_005A68F9:   mov ecx, var_58
loc_005A68FC:   mov var_54, eax
loc_005A68FF:   push esi
loc_005A6900:   mov [edx+00000004h], ecx
loc_005A6903:   mov ecx, [esi]
loc_005A6905:   mov [edx+00000008h], eax
loc_005A6908:   mov eax, var_50
loc_005A690B:   mov [edx+0000000Ch], eax
loc_005A690E:   Call [ecx+00000380h]
loc_005A6914:   lea edx, var_1C
loc_005A6917:   push eax
loc_005A6918:   push edx
loc_005A6919:   Call ebx
loc_005A691B:   push eax
  loc_005A691C: call [00401280h] ; __vbaLateIdSt
loc_005A6922:   lea ecx, var_1C
loc_005A6925:   Call edi
  loc_005A6927: sub esp, 00000010h
  loc_005A692A: mov ecx, 00000008h
loc_005A692F:   mov edx, esp
loc_005A6931:   mov var_5C, ecx
  loc_005A6934: mov eax, 004341FCh ; "&Batal"
loc_005A6939:   push FFFFFDFAh
loc_005A693E:   mov [edx], ecx
loc_005A6940:   mov ecx, var_58
loc_005A6943:   mov var_54, eax
loc_005A6946:   push esi
loc_005A6947:   mov [edx+00000004h], ecx
loc_005A694A:   mov ecx, [esi]
loc_005A694C:   mov [edx+00000008h], eax
loc_005A694F:   mov eax, var_50
loc_005A6952:   mov [edx+0000000Ch], eax
loc_005A6955:   Call [ecx+00000398h]
loc_005A695B:   lea edx, var_1C
loc_005A695E:   push eax
loc_005A695F:   push edx
loc_005A6960:   Call ebx
loc_005A6962:   push eax
  loc_005A6963: call [00401280h] ; __vbaLateIdSt
loc_005A6969:   lea ecx, var_1C
loc_005A696C:   Call edi
  loc_005A696E: or eax, FFFFFFFFh
  loc_005A6971: sub esp, 00000010h
loc_005A6974:   mov edx, esp
  loc_005A6976: mov ecx, 0000000Bh
loc_005A697B:   mov var_5C, ecx
loc_005A697E:   mov var_54, eax
loc_005A6981:   mov [edx], ecx
loc_005A6983:   mov ecx, var_58
loc_005A6986:   mov [edx+00000004h], ecx
loc_005A6989:   mov ecx, [esi]
loc_005A698B:   mov [edx+00000008h], eax
loc_005A698E:   mov eax, var_50
  loc_005A6991: push 8001000Dh
loc_005A6996:   push esi
loc_005A6997:   mov [edx+0000000Ch], eax
loc_005A699A:   Call [ecx+00000394h]
loc_005A69A0:   lea edx, var_1C
loc_005A69A3:   push eax
loc_005A69A4:   push edx
loc_005A69A5:   Call ebx
loc_005A69A7:   push eax
  loc_005A69A8: call [00401280h] ; __vbaLateIdSt
loc_005A69AE:   lea ecx, var_1C
loc_005A69B1:   Call edi
  loc_005A69B3: xor edi, edi
loc_005A69B5:   lea edx, var_5C
loc_005A69B8:   lea ecx, [esi+00000080h]
  loc_005A69BE: mov [esi+00000034h], 0001h
loc_005A69C4:   mov var_54, edi
  loc_005A69C7: mov var_5C, 00000002h
  loc_005A69CE: call [0040101Ch] ; __vbaVarMove
loc_005A69D4:   mov var_4, edi
  loc_005A69D7: push 005A6A09h
  loc_005A69DC: jmp 005A6A08h
loc_005A69DE:   lea ecx, var_18
  loc_005A69E1: call [0040129Ch] ; __vbaFreeStr
loc_005A69E7:   lea ecx, var_1C
  loc_005A69EA: call [004012A0h] ; __vbaFreeObj
loc_005A69F0:   lea eax, var_4C
loc_005A69F3:   lea ecx, var_3C
loc_005A69F6:   push eax
loc_005A69F7:   lea edx, var_2C
loc_005A69FA:   push ecx
loc_005A69FB:   push edx
  loc_005A69FC: push 00000003h
  loc_005A69FE: call [0040103Ch] ; __vbaFreeVarList
  loc_005A6A04: add esp, 00000010h
loc_005A6A07:   ret
loc_005A6A08:   ret
loc_005A6A09:   mov eax, Me
loc_005A6A0C:   push eax
loc_005A6A0D:   mov ecx, [eax]
loc_005A6A0F:   Call [ecx+00000008h]
loc_005A6A12:   mov eax, var_4
loc_005A6A15:   mov ecx, var_14
loc_005A6A18:   pop edi
loc_005A6A19:   pop esi
loc_005A6A1A:   mov fs: [00000000h] , ecx
loc_005A6A21:   pop ebx
loc_005A6A22:   mov esp, ebp
loc_005A6A24:   pop ebp
  loc_005A6A25: retn 0004h
End Sub

Private Sub Form_Unload(Cancel As Integer) '5A6A30
loc_005A6A30:   push ebp
loc_005A6A31:   mov ebp, esp
  loc_005A6A33: sub esp, 0000000Ch
  loc_005A6A36: push 00405696h ; __vbaExceptHandler
loc_005A6A3B:   mov eax, fs: [00000000h]
loc_005A6A41:   push eax
loc_005A6A42:   mov fs: [00000000h] , esp
  loc_005A6A49: sub esp, 00000098h
loc_005A6A4F:   push ebx
loc_005A6A50:   push esi
loc_005A6A51:   push edi
loc_005A6A52:   mov var_C, esp
  loc_005A6A55: mov var_8, 00402D50h
loc_005A6A5C:   mov edi, Me
loc_005A6A5F:   mov eax, edi
  loc_005A6A61: and eax, 00000001h
loc_005A6A64:   mov var_4, eax
  loc_005A6A67: and edi, FFFFFFFEh
loc_005A6A6A:   push edi
loc_005A6A6B:   mov Me, edi
loc_005A6A6E:   mov ecx, [edi]
loc_005A6A70:   Call [ecx+00000004h]
  loc_005A6A73: mov ebx, [00401238h] ; __vbaVarDup
  loc_005A6A79: mov ecx, 80020004h
  loc_005A6A7E: xor esi, esi
loc_005A6A80:   mov var_50, ecx
  loc_005A6A83: mov eax, 0000000Ah
loc_005A6A88:   mov var_40, ecx
loc_005A6A8B:   mov var_48, esi
loc_005A6A8E:   mov var_58, esi
loc_005A6A91:   mov var_78, esi
loc_005A6A94:   lea edx, var_78
loc_005A6A97:   lea ecx, var_38
loc_005A6A9A:   mov var_18, esi
loc_005A6A9D:   mov var_28, esi
loc_005A6AA0:   mov var_38, esi
loc_005A6AA3:   mov var_68, esi
loc_005A6AA6:   mov var_58, eax
loc_005A6AA9:   mov var_48, eax
  loc_005A6AAC: mov var_70, 0042ED7Ch ; "Konfirmasi"
  loc_005A6AB3: mov var_78, 00000008h
loc_005A6ABA:   Call ebx
loc_005A6ABC:   lea edx, var_68
loc_005A6ABF:   lea ecx, var_28
  loc_005A6AC2: mov var_60, 00434874h ; "APAKAH KAMU INGIN MENUTUP APLIKASI INI..?"
  loc_005A6AC9: mov var_68, 00000008h
loc_005A6AD0:   Call ebx
loc_005A6AD2:   lea edx, var_58
loc_005A6AD5:   lea eax, var_48
loc_005A6AD8:   push edx
loc_005A6AD9:   lea ecx, var_38
loc_005A6ADC:   push eax
loc_005A6ADD:   push ecx
loc_005A6ADE:   lea edx, var_28
  loc_005A6AE1: push 00000024h
loc_005A6AE3:   push edx
  loc_005A6AE4: call [004010B4h] ; rtcMsgBox
  loc_005A6AEA: xor ebx, ebx
  loc_005A6AEC: cmp eax, 00000007h
loc_005A6AEF:   lea eax, var_58
loc_005A6AF2:   lea ecx, var_48
loc_005A6AF5:   push eax
loc_005A6AF6:   lea edx, var_38
loc_005A6AF9:   push ecx
loc_005A6AFA:   lea eax, var_28
loc_005A6AFD:   push edx
loc_005A6AFE:   push eax
loc_005A6AFF:   setz bl
  loc_005A6B02: push 00000004h
loc_005A6B04:   neg ebx
  loc_005A6B06: call [0040103Ch] ; __vbaFreeVarList
  loc_005A6B0C: add esp, 00000014h
loc_005A6B0F:   cmp bx, si
  loc_005A6B12: jz 005A6B1Eh
loc_005A6B14:   mov ecx, Cancel
  loc_005A6B17: mov [ecx], 0001h
  loc_005A6B1C: jmp 005A6B78h
loc_005A6B1E:   cmp [0066A318h], esi
  loc_005A6B24: jnz 005A6B36h
  loc_005A6B26: push 0066A318h
  loc_005A6B2B: push 0042E390h
  loc_005A6B30: call [004011E8h] ; __vbaNew2
loc_005A6B36:   mov ebx, [0066A318h]
loc_005A6B3C:   lea eax, var_18
loc_005A6B3F:   push edi
loc_005A6B40:   push eax
loc_005A6B41:   mov edx, [ebx]
loc_005A6B43:   mov var_AC, edx
  loc_005A6B49: call [004010C8h] ; __vbaObjSetAddref
loc_005A6B4F:   mov ecx, var_AC
loc_005A6B55:   push eax
loc_005A6B56:   push ebx
loc_005A6B57:   Call [ecx+00000010h]
loc_005A6B5A:   cmp eax, esi
loc_005A6B5C:   fnclex
  loc_005A6B5E: jge 005A6B6Fh
  loc_005A6B60: push 00000010h
  loc_005A6B62: push 0042E380h
loc_005A6B67:   push ebx
loc_005A6B68:   push eax
  loc_005A6B69: call [00401080h] ; __vbaHresultCheckObj
loc_005A6B6F:   lea ecx, var_18
  loc_005A6B72: call [004012A0h] ; __vbaFreeObj
loc_005A6B78:   mov var_4, esi
  loc_005A6B7B: push 005A6BA8h
  loc_005A6B80: jmp 005A6BA7h
loc_005A6B82:   lea ecx, var_18
  loc_005A6B85: call [004012A0h] ; __vbaFreeObj
loc_005A6B8B:   lea edx, var_58
loc_005A6B8E:   lea eax, var_48
loc_005A6B91:   push edx
loc_005A6B92:   lea ecx, var_38
loc_005A6B95:   push eax
loc_005A6B96:   lea edx, var_28
loc_005A6B99:   push ecx
loc_005A6B9A:   push edx
  loc_005A6B9B: push 00000004h
  loc_005A6B9D: call [0040103Ch] ; __vbaFreeVarList
  loc_005A6BA3: add esp, 00000014h
loc_005A6BA6:   ret
loc_005A6BA7:   ret
loc_005A6BA8:   mov eax, Me
loc_005A6BAB:   push eax
loc_005A6BAC:   mov ecx, [eax]
loc_005A6BAE:   Call [ecx+00000008h]
loc_005A6BB1:   mov eax, var_4
loc_005A6BB4:   mov ecx, var_14
loc_005A6BB7:   pop edi
loc_005A6BB8:   pop esi
loc_005A6BB9:   mov fs: [00000000h] , ecx
loc_005A6BC0:   pop ebx
loc_005A6BC1:   mov esp, ebp
loc_005A6BC3:   pop ebp
  loc_005A6BC4: retn 0008h
End Sub

Private Sub GridJual_UnknownEvent_10() '5AEEF0
loc_005AEEF0:   push ebp
loc_005AEEF1:   mov ebp, esp
  loc_005AEEF3: sub esp, 0000000Ch
  loc_005AEEF6: push 00405696h ; __vbaExceptHandler
loc_005AEEFB:   mov eax, fs: [00000000h]
loc_005AEF01:   push eax
loc_005AEF02:   mov fs: [00000000h] , esp
  loc_005AEF09: sub esp, 00000068h
loc_005AEF0C:   push ebx
loc_005AEF0D:   push esi
loc_005AEF0E:   push edi
loc_005AEF0F:   mov var_C, esp
  loc_005AEF12: mov var_8, 00402F68h
loc_005AEF19:   mov esi, Me
loc_005AEF1C:   mov eax, esi
  loc_005AEF1E: and eax, 00000001h
loc_005AEF21:   mov var_4, eax
  loc_005AEF24: and esi, FFFFFFFEh
loc_005AEF27:   push esi
loc_005AEF28:   mov Me, esi
loc_005AEF2B:   mov ecx, [esi]
loc_005AEF2D:   Call [ecx+00000004h]
loc_005AEF30:   mov edx, Cancel
  loc_005AEF33: xor eax, eax
loc_005AEF35:   mov var_18, eax
loc_005AEF38:   mov var_28, eax
  loc_005AEF3B: cmp [edx], 0002h
loc_005AEF3F:   mov var_38, eax
loc_005AEF42:   mov var_48, eax
loc_005AEF45:   mov var_58, eax
loc_005AEF48:   mov var_68, eax
  loc_005AEF4B: jnz 005AF228h
loc_005AEF51:   push eax
loc_005AEF52:   mov eax, [esi]
  loc_005AEF54: push 00000004h
loc_005AEF56:   push esi
loc_005AEF57:   Call [eax+00000390h]
  loc_005AEF5D: mov edi, [004010B8h] ; __vbaObjSet
loc_005AEF63:   lea ecx, var_18
loc_005AEF66:   push eax
loc_005AEF67:   push ecx
loc_005AEF68:   Call edi
  loc_005AEF6A: mov ebx, [0040114Ch] ; __vbaLateIdCallLd
loc_005AEF70:   lea edx, var_28
loc_005AEF73:   push eax
loc_005AEF74:   push edx
loc_005AEF75:   Call ebx
  loc_005AEF77: add esp, 00000010h
loc_005AEF7A:   push eax
  loc_005AEF7B: call [0040121Ch] ; __vbaI4Var
  loc_005AEF81: xor ecx, ecx
  loc_005AEF83: cmp eax, 00000001h
loc_005AEF86:   setl cl
loc_005AEF89:   neg ecx
loc_005AEF8B:   mov var_6C, cx
loc_005AEF8F:   lea ecx, var_18
  loc_005AEF92: call [004012A0h] ; __vbaFreeObj
loc_005AEF98:   lea ecx, var_28
  loc_005AEF9B: call [00401020h] ; __vbaFreeVar
  loc_005AEFA1: cmp var_6C, 0000h
  loc_005AEFA6: jz 005AF04Bh
loc_005AEFAC:   mov eax, [00668264h]
loc_005AEFB1:   test eax, eax
  loc_005AEFB3: jnz 005AEFCAh
  loc_005AEFB5: push 00668264h
  loc_005AEFBA: push 0040896Ch
  loc_005AEFBF: call [004011E8h] ; __vbaNew2
loc_005AEFC5:   mov eax, [00668264h]
loc_005AEFCA:   mov edx, [eax]
loc_005AEFCC:   push eax
loc_005AEFCD:   Call [edx+00000308h]
loc_005AEFD3:   push eax
loc_005AEFD4:   lea eax, var_18
loc_005AEFD7:   push eax
loc_005AEFD8:   Call edi
loc_005AEFDA:   mov esi, eax
loc_005AEFDC:   push FFFFFFFFh
loc_005AEFDE:   push esi
loc_005AEFDF:   mov ecx, [esi]
loc_005AEFE1:   Call [ecx+00000074h]
loc_005AEFE4:   test eax, eax
loc_005AEFE6:   fnclex
  loc_005AEFE8: jge 005AEFF9h
  loc_005AEFEA: push 00000074h
  loc_005AEFEC: push 00434274h
loc_005AEFF1:   push esi
loc_005AEFF2:   push eax
  loc_005AEFF3: call [00401080h] ; __vbaHresultCheckObj
  loc_005AEFF9: mov ebx, [004012A0h] ; __vbaFreeObj
loc_005AEFFF:   lea ecx, var_18
loc_005AF002:   Call ebx
loc_005AF004:   mov eax, [00668264h]
loc_005AF009:   test eax, eax
  loc_005AF00B: jnz 005AF022h
  loc_005AF00D: push 00668264h
  loc_005AF012: push 0040896Ch
  loc_005AF017: call [004011E8h] ; __vbaNew2
loc_005AF01D:   mov eax, [00668264h]
loc_005AF022:   mov edx, [eax]
loc_005AF024:   push eax
loc_005AF025:   Call [edx+00000308h]
loc_005AF02B:   push eax
loc_005AF02C:   lea eax, var_18
loc_005AF02F:   push eax
loc_005AF030:   Call edi
loc_005AF032:   mov esi, eax
loc_005AF034:   push FFFFFFFFh
loc_005AF036:   push esi
loc_005AF037:   mov ecx, [esi]
loc_005AF039:   Call [ecx+00000074h]
loc_005AF03C:   test eax, eax
loc_005AF03E:   fnclex
  loc_005AF040: jge 005AF139h
  loc_005AF046: jmp 005AF12Ah
loc_005AF04B:   mov edx, [esi]
  loc_005AF04D: push 00000000h
  loc_005AF04F: push 0000000Ah
loc_005AF051:   push esi
loc_005AF052:   Call [edx+00000390h]
loc_005AF058:   push eax
loc_005AF059:   lea eax, var_18
loc_005AF05C:   push eax
loc_005AF05D:   Call edi
loc_005AF05F:   lea ecx, var_28
loc_005AF062:   push eax
loc_005AF063:   push ecx
loc_005AF064:   Call ebx
  loc_005AF066: add esp, 00000010h
loc_005AF069:   push eax
  loc_005AF06A: call [0040121Ch] ; __vbaI4Var
  loc_005AF070: mov ebx, [004012A0h] ; __vbaFreeObj
  loc_005AF076: xor edx, edx
  loc_005AF078: cmp eax, 00000001h
loc_005AF07B:   lea ecx, var_18
loc_005AF07E:   setge dl
loc_005AF081:   neg edx
loc_005AF083:   mov si, dx
loc_005AF086:   Call ebx
loc_005AF088:   lea ecx, var_28
  loc_005AF08B: call [00401020h] ; __vbaFreeVar
loc_005AF091:   test si, si
  loc_005AF094: jz 005AF13Eh
loc_005AF09A:   mov eax, [00668264h]
loc_005AF09F:   test eax, eax
  loc_005AF0A1: jnz 005AF0B8h
  loc_005AF0A3: push 00668264h
  loc_005AF0A8: push 0040896Ch
  loc_005AF0AD: call [004011E8h] ; __vbaNew2
loc_005AF0B3:   mov eax, [00668264h]
loc_005AF0B8:   mov ecx, [eax]
loc_005AF0BA:   push eax
loc_005AF0BB:   Call [ecx+00000308h]
loc_005AF0C1:   lea edx, var_18
loc_005AF0C4:   push eax
loc_005AF0C5:   push edx
loc_005AF0C6:   Call edi
loc_005AF0C8:   mov esi, eax
loc_005AF0CA:   push FFFFFFFFh
loc_005AF0CC:   push esi
loc_005AF0CD:   mov eax, [esi]
loc_005AF0CF:   Call [eax+00000074h]
loc_005AF0D2:   test eax, eax
loc_005AF0D4:   fnclex
  loc_005AF0D6: jge 005AF0E7h
  loc_005AF0D8: push 00000074h
  loc_005AF0DA: push 00434274h
loc_005AF0DF:   push esi
loc_005AF0E0:   push eax
  loc_005AF0E1: call [00401080h] ; __vbaHresultCheckObj
loc_005AF0E7:   lea ecx, var_18
loc_005AF0EA:   Call ebx
loc_005AF0EC:   mov eax, [00668264h]
loc_005AF0F1:   test eax, eax
  loc_005AF0F3: jnz 005AF10Ah
  loc_005AF0F5: push 00668264h
  loc_005AF0FA: push 0040896Ch
  loc_005AF0FF: call [004011E8h] ; __vbaNew2
loc_005AF105:   mov eax, [00668264h]
loc_005AF10A:   mov ecx, [eax]
loc_005AF10C:   push eax
loc_005AF10D:   Call [ecx+00000308h]
loc_005AF113:   lea edx, var_18
loc_005AF116:   push eax
loc_005AF117:   push edx
loc_005AF118:   Call edi
loc_005AF11A:   mov esi, eax
loc_005AF11C:   push FFFFFFFFh
loc_005AF11E:   push esi
loc_005AF11F:   mov eax, [esi]
loc_005AF121:   Call [eax+00000074h]
loc_005AF124:   test eax, eax
loc_005AF126:   fnclex
  loc_005AF128: jge 005AF139h
  loc_005AF12A: push 00000074h
  loc_005AF12C: push 00434274h
loc_005AF131:   push esi
loc_005AF132:   push eax
  loc_005AF133: call [00401080h] ; __vbaHresultCheckObj
loc_005AF139:   lea ecx, var_18
loc_005AF13C:   Call ebx
loc_005AF13E:   mov eax, [00668264h]
  loc_005AF143: mov ebx, 80020004h
  loc_005AF148: mov edi, 0000000Ah
loc_005AF14D:   mov esi, ebx
loc_005AF14F:   test eax, eax
loc_005AF151:   mov var_30, ebx
loc_005AF154:   mov var_38, edi
  loc_005AF157: jnz 005AF16Eh
  loc_005AF159: push 00668264h
  loc_005AF15E: push 0040896Ch
  loc_005AF163: call [004011E8h] ; __vbaNew2
loc_005AF169:   mov eax, [00668264h]
loc_005AF16E:   mov ecx, Me
  loc_005AF171: sub esp, 00000010h
loc_005AF174:   mov edx, [ecx]
loc_005AF176:   mov ecx, esp
  loc_005AF178: sub esp, 00000010h
loc_005AF17B:   mov var_7C, edx
loc_005AF17E:   mov [ecx], edi
loc_005AF180:   mov edi, var_64
loc_005AF183:   mov [ecx+00000004h], edi
loc_005AF186:   mov [ecx+00000008h], esi
loc_005AF189:   mov esi, var_5C
loc_005AF18C:   mov [ecx+0000000Ch], esi
loc_005AF18F:   mov esi, esp
  loc_005AF191: mov ecx, 0000000Ah
  loc_005AF196: sub esp, 00000010h
loc_005AF199:   mov [esi], ecx
loc_005AF19B:   mov ecx, var_54
loc_005AF19E:   mov [esi+00000004h], ecx
loc_005AF1A1:   mov ecx, var_4C
loc_005AF1A4:   mov [esi+00000008h], ebx
loc_005AF1A7:   mov [esi+0000000Ch], ecx
loc_005AF1AA:   mov esi, esp
  loc_005AF1AC: mov ecx, 0000000Ah
  loc_005AF1B1: sub esp, 00000010h
loc_005AF1B4:   mov [esi], ecx
loc_005AF1B6:   mov ecx, var_44
loc_005AF1B9:   mov [esi+00000004h], ecx
  loc_005AF1BC: mov ecx, 80020004h
loc_005AF1C1:   mov [esi+00000008h], ecx
loc_005AF1C4:   mov ecx, var_3C
loc_005AF1C7:   mov [esi+0000000Ch], ecx
loc_005AF1CA:   mov esi, var_38
loc_005AF1CD:   mov ecx, esp
loc_005AF1CF:   push eax
loc_005AF1D0:   mov [ecx], esi
loc_005AF1D2:   mov esi, var_34
loc_005AF1D5:   mov [ecx+00000004h], esi
loc_005AF1D8:   mov esi, var_30
loc_005AF1DB:   mov [ecx+00000008h], esi
loc_005AF1DE:   mov esi, var_2C
loc_005AF1E1:   mov [ecx+0000000Ch], esi
loc_005AF1E4:   mov ecx, [eax]
loc_005AF1E6:   Call [ecx+00000304h]
loc_005AF1EC:   lea edx, var_18
loc_005AF1EF:   push eax
loc_005AF1F0:   push edx
  loc_005AF1F1: call [004010B8h] ; __vbaObjSet
loc_005AF1F7:   mov esi, Me
loc_005AF1FA:   push eax
loc_005AF1FB:   mov eax, var_7C
loc_005AF1FE:   push esi
loc_005AF1FF:   Call [eax+000002BCh]
loc_005AF205:   test eax, eax
loc_005AF207:   fnclex
  loc_005AF209: jge 005AF21Dh
  loc_005AF20B: push 000002BCh
  loc_005AF210: push 00434B48h
loc_005AF215:   push esi
loc_005AF216:   push eax
  loc_005AF217: call [00401080h] ; __vbaHresultCheckObj
loc_005AF21D:   lea ecx, var_18
  loc_005AF220: call [004012A0h] ; __vbaFreeObj
  loc_005AF226: xor eax, eax
loc_005AF228:   mov var_4, eax
  loc_005AF22B: push 005AF246h
  loc_005AF230: jmp 005AF245h
loc_005AF232:   lea ecx, var_18
  loc_005AF235: call [004012A0h] ; __vbaFreeObj
loc_005AF23B:   lea ecx, var_28
  loc_005AF23E: call [00401020h] ; __vbaFreeVar
loc_005AF244:   ret
loc_005AF245:   ret
loc_005AF246:   mov eax, Me
loc_005AF249:   push eax
loc_005AF24A:   mov ecx, [eax]
loc_005AF24C:   Call [ecx+00000008h]
loc_005AF24F:   mov eax, var_4
loc_005AF252:   mov ecx, var_14
loc_005AF255:   pop edi
loc_005AF256:   pop esi
loc_005AF257:   mov fs: [00000000h] , ecx
loc_005AF25E:   pop ebx
loc_005AF25F:   mov esp, ebp
loc_005AF261:   pop ebp
  loc_005AF262: retn 0014h
End Sub

Private Sub txtNoNota_KeyPress(KeyAscii As Integer) '5B14B0
loc_005B14B0:   push ebp
loc_005B14B1:   mov ebp, esp
  loc_005B14B3: sub esp, 0000000Ch
  loc_005B14B6: push 00405696h ; __vbaExceptHandler
loc_005B14BB:   mov eax, fs: [00000000h]
loc_005B14C1:   push eax
loc_005B14C2:   mov fs: [00000000h] , esp
  loc_005B14C9: sub esp, 00000028h
loc_005B14CC:   push ebx
loc_005B14CD:   push esi
loc_005B14CE:   push edi
loc_005B14CF:   mov var_C, esp
  loc_005B14D2: mov var_8, 00402FC8h
loc_005B14D9:   mov eax, Me
loc_005B14DC:   mov ecx, eax
  loc_005B14DE: and ecx, 00000001h
loc_005B14E1:   mov var_4, ecx
  loc_005B14E4: and al, FEh
loc_005B14E6:   push eax
loc_005B14E7:   mov Me, eax
loc_005B14EA:   mov edx, [eax]
loc_005B14EC:   Call [edx+00000004h]
loc_005B14EF:   mov esi, KeyAscii
  loc_005B14F2: xor edi, edi
loc_005B14F4:   mov var_24, edi
  loc_005B14F7: cmp [esi], 000Dh
  loc_005B14FB: jnz 005B1526h
loc_005B14FD:   lea eax, var_24
  loc_005B1500: mov var_1C, 80020004h
loc_005B1507:   push eax
  loc_005B1508: push 004300CCh ; "{tab}"
  loc_005B150D: mov var_24, 0000000Ah
  loc_005B1514: call [004010D4h] ; rtcSendKeys
loc_005B151A:   lea ecx, var_24
  loc_005B151D: call [00401020h] ; __vbaFreeVar
loc_005B1523:   mov [esi], di
loc_005B1526:   mov var_4, edi
  loc_005B1529: push 005B153Bh
  loc_005B152E: jmp 005B153Ah
loc_005B1530:   lea ecx, var_24
  loc_005B1533: call [00401020h] ; __vbaFreeVar
loc_005B1539:   ret
loc_005B153A:   ret
loc_005B153B:   mov eax, Me
loc_005B153E:   push eax
loc_005B153F:   mov ecx, [eax]
loc_005B1541:   Call [ecx+00000008h]
loc_005B1544:   mov eax, var_4
loc_005B1547:   mov ecx, var_14
loc_005B154A:   pop edi
loc_005B154B:   pop esi
loc_005B154C:   mov fs: [00000000h] , ecx
loc_005B1553:   pop ebx
loc_005B1554:   mov esp, ebp
loc_005B1556:   pop ebp
  loc_005B1557: retn 0008h
End Sub

Private Sub cmbPelanggan_UnknownEvent_C() '5A8D90
loc_005A8D90:   push ebp
loc_005A8D91:   mov ebp, esp
  loc_005A8D93: sub esp, 0000000Ch
  loc_005A8D96: push 00405696h ; __vbaExceptHandler
loc_005A8D9B:   mov eax, fs: [00000000h]
loc_005A8DA1:   push eax
loc_005A8DA2:   mov fs: [00000000h] , esp
  loc_005A8DA9: sub esp, 000000E8h
loc_005A8DAF:   push ebx
loc_005A8DB0:   push esi
loc_005A8DB1:   push edi
loc_005A8DB2:   mov var_C, esp
  loc_005A8DB5: mov var_8, 00402DE0h
loc_005A8DBC:   mov esi, Me
loc_005A8DBF:   mov eax, esi
  loc_005A8DC1: and eax, 00000001h
loc_005A8DC4:   mov var_4, eax
  loc_005A8DC7: and esi, FFFFFFFEh
loc_005A8DCA:   push esi
loc_005A8DCB:   mov Me, esi
loc_005A8DCE:   mov ecx, [esi]
loc_005A8DD0:   Call [ecx+00000004h]
  loc_005A8DD3: xor edi, edi
  loc_005A8DD5: push 0042DF30h
loc_005A8DDA:   mov var_18, edi
loc_005A8DDD:   mov var_1C, edi
loc_005A8DE0:   mov var_20, edi
loc_005A8DE3:   mov var_24, edi
loc_005A8DE6:   mov var_34, edi
loc_005A8DE9:   mov var_44, edi
loc_005A8DEC:   mov var_54, edi
loc_005A8DEF:   mov var_64, edi
loc_005A8DF2:   mov var_74, edi
loc_005A8DF5:   mov var_84, edi
loc_005A8DFB:   mov var_94, edi
loc_005A8E01:   mov var_A4, edi
loc_005A8E07:   mov var_C8, edi
loc_005A8E0D:   mov var_CC, edi
loc_005A8E13:   mov var_EC, edi
  loc_005A8E19: call [00401168h] ; __vbaNew
  loc_005A8E1F: mov ebx, [004010B8h] ; __vbaObjSet
loc_005A8E25:   push eax
  loc_005A8E26: push 0066805Ch
loc_005A8E2B:   Call ebx
loc_005A8E2D:   mov edx, [esi]
loc_005A8E2F:   push edi
loc_005A8E30:   push FFFFFDFBh
loc_005A8E35:   push esi
loc_005A8E36:   Call [edx+0000038Ch]
loc_005A8E3C:   push eax
loc_005A8E3D:   lea eax, var_1C
loc_005A8E40:   push eax
loc_005A8E41:   Call ebx
loc_005A8E43:   lea ecx, var_34
loc_005A8E46:   push eax
loc_005A8E47:   push ecx
  loc_005A8E48: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005A8E4E: add esp, 00000010h
loc_005A8E51:   lea edx, var_34
loc_005A8E54:   push edx
  loc_005A8E55: call [00401034h] ; __vbaStrVarMove
loc_005A8E5B:   mov var_3C, eax
loc_005A8E5E:   lea eax, var_44
  loc_005A8E61: push 00000006h
loc_005A8E63:   lea ecx, var_54
loc_005A8E66:   push eax
loc_005A8E67:   push ecx
  loc_005A8E68: mov var_44, 00000008h
  loc_005A8E6F: call [00401260h] ; rtcLeftCharVar
loc_005A8E75:   cmp [0066805Ch], edi
  loc_005A8E7B: jnz 005A8E8Dh
  loc_005A8E7D: push 0066805Ch
  loc_005A8E82: push 0042DF30h
  loc_005A8E87: call [004011E8h] ; __vbaNew2
loc_005A8E8D:   mov eax, [0066803Ch]
loc_005A8E92:   mov ebx, [0066805Ch]
loc_005A8E98:   cmp eax, edi
loc_005A8E9A:   mov var_D0, ebx
  loc_005A8EA0: jnz 005A8EB2h
  loc_005A8EA2: push 0066803Ch
  loc_005A8EA7: push 0042DEFCh
  loc_005A8EAC: call [004011E8h] ; __vbaNew2
loc_005A8EB2:   mov edx, [0066803Ch]
  loc_005A8EB8: push 00435BBCh ; "SELECT * FROM Pelanggan WHERE "
  loc_005A8EBD: push 00435C00h ; " kode_pelanggan ='"
loc_005A8EC2:   mov var_9C, edx
  loc_005A8EC8: mov var_A4, 00000009h
  loc_005A8ED2: call [00401070h] ; __vbaStrCat
loc_005A8ED8:   mov ecx, var_A4
loc_005A8EDE:   push FFFFFFFFh
  loc_005A8EE0: push 00000004h
loc_005A8EE2:   mov var_5C, eax
  loc_005A8EE5: push 00000002h
loc_005A8EE7:   mov edx, var_A0
  loc_005A8EED: mov eax, 00000008h
  loc_005A8EF2: sub esp, 00000010h
loc_005A8EF5:   mov var_64, eax
loc_005A8EF8:   mov var_94, eax
loc_005A8EFE:   mov eax, esp
  loc_005A8F00: mov var_8C, 0042EBA8h ; "'"
loc_005A8F0A:   mov ebx, [ebx]
loc_005A8F0C:   mov [eax], ecx
loc_005A8F0E:   mov ecx, var_9C
loc_005A8F14:   mov [eax+00000004h], edx
loc_005A8F17:   mov edx, var_98
loc_005A8F1D:   mov [eax+00000008h], ecx
loc_005A8F20:   lea ecx, var_54
loc_005A8F23:   mov [eax+0000000Ch], edx
loc_005A8F26:   lea eax, var_64
loc_005A8F29:   push eax
loc_005A8F2A:   lea edx, var_74
loc_005A8F2D:   push ecx
loc_005A8F2E:   push edx
  loc_005A8F2F: call [004011C4h] ; __vbaVarCat
loc_005A8F35:   push eax
loc_005A8F36:   lea eax, var_94
loc_005A8F3C:   lea ecx, var_84
loc_005A8F42:   push eax
loc_005A8F43:   push ecx
  loc_005A8F44: call [004011C4h] ; __vbaVarCat
loc_005A8F4A:   mov ecx, [eax]
  loc_005A8F4C: sub esp, 00000010h
loc_005A8F4F:   mov edx, esp
loc_005A8F51:   mov var_F8, ebx
loc_005A8F57:   mov ebx, var_D0
loc_005A8F5D:   mov [edx], ecx
loc_005A8F5F:   mov ecx, [eax+00000004h]
loc_005A8F62:   push ebx
loc_005A8F63:   mov [edx+00000004h], ecx
loc_005A8F66:   mov ecx, [eax+00000008h]
loc_005A8F69:   mov eax, [eax+0000000Ch]
loc_005A8F6C:   mov [edx+00000008h], ecx
loc_005A8F6F:   mov ecx, var_F8
loc_005A8F75:   mov [edx+0000000Ch], eax
loc_005A8F78:   Call [ecx+000000A0h]
loc_005A8F7E:   cmp eax, edi
loc_005A8F80:   fnclex
  loc_005A8F82: jge 005A8F9Ah
  loc_005A8F84: push 000000A0h
  loc_005A8F89: push 0042DF44h
loc_005A8F8E:   push ebx
  loc_005A8F8F: mov ebx, [00401080h] ; __vbaHresultCheckObj
loc_005A8F95:   push eax
loc_005A8F96:   Call ebx
  loc_005A8F98: jmp 005A8FA0h
  loc_005A8F9A: mov ebx, [00401080h] ; __vbaHresultCheckObj
loc_005A8FA0:   lea ecx, var_1C
  loc_005A8FA3: call [004012A0h] ; __vbaFreeObj
loc_005A8FA9:   lea edx, var_84
loc_005A8FAF:   lea eax, var_74
loc_005A8FB2:   push edx
loc_005A8FB3:   lea ecx, var_54
loc_005A8FB6:   push eax
loc_005A8FB7:   lea edx, var_64
loc_005A8FBA:   push ecx
loc_005A8FBB:   lea eax, var_44
loc_005A8FBE:   push edx
loc_005A8FBF:   lea ecx, var_34
loc_005A8FC2:   push eax
loc_005A8FC3:   push ecx
  loc_005A8FC4: push 00000006h
  loc_005A8FC6: call [0040103Ch] ; __vbaFreeVarList
loc_005A8FCC:   mov eax, [0066805Ch]
  loc_005A8FD1: add esp, 0000001Ch
loc_005A8FD4:   cmp eax, edi
  loc_005A8FD6: jnz 005A8FE8h
  loc_005A8FD8: push 0066805Ch
  loc_005A8FDD: push 0042DF30h
  loc_005A8FE2: call [004011E8h] ; __vbaNew2
loc_005A8FE8:   mov edx, [0066805Ch]
loc_005A8FEE:   lea eax, var_EC
loc_005A8FF4:   push edx
loc_005A8FF5:   push eax
  loc_005A8FF6: call [004010C8h] ; __vbaObjSetAddref
loc_005A8FFC:   mov eax, var_EC
loc_005A9002:   lea edx, var_C8
loc_005A9008:   push edx
loc_005A9009:   push eax
loc_005A900A:   mov ecx, [eax]
loc_005A900C:   Call [ecx+00000050h]
loc_005A900F:   cmp eax, edi
loc_005A9011:   fnclex
  loc_005A9013: jge 005A9026h
loc_005A9015:   mov ecx, var_EC
  loc_005A901B: push 00000050h
  loc_005A901D: push 0042DF44h
loc_005A9022:   push ecx
loc_005A9023:   push eax
loc_005A9024:   Call ebx
loc_005A9026:   mov eax, var_EC
loc_005A902C:   lea ecx, var_CC
loc_005A9032:   push ecx
loc_005A9033:   push eax
loc_005A9034:   mov edx, [eax]
loc_005A9036:   Call [edx+00000034h]
loc_005A9039:   cmp eax, edi
loc_005A903B:   fnclex
  loc_005A903D: jge 005A9050h
loc_005A903F:   mov edx, var_EC
  loc_005A9045: push 00000034h
  loc_005A9047: push 0042DF44h
loc_005A904C:   push edx
loc_005A904D:   push eax
loc_005A904E:   Call ebx
  loc_005A9050: xor eax, eax
loc_005A9052:   cmp var_CC, di
loc_005A9059:   setz al
  loc_005A905C: xor ecx, ecx
loc_005A905E:   cmp var_C8, di
loc_005A9065:   setz cl
  loc_005A9068: or eax, ecx
  loc_005A906A: jnz 005A90FFh
  loc_005A9070: mov esi, [00401238h] ; __vbaVarDup
  loc_005A9076: mov ecx, 80020004h
loc_005A907B:   mov var_5C, ecx
  loc_005A907E: mov eax, 0000000Ah
loc_005A9083:   mov var_4C, ecx
  loc_005A9086: mov ebx, 00000008h
loc_005A908B:   lea edx, var_A4
loc_005A9091:   lea ecx, var_44
loc_005A9094:   mov var_64, eax
loc_005A9097:   mov var_54, eax
  loc_005A909A: mov var_9C, 0042E624h ; "Error"
loc_005A90A4:   mov var_A4, ebx
  loc_005A90AA: call __vbaVarDup
loc_005A90AC:   lea edx, var_94
loc_005A90B2:   lea ecx, var_34
  loc_005A90B5: mov var_8C, 00435C2Ch ; "KODE PELANGGAN TIDAK DITEMUKAN"
loc_005A90BF:   mov var_94, ebx
  loc_005A90C5: call __vbaVarDup
loc_005A90C7:   lea edx, var_64
loc_005A90CA:   lea eax, var_54
loc_005A90CD:   push edx
loc_005A90CE:   lea ecx, var_44
loc_005A90D1:   push eax
loc_005A90D2:   push ecx
loc_005A90D3:   lea edx, var_34
  loc_005A90D6: push 00000010h
loc_005A90D8:   push edx
  loc_005A90D9: call [004010B4h] ; rtcMsgBox
loc_005A90DF:   lea eax, var_64
loc_005A90E2:   lea ecx, var_54
loc_005A90E5:   push eax
loc_005A90E6:   lea edx, var_44
loc_005A90E9:   push ecx
loc_005A90EA:   lea eax, var_34
loc_005A90ED:   push edx
loc_005A90EE:   push eax
  loc_005A90EF: push 00000004h
  loc_005A90F1: call [0040103Ch] ; __vbaFreeVarList
  loc_005A90F7: add esp, 00000014h
  loc_005A90FA: jmp 005A93A9h
loc_005A90FF:   mov eax, var_EC
loc_005A9105:   lea edx, var_1C
loc_005A9108:   push edx
loc_005A9109:   push eax
loc_005A910A:   mov ecx, [eax]
loc_005A910C:   Call [ecx+00000054h]
loc_005A910F:   cmp eax, edi
loc_005A9111:   fnclex
  loc_005A9113: jge 005A9126h
loc_005A9115:   mov ecx, var_EC
  loc_005A911B: push 00000054h
  loc_005A911D: push 0042DF44h
loc_005A9122:   push ecx
loc_005A9123:   push eax
loc_005A9124:   Call ebx
loc_005A9126:   lea ebx, var_20
loc_005A9129:   mov eax, var_1C
loc_005A912C:   push ebx
  loc_005A912D: mov ecx, 00000008h
  loc_005A9132: sub esp, 00000010h
loc_005A9135:   mov var_94, ecx
loc_005A913B:   mov ebx, esp
  loc_005A913D: mov var_8C, 0042F1B8h ; "kode_Pelanggan"
loc_005A9147:   mov edx, [eax]
loc_005A9149:   push eax
loc_005A914A:   mov [ebx], ecx
loc_005A914C:   mov ecx, var_90
loc_005A9152:   mov var_D4, eax
loc_005A9158:   mov [ebx+00000004h], ecx
loc_005A915B:   mov ecx, var_8C
loc_005A9161:   mov [ebx+00000008h], ecx
loc_005A9164:   mov ecx, var_88
loc_005A916A:   mov [ebx+0000000Ch], ecx
loc_005A916D:   Call [edx+00000028h]
loc_005A9170:   cmp eax, edi
loc_005A9172:   fnclex
  loc_005A9174: jge 005A918Bh
loc_005A9176:   mov edx, var_D4
  loc_005A917C: push 00000028h
  loc_005A917E: push 0042DFACh
loc_005A9183:   push edx
loc_005A9184:   push eax
  loc_005A9185: call [00401080h] ; __vbaHresultCheckObj
loc_005A918B:   mov eax, var_20
loc_005A918E:   lea edx, var_34
loc_005A9191:   push edx
loc_005A9192:   push eax
loc_005A9193:   mov ecx, [eax]
loc_005A9195:   mov ebx, eax
loc_005A9197:   Call [ecx+00000034h]
loc_005A919A:   cmp eax, edi
loc_005A919C:   fnclex
  loc_005A919E: jge 005A91AFh
  loc_005A91A0: push 00000034h
  loc_005A91A2: push 0042DFBCh
loc_005A91A7:   push ebx
loc_005A91A8:   push eax
  loc_005A91A9: call [00401080h] ; __vbaHresultCheckObj
loc_005A91AF:   lea eax, var_34
loc_005A91B2:   push eax
  loc_005A91B3: call [00401034h] ; __vbaStrVarMove
  loc_005A91B9: sub esp, 00000010h
  loc_005A91BC: mov ecx, 00000008h
loc_005A91C1:   mov edx, esp
loc_005A91C3:   mov var_44, ecx
loc_005A91C6:   mov var_3C, eax
loc_005A91C9:   push FFFFFDFBh
loc_005A91CE:   mov [edx], ecx
loc_005A91D0:   mov ecx, var_40
loc_005A91D3:   push esi
loc_005A91D4:   mov [edx+00000004h], ecx
loc_005A91D7:   mov ecx, [esi]
loc_005A91D9:   mov [edx+00000008h], eax
loc_005A91DC:   mov eax, var_38
loc_005A91DF:   mov [edx+0000000Ch], eax
loc_005A91E2:   Call [ecx+0000038Ch]
  loc_005A91E8: mov ebx, [004010B8h] ; __vbaObjSet
loc_005A91EE:   lea edx, var_24
loc_005A91F1:   push eax
loc_005A91F2:   push edx
loc_005A91F3:   Call ebx
loc_005A91F5:   push eax
  loc_005A91F6: call [00401280h] ; __vbaLateIdSt
loc_005A91FC:   lea eax, var_24
loc_005A91FF:   lea ecx, var_20
loc_005A9202:   push eax
loc_005A9203:   lea edx, var_1C
loc_005A9206:   push ecx
loc_005A9207:   push edx
  loc_005A9208: push 00000003h
  loc_005A920A: call [0040104Ch] ; __vbaFreeObjList
loc_005A9210:   lea eax, var_44
loc_005A9213:   lea ecx, var_34
loc_005A9216:   push eax
loc_005A9217:   push ecx
  loc_005A9218: push 00000002h
  loc_005A921A: call [0040103Ch] ; __vbaFreeVarList
loc_005A9220:   mov edx, [esi]
  loc_005A9222: add esp, 0000001Ch
loc_005A9225:   push esi
loc_005A9226:   Call [edx+00000310h]
loc_005A922C:   push eax
loc_005A922D:   lea eax, var_24
loc_005A9230:   push eax
loc_005A9231:   Call ebx
loc_005A9233:   mov var_E4, eax
loc_005A9239:   mov eax, var_EC
loc_005A923F:   lea edx, var_1C
loc_005A9242:   mov ecx, [eax]
loc_005A9244:   push edx
loc_005A9245:   push eax
loc_005A9246:   Call [ecx+00000054h]
loc_005A9249:   cmp eax, edi
loc_005A924B:   fnclex
  loc_005A924D: jge 005A9264h
loc_005A924F:   mov ecx, var_EC
  loc_005A9255: push 00000054h
  loc_005A9257: push 0042DF44h
loc_005A925C:   push ecx
loc_005A925D:   push eax
  loc_005A925E: call [00401080h] ; __vbaHresultCheckObj
loc_005A9264:   lea ebx, var_20
loc_005A9267:   mov eax, var_1C
loc_005A926A:   push ebx
  loc_005A926B: mov ecx, 00000008h
  loc_005A9270: sub esp, 00000010h
loc_005A9273:   mov var_94, ecx
loc_005A9279:   mov ebx, esp
  loc_005A927B: mov var_8C, 0042F358h ; "nama_pelanggan"
loc_005A9285:   mov edx, [eax]
loc_005A9287:   push eax
loc_005A9288:   mov [ebx], ecx
loc_005A928A:   mov ecx, var_90
loc_005A9290:   mov var_D4, eax
loc_005A9296:   mov [ebx+00000004h], ecx
loc_005A9299:   mov ecx, var_8C
loc_005A929F:   mov [ebx+00000008h], ecx
loc_005A92A2:   mov ecx, var_88
loc_005A92A8:   mov [ebx+0000000Ch], ecx
loc_005A92AB:   Call [edx+00000028h]
loc_005A92AE:   cmp eax, edi
loc_005A92B0:   fnclex
  loc_005A92B2: jge 005A92C9h
loc_005A92B4:   mov edx, var_D4
  loc_005A92BA: push 00000028h
  loc_005A92BC: push 0042DFACh
loc_005A92C1:   push edx
loc_005A92C2:   push eax
  loc_005A92C3: call [00401080h] ; __vbaHresultCheckObj
loc_005A92C9:   mov eax, var_20
loc_005A92CC:   lea edx, var_34
loc_005A92CF:   push edx
loc_005A92D0:   push eax
loc_005A92D1:   mov ecx, [eax]
loc_005A92D3:   mov ebx, eax
loc_005A92D5:   Call [ecx+00000034h]
loc_005A92D8:   cmp eax, edi
loc_005A92DA:   fnclex
  loc_005A92DC: jge 005A92EDh
  loc_005A92DE: push 00000034h
  loc_005A92E0: push 0042DFBCh
loc_005A92E5:   push ebx
loc_005A92E6:   push eax
  loc_005A92E7: call [00401080h] ; __vbaHresultCheckObj
loc_005A92ED:   mov eax, var_E4
loc_005A92F3:   lea ecx, var_34
loc_005A92F6:   push ecx
loc_005A92F7:   mov ebx, [eax]
  loc_005A92F9: call [00401034h] ; __vbaStrVarMove
loc_005A92FF:   mov edx, eax
loc_005A9301:   lea ecx, var_18
  loc_005A9304: call [0040126Ch] ; __vbaStrMove
loc_005A930A:   mov edx, ebx
loc_005A930C:   mov ebx, var_E4
loc_005A9312:   push eax
loc_005A9313:   push ebx
loc_005A9314:   Call [edx+000000A4h]
loc_005A931A:   cmp eax, edi
loc_005A931C:   fnclex
  loc_005A931E: jge 005A9332h
  loc_005A9320: push 000000A4h
  loc_005A9325: push 0042DFCCh
loc_005A932A:   push ebx
loc_005A932B:   push eax
  loc_005A932C: call [00401080h] ; __vbaHresultCheckObj
loc_005A9332:   lea ecx, var_18
  loc_005A9335: call [0040129Ch] ; __vbaFreeStr
loc_005A933B:   lea eax, var_24
loc_005A933E:   lea ecx, var_20
loc_005A9341:   push eax
loc_005A9342:   lea edx, var_1C
loc_005A9345:   push ecx
loc_005A9346:   push edx
  loc_005A9347: push 00000003h
  loc_005A9349: call [0040104Ch] ; __vbaFreeObjList
  loc_005A934F: add esp, 00000010h
loc_005A9352:   lea ecx, var_34
  loc_005A9355: call [00401020h] ; __vbaFreeVar
loc_005A935B:   lea eax, var_EC
loc_005A9361:   push edi
loc_005A9362:   push eax
  loc_005A9363: call [004010C8h] ; __vbaObjSetAddref
loc_005A9369:   mov ecx, [esi]
loc_005A936B:   push esi
loc_005A936C:   Call [ecx+00000348h]
loc_005A9372:   lea edx, var_1C
loc_005A9375:   push eax
loc_005A9376:   push edx
  loc_005A9377: call [004010B8h] ; __vbaObjSet
loc_005A937D:   mov esi, eax
loc_005A937F:   push esi
loc_005A9380:   mov eax, [esi]
loc_005A9382:   Call [eax+00000204h]
loc_005A9388:   cmp eax, edi
loc_005A938A:   fnclex
  loc_005A938C: jge 005A93A0h
  loc_005A938E: push 00000204h
  loc_005A9393: push 0042DFCCh
loc_005A9398:   push esi
loc_005A9399:   push eax
  loc_005A939A: call [00401080h] ; __vbaHresultCheckObj
loc_005A93A0:   lea ecx, var_1C
  loc_005A93A3: call [004012A0h] ; __vbaFreeObj
loc_005A93A9:   mov var_4, edi
  loc_005A93AC: push 005A9404h
  loc_005A93B1: jmp 005A93F7h
loc_005A93B3:   lea ecx, var_18
  loc_005A93B6: call [0040129Ch] ; __vbaFreeStr
loc_005A93BC:   lea ecx, var_24
loc_005A93BF:   lea edx, var_20
loc_005A93C2:   push ecx
loc_005A93C3:   lea eax, var_1C
loc_005A93C6:   push edx
loc_005A93C7:   push eax
  loc_005A93C8: push 00000003h
  loc_005A93CA: call [0040104Ch] ; __vbaFreeObjList
loc_005A93D0:   lea ecx, var_84
loc_005A93D6:   lea edx, var_74
loc_005A93D9:   push ecx
loc_005A93DA:   lea eax, var_64
loc_005A93DD:   push edx
loc_005A93DE:   lea ecx, var_54
loc_005A93E1:   push eax
loc_005A93E2:   lea edx, var_44
loc_005A93E5:   push ecx
loc_005A93E6:   lea eax, var_34
loc_005A93E9:   push edx
loc_005A93EA:   push eax
  loc_005A93EB: push 00000006h
  loc_005A93ED: call [0040103Ch] ; __vbaFreeVarList
  loc_005A93F3: add esp, 0000002Ch
loc_005A93F6:   ret
loc_005A93F7:   lea ecx, var_EC
  loc_005A93FD: call [004012A0h] ; __vbaFreeObj
loc_005A9403:   ret
loc_005A9404:   mov eax, Me
loc_005A9407:   push eax
loc_005A9408:   mov ecx, [eax]
loc_005A940A:   Call [ecx+00000008h]
loc_005A940D:   mov eax, var_4
loc_005A9410:   mov ecx, var_14
loc_005A9413:   pop edi
loc_005A9414:   pop esi
loc_005A9415:   mov fs: [00000000h] , ecx
loc_005A941C:   pop ebx
loc_005A941D:   mov esp, ebp
loc_005A941F:   pop ebp
  loc_005A9420: retn 0004h
End Sub

Private Sub txtPelanggan_KeyPress(KeyAscii As Integer) '5B1560
loc_005B1560:   push ebp
loc_005B1561:   mov ebp, esp
  loc_005B1563: sub esp, 0000000Ch
  loc_005B1566: push 00405696h ; __vbaExceptHandler
loc_005B156B:   mov eax, fs: [00000000h]
loc_005B1571:   push eax
loc_005B1572:   mov fs: [00000000h] , esp
  loc_005B1579: sub esp, 00000028h
loc_005B157C:   push ebx
loc_005B157D:   push esi
loc_005B157E:   push edi
loc_005B157F:   mov var_C, esp
  loc_005B1582: mov var_8, 00402FD8h
loc_005B1589:   mov eax, Me
loc_005B158C:   mov ecx, eax
  loc_005B158E: and ecx, 00000001h
loc_005B1591:   mov var_4, ecx
  loc_005B1594: and al, FEh
loc_005B1596:   push eax
loc_005B1597:   mov Me, eax
loc_005B159A:   mov edx, [eax]
loc_005B159C:   Call [edx+00000004h]
loc_005B159F:   mov esi, KeyAscii
  loc_005B15A2: xor edi, edi
loc_005B15A4:   mov var_24, edi
  loc_005B15A7: cmp [esi], 000Dh
  loc_005B15AB: jnz 005B15D6h
loc_005B15AD:   lea eax, var_24
  loc_005B15B0: mov var_1C, 80020004h
loc_005B15B7:   push eax
  loc_005B15B8: push 004300CCh ; "{tab}"
  loc_005B15BD: mov var_24, 0000000Ah
  loc_005B15C4: call [004010D4h] ; rtcSendKeys
loc_005B15CA:   lea ecx, var_24
  loc_005B15CD: call [00401020h] ; __vbaFreeVar
loc_005B15D3:   mov [esi], di
loc_005B15D6:   mov var_4, edi
  loc_005B15D9: push 005B15EBh
  loc_005B15DE: jmp 005B15EAh
loc_005B15E0:   lea ecx, var_24
  loc_005B15E3: call [00401020h] ; __vbaFreeVar
loc_005B15E9:   ret
loc_005B15EA:   ret
loc_005B15EB:   mov eax, Me
loc_005B15EE:   push eax
loc_005B15EF:   mov ecx, [eax]
loc_005B15F1:   Call [ecx+00000008h]
loc_005B15F4:   mov eax, var_4
loc_005B15F7:   mov ecx, var_14
loc_005B15FA:   pop edi
loc_005B15FB:   pop esi
loc_005B15FC:   mov fs: [00000000h] , ecx
loc_005B1603:   pop ebx
loc_005B1604:   mov esp, ebp
loc_005B1606:   pop ebp
  loc_005B1607: retn 0008h
End Sub

Private Sub txtNamaBarang_KeyPress(KeyAscii As Integer) '5B1610
loc_005B1610:   push ebp
loc_005B1611:   mov ebp, esp
  loc_005B1613: sub esp, 0000000Ch
  loc_005B1616: push 00405696h ; __vbaExceptHandler
loc_005B161B:   mov eax, fs: [00000000h]
loc_005B1621:   push eax
loc_005B1622:   mov fs: [00000000h] , esp
  loc_005B1629: sub esp, 00000028h
loc_005B162C:   push ebx
loc_005B162D:   push esi
loc_005B162E:   push edi
loc_005B162F:   mov var_C, esp
  loc_005B1632: mov var_8, 00402FE8h
loc_005B1639:   mov eax, Me
loc_005B163C:   mov ecx, eax
  loc_005B163E: and ecx, 00000001h
loc_005B1641:   mov var_4, ecx
  loc_005B1644: and al, FEh
loc_005B1646:   push eax
loc_005B1647:   mov Me, eax
loc_005B164A:   mov edx, [eax]
loc_005B164C:   Call [edx+00000004h]
loc_005B164F:   mov esi, KeyAscii
  loc_005B1652: xor edi, edi
loc_005B1654:   mov var_24, edi
  loc_005B1657: cmp [esi], 000Dh
  loc_005B165B: jnz 005B1686h
loc_005B165D:   lea eax, var_24
  loc_005B1660: mov var_1C, 80020004h
loc_005B1667:   push eax
  loc_005B1668: push 004300CCh ; "{tab}"
  loc_005B166D: mov var_24, 0000000Ah
  loc_005B1674: call [004010D4h] ; rtcSendKeys
loc_005B167A:   lea ecx, var_24
  loc_005B167D: call [00401020h] ; __vbaFreeVar
loc_005B1683:   mov [esi], di
loc_005B1686:   mov var_4, edi
  loc_005B1689: push 005B169Bh
  loc_005B168E: jmp 005B169Ah
loc_005B1690:   lea ecx, var_24
  loc_005B1693: call [00401020h] ; __vbaFreeVar
loc_005B1699:   ret
loc_005B169A:   ret
loc_005B169B:   mov eax, Me
loc_005B169E:   push eax
loc_005B169F:   mov ecx, [eax]
loc_005B16A1:   Call [ecx+00000008h]
loc_005B16A4:   mov eax, var_4
loc_005B16A7:   mov ecx, var_14
loc_005B16AA:   pop edi
loc_005B16AB:   pop esi
loc_005B16AC:   mov fs: [00000000h] , ecx
loc_005B16B3:   pop ebx
loc_005B16B4:   mov esp, ebp
loc_005B16B6:   pop ebp
  loc_005B16B7: retn 0008h
End Sub

Private Sub txtDiskon_KeyPress(KeyAscii As Integer) '5B16C0
loc_005B16C0:   push ebp
loc_005B16C1:   mov ebp, esp
  loc_005B16C3: sub esp, 0000000Ch
  loc_005B16C6: push 00405696h ; __vbaExceptHandler
loc_005B16CB:   mov eax, fs: [00000000h]
loc_005B16D1:   push eax
loc_005B16D2:   mov fs: [00000000h] , esp
  loc_005B16D9: sub esp, 00000028h
loc_005B16DC:   push ebx
loc_005B16DD:   push esi
loc_005B16DE:   push edi
loc_005B16DF:   mov var_C, esp
  loc_005B16E2: mov var_8, 00402FF8h
loc_005B16E9:   mov eax, Me
loc_005B16EC:   mov ecx, eax
  loc_005B16EE: and ecx, 00000001h
loc_005B16F1:   mov var_4, ecx
  loc_005B16F4: and al, FEh
loc_005B16F6:   push eax
loc_005B16F7:   mov Me, eax
loc_005B16FA:   mov edx, [eax]
loc_005B16FC:   Call [edx+00000004h]
loc_005B16FF:   mov esi, KeyAscii
  loc_005B1702: xor edi, edi
loc_005B1704:   mov var_24, edi
  loc_005B1707: cmp [esi], 000Dh
  loc_005B170B: jnz 005B1736h
loc_005B170D:   lea eax, var_24
  loc_005B1710: mov var_1C, 80020004h
loc_005B1717:   push eax
  loc_005B1718: push 004300CCh ; "{tab}"
  loc_005B171D: mov var_24, 0000000Ah
  loc_005B1724: call [004010D4h] ; rtcSendKeys
loc_005B172A:   lea ecx, var_24
  loc_005B172D: call [00401020h] ; __vbaFreeVar
loc_005B1733:   mov [esi], di
loc_005B1736:   mov var_4, edi
  loc_005B1739: push 005B174Bh
  loc_005B173E: jmp 005B174Ah
loc_005B1740:   lea ecx, var_24
  loc_005B1743: call [00401020h] ; __vbaFreeVar
loc_005B1749:   ret
loc_005B174A:   ret
loc_005B174B:   mov eax, Me
loc_005B174E:   push eax
loc_005B174F:   mov ecx, [eax]
loc_005B1751:   Call [ecx+00000008h]
loc_005B1754:   mov eax, var_4
loc_005B1757:   mov ecx, var_14
loc_005B175A:   pop edi
loc_005B175B:   pop esi
loc_005B175C:   mov fs: [00000000h] , ecx
loc_005B1763:   pop ebx
loc_005B1764:   mov esp, ebp
loc_005B1766:   pop ebp
  loc_005B1767: retn 0008h
End Sub

Private Sub txtHarga_KeyPress(KeyAscii As Integer) '5B1770
loc_005B1770:   push ebp
loc_005B1771:   mov ebp, esp
  loc_005B1773: sub esp, 0000000Ch
  loc_005B1776: push 00405696h ; __vbaExceptHandler
loc_005B177B:   mov eax, fs: [00000000h]
loc_005B1781:   push eax
loc_005B1782:   mov fs: [00000000h] , esp
  loc_005B1789: sub esp, 00000028h
loc_005B178C:   push ebx
loc_005B178D:   push esi
loc_005B178E:   push edi
loc_005B178F:   mov var_C, esp
  loc_005B1792: mov var_8, 00403008h
loc_005B1799:   mov eax, Me
loc_005B179C:   mov ecx, eax
  loc_005B179E: and ecx, 00000001h
loc_005B17A1:   mov var_4, ecx
  loc_005B17A4: and al, FEh
loc_005B17A6:   push eax
loc_005B17A7:   mov Me, eax
loc_005B17AA:   mov edx, [eax]
loc_005B17AC:   Call [edx+00000004h]
loc_005B17AF:   mov esi, KeyAscii
  loc_005B17B2: xor edi, edi
loc_005B17B4:   mov var_24, edi
  loc_005B17B7: cmp [esi], 000Dh
  loc_005B17BB: jnz 005B17E6h
loc_005B17BD:   lea eax, var_24
  loc_005B17C0: mov var_1C, 80020004h
loc_005B17C7:   push eax
  loc_005B17C8: push 004300CCh ; "{tab}"
  loc_005B17CD: mov var_24, 0000000Ah
  loc_005B17D4: call [004010D4h] ; rtcSendKeys
loc_005B17DA:   lea ecx, var_24
  loc_005B17DD: call [00401020h] ; __vbaFreeVar
loc_005B17E3:   mov [esi], di
loc_005B17E6:   mov var_4, edi
  loc_005B17E9: push 005B17FBh
  loc_005B17EE: jmp 005B17FAh
loc_005B17F0:   lea ecx, var_24
  loc_005B17F3: call [00401020h] ; __vbaFreeVar
loc_005B17F9:   ret
loc_005B17FA:   ret
loc_005B17FB:   mov eax, Me
loc_005B17FE:   push eax
loc_005B17FF:   mov ecx, [eax]
loc_005B1801:   Call [ecx+00000008h]
loc_005B1804:   mov eax, var_4
loc_005B1807:   mov ecx, var_14
loc_005B180A:   pop edi
loc_005B180B:   pop esi
loc_005B180C:   mov fs: [00000000h] , ecx
loc_005B1813:   pop ebx
loc_005B1814:   mov esp, ebp
loc_005B1816:   pop ebp
  loc_005B1817: retn 0008h
End Sub

Private Sub txtQty_GotFocus() '5B10A0
loc_005B10A0:   push ebp
loc_005B10A1:   mov ebp, esp
  loc_005B10A3: sub esp, 0000000Ch
  loc_005B10A6: push 00405696h ; __vbaExceptHandler
loc_005B10AB:   mov eax, fs: [00000000h]
loc_005B10B1:   push eax
loc_005B10B2:   mov fs: [00000000h] , esp
  loc_005B10B9: sub esp, 0000001Ch
loc_005B10BC:   push ebx
loc_005B10BD:   push esi
loc_005B10BE:   push edi
loc_005B10BF:   mov var_C, esp
  loc_005B10C2: mov var_8, 00402F98h
loc_005B10C9:   mov esi, Me
loc_005B10CC:   mov eax, esi
  loc_005B10CE: and eax, 00000001h
loc_005B10D1:   mov var_4, eax
  loc_005B10D4: and esi, FFFFFFFEh
loc_005B10D7:   push esi
loc_005B10D8:   mov Me, esi
loc_005B10DB:   mov ecx, [esi]
loc_005B10DD:   Call [ecx+00000004h]
loc_005B10E0:   mov edx, [esi]
  loc_005B10E2: xor eax, eax
loc_005B10E4:   push esi
loc_005B10E5:   mov var_18, eax
loc_005B10E8:   mov var_1C, eax
loc_005B10EB:   Call [edx+0000033Ch]
  loc_005B10F1: mov ebx, [004010B8h] ; __vbaObjSet
loc_005B10F7:   push eax
loc_005B10F8:   lea eax, var_1C
loc_005B10FB:   push eax
loc_005B10FC:   Call ebx
loc_005B10FE:   mov edi, eax
loc_005B1100:   lea edx, var_18
loc_005B1103:   push edx
loc_005B1104:   push edi
loc_005B1105:   mov ecx, [edi]
loc_005B1107:   Call [ecx+000000A0h]
loc_005B110D:   test eax, eax
loc_005B110F:   fnclex
  loc_005B1111: jge 005B1125h
  loc_005B1113: push 000000A0h
  loc_005B1118: push 0042DFCCh
loc_005B111D:   push edi
loc_005B111E:   push eax
  loc_005B111F: call [00401080h] ; __vbaHresultCheckObj
loc_005B1125:   mov eax, var_18
loc_005B1128:   push eax
  loc_005B1129: push 00434A94h
  loc_005B112E: call [0040112Ch] ; __vbaStrCmp
loc_005B1134:   mov edi, eax
loc_005B1136:   lea ecx, var_18
loc_005B1139:   neg edi
loc_005B113B:   sbb edi, edi
loc_005B113D:   inc edi
loc_005B113E:   neg edi
  loc_005B1140: call [0040129Ch] ; __vbaFreeStr
loc_005B1146:   lea ecx, var_1C
  loc_005B1149: call [004012A0h] ; __vbaFreeObj
loc_005B114F:   test di, di
  loc_005B1152: jz 005B1195h
loc_005B1154:   mov ecx, [esi]
loc_005B1156:   push esi
loc_005B1157:   Call [ecx+0000033Ch]
loc_005B115D:   lea edx, var_1C
loc_005B1160:   push eax
loc_005B1161:   push edx
loc_005B1162:   Call ebx
loc_005B1164:   mov esi, eax
  loc_005B1166: push 0042DFECh
loc_005B116B:   push esi
loc_005B116C:   mov eax, [esi]
loc_005B116E:   Call [eax+000000A4h]
loc_005B1174:   test eax, eax
loc_005B1176:   fnclex
  loc_005B1178: jge 005B118Ch
  loc_005B117A: push 000000A4h
  loc_005B117F: push 0042DFCCh
loc_005B1184:   push esi
loc_005B1185:   push eax
  loc_005B1186: call [00401080h] ; __vbaHresultCheckObj
loc_005B118C:   lea ecx, var_1C
  loc_005B118F: call [004012A0h] ; __vbaFreeObj
  loc_005B1195: mov var_4, 00000000h
  loc_005B119C: push 005B11B7h
  loc_005B11A1: jmp 005B11B6h
loc_005B11A3:   lea ecx, var_18
  loc_005B11A6: call [0040129Ch] ; __vbaFreeStr
loc_005B11AC:   lea ecx, var_1C
  loc_005B11AF: call [004012A0h] ; __vbaFreeObj
loc_005B11B5:   ret
loc_005B11B6:   ret
loc_005B11B7:   mov eax, Me
loc_005B11BA:   push eax
loc_005B11BB:   mov ecx, [eax]
loc_005B11BD:   Call [ecx+00000008h]
loc_005B11C0:   mov eax, var_4
loc_005B11C3:   mov ecx, var_14
loc_005B11C6:   pop edi
loc_005B11C7:   pop esi
loc_005B11C8:   mov fs: [00000000h] , ecx
loc_005B11CF:   pop ebx
loc_005B11D0:   mov esp, ebp
loc_005B11D2:   pop ebp
  loc_005B11D3: retn 0004h
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer) '5B11E0
loc_005B11E0:   push ebp
loc_005B11E1:   mov ebp, esp
  loc_005B11E3: sub esp, 0000000Ch
  loc_005B11E6: push 00405696h ; __vbaExceptHandler
loc_005B11EB:   mov eax, fs: [00000000h]
loc_005B11F1:   push eax
loc_005B11F2:   mov fs: [00000000h] , esp
  loc_005B11F9: sub esp, 00000014h
loc_005B11FC:   push ebx
loc_005B11FD:   push esi
loc_005B11FE:   push edi
loc_005B11FF:   mov var_C, esp
  loc_005B1202: mov var_8, 00402FA8h
loc_005B1209:   mov esi, Me
loc_005B120C:   mov eax, esi
  loc_005B120E: and eax, 00000001h
loc_005B1211:   mov var_4, eax
  loc_005B1214: and esi, FFFFFFFEh
loc_005B1217:   push esi
loc_005B1218:   mov Me, esi
loc_005B121B:   mov ecx, [esi]
loc_005B121D:   Call [ecx+00000004h]
loc_005B1220:   mov edi, KeyAscii
  loc_005B1223: mov var_18, 00000000h
  loc_005B122A: cmp [edi], 000Dh
  loc_005B122E: jnz 005B1272h
loc_005B1230:   mov edx, [esi]
loc_005B1232:   push esi
loc_005B1233:   Call [edx+00000348h]
loc_005B1239:   push eax
loc_005B123A:   lea eax, var_18
loc_005B123D:   push eax
  loc_005B123E: call [004010B8h] ; __vbaObjSet
loc_005B1244:   mov esi, eax
loc_005B1246:   push esi
loc_005B1247:   mov ecx, [esi]
loc_005B1249:   Call [ecx+00000204h]
loc_005B124F:   test eax, eax
loc_005B1251:   fnclex
  loc_005B1253: jge 005B1267h
  loc_005B1255: push 00000204h
  loc_005B125A: push 0042DFCCh
loc_005B125F:   push esi
loc_005B1260:   push eax
  loc_005B1261: call [00401080h] ; __vbaHresultCheckObj
loc_005B1267:   lea ecx, var_18
  loc_005B126A: call [004012A0h] ; __vbaFreeObj
  loc_005B1270: jmp 005B12BBh
loc_005B1272:   mov si, [edi]
  loc_005B1275: push 0042E3ACh
  loc_005B127A: call [00401050h] ; rtcAnsiValueBstr
  loc_005B1280: xor ebx, ebx
loc_005B1282:   cmp [edi], ax
  loc_005B1285: push 0042E3A4h
loc_005B128A:   setg bl
  loc_005B128D: call [00401050h] ; rtcAnsiValueBstr
  loc_005B1293: xor edx, edx
loc_005B1295:   cmp si, ax
loc_005B1298:   setl dl
  loc_005B129B: or ebx, edx
loc_005B129D:   neg ebx
loc_005B129F:   sbb ebx, ebx
  loc_005B12A1: xor eax, eax
loc_005B12A3:   neg ebx
  loc_005B12A5: cmp si, 0008h
loc_005B12A9:   setnz al
loc_005B12AC:   test eax, ebx
  loc_005B12AE: jz 005B12BBh
  loc_005B12B0: call [004011C8h] ; rtcBeep
  loc_005B12B6: mov [edi], 0000h
  loc_005B12BB: mov var_4, 00000000h
  loc_005B12C2: push 005B12D4h
  loc_005B12C7: jmp 005B12D3h
loc_005B12C9:   lea ecx, var_18
  loc_005B12CC: call [004012A0h] ; __vbaFreeObj
loc_005B12D2:   ret
loc_005B12D3:   ret
loc_005B12D4:   mov eax, Me
loc_005B12D7:   push eax
loc_005B12D8:   mov ecx, [eax]
loc_005B12DA:   Call [ecx+00000008h]
loc_005B12DD:   mov eax, var_4
loc_005B12E0:   mov ecx, var_14
loc_005B12E3:   pop edi
loc_005B12E4:   pop esi
loc_005B12E5:   mov fs: [00000000h] , ecx
loc_005B12EC:   pop ebx
loc_005B12ED:   mov esp, ebp
loc_005B12EF:   pop ebp
  loc_005B12F0: retn 0008h
End Sub

Private Sub txtQty_LostFocus() '5B1300
loc_005B1300:   push ebp
loc_005B1301:   mov ebp, esp
  loc_005B1303: sub esp, 0000000Ch
  loc_005B1306: push 00405696h ; __vbaExceptHandler
loc_005B130B:   mov eax, fs: [00000000h]
loc_005B1311:   push eax
loc_005B1312:   mov fs: [00000000h] , esp
  loc_005B1319: sub esp, 0000002Ch
loc_005B131C:   push ebx
loc_005B131D:   push esi
loc_005B131E:   push edi
loc_005B131F:   mov var_C, esp
  loc_005B1322: mov var_8, 00402FB8h
loc_005B1329:   mov edi, Me
loc_005B132C:   mov eax, edi
  loc_005B132E: and eax, 00000001h
loc_005B1331:   mov var_4, eax
  loc_005B1334: and edi, FFFFFFFEh
loc_005B1337:   push edi
loc_005B1338:   mov Me, edi
loc_005B133B:   mov ecx, [edi]
loc_005B133D:   Call [ecx+00000004h]
loc_005B1340:   mov edx, [edi]
  loc_005B1342: xor eax, eax
loc_005B1344:   push edi
loc_005B1345:   mov var_18, eax
loc_005B1348:   mov var_1C, eax
loc_005B134B:   mov var_20, eax
loc_005B134E:   mov var_24, eax
loc_005B1351:   Call [edx+0000033Ch]
  loc_005B1357: mov ebx, [004010B8h] ; __vbaObjSet
loc_005B135D:   push eax
loc_005B135E:   lea eax, var_20
loc_005B1361:   push eax
loc_005B1362:   Call ebx
loc_005B1364:   mov esi, eax
loc_005B1366:   lea edx, var_18
loc_005B1369:   push edx
loc_005B136A:   push esi
loc_005B136B:   mov ecx, [esi]
loc_005B136D:   Call [ecx+000000A0h]
loc_005B1373:   test eax, eax
loc_005B1375:   fnclex
  loc_005B1377: jge 005B138Bh
  loc_005B1379: push 000000A0h
  loc_005B137E: push 0042DFCCh
loc_005B1383:   push esi
loc_005B1384:   push eax
  loc_005B1385: call [00401080h] ; __vbaHresultCheckObj
loc_005B138B:   mov eax, [edi]
loc_005B138D:   push edi
loc_005B138E:   Call [eax+0000033Ch]
loc_005B1394:   lea ecx, var_24
loc_005B1397:   push eax
loc_005B1398:   push ecx
loc_005B1399:   Call ebx
loc_005B139B:   mov esi, eax
loc_005B139D:   lea eax, var_1C
loc_005B13A0:   push eax
loc_005B13A1:   push esi
loc_005B13A2:   mov edx, [esi]
loc_005B13A4:   Call [edx+000000A0h]
loc_005B13AA:   test eax, eax
loc_005B13AC:   fnclex
  loc_005B13AE: jge 005B13C2h
  loc_005B13B0: push 000000A0h
  loc_005B13B5: push 0042DFCCh
loc_005B13BA:   push esi
loc_005B13BB:   push eax
  loc_005B13BC: call [00401080h] ; __vbaHresultCheckObj
loc_005B13C2:   mov ecx, var_1C
  loc_005B13C5: mov ebx, [0040112Ch] ; __vbaStrCmp
loc_005B13CB:   push ecx
  loc_005B13CC: push 0042E3A4h
loc_005B13D1:   Call ebx
loc_005B13D3:   mov edx, var_18
loc_005B13D6:   mov esi, eax
loc_005B13D8:   neg esi
loc_005B13DA:   sbb esi, esi
loc_005B13DC:   push edx
loc_005B13DD:   inc esi
  loc_005B13DE: push 0042DFECh
loc_005B13E3:   neg esi
loc_005B13E5:   Call ebx
loc_005B13E7:   neg eax
loc_005B13E9:   sbb eax, eax
loc_005B13EB:   lea ecx, var_18
loc_005B13EE:   inc eax
loc_005B13EF:   neg eax
  loc_005B13F1: or esi, eax
loc_005B13F3:   lea eax, var_1C
loc_005B13F6:   push eax
loc_005B13F7:   push ecx
  loc_005B13F8: push 00000002h
  loc_005B13FA: call [00401204h] ; __vbaFreeStrList
loc_005B1400:   lea edx, var_24
loc_005B1403:   lea eax, var_20
loc_005B1406:   push edx
loc_005B1407:   push eax
  loc_005B1408: push 00000002h
  loc_005B140A: call [0040104Ch] ; __vbaFreeObjList
  loc_005B1410: add esp, 00000018h
loc_005B1413:   test si, si
  loc_005B1416: jz 005B145Dh
loc_005B1418:   mov ecx, [edi]
loc_005B141A:   push edi
loc_005B141B:   Call [ecx+0000033Ch]
loc_005B1421:   lea edx, var_20
loc_005B1424:   push eax
loc_005B1425:   push edx
  loc_005B1426: call [004010B8h] ; __vbaObjSet
loc_005B142C:   mov esi, eax
  loc_005B142E: push 00434A94h
loc_005B1433:   push esi
loc_005B1434:   mov eax, [esi]
loc_005B1436:   Call [eax+000000A4h]
loc_005B143C:   test eax, eax
loc_005B143E:   fnclex
  loc_005B1440: jge 005B1454h
  loc_005B1442: push 000000A4h
  loc_005B1447: push 0042DFCCh
loc_005B144C:   push esi
loc_005B144D:   push eax
  loc_005B144E: call [00401080h] ; __vbaHresultCheckObj
loc_005B1454:   lea ecx, var_20
  loc_005B1457: call [004012A0h] ; __vbaFreeObj
  loc_005B145D: mov var_4, 00000000h
  loc_005B1464: push 005B1490h
  loc_005B1469: jmp 005B148Fh
loc_005B146B:   lea ecx, var_1C
loc_005B146E:   lea edx, var_18
loc_005B1471:   push ecx
loc_005B1472:   push edx
  loc_005B1473: push 00000002h
  loc_005B1475: call [00401204h] ; __vbaFreeStrList
loc_005B147B:   lea eax, var_24
loc_005B147E:   lea ecx, var_20
loc_005B1481:   push eax
loc_005B1482:   push ecx
  loc_005B1483: push 00000002h
  loc_005B1485: call [0040104Ch] ; __vbaFreeObjList
  loc_005B148B: add esp, 00000018h
loc_005B148E:   ret
loc_005B148F:   ret
loc_005B1490:   mov eax, Me
loc_005B1493:   push eax
loc_005B1494:   mov edx, [eax]
loc_005B1496:   Call [edx+00000008h]
loc_005B1499:   mov eax, var_4
loc_005B149C:   mov ecx, var_14
loc_005B149F:   pop edi
loc_005B14A0:   pop esi
loc_005B14A1:   mov fs: [00000000h] , ecx
loc_005B14A8:   pop ebx
loc_005B14A9:   mov esp, ebp
loc_005B14AB:   pop ebp
  loc_005B14AC: retn 0004h
End Sub

Public Sub Hitung() '5A5E90
loc_005A5E90:   push ebp
loc_005A5E91:   mov ebp, esp
  loc_005A5E93: sub esp, 0000000Ch
  loc_005A5E96: push 00405696h ; __vbaExceptHandler
loc_005A5E9B:   mov eax, fs: [00000000h]
loc_005A5EA1:   push eax
loc_005A5EA2:   mov fs: [00000000h] , esp
  loc_005A5EA9: sub esp, 000000C8h
loc_005A5EAF:   push ebx
loc_005A5EB0:   push esi
loc_005A5EB1:   push edi
loc_005A5EB2:   mov var_C, esp
  loc_005A5EB5: mov var_8, 00402D20h
  loc_005A5EBC: xor ebx, ebx
loc_005A5EBE:   mov var_4, ebx
loc_005A5EC1:   mov esi, Me
loc_005A5EC4:   push esi
loc_005A5EC5:   mov eax, [esi]
loc_005A5EC7:   Call [eax+00000004h]
loc_005A5ECA:   mov ecx, [esi]
loc_005A5ECC:   push esi
loc_005A5ECD:   mov var_24, ebx
loc_005A5ED0:   mov var_28, ebx
loc_005A5ED3:   mov var_2C, ebx
loc_005A5ED6:   mov var_30, ebx
loc_005A5ED9:   mov var_40, ebx
loc_005A5EDC:   mov var_50, ebx
loc_005A5EDF:   mov var_60, ebx
loc_005A5EE2:   mov var_70, ebx
loc_005A5EE5:   mov var_80, ebx
loc_005A5EE8:   mov var_90, ebx
loc_005A5EEE:   mov var_A0, ebx
loc_005A5EF4:   mov var_C0, ebx
loc_005A5EFA:   Call [ecx+00000338h]
loc_005A5F00:   lea edx, var_30
loc_005A5F03:   push eax
loc_005A5F04:   push edx
  loc_005A5F05: call [004010B8h] ; __vbaObjSet
loc_005A5F0B:   mov edi, eax
  loc_005A5F0D: push 0042E3A4h
loc_005A5F12:   push edi
loc_005A5F13:   mov eax, [edi]
loc_005A5F15:   Call [eax+00000054h]
loc_005A5F18:   cmp eax, ebx
loc_005A5F1A:   fnclex
  loc_005A5F1C: jge 005A5F2Dh
  loc_005A5F1E: push 00000054h
  loc_005A5F20: push 00433F80h
loc_005A5F25:   push edi
loc_005A5F26:   push eax
  loc_005A5F27: call [00401080h] ; __vbaHresultCheckObj
loc_005A5F2D:   lea ecx, var_30
  loc_005A5F30: call [004012A0h] ; __vbaFreeObj
loc_005A5F36:   mov ecx, [esi]
loc_005A5F38:   push esi
loc_005A5F39:   Call [ecx+00000300h]
loc_005A5F3F:   lea edx, var_30
loc_005A5F42:   push eax
loc_005A5F43:   push edx
  loc_005A5F44: call [004010B8h] ; __vbaObjSet
loc_005A5F4A:   mov edi, eax
  loc_005A5F4C: push 0042E3A4h
loc_005A5F51:   push edi
loc_005A5F52:   mov eax, [edi]
loc_005A5F54:   Call [eax+00000054h]
loc_005A5F57:   cmp eax, ebx
loc_005A5F59:   fnclex
  loc_005A5F5B: jge 005A5F6Ch
  loc_005A5F5D: push 00000054h
  loc_005A5F5F: push 00433F80h
loc_005A5F64:   push edi
loc_005A5F65:   push eax
  loc_005A5F66: call [00401080h] ; __vbaHresultCheckObj
loc_005A5F6C:   lea ecx, var_30
  loc_005A5F6F: call [004012A0h] ; __vbaFreeObj
loc_005A5F75:   lea ecx, [esi+00000090h]
loc_005A5F7B:   lea edx, var_80
loc_005A5F7E:   mov var_78, ebx
  loc_005A5F81: mov var_80, 00000002h
  loc_005A5F88: call [0040101Ch] ; __vbaVarMove
loc_005A5F8E:   push ebx
loc_005A5F8F:   lea edi, [esi+000000C0h]
  loc_005A5F95: call [00401128h] ; __vbaCyI2
loc_005A5F9B:   mov ecx, [esi]
loc_005A5F9D:   push ebx
  loc_005A5F9E: push 00000004h
loc_005A5FA0:   mov [edi], eax
loc_005A5FA2:   push esi
loc_005A5FA3:   mov [edi+00000004h], edx
loc_005A5FA6:   Call [ecx+00000390h]
loc_005A5FAC:   lea edx, var_30
loc_005A5FAF:   push eax
loc_005A5FB0:   push edx
  loc_005A5FB1: call [004010B8h] ; __vbaObjSet
loc_005A5FB7:   push eax
loc_005A5FB8:   lea eax, var_40
loc_005A5FBB:   push eax
  loc_005A5FBC: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005A5FC2: add esp, 00000010h
loc_005A5FC5:   push eax
  loc_005A5FC6: call [0040121Ch] ; __vbaI4Var
loc_005A5FCC:   mov ecx, eax
  loc_005A5FCE: sub ecx, 00000001h
  loc_005A5FD1: jo 005A6432h
  loc_005A5FD7: call [00401138h] ; __vbaI2I4
loc_005A5FDD:   lea ecx, var_30
loc_005A5FE0:   mov var_D0, eax
  loc_005A5FE6: mov [esi+00000048h], 0001h
  loc_005A5FEC: call [004012A0h] ; __vbaFreeObj
loc_005A5FF2:   lea ecx, var_40
  loc_005A5FF5: call [00401020h] ; __vbaFreeVar
  loc_005A5FFB: mov ebx, [00401034h] ; __vbaStrVarMove
loc_005A6001:   mov ax, [esi+00000048h]
loc_005A6005:   cmp ax, var_D0
  loc_005A600C: jg 005A627Eh
  loc_005A6012: sub esp, 00000010h
  loc_005A6015: mov ecx, 00000003h
loc_005A601A:   mov edx, esp
loc_005A601C:   mov var_80, ecx
loc_005A601F:   mov var_A0, ecx
  loc_005A6025: sub esp, 00000010h
loc_005A6028:   mov [edx], ecx
loc_005A602A:   mov ecx, var_7C
loc_005A602D:   movsx eax, ax
loc_005A6030:   mov [edx+00000004h], ecx
loc_005A6033:   mov var_78, eax
loc_005A6036:   mov ecx, esp
  loc_005A6038: push 00000002h
loc_005A603A:   mov [edx+00000008h], eax
loc_005A603D:   mov eax, var_74
  loc_005A6040: push 00000041h
loc_005A6042:   push esi
loc_005A6043:   mov [edx+0000000Ch], eax
loc_005A6046:   mov edx, var_A0
loc_005A604C:   mov eax, var_9C
loc_005A6052:   mov [ecx], edx
loc_005A6054:   mov edx, var_94
loc_005A605A:   mov [ecx+00000004h], eax
  loc_005A605D: mov eax, 00000005h
loc_005A6062:   mov [ecx+00000008h], eax
loc_005A6065:   mov eax, [esi]
loc_005A6067:   mov [ecx+0000000Ch], edx
loc_005A606A:   Call [eax+00000390h]
loc_005A6070:   lea ecx, var_30
loc_005A6073:   push eax
loc_005A6074:   push ecx
  loc_005A6075: call [004010B8h] ; __vbaObjSet
loc_005A607B:   lea edx, var_40
loc_005A607E:   push eax
loc_005A607F:   push edx
  loc_005A6080: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005A6086: add esp, 00000030h
loc_005A6089:   lea edx, var_C0
loc_005A608F:   lea ecx, var_60
  loc_005A6092: mov var_B8, 00435A20h ; "00"
  loc_005A609C: mov var_C0, 00000008h
  loc_005A60A6: call [00401238h] ; __vbaVarDup
loc_005A60AC:   lea eax, var_40
loc_005A60AF:   push eax
loc_005A60B0:   Call ebx
  loc_005A60B2: push 00000001h
loc_005A60B4:   lea ecx, var_60
loc_005A60B7:   mov var_48, eax
  loc_005A60BA: push 00000001h
loc_005A60BC:   lea edx, var_50
loc_005A60BF:   push ecx
loc_005A60C0:   lea eax, var_70
loc_005A60C3:   push edx
loc_005A60C4:   push eax
  loc_005A60C5: mov var_50, 00000008h
  loc_005A60CC: call [00401078h] ; rtcVarFromFormatVar
loc_005A60D2:   lea edx, var_70
loc_005A60D5:   lea ecx, var_24
  loc_005A60D8: call [0040101Ch] ; __vbaVarMove
loc_005A60DE:   lea ecx, var_30
  loc_005A60E1: call [004012A0h] ; __vbaFreeObj
loc_005A60E7:   lea ecx, var_60
loc_005A60EA:   lea edx, var_50
loc_005A60ED:   push ecx
loc_005A60EE:   lea eax, var_40
loc_005A60F1:   push edx
loc_005A60F2:   push eax
  loc_005A60F3: push 00000003h
  loc_005A60F5: call [0040103Ch] ; __vbaFreeVarList
loc_005A60FB:   movsx eax, [esi+00000048h]
loc_005A60FF:   mov edx, esp
  loc_005A6101: mov ecx, 00000003h
loc_005A6106:   mov var_80, ecx
loc_005A6109:   mov var_A0, ecx
loc_005A610F:   mov [edx], ecx
loc_005A6111:   mov ecx, var_7C
loc_005A6114:   mov var_78, eax
loc_005A6117:   mov [edx+00000004h], ecx
loc_005A611A:   mov [edx+00000008h], eax
loc_005A611D:   mov eax, var_74
loc_005A6120:   mov [edx+0000000Ch], eax
loc_005A6123:   mov edx, var_A0
loc_005A6129:   mov eax, var_9C
  loc_005A612F: sub esp, 00000010h
loc_005A6132:   mov ecx, esp
  loc_005A6134: push 00000002h
  loc_005A6136: push 00000041h
loc_005A6138:   mov [ecx], edx
loc_005A613A:   mov edx, var_94
loc_005A6140:   push esi
loc_005A6141:   mov [ecx+00000004h], eax
  loc_005A6144: mov eax, 00000006h
loc_005A6149:   mov [ecx+00000008h], eax
loc_005A614C:   mov eax, [esi]
loc_005A614E:   mov [ecx+0000000Ch], edx
loc_005A6151:   Call [eax+00000390h]
loc_005A6157:   lea ecx, var_30
loc_005A615A:   push eax
loc_005A615B:   push ecx
  loc_005A615C: call [004010B8h] ; __vbaObjSet
loc_005A6162:   lea edx, var_40
loc_005A6165:   push eax
loc_005A6166:   push edx
  loc_005A6167: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005A616D: add esp, 00000030h
loc_005A6170:   lea edx, var_C0
loc_005A6176:   lea ecx, var_60
  loc_005A6179: mov var_B8, 00435A20h ; "00"
  loc_005A6183: mov var_C0, 00000008h
  loc_005A618D: call [00401238h] ; __vbaVarDup
loc_005A6193:   lea eax, var_40
loc_005A6196:   push eax
loc_005A6197:   Call ebx
  loc_005A6199: push 00000001h
loc_005A619B:   lea ecx, var_60
loc_005A619E:   mov var_48, eax
  loc_005A61A1: push 00000001h
loc_005A61A3:   lea edx, var_50
loc_005A61A6:   push ecx
loc_005A61A7:   lea eax, var_70
loc_005A61AA:   push edx
loc_005A61AB:   push eax
  loc_005A61AC: mov var_50, 00000008h
  loc_005A61B3: call [00401078h] ; rtcVarFromFormatVar
loc_005A61B9:   lea ecx, var_70
loc_005A61BC:   push ecx
loc_005A61BD:   Call ebx
loc_005A61BF:   mov edx, eax
loc_005A61C1:   lea ecx, var_28
  loc_005A61C4: call [0040126Ch] ; __vbaStrMove
loc_005A61CA:   lea ecx, var_30
  loc_005A61CD: call [004012A0h] ; __vbaFreeObj
loc_005A61D3:   lea edx, var_70
loc_005A61D6:   lea eax, var_60
loc_005A61D9:   push edx
loc_005A61DA:   lea ecx, var_50
loc_005A61DD:   push eax
loc_005A61DE:   lea edx, var_40
loc_005A61E1:   push ecx
loc_005A61E2:   push edx
  loc_005A61E3: push 00000004h
  loc_005A61E5: call [0040103Ch] ; __vbaFreeVarList
  loc_005A61EB: add esp, 00000014h
loc_005A61EE:   lea eax, var_24
loc_005A61F1:   lea ecx, var_2C
loc_005A61F4:   push eax
loc_005A61F5:   push ecx
  loc_005A61F6: call [004011C0h] ; __vbaStrVarVal
loc_005A61FC:   push eax
  loc_005A61FD: call [004012A4h] ; rtcR8ValFromBstr
  loc_005A6203: fstp real8 ptr var_78
loc_005A6206:   lea eax, [esi+00000090h]
loc_005A620C:   lea edx, var_80
loc_005A620F:   push eax
loc_005A6210:   lea eax, var_40
loc_005A6213:   push edx
loc_005A6214:   push eax
  loc_005A6215: mov var_80, 00000005h
  loc_005A621C: call [0040122Ch] ; __vbaVarAdd
loc_005A6222:   mov edx, eax
loc_005A6224:   lea ecx, [esi+00000090h]
  loc_005A622A: call [0040101Ch] ; __vbaVarMove
loc_005A6230:   lea ecx, var_2C
  loc_005A6233: call [0040129Ch] ; __vbaFreeStr
loc_005A6239:   lea ecx, var_40
  loc_005A623C: call [00401020h] ; __vbaFreeVar
loc_005A6242:   mov ecx, var_28
loc_005A6245:   push ecx
  loc_005A6246: call [004012A4h] ; rtcR8ValFromBstr
  loc_005A624C: call [00401224h] ; __vbaFpCy
loc_005A6252:   push edx
loc_005A6253:   mov edx, [edi+00000004h]
loc_005A6256:   push eax
loc_005A6257:   mov eax, [edi]
loc_005A6259:   push edx
loc_005A625A:   push eax
  loc_005A625B: call [004010B0h] ; __vbaCyAdd
loc_005A6261:   mov [edi], eax
  loc_005A6263: mov eax, 00000001h
loc_005A6268:   Add ax, [esi+00000048h]
loc_005A626C:   mov [edi+00000004h], edx
  loc_005A626F: jo 005A6432h
loc_005A6275:   mov [esi+00000048h], ax
  loc_005A6279: jmp 005A6001h
loc_005A627E:   mov ecx, [esi]
loc_005A6280:   push esi
loc_005A6281:   Call [ecx+00000338h]
loc_005A6287:   lea edx, var_30
loc_005A628A:   push eax
loc_005A628B:   push edx
  loc_005A628C: call [004010B8h] ; __vbaObjSet
loc_005A6292:   mov ebx, eax
loc_005A6294:   lea edx, var_80
loc_005A6297:   lea ecx, var_40
loc_005A629A:   mov var_C4, ebx
  loc_005A62A0: mov var_78, 00435A2Ch ; "##,##0"
  loc_005A62A7: mov var_80, 00000008h
  loc_005A62AE: call [00401238h] ; __vbaVarDup
  loc_005A62B4: push 00000001h
loc_005A62B6:   lea eax, var_40
  loc_005A62B9: push 00000001h
loc_005A62BB:   push eax
loc_005A62BC:   lea eax, [esi+00000090h]
loc_005A62C2:   lea ecx, var_50
loc_005A62C5:   push eax
loc_005A62C6:   push ecx
  loc_005A62C7: call [00401078h] ; rtcVarFromFormatVar
loc_005A62CD:   mov ebx, [ebx]
loc_005A62CF:   lea edx, var_50
loc_005A62D2:   lea eax, var_2C
loc_005A62D5:   push edx
loc_005A62D6:   push eax
  loc_005A62D7: call [004011C0h] ; __vbaStrVarVal
loc_005A62DD:   mov ecx, ebx
loc_005A62DF:   mov ebx, var_C4
loc_005A62E5:   push eax
loc_005A62E6:   push ebx
loc_005A62E7:   Call [ecx+00000054h]
loc_005A62EA:   test eax, eax
loc_005A62EC:   fnclex
  loc_005A62EE: jge 005A62FFh
  loc_005A62F0: push 00000054h
  loc_005A62F2: push 00433F80h
loc_005A62F7:   push ebx
loc_005A62F8:   push eax
  loc_005A62F9: call [00401080h] ; __vbaHresultCheckObj
loc_005A62FF:   lea ecx, var_2C
  loc_005A6302: call [0040129Ch] ; __vbaFreeStr
loc_005A6308:   lea ecx, var_30
  loc_005A630B: call [004012A0h] ; __vbaFreeObj
  loc_005A6311: mov ebx, [0040103Ch] ; __vbaFreeVarList
loc_005A6317:   lea edx, var_50
loc_005A631A:   lea eax, var_40
loc_005A631D:   push edx
loc_005A631E:   push eax
  loc_005A631F: push 00000002h
loc_005A6321:   Call ebx
loc_005A6323:   mov ecx, [esi]
  loc_005A6325: add esp, 0000000Ch
loc_005A6328:   push esi
loc_005A6329:   Call [ecx+00000300h]
loc_005A632F:   lea edx, var_30
loc_005A6332:   push eax
loc_005A6333:   push edx
  loc_005A6334: call [004010B8h] ; __vbaObjSet
loc_005A633A:   lea edx, var_90
loc_005A6340:   lea ecx, var_40
loc_005A6343:   mov esi, eax
  loc_005A6345: mov var_88, 00435A2Ch ; "##,##0"
  loc_005A634F: mov var_90, 00000008h
  loc_005A6359: call [00401238h] ; __vbaVarDup
  loc_005A635F: push 00000001h
loc_005A6361:   lea eax, var_40
  loc_005A6364: push 00000001h
loc_005A6366:   lea ecx, var_80
loc_005A6369:   push eax
loc_005A636A:   lea edx, var_50
loc_005A636D:   push ecx
loc_005A636E:   push edx
loc_005A636F:   mov var_78, edi
  loc_005A6372: mov var_80, 00004006h
  loc_005A6379: call [00401078h] ; rtcVarFromFormatVar
loc_005A637F:   mov edi, [esi]
loc_005A6381:   lea eax, var_50
loc_005A6384:   lea ecx, var_2C
loc_005A6387:   push eax
loc_005A6388:   push ecx
  loc_005A6389: call [004011C0h] ; __vbaStrVarVal
loc_005A638F:   push eax
loc_005A6390:   push esi
loc_005A6391:   Call [edi+00000054h]
loc_005A6394:   test eax, eax
loc_005A6396:   fnclex
  loc_005A6398: jge 005A63A9h
  loc_005A639A: push 00000054h
  loc_005A639C: push 00433F80h
loc_005A63A1:   push esi
loc_005A63A2:   push eax
  loc_005A63A3: call [00401080h] ; __vbaHresultCheckObj
loc_005A63A9:   lea ecx, var_2C
  loc_005A63AC: call [0040129Ch] ; __vbaFreeStr
loc_005A63B2:   lea ecx, var_30
  loc_005A63B5: call [004012A0h] ; __vbaFreeObj
loc_005A63BB:   lea edx, var_50
loc_005A63BE:   lea eax, var_40
loc_005A63C1:   push edx
loc_005A63C2:   push eax
  loc_005A63C3: push 00000002h
loc_005A63C5:   Call ebx
  loc_005A63C7: add esp, 0000000Ch
loc_005A63CA:   fwait
  loc_005A63CB: push 005A6413h
  loc_005A63D0: jmp 005A6400h
loc_005A63D2:   lea ecx, var_2C
  loc_005A63D5: call [0040129Ch] ; __vbaFreeStr
loc_005A63DB:   lea ecx, var_30
  loc_005A63DE: call [004012A0h] ; __vbaFreeObj
loc_005A63E4:   lea ecx, var_70
loc_005A63E7:   lea edx, var_60
loc_005A63EA:   push ecx
loc_005A63EB:   lea eax, var_50
loc_005A63EE:   push edx
loc_005A63EF:   lea ecx, var_40
loc_005A63F2:   push eax
loc_005A63F3:   push ecx
  loc_005A63F4: push 00000004h
  loc_005A63F6: call [0040103Ch] ; __vbaFreeVarList
  loc_005A63FC: add esp, 00000014h
loc_005A63FF:   ret
loc_005A6400:   lea ecx, var_24
  loc_005A6403: call [00401020h] ; __vbaFreeVar
loc_005A6409:   lea ecx, var_28
  loc_005A640C: call [0040129Ch] ; __vbaFreeStr
loc_005A6412:   ret
loc_005A6413:   mov eax, Me
loc_005A6416:   push eax
loc_005A6417:   mov edx, [eax]
loc_005A6419:   Call [edx+00000008h]
loc_005A641C:   mov eax, var_4
loc_005A641F:   mov ecx, var_14
loc_005A6422:   pop edi
loc_005A6423:   pop esi
loc_005A6424:   mov fs: [00000000h] , ecx
loc_005A642B:   pop ebx
loc_005A642C:   mov esp, ebp
loc_005A642E:   pop ebp
  loc_005A642F: retn 0004h
End Sub

Public Sub FormKosong() '5A6BD0
loc_005A6BD0:   push ebp
loc_005A6BD1:   mov ebp, esp
  loc_005A6BD3: sub esp, 0000000Ch
  loc_005A6BD6: push 00405696h ; __vbaExceptHandler
loc_005A6BDB:   mov eax, fs: [00000000h]
loc_005A6BE1:   push eax
loc_005A6BE2:   mov fs: [00000000h] , esp
  loc_005A6BE9: sub esp, 00000064h
loc_005A6BEC:   push ebx
loc_005A6BED:   push esi
loc_005A6BEE:   push edi
loc_005A6BEF:   mov var_C, esp
  loc_005A6BF2: mov var_8, 00402D60h
  loc_005A6BF9: xor edi, edi
loc_005A6BFB:   mov var_4, edi
loc_005A6BFE:   mov esi, Me
loc_005A6C01:   push esi
loc_005A6C02:   mov eax, [esi]
loc_005A6C04:   Call [eax+00000004h]
  loc_005A6C07: sub esp, 00000010h
  loc_005A6C0A: mov ecx, 00000008h
loc_005A6C0F:   mov edx, esp
loc_005A6C11:   mov var_58, edi
loc_005A6C14:   mov var_58, ecx
  loc_005A6C17: mov eax, 0042F1A4h ; "PL-001"
loc_005A6C1C:   mov [edx], ecx
loc_005A6C1E:   mov ecx, var_54
loc_005A6C21:   mov var_50, eax
loc_005A6C24:   push FFFFFDFBh
loc_005A6C29:   mov [edx+00000004h], ecx
loc_005A6C2C:   mov ecx, [esi]
loc_005A6C2E:   push esi
loc_005A6C2F:   mov var_18, edi
loc_005A6C32:   mov [edx+00000008h], eax
loc_005A6C35:   mov eax, var_4C
loc_005A6C38:   mov var_28, edi
loc_005A6C3B:   mov var_38, edi
loc_005A6C3E:   mov var_48, edi
loc_005A6C41:   mov [edx+0000000Ch], eax
loc_005A6C44:   Call [ecx+0000038Ch]
  loc_005A6C4A: mov edi, [004010B8h] ; __vbaObjSet
loc_005A6C50:   lea edx, var_18
loc_005A6C53:   push eax
loc_005A6C54:   push edx
loc_005A6C55:   Call edi
loc_005A6C57:   push eax
  loc_005A6C58: call [00401280h] ; __vbaLateIdSt
  loc_005A6C5E: mov ebx, [004012A0h] ; __vbaFreeObj
loc_005A6C64:   lea ecx, var_18
loc_005A6C67:   Call ebx
loc_005A6C69:   mov eax, [esi]
loc_005A6C6B:   push esi
loc_005A6C6C:   Call [eax+00000310h]
loc_005A6C72:   lea ecx, var_18
loc_005A6C75:   push eax
loc_005A6C76:   push ecx
loc_005A6C77:   Call edi
loc_005A6C79:   mov edx, [eax]
  loc_005A6C7B: push 0042F340h ; "Umum"
loc_005A6C80:   push eax
loc_005A6C81:   mov var_6C, eax
loc_005A6C84:   Call [edx+000000A4h]
loc_005A6C8A:   test eax, eax
loc_005A6C8C:   fnclex
  loc_005A6C8E: jge 005A6CA5h
loc_005A6C90:   mov ecx, var_6C
  loc_005A6C93: push 000000A4h
  loc_005A6C98: push 0042DFCCh
loc_005A6C9D:   push ecx
loc_005A6C9E:   push eax
  loc_005A6C9F: call [00401080h] ; __vbaHresultCheckObj
loc_005A6CA5:   lea ecx, var_18
loc_005A6CA8:   Call ebx
loc_005A6CAA:   mov edx, [esi]
loc_005A6CAC:   push esi
loc_005A6CAD:   Call [edx+00000314h]
loc_005A6CB3:   push eax
loc_005A6CB4:   lea eax, var_18
loc_005A6CB7:   push eax
loc_005A6CB8:   Call edi
loc_005A6CBA:   mov ecx, [eax]
  loc_005A6CBC: push 0042F350h ; "-"
loc_005A6CC1:   push eax
loc_005A6CC2:   mov var_6C, eax
loc_005A6CC5:   Call [ecx+000000A4h]
loc_005A6CCB:   test eax, eax
loc_005A6CCD:   fnclex
  loc_005A6CCF: jge 005A6CE6h
loc_005A6CD1:   mov edx, var_6C
  loc_005A6CD4: push 000000A4h
  loc_005A6CD9: push 0042DFCCh
loc_005A6CDE:   push edx
loc_005A6CDF:   push eax
  loc_005A6CE0: call [00401080h] ; __vbaHresultCheckObj
loc_005A6CE6:   lea ecx, var_18
loc_005A6CE9:   Call ebx
  loc_005A6CEB: sub esp, 00000010h
  loc_005A6CEE: mov ecx, 00000008h
loc_005A6CF3:   mov edx, esp
loc_005A6CF5:   mov var_58, ecx
  loc_005A6CF8: mov eax, 00433F74h ; "Tunai"
loc_005A6CFD:   push FFFFFDFBh
loc_005A6D02:   mov [edx], ecx
loc_005A6D04:   mov ecx, var_54
loc_005A6D07:   mov var_50, eax
loc_005A6D0A:   push esi
loc_005A6D0B:   mov [edx+00000004h], ecx
loc_005A6D0E:   mov ecx, [esi]
loc_005A6D10:   mov [edx+00000008h], eax
loc_005A6D13:   mov eax, var_4C
loc_005A6D16:   mov [edx+0000000Ch], eax
loc_005A6D19:   Call [ecx+00000388h]
loc_005A6D1F:   lea edx, var_18
loc_005A6D22:   push eax
loc_005A6D23:   push edx
loc_005A6D24:   Call edi
loc_005A6D26:   push eax
  loc_005A6D27: call [00401280h] ; __vbaLateIdSt
loc_005A6D2D:   lea ecx, var_18
loc_005A6D30:   Call ebx
  loc_005A6D32: sub esp, 00000010h
  loc_005A6D35: mov ecx, 0000000Bh
loc_005A6D3A:   mov edx, esp
loc_005A6D3C:   mov var_58, ecx
  loc_005A6D3F: xor eax, eax
  loc_005A6D41: push 80010007h
loc_005A6D46:   mov [edx], ecx
loc_005A6D48:   mov ecx, var_54
loc_005A6D4B:   mov var_50, eax
loc_005A6D4E:   push esi
loc_005A6D4F:   mov [edx+00000004h], ecx
loc_005A6D52:   mov ecx, [esi]
loc_005A6D54:   mov [edx+00000008h], eax
loc_005A6D57:   mov eax, var_4C
loc_005A6D5A:   mov [edx+0000000Ch], eax
loc_005A6D5D:   Call [ecx+00000384h]
loc_005A6D63:   lea edx, var_18
loc_005A6D66:   push eax
loc_005A6D67:   push edx
loc_005A6D68:   Call edi
loc_005A6D6A:   push eax
  loc_005A6D6B: call [00401280h] ; __vbaLateIdSt
loc_005A6D71:   lea ecx, var_18
loc_005A6D74:   Call ebx
loc_005A6D76:   mov eax, [esi]
loc_005A6D78:   push esi
loc_005A6D79:   Call [eax+00000320h]
loc_005A6D7F:   lea ecx, var_18
loc_005A6D82:   push eax
loc_005A6D83:   push ecx
loc_005A6D84:   Call edi
loc_005A6D86:   mov edx, [eax]
  loc_005A6D88: push 00000000h
loc_005A6D8A:   push eax
loc_005A6D8B:   mov var_6C, eax
loc_005A6D8E:   Call [edx+0000009Ch]
loc_005A6D94:   test eax, eax
loc_005A6D96:   fnclex
  loc_005A6D98: jge 005A6DAFh
loc_005A6D9A:   mov ecx, var_6C
  loc_005A6D9D: push 0000009Ch
  loc_005A6DA2: push 00433F80h
loc_005A6DA7:   push ecx
loc_005A6DA8:   push eax
  loc_005A6DA9: call [00401080h] ; __vbaHresultCheckObj
loc_005A6DAF:   lea ecx, var_18
loc_005A6DB2:   Call ebx
loc_005A6DB4:   lea edx, var_28
loc_005A6DB7:   push edx
  loc_005A6DB8: call [00401220h] ; rtcGetDateVar
loc_005A6DBE:   lea edx, var_58
loc_005A6DC1:   lea ecx, var_38
  loc_005A6DC4: mov var_50, 00433FA8h ; "dd/MM/yyyy"
  loc_005A6DCB: mov var_58, 00000008h
  loc_005A6DD2: call [00401238h] ; __vbaVarDup
  loc_005A6DD8: push 00000001h
loc_005A6DDA:   lea eax, var_38
  loc_005A6DDD: push 00000001h
loc_005A6DDF:   lea ecx, var_28
loc_005A6DE2:   push eax
loc_005A6DE3:   lea edx, var_48
loc_005A6DE6:   push ecx
loc_005A6DE7:   push edx
  loc_005A6DE8: call [00401078h] ; rtcVarFromFormatVar
loc_005A6DEE:   mov ecx, var_48
loc_005A6DF1:   mov edx, var_44
  loc_005A6DF4: sub esp, 00000010h
loc_005A6DF7:   mov eax, esp
  loc_005A6DF9: push 00000014h
loc_005A6DFB:   push esi
loc_005A6DFC:   mov [eax], ecx
loc_005A6DFE:   mov ecx, var_40
loc_005A6E01:   mov [eax+00000004h], edx
loc_005A6E04:   mov edx, var_3C
loc_005A6E07:   mov [eax+00000008h], ecx
loc_005A6E0A:   mov [eax+0000000Ch], edx
loc_005A6E0D:   mov eax, [esi]
loc_005A6E0F:   Call [eax+00000384h]
loc_005A6E15:   lea ecx, var_18
loc_005A6E18:   push eax
loc_005A6E19:   push ecx
loc_005A6E1A:   Call edi
loc_005A6E1C:   push eax
  loc_005A6E1D: call [00401280h] ; __vbaLateIdSt
loc_005A6E23:   lea ecx, var_18
loc_005A6E26:   Call ebx
loc_005A6E28:   lea edx, var_48
loc_005A6E2B:   lea eax, var_38
loc_005A6E2E:   push edx
loc_005A6E2F:   lea ecx, var_28
loc_005A6E32:   push eax
loc_005A6E33:   push ecx
  loc_005A6E34: push 00000003h
  loc_005A6E36: call [0040103Ch] ; __vbaFreeVarList
loc_005A6E3C:   mov edx, [esi]
  loc_005A6E3E: add esp, 00000010h
loc_005A6E41:   push esi
loc_005A6E42:   Call [edx+00000348h]
loc_005A6E48:   push eax
loc_005A6E49:   lea eax, var_18
loc_005A6E4C:   push eax
loc_005A6E4D:   Call edi
loc_005A6E4F:   mov ecx, [eax]
  loc_005A6E51: push 0042DFECh
loc_005A6E56:   push eax
loc_005A6E57:   mov var_6C, eax
loc_005A6E5A:   Call [ecx+000000A4h]
loc_005A6E60:   test eax, eax
loc_005A6E62:   fnclex
  loc_005A6E64: jge 005A6E7Bh
loc_005A6E66:   mov edx, var_6C
  loc_005A6E69: push 000000A4h
  loc_005A6E6E: push 0042DFCCh
loc_005A6E73:   push edx
loc_005A6E74:   push eax
  loc_005A6E75: call [00401080h] ; __vbaHresultCheckObj
loc_005A6E7B:   lea ecx, var_18
loc_005A6E7E:   Call ebx
loc_005A6E80:   mov eax, [esi]
loc_005A6E82:   push esi
loc_005A6E83:   Call [eax+00000344h]
loc_005A6E89:   lea ecx, var_18
loc_005A6E8C:   push eax
loc_005A6E8D:   push ecx
loc_005A6E8E:   Call edi
loc_005A6E90:   mov edx, [eax]
  loc_005A6E92: push 0042DFECh
loc_005A6E97:   push eax
loc_005A6E98:   mov var_6C, eax
loc_005A6E9B:   Call [edx+000000A4h]
loc_005A6EA1:   test eax, eax
loc_005A6EA3:   fnclex
  loc_005A6EA5: jge 005A6EBCh
loc_005A6EA7:   mov ecx, var_6C
  loc_005A6EAA: push 000000A4h
  loc_005A6EAF: push 0042DFCCh
loc_005A6EB4:   push ecx
loc_005A6EB5:   push eax
  loc_005A6EB6: call [00401080h] ; __vbaHresultCheckObj
loc_005A6EBC:   lea ecx, var_18
loc_005A6EBF:   Call ebx
loc_005A6EC1:   mov edx, [esi]
loc_005A6EC3:   push esi
loc_005A6EC4:   Call [edx+00000330h]
loc_005A6ECA:   push eax
loc_005A6ECB:   lea eax, var_18
loc_005A6ECE:   push eax
loc_005A6ECF:   Call edi
loc_005A6ED1:   mov ecx, [eax]
  loc_005A6ED3: push 0042E3A4h
loc_005A6ED8:   push eax
loc_005A6ED9:   mov var_6C, eax
loc_005A6EDC:   Call [ecx+000000A4h]
loc_005A6EE2:   test eax, eax
loc_005A6EE4:   fnclex
  loc_005A6EE6: jge 005A6EFDh
loc_005A6EE8:   mov edx, var_6C
  loc_005A6EEB: push 000000A4h
  loc_005A6EF0: push 0042DFCCh
loc_005A6EF5:   push edx
loc_005A6EF6:   push eax
  loc_005A6EF7: call [00401080h] ; __vbaHresultCheckObj
loc_005A6EFD:   lea ecx, var_18
loc_005A6F00:   Call ebx
loc_005A6F02:   mov eax, [esi]
loc_005A6F04:   push esi
loc_005A6F05:   Call [eax+00000340h]
loc_005A6F0B:   lea ecx, var_18
loc_005A6F0E:   push eax
loc_005A6F0F:   push ecx
loc_005A6F10:   Call edi
loc_005A6F12:   mov edx, [eax]
  loc_005A6F14: push 0042E3A4h
loc_005A6F19:   push eax
loc_005A6F1A:   mov var_6C, eax
loc_005A6F1D:   Call [edx+000000A4h]
loc_005A6F23:   test eax, eax
loc_005A6F25:   fnclex
  loc_005A6F27: jge 005A6F3Eh
loc_005A6F29:   mov ecx, var_6C
  loc_005A6F2C: push 000000A4h
  loc_005A6F31: push 0042DFCCh
loc_005A6F36:   push ecx
loc_005A6F37:   push eax
  loc_005A6F38: call [00401080h] ; __vbaHresultCheckObj
loc_005A6F3E:   lea ecx, var_18
loc_005A6F41:   Call ebx
loc_005A6F43:   mov edx, [esi]
loc_005A6F45:   push esi
loc_005A6F46:   Call [edx+0000033Ch]
loc_005A6F4C:   push eax
loc_005A6F4D:   lea eax, var_18
loc_005A6F50:   push eax
loc_005A6F51:   Call edi
loc_005A6F53:   mov ecx, [eax]
  loc_005A6F55: push 00434A94h
loc_005A6F5A:   push eax
loc_005A6F5B:   mov var_6C, eax
loc_005A6F5E:   Call [ecx+000000A4h]
loc_005A6F64:   test eax, eax
loc_005A6F66:   fnclex
  loc_005A6F68: jge 005A6F7Fh
loc_005A6F6A:   mov edx, var_6C
  loc_005A6F6D: push 000000A4h
  loc_005A6F72: push 0042DFCCh
loc_005A6F77:   push edx
loc_005A6F78:   push eax
  loc_005A6F79: call [00401080h] ; __vbaHresultCheckObj
loc_005A6F7F:   lea ecx, var_18
loc_005A6F82:   Call ebx
loc_005A6F84:   mov eax, [esi]
loc_005A6F86:   push esi
loc_005A6F87:   Call [eax+00000338h]
loc_005A6F8D:   lea ecx, var_18
loc_005A6F90:   push eax
loc_005A6F91:   push ecx
loc_005A6F92:   Call edi
loc_005A6F94:   mov edx, [eax]
  loc_005A6F96: push 0042E3A4h
loc_005A6F9B:   push eax
loc_005A6F9C:   mov var_6C, eax
loc_005A6F9F:   Call [edx+00000054h]
loc_005A6FA2:   test eax, eax
loc_005A6FA4:   fnclex
  loc_005A6FA6: jge 005A6FBAh
loc_005A6FA8:   mov ecx, var_6C
  loc_005A6FAB: push 00000054h
  loc_005A6FAD: push 00433F80h
loc_005A6FB2:   push ecx
loc_005A6FB3:   push eax
  loc_005A6FB4: call [00401080h] ; __vbaHresultCheckObj
loc_005A6FBA:   lea ecx, var_18
loc_005A6FBD:   Call ebx
loc_005A6FBF:   mov edx, [esi]
loc_005A6FC1:   push esi
loc_005A6FC2:   Call [edx+00000304h]
loc_005A6FC8:   push eax
loc_005A6FC9:   lea eax, var_18
loc_005A6FCC:   push eax
loc_005A6FCD:   Call edi
loc_005A6FCF:   mov ecx, [eax]
  loc_005A6FD1: push 0042E3A4h
loc_005A6FD6:   push eax
loc_005A6FD7:   mov var_6C, eax
loc_005A6FDA:   Call [ecx+000000A4h]
loc_005A6FE0:   test eax, eax
loc_005A6FE2:   fnclex
  loc_005A6FE4: jge 005A6FFBh
loc_005A6FE6:   mov edx, var_6C
  loc_005A6FE9: push 000000A4h
  loc_005A6FEE: push 0042DFCCh
loc_005A6FF3:   push edx
loc_005A6FF4:   push eax
  loc_005A6FF5: call [00401080h] ; __vbaHresultCheckObj
loc_005A6FFB:   lea ecx, var_18
loc_005A6FFE:   Call ebx
loc_005A7000:   mov eax, [esi]
loc_005A7002:   push esi
loc_005A7003:   Call [eax+00000300h]
loc_005A7009:   lea ecx, var_18
loc_005A700C:   push eax
loc_005A700D:   push ecx
loc_005A700E:   Call edi
loc_005A7010:   mov edx, [eax]
  loc_005A7012: push 0042E3A4h
loc_005A7017:   push eax
loc_005A7018:   mov var_6C, eax
loc_005A701B:   Call [edx+00000054h]
loc_005A701E:   test eax, eax
loc_005A7020:   fnclex
  loc_005A7022: jge 005A7036h
loc_005A7024:   mov ecx, var_6C
  loc_005A7027: push 00000054h
  loc_005A7029: push 00433F80h
loc_005A702E:   push ecx
loc_005A702F:   push eax
  loc_005A7030: call [00401080h] ; __vbaHresultCheckObj
loc_005A7036:   lea ecx, var_18
loc_005A7039:   Call ebx
loc_005A703B:   mov edx, [esi]
  loc_005A703D: push 00000000h
  loc_005A703F: push 00000044h
loc_005A7041:   push esi
  loc_005A7042: mov [esi+00000034h], 0001h
loc_005A7048:   Call [edx+00000390h]
loc_005A704E:   push eax
loc_005A704F:   lea eax, var_18
loc_005A7052:   push eax
loc_005A7053:   Call edi
loc_005A7055:   push eax
  loc_005A7056: call [0040102Ch] ; __vbaLateIdCall
  loc_005A705C: add esp, 0000000Ch
loc_005A705F:   lea ecx, var_18
loc_005A7062:   Call ebx
  loc_005A7064: sub esp, 00000010h
  loc_005A7067: mov ecx, 00000003h
loc_005A706C:   mov edx, esp
loc_005A706E:   mov var_58, ecx
  loc_005A7071: mov eax, 00000002h
  loc_005A7076: push 00000004h
loc_005A7078:   mov [edx], ecx
loc_005A707A:   mov ecx, var_54
loc_005A707D:   mov var_50, eax
loc_005A7080:   push esi
loc_005A7081:   mov [edx+00000004h], ecx
loc_005A7084:   mov ecx, [esi]
loc_005A7086:   mov [edx+00000008h], eax
loc_005A7089:   mov eax, var_4C
loc_005A708C:   mov [edx+0000000Ch], eax
loc_005A708F:   Call [ecx+00000390h]
loc_005A7095:   lea edx, var_18
loc_005A7098:   push eax
loc_005A7099:   push edx
loc_005A709A:   Call edi
loc_005A709C:   push eax
  loc_005A709D: call [00401280h] ; __vbaLateIdSt
loc_005A70A3:   lea ecx, var_18
loc_005A70A6:   Call ebx
loc_005A70A8:   mov eax, [esi]
loc_005A70AA:   push esi
loc_005A70AB:   Call [eax+00000714h]
loc_005A70B1:   test eax, eax
  loc_005A70B3: jge 005A70C7h
  loc_005A70B5: push 00000714h
  loc_005A70BA: push 00434B78h
loc_005A70BF:   push esi
loc_005A70C0:   push eax
  loc_005A70C1: call [00401080h] ; __vbaHresultCheckObj
  loc_005A70C7: push 005A70F0h
  loc_005A70CC: jmp 005A70EFh
loc_005A70CE:   lea ecx, var_18
  loc_005A70D1: call [004012A0h] ; __vbaFreeObj
loc_005A70D7:   lea ecx, var_48
loc_005A70DA:   lea edx, var_38
loc_005A70DD:   push ecx
loc_005A70DE:   lea eax, var_28
loc_005A70E1:   push edx
loc_005A70E2:   push eax
  loc_005A70E3: push 00000003h
  loc_005A70E5: call [0040103Ch] ; __vbaFreeVarList
  loc_005A70EB: add esp, 00000010h
loc_005A70EE:   ret
loc_005A70EF:   ret
loc_005A70F0:   mov eax, Me
loc_005A70F3:   push eax
loc_005A70F4:   mov ecx, [eax]
loc_005A70F6:   Call [ecx+00000008h]
loc_005A70F9:   mov eax, var_4
loc_005A70FC:   mov ecx, var_14
loc_005A70FF:   pop edi
loc_005A7100:   pop esi
loc_005A7101:   mov fs: [00000000h] , ecx
loc_005A7108:   pop ebx
loc_005A7109:   mov esp, ebp
loc_005A710B:   pop ebp
  loc_005A710C: retn 0004h
End Sub

Public Sub BersihBarang() '5A7110
loc_005A7110:   push ebp
loc_005A7111:   mov ebp, esp
  loc_005A7113: sub esp, 0000000Ch
  loc_005A7116: push 00405696h ; __vbaExceptHandler
loc_005A711B:   mov eax, fs: [00000000h]
loc_005A7121:   push eax
loc_005A7122:   mov fs: [00000000h] , esp
  loc_005A7129: sub esp, 00000014h
loc_005A712C:   push ebx
loc_005A712D:   push esi
loc_005A712E:   push edi
loc_005A712F:   mov var_C, esp
  loc_005A7132: mov var_8, 00402D70h
  loc_005A7139: xor edi, edi
loc_005A713B:   mov var_4, edi
loc_005A713E:   mov esi, Me
loc_005A7141:   push esi
loc_005A7142:   mov eax, [esi]
loc_005A7144:   Call [eax+00000004h]
loc_005A7147:   mov ecx, [esi]
loc_005A7149:   push esi
loc_005A714A:   mov var_18, edi
loc_005A714D:   Call [ecx+00000348h]
  loc_005A7153: mov edi, [004010B8h] ; __vbaObjSet
loc_005A7159:   lea edx, var_18
loc_005A715C:   push eax
loc_005A715D:   push edx
loc_005A715E:   Call edi
loc_005A7160:   mov ebx, eax
  loc_005A7162: push 0042DFECh
loc_005A7167:   push ebx
loc_005A7168:   mov eax, [ebx]
loc_005A716A:   Call [eax+000000A4h]
loc_005A7170:   test eax, eax
loc_005A7172:   fnclex
  loc_005A7174: jge 005A7188h
  loc_005A7176: push 000000A4h
  loc_005A717B: push 0042DFCCh
loc_005A7180:   push ebx
loc_005A7181:   push eax
  loc_005A7182: call [00401080h] ; __vbaHresultCheckObj
loc_005A7188:   lea ecx, var_18
  loc_005A718B: call [004012A0h] ; __vbaFreeObj
loc_005A7191:   mov ecx, [esi]
loc_005A7193:   push esi
loc_005A7194:   Call [ecx+00000344h]
loc_005A719A:   lea edx, var_18
loc_005A719D:   push eax
loc_005A719E:   push edx
loc_005A719F:   Call edi
loc_005A71A1:   mov ebx, eax
  loc_005A71A3: push 0042DFECh
loc_005A71A8:   push ebx
loc_005A71A9:   mov eax, [ebx]
loc_005A71AB:   Call [eax+000000A4h]
loc_005A71B1:   test eax, eax
loc_005A71B3:   fnclex
  loc_005A71B5: jge 005A71C9h
  loc_005A71B7: push 000000A4h
  loc_005A71BC: push 0042DFCCh
loc_005A71C1:   push ebx
loc_005A71C2:   push eax
  loc_005A71C3: call [00401080h] ; __vbaHresultCheckObj
loc_005A71C9:   lea ecx, var_18
  loc_005A71CC: call [004012A0h] ; __vbaFreeObj
loc_005A71D2:   mov ecx, [esi]
loc_005A71D4:   push esi
loc_005A71D5:   Call [ecx+00000340h]
loc_005A71DB:   lea edx, var_18
loc_005A71DE:   push eax
loc_005A71DF:   push edx
loc_005A71E0:   Call edi
loc_005A71E2:   mov ebx, eax
  loc_005A71E4: push 0042E3A4h
loc_005A71E9:   push ebx
loc_005A71EA:   mov eax, [ebx]
loc_005A71EC:   Call [eax+000000A4h]
loc_005A71F2:   test eax, eax
loc_005A71F4:   fnclex
  loc_005A71F6: jge 005A720Ah
  loc_005A71F8: push 000000A4h
  loc_005A71FD: push 0042DFCCh
loc_005A7202:   push ebx
loc_005A7203:   push eax
  loc_005A7204: call [00401080h] ; __vbaHresultCheckObj
loc_005A720A:   lea ecx, var_18
  loc_005A720D: call [004012A0h] ; __vbaFreeObj
loc_005A7213:   mov ecx, [esi]
loc_005A7215:   push esi
loc_005A7216:   Call [ecx+00000330h]
loc_005A721C:   lea edx, var_18
loc_005A721F:   push eax
loc_005A7220:   push edx
loc_005A7221:   Call edi
loc_005A7223:   mov ebx, eax
  loc_005A7225: push 0042E3A4h
loc_005A722A:   push ebx
loc_005A722B:   mov eax, [ebx]
loc_005A722D:   Call [eax+000000A4h]
loc_005A7233:   test eax, eax
loc_005A7235:   fnclex
  loc_005A7237: jge 005A724Bh
  loc_005A7239: push 000000A4h
  loc_005A723E: push 0042DFCCh
loc_005A7243:   push ebx
loc_005A7244:   push eax
  loc_005A7245: call [00401080h] ; __vbaHresultCheckObj
loc_005A724B:   lea ecx, var_18
  loc_005A724E: call [004012A0h] ; __vbaFreeObj
loc_005A7254:   mov ecx, [esi]
loc_005A7256:   push esi
loc_005A7257:   Call [ecx+0000033Ch]
loc_005A725D:   lea edx, var_18
loc_005A7260:   push eax
loc_005A7261:   push edx
loc_005A7262:   Call edi
loc_005A7264:   mov ebx, eax
  loc_005A7266: push 0042E3A4h
loc_005A726B:   push ebx
loc_005A726C:   mov eax, [ebx]
loc_005A726E:   Call [eax+000000A4h]
loc_005A7274:   test eax, eax
loc_005A7276:   fnclex
  loc_005A7278: jge 005A728Ch
  loc_005A727A: push 000000A4h
  loc_005A727F: push 0042DFCCh
loc_005A7284:   push ebx
loc_005A7285:   push eax
  loc_005A7286: call [00401080h] ; __vbaHresultCheckObj
  loc_005A728C: mov ebx, [004012A0h] ; __vbaFreeObj
loc_005A7292:   lea ecx, var_18
loc_005A7295:   Call ebx
loc_005A7297:   mov ecx, [esi]
loc_005A7299:   push esi
loc_005A729A:   Call [ecx+00000304h]
loc_005A72A0:   lea edx, var_18
loc_005A72A3:   push eax
loc_005A72A4:   push edx
loc_005A72A5:   Call edi
loc_005A72A7:   mov esi, eax
  loc_005A72A9: push 0042E3A4h
loc_005A72AE:   push esi
loc_005A72AF:   mov eax, [esi]
loc_005A72B1:   Call [eax+000000A4h]
loc_005A72B7:   test eax, eax
loc_005A72B9:   fnclex
  loc_005A72BB: jge 005A72CFh
  loc_005A72BD: push 000000A4h
  loc_005A72C2: push 0042DFCCh
loc_005A72C7:   push esi
loc_005A72C8:   push eax
  loc_005A72C9: call [00401080h] ; __vbaHresultCheckObj
loc_005A72CF:   lea ecx, var_18
loc_005A72D2:   Call ebx
  loc_005A72D4: push 005A72E6h
  loc_005A72D9: jmp 005A72E5h
loc_005A72DB:   lea ecx, var_18
  loc_005A72DE: call [004012A0h] ; __vbaFreeObj
loc_005A72E4:   ret
loc_005A72E5:   ret
loc_005A72E6:   mov eax, Me
loc_005A72E9:   push eax
loc_005A72EA:   mov ecx, [eax]
loc_005A72EC:   Call [ecx+00000008h]
loc_005A72EF:   mov eax, var_4
loc_005A72F2:   mov ecx, var_14
loc_005A72F5:   pop edi
loc_005A72F6:   pop esi
loc_005A72F7:   mov fs: [00000000h] , ecx
loc_005A72FE:   pop ebx
loc_005A72FF:   mov esp, ebp
loc_005A7301:   pop ebp
  loc_005A7302: retn 0004h
End Sub

Public Sub FormMati() '5A7310
loc_005A7310:   push ebp
loc_005A7311:   mov ebp, esp
  loc_005A7313: sub esp, 0000000Ch
  loc_005A7316: push 00405696h ; __vbaExceptHandler
loc_005A731B:   mov eax, fs: [00000000h]
loc_005A7321:   push eax
loc_005A7322:   mov fs: [00000000h] , esp
  loc_005A7329: sub esp, 00000034h
loc_005A732C:   push ebx
loc_005A732D:   push esi
loc_005A732E:   push edi
loc_005A732F:   mov var_C, esp
  loc_005A7332: mov var_8, 00402D80h
  loc_005A7339: xor edi, edi
loc_005A733B:   mov var_4, edi
loc_005A733E:   mov esi, Me
loc_005A7341:   push esi
loc_005A7342:   mov eax, [esi]
loc_005A7344:   Call [eax+00000004h]
loc_005A7347:   mov ecx, [esi]
loc_005A7349:   push esi
loc_005A734A:   mov var_18, edi
loc_005A734D:   Call [ecx+0000030Ch]
  loc_005A7353: mov edi, [004010B8h] ; __vbaObjSet
loc_005A7359:   lea edx, var_18
loc_005A735C:   push eax
loc_005A735D:   push edx
loc_005A735E:   Call edi
loc_005A7360:   mov ebx, eax
  loc_005A7362: push 00000000h
loc_005A7364:   push ebx
loc_005A7365:   mov eax, [ebx]
loc_005A7367:   Call [eax+0000008Ch]
loc_005A736D:   test eax, eax
loc_005A736F:   fnclex
  loc_005A7371: jge 005A7385h
  loc_005A7373: push 0000008Ch
  loc_005A7378: push 0042DFCCh
loc_005A737D:   push ebx
loc_005A737E:   push eax
  loc_005A737F: call [00401080h] ; __vbaHresultCheckObj
  loc_005A7385: mov ebx, [004012A0h] ; __vbaFreeObj
loc_005A738B:   lea ecx, var_18
loc_005A738E:   Call ebx
  loc_005A7390: sub esp, 00000010h
  loc_005A7393: mov ecx, 0000000Bh
loc_005A7398:   mov edx, esp
  loc_005A739A: xor eax, eax
  loc_005A739C: push 8001000Dh
loc_005A73A1:   push esi
loc_005A73A2:   mov [edx], ecx
loc_005A73A4:   mov ecx, var_24
loc_005A73A7:   mov [edx+00000004h], ecx
loc_005A73AA:   mov ecx, [esi]
loc_005A73AC:   mov [edx+00000008h], eax
loc_005A73AF:   mov eax, var_1C
loc_005A73B2:   mov [edx+0000000Ch], eax
loc_005A73B5:   Call [ecx+00000388h]
loc_005A73BB:   lea edx, var_18
loc_005A73BE:   push eax
loc_005A73BF:   push edx
loc_005A73C0:   Call edi
loc_005A73C2:   push eax
  loc_005A73C3: call [00401280h] ; __vbaLateIdSt
loc_005A73C9:   lea ecx, var_18
loc_005A73CC:   Call ebx
loc_005A73CE:   mov eax, [esi]
loc_005A73D0:   push esi
loc_005A73D1:   Call [eax+00000348h]
loc_005A73D7:   lea ecx, var_18
loc_005A73DA:   push eax
loc_005A73DB:   push ecx
loc_005A73DC:   Call edi
loc_005A73DE:   mov edx, [eax]
  loc_005A73E0: push 00000000h
loc_005A73E2:   push eax
loc_005A73E3:   mov var_3C, eax
loc_005A73E6:   Call [edx+0000008Ch]
loc_005A73EC:   test eax, eax
loc_005A73EE:   fnclex
  loc_005A73F0: jge 005A7407h
loc_005A73F2:   mov ecx, var_3C
  loc_005A73F5: push 0000008Ch
  loc_005A73FA: push 0042DFCCh
loc_005A73FF:   push ecx
loc_005A7400:   push eax
  loc_005A7401: call [00401080h] ; __vbaHresultCheckObj
loc_005A7407:   lea ecx, var_18
loc_005A740A:   Call ebx
loc_005A740C:   mov edx, [esi]
loc_005A740E:   push esi
loc_005A740F:   Call [edx+0000033Ch]
loc_005A7415:   push eax
loc_005A7416:   lea eax, var_18
loc_005A7419:   push eax
loc_005A741A:   Call edi
loc_005A741C:   mov ecx, [eax]
  loc_005A741E: push 00000000h
loc_005A7420:   push eax
loc_005A7421:   mov var_3C, eax
loc_005A7424:   Call [ecx+0000008Ch]
loc_005A742A:   test eax, eax
loc_005A742C:   fnclex
  loc_005A742E: jge 005A7445h
loc_005A7430:   mov edx, var_3C
  loc_005A7433: push 0000008Ch
  loc_005A7438: push 0042DFCCh
loc_005A743D:   push edx
loc_005A743E:   push eax
  loc_005A743F: call [00401080h] ; __vbaHresultCheckObj
loc_005A7445:   lea ecx, var_18
loc_005A7448:   Call ebx
  loc_005A744A: sub esp, 00000010h
  loc_005A744D: mov ecx, 0000000Bh
loc_005A7452:   mov edx, esp
  loc_005A7454: xor eax, eax
  loc_005A7456: push 8001000Dh
loc_005A745B:   push esi
loc_005A745C:   mov [edx], ecx
loc_005A745E:   mov ecx, var_24
loc_005A7461:   mov [edx+00000004h], ecx
loc_005A7464:   mov ecx, [esi]
loc_005A7466:   mov [edx+00000008h], eax
loc_005A7469:   mov eax, var_1C
loc_005A746C:   mov [edx+0000000Ch], eax
loc_005A746F:   Call [ecx+0000038Ch]
loc_005A7475:   lea edx, var_18
loc_005A7478:   push eax
loc_005A7479:   push edx
loc_005A747A:   Call edi
loc_005A747C:   push eax
  loc_005A747D: call [00401280h] ; __vbaLateIdSt
loc_005A7483:   lea ecx, var_18
loc_005A7486:   Call ebx
loc_005A7488:   mov eax, [esi]
loc_005A748A:   push esi
loc_005A748B:   Call [eax+00000310h]
loc_005A7491:   lea ecx, var_18
loc_005A7494:   push eax
loc_005A7495:   push ecx
loc_005A7496:   Call edi
loc_005A7498:   mov edx, [eax]
  loc_005A749A: push 00000000h
loc_005A749C:   push eax
loc_005A749D:   mov var_3C, eax
loc_005A74A0:   Call [edx+0000008Ch]
loc_005A74A6:   test eax, eax
loc_005A74A8:   fnclex
  loc_005A74AA: jge 005A74C1h
loc_005A74AC:   mov ecx, var_3C
  loc_005A74AF: push 0000008Ch
  loc_005A74B4: push 0042DFCCh
loc_005A74B9:   push ecx
loc_005A74BA:   push eax
  loc_005A74BB: call [00401080h] ; __vbaHresultCheckObj
loc_005A74C1:   lea ecx, var_18
loc_005A74C4:   Call ebx
loc_005A74C6:   mov edx, [esi]
loc_005A74C8:   push esi
loc_005A74C9:   Call [edx+00000314h]
loc_005A74CF:   push eax
loc_005A74D0:   lea eax, var_18
loc_005A74D3:   push eax
loc_005A74D4:   Call edi
loc_005A74D6:   mov esi, eax
  loc_005A74D8: push 00000000h
loc_005A74DA:   push esi
loc_005A74DB:   mov ecx, [esi]
loc_005A74DD:   Call [ecx+0000008Ch]
loc_005A74E3:   test eax, eax
loc_005A74E5:   fnclex
  loc_005A74E7: jge 005A74FBh
  loc_005A74E9: push 0000008Ch
  loc_005A74EE: push 0042DFCCh
loc_005A74F3:   push esi
loc_005A74F4:   push eax
  loc_005A74F5: call [00401080h] ; __vbaHresultCheckObj
loc_005A74FB:   lea ecx, var_18
loc_005A74FE:   Call ebx
  loc_005A7500: push 005A7512h
  loc_005A7505: jmp 005A7511h
loc_005A7507:   lea ecx, var_18
  loc_005A750A: call [004012A0h] ; __vbaFreeObj
loc_005A7510:   ret
loc_005A7511:   ret
loc_005A7512:   mov eax, Me
loc_005A7515:   push eax
loc_005A7516:   mov edx, [eax]
loc_005A7518:   Call [edx+00000008h]
loc_005A751B:   mov eax, var_4
loc_005A751E:   mov ecx, var_14
loc_005A7521:   pop edi
loc_005A7522:   pop esi
loc_005A7523:   mov fs: [00000000h] , ecx
loc_005A752A:   pop ebx
loc_005A752B:   mov esp, ebp
loc_005A752D:   pop ebp
  loc_005A752E: retn 0004h
End Sub

Public Sub FormHidup() '5A7540
loc_005A7540:   push ebp
loc_005A7541:   mov ebp, esp
  loc_005A7543: sub esp, 0000000Ch
  loc_005A7546: push 00405696h ; __vbaExceptHandler
loc_005A754B:   mov eax, fs: [00000000h]
loc_005A7551:   push eax
loc_005A7552:   mov fs: [00000000h] , esp
  loc_005A7559: sub esp, 00000034h
loc_005A755C:   push ebx
loc_005A755D:   push esi
loc_005A755E:   push edi
loc_005A755F:   mov var_C, esp
  loc_005A7562: mov var_8, 00402D90h
  loc_005A7569: xor edi, edi
loc_005A756B:   mov var_4, edi
loc_005A756E:   mov esi, Me
loc_005A7571:   push esi
loc_005A7572:   mov eax, [esi]
loc_005A7574:   Call [eax+00000004h]
  loc_005A7577: sub esp, 00000010h
  loc_005A757A: mov ecx, 0000000Bh
loc_005A757F:   mov edx, esp
  loc_005A7581: or eax, FFFFFFFFh
  loc_005A7584: push 8001000Dh
loc_005A7589:   push esi
loc_005A758A:   mov [edx], ecx
loc_005A758C:   mov ecx, var_24
loc_005A758F:   mov var_18, edi
loc_005A7592:   mov [edx+00000004h], ecx
loc_005A7595:   mov ecx, [esi]
loc_005A7597:   mov [edx+00000008h], eax
loc_005A759A:   mov eax, var_1C
loc_005A759D:   mov [edx+0000000Ch], eax
loc_005A75A0:   Call [ecx+00000388h]
  loc_005A75A6: mov edi, [004010B8h] ; __vbaObjSet
loc_005A75AC:   lea edx, var_18
loc_005A75AF:   push eax
loc_005A75B0:   push edx
loc_005A75B1:   Call edi
loc_005A75B3:   push eax
  loc_005A75B4: call [00401280h] ; __vbaLateIdSt
loc_005A75BA:   lea ecx, var_18
  loc_005A75BD: call [004012A0h] ; __vbaFreeObj
loc_005A75C3:   mov eax, [esi]
loc_005A75C5:   push esi
loc_005A75C6:   Call [eax+00000348h]
loc_005A75CC:   lea ecx, var_18
loc_005A75CF:   push eax
loc_005A75D0:   push ecx
loc_005A75D1:   Call edi
loc_005A75D3:   mov ebx, eax
loc_005A75D5:   push FFFFFFFFh
loc_005A75D7:   push ebx
loc_005A75D8:   mov edx, [ebx]
loc_005A75DA:   Call [edx+0000008Ch]
loc_005A75E0:   test eax, eax
loc_005A75E2:   fnclex
  loc_005A75E4: jge 005A75F8h
  loc_005A75E6: push 0000008Ch
  loc_005A75EB: push 0042DFCCh
loc_005A75F0:   push ebx
loc_005A75F1:   push eax
  loc_005A75F2: call [00401080h] ; __vbaHresultCheckObj
loc_005A75F8:   lea ecx, var_18
  loc_005A75FB: call [004012A0h] ; __vbaFreeObj
loc_005A7601:   mov eax, [esi]
loc_005A7603:   push esi
loc_005A7604:   Call [eax+0000033Ch]
loc_005A760A:   lea ecx, var_18
loc_005A760D:   push eax
loc_005A760E:   push ecx
loc_005A760F:   Call edi
loc_005A7611:   mov ebx, eax
loc_005A7613:   push FFFFFFFFh
loc_005A7615:   push ebx
loc_005A7616:   mov edx, [ebx]
loc_005A7618:   Call [edx+0000008Ch]
loc_005A761E:   test eax, eax
loc_005A7620:   fnclex
  loc_005A7622: jge 005A7636h
  loc_005A7624: push 0000008Ch
  loc_005A7629: push 0042DFCCh
loc_005A762E:   push ebx
loc_005A762F:   push eax
  loc_005A7630: call [00401080h] ; __vbaHresultCheckObj
  loc_005A7636: mov ebx, [004012A0h] ; __vbaFreeObj
loc_005A763C:   lea ecx, var_18
loc_005A763F:   Call ebx
  loc_005A7641: sub esp, 00000010h
  loc_005A7644: mov ecx, 0000000Bh
loc_005A7649:   mov edx, esp
  loc_005A764B: or eax, FFFFFFFFh
  loc_005A764E: push 8001000Dh
loc_005A7653:   push esi
loc_005A7654:   mov [edx], ecx
loc_005A7656:   mov ecx, var_24
loc_005A7659:   mov [edx+00000004h], ecx
loc_005A765C:   mov ecx, [esi]
loc_005A765E:   mov [edx+00000008h], eax
loc_005A7661:   mov eax, var_1C
loc_005A7664:   mov [edx+0000000Ch], eax
loc_005A7667:   Call [ecx+0000038Ch]
loc_005A766D:   lea edx, var_18
loc_005A7670:   push eax
loc_005A7671:   push edx
loc_005A7672:   Call edi
loc_005A7674:   push eax
  loc_005A7675: call [00401280h] ; __vbaLateIdSt
loc_005A767B:   lea ecx, var_18
loc_005A767E:   Call ebx
loc_005A7680:   mov eax, [esi]
loc_005A7682:   push esi
loc_005A7683:   Call [eax+00000310h]
loc_005A7689:   lea ecx, var_18
loc_005A768C:   push eax
loc_005A768D:   push ecx
loc_005A768E:   Call edi
loc_005A7690:   mov ebx, eax
loc_005A7692:   push FFFFFFFFh
loc_005A7694:   push ebx
loc_005A7695:   mov edx, [ebx]
loc_005A7697:   Call [edx+0000008Ch]
loc_005A769D:   test eax, eax
loc_005A769F:   fnclex
  loc_005A76A1: jge 005A76B5h
  loc_005A76A3: push 0000008Ch
  loc_005A76A8: push 0042DFCCh
loc_005A76AD:   push ebx
loc_005A76AE:   push eax
  loc_005A76AF: call [00401080h] ; __vbaHresultCheckObj
  loc_005A76B5: mov ebx, [004012A0h] ; __vbaFreeObj
loc_005A76BB:   lea ecx, var_18
loc_005A76BE:   Call ebx
loc_005A76C0:   mov eax, [esi]
loc_005A76C2:   push esi
loc_005A76C3:   Call [eax+00000314h]
loc_005A76C9:   lea ecx, var_18
loc_005A76CC:   push eax
loc_005A76CD:   push ecx
loc_005A76CE:   Call edi
loc_005A76D0:   mov esi, eax
loc_005A76D2:   push FFFFFFFFh
loc_005A76D4:   push esi
loc_005A76D5:   mov edx, [esi]
loc_005A76D7:   Call [edx+0000008Ch]
loc_005A76DD:   test eax, eax
loc_005A76DF:   fnclex
  loc_005A76E1: jge 005A76F5h
  loc_005A76E3: push 0000008Ch
  loc_005A76E8: push 0042DFCCh
loc_005A76ED:   push esi
loc_005A76EE:   push eax
  loc_005A76EF: call [00401080h] ; __vbaHresultCheckObj
loc_005A76F5:   lea ecx, var_18
loc_005A76F8:   Call ebx
  loc_005A76FA: push 005A770Ch
  loc_005A76FF: jmp 005A770Bh
loc_005A7701:   lea ecx, var_18
  loc_005A7704: call [004012A0h] ; __vbaFreeObj
loc_005A770A:   ret
loc_005A770B:   ret
loc_005A770C:   mov eax, Me
loc_005A770F:   push eax
loc_005A7710:   mov ecx, [eax]
loc_005A7712:   Call [ecx+00000008h]
loc_005A7715:   mov eax, var_4
loc_005A7718:   mov ecx, var_14
loc_005A771B:   pop edi
loc_005A771C:   pop esi
loc_005A771D:   mov fs: [00000000h] , ecx
loc_005A7724:   pop ebx
loc_005A7725:   mov esp, ebp
loc_005A7727:   pop ebp
  loc_005A7728: retn 0004h
End Sub

Public Sub FormNormal() '5A7730
loc_005A7730:   push ebp
loc_005A7731:   mov ebp, esp
  loc_005A7733: sub esp, 0000000Ch
  loc_005A7736: push 00405696h ; __vbaExceptHandler
loc_005A773B:   mov eax, fs: [00000000h]
loc_005A7741:   push eax
loc_005A7742:   mov fs: [00000000h] , esp
  loc_005A7749: sub esp, 00000034h
loc_005A774C:   push ebx
loc_005A774D:   push esi
loc_005A774E:   push edi
loc_005A774F:   mov var_C, esp
  loc_005A7752: mov var_8, 00402DA0h
  loc_005A7759: xor edi, edi
loc_005A775B:   mov var_4, edi
loc_005A775E:   mov esi, Me
loc_005A7761:   push esi
loc_005A7762:   mov eax, [esi]
loc_005A7764:   Call [eax+00000004h]
loc_005A7767:   mov ecx, [esi]
loc_005A7769:   push esi
loc_005A776A:   mov var_18, edi
loc_005A776D:   Call [ecx+00000704h]
loc_005A7773:   cmp eax, edi
  loc_005A7775: jge 005A7789h
  loc_005A7777: push 00000704h
  loc_005A777C: push 00434B78h
loc_005A7781:   push esi
loc_005A7782:   push eax
  loc_005A7783: call [00401080h] ; __vbaHresultCheckObj
loc_005A7789:   mov edx, [esi]
loc_005A778B:   push esi
loc_005A778C:   Call [edx+000006FCh]
loc_005A7792:   cmp eax, edi
  loc_005A7794: jge 005A77A8h
  loc_005A7796: push 000006FCh
  loc_005A779B: push 00434B78h
loc_005A77A0:   push esi
loc_005A77A1:   push eax
  loc_005A77A2: call [00401080h] ; __vbaHresultCheckObj
  loc_005A77A8: sub esp, 00000010h
  loc_005A77AB: mov ecx, 0000000Bh
loc_005A77B0:   mov edx, esp
  loc_005A77B2: or eax, FFFFFFFFh
  loc_005A77B5: push 8001000Dh
loc_005A77BA:   push esi
loc_005A77BB:   mov [edx], ecx
loc_005A77BD:   mov ecx, var_24
loc_005A77C0:   mov [edx+00000004h], ecx
loc_005A77C3:   mov ecx, [esi]
loc_005A77C5:   mov [edx+00000008h], eax
loc_005A77C8:   mov eax, var_1C
loc_005A77CB:   mov [edx+0000000Ch], eax
loc_005A77CE:   Call [ecx+0000037Ch]
  loc_005A77D4: mov edi, [004010B8h] ; __vbaObjSet
loc_005A77DA:   lea edx, var_18
loc_005A77DD:   push eax
loc_005A77DE:   push edx
loc_005A77DF:   Call edi
loc_005A77E1:   push eax
  loc_005A77E2: call [00401280h] ; __vbaLateIdSt
  loc_005A77E8: mov ebx, [004012A0h] ; __vbaFreeObj
loc_005A77EE:   lea ecx, var_18
loc_005A77F1:   Call ebx
  loc_005A77F3: sub esp, 00000010h
  loc_005A77F6: mov ecx, 00000008h
loc_005A77FB:   mov edx, esp
  loc_005A77FD: mov eax, 0042E930h ; "&Keluar"
loc_005A7802:   push FFFFFDFAh
loc_005A7807:   push esi
loc_005A7808:   mov [edx], ecx
loc_005A780A:   mov ecx, var_24
loc_005A780D:   mov [edx+00000004h], ecx
loc_005A7810:   mov ecx, [esi]
loc_005A7812:   mov [edx+00000008h], eax
loc_005A7815:   mov eax, var_1C
loc_005A7818:   mov [edx+0000000Ch], eax
loc_005A781B:   Call [ecx+00000398h]
loc_005A7821:   lea edx, var_18
loc_005A7824:   push eax
loc_005A7825:   push edx
loc_005A7826:   Call edi
loc_005A7828:   push eax
  loc_005A7829: call [00401280h] ; __vbaLateIdSt
loc_005A782F:   lea ecx, var_18
loc_005A7832:   Call ebx
  loc_005A7834: sub esp, 00000010h
  loc_005A7837: mov ecx, 0000000Bh
loc_005A783C:   mov edx, esp
  loc_005A783E: xor eax, eax
  loc_005A7840: push 8001000Dh
loc_005A7845:   push esi
loc_005A7846:   mov [edx], ecx
loc_005A7848:   mov ecx, var_24
loc_005A784B:   mov [edx+00000004h], ecx
loc_005A784E:   mov ecx, [esi]
loc_005A7850:   mov [edx+00000008h], eax
loc_005A7853:   mov eax, var_1C
loc_005A7856:   mov [edx+0000000Ch], eax
loc_005A7859:   Call [ecx+00000394h]
loc_005A785F:   lea edx, var_18
loc_005A7862:   push eax
loc_005A7863:   push edx
loc_005A7864:   Call edi
loc_005A7866:   push eax
  loc_005A7867: call [00401280h] ; __vbaLateIdSt
loc_005A786D:   lea ecx, var_18
loc_005A7870:   Call ebx
  loc_005A7872: xor eax, eax
  loc_005A7874: sub esp, 00000010h
loc_005A7877:   mov edx, esp
  loc_005A7879: mov ecx, 0000000Bh
  loc_005A787E: push 8001000Dh
loc_005A7883:   push esi
loc_005A7884:   mov [edx], ecx
loc_005A7886:   mov ecx, var_24
loc_005A7889:   mov [edx+00000004h], ecx
loc_005A788C:   mov ecx, [esi]
loc_005A788E:   mov [edx+00000008h], eax
loc_005A7891:   mov eax, var_1C
loc_005A7894:   mov [edx+0000000Ch], eax
loc_005A7897:   Call [ecx+00000380h]
loc_005A789D:   lea edx, var_18
loc_005A78A0:   push eax
loc_005A78A1:   push edx
loc_005A78A2:   Call edi
loc_005A78A4:   push eax
  loc_005A78A5: call [00401280h] ; __vbaLateIdSt
loc_005A78AB:   lea ecx, var_18
loc_005A78AE:   Call ebx
loc_005A78B0:   mov eax, [esi]
loc_005A78B2:   push esi
loc_005A78B3:   Call [eax+0000030Ch]
loc_005A78B9:   lea ecx, var_18
loc_005A78BC:   push eax
loc_005A78BD:   push ecx
loc_005A78BE:   Call edi
loc_005A78C0:   mov edx, [eax]
loc_005A78C2:   push FFFFFFFFh
loc_005A78C4:   push eax
loc_005A78C5:   mov var_3C, eax
loc_005A78C8:   Call [edx+000001B4h]
loc_005A78CE:   test eax, eax
loc_005A78D0:   fnclex
  loc_005A78D2: jge 005A78E9h
loc_005A78D4:   mov ecx, var_3C
  loc_005A78D7: push 000001B4h
  loc_005A78DC: push 0042DFCCh
loc_005A78E1:   push ecx
loc_005A78E2:   push eax
  loc_005A78E3: call [00401080h] ; __vbaHresultCheckObj
loc_005A78E9:   lea ecx, var_18
loc_005A78EC:   Call ebx
loc_005A78EE:   mov edx, [esi]
loc_005A78F0:   push esi
loc_005A78F1:   Call [edx+00000340h]
loc_005A78F7:   push eax
loc_005A78F8:   lea eax, var_18
loc_005A78FB:   push eax
loc_005A78FC:   Call edi
loc_005A78FE:   mov esi, eax
loc_005A7900:   push FFFFFFFFh
loc_005A7902:   push esi
loc_005A7903:   mov ecx, [esi]
loc_005A7905:   Call [ecx+000001B4h]
loc_005A790B:   test eax, eax
loc_005A790D:   fnclex
  loc_005A790F: jge 005A7923h
  loc_005A7911: push 000001B4h
  loc_005A7916: push 0042DFCCh
loc_005A791B:   push esi
loc_005A791C:   push eax
  loc_005A791D: call [00401080h] ; __vbaHresultCheckObj
loc_005A7923:   lea ecx, var_18
loc_005A7926:   Call ebx
  loc_005A7928: push 005A793Ah
  loc_005A792D: jmp 005A7939h
loc_005A792F:   lea ecx, var_18
  loc_005A7932: call [004012A0h] ; __vbaFreeObj
loc_005A7938:   ret
loc_005A7939:   ret
loc_005A793A:   mov eax, Me
loc_005A793D:   push eax
loc_005A793E:   mov edx, [eax]
loc_005A7940:   Call [edx+00000008h]
loc_005A7943:   mov eax, var_4
loc_005A7946:   mov ecx, var_14
loc_005A7949:   pop edi
loc_005A794A:   pop esi
loc_005A794B:   mov fs: [00000000h] , ecx
loc_005A7952:   pop ebx
loc_005A7953:   mov esp, ebp
loc_005A7955:   pop ebp
  loc_005A7956: retn 0004h
End Sub

Public Sub BuatKode() '5A7960
loc_005A7960:   push ebp
loc_005A7961:   mov ebp, esp
  loc_005A7963: sub esp, 0000000Ch
  loc_005A7966: push 00405696h ; __vbaExceptHandler
loc_005A796B:   mov eax, fs: [00000000h]
loc_005A7971:   push eax
loc_005A7972:   mov fs: [00000000h] , esp
  loc_005A7979: sub esp, 00000094h
loc_005A797F:   push ebx
loc_005A7980:   push esi
loc_005A7981:   push edi
loc_005A7982:   mov var_C, esp
  loc_005A7985: mov var_8, 00402DB0h
  loc_005A798C: xor edi, edi
loc_005A798E:   mov var_4, edi
loc_005A7991:   mov eax, Me
loc_005A7994:   push eax
loc_005A7995:   mov ecx, [eax]
loc_005A7997:   Call [ecx+00000004h]
  loc_005A799A: mov edx, 00435A68h ; "SELECT * FROM Penjualan ORDER BY no_nota"
  loc_005A799F: mov ecx, 0066809Ch
loc_005A79A4:   mov var_18, edi
loc_005A79A7:   mov var_1C, edi
loc_005A79AA:   mov var_20, edi
loc_005A79AD:   mov var_24, edi
loc_005A79B0:   mov var_34, edi
loc_005A79B3:   mov var_44, edi
loc_005A79B6:   mov var_54, edi
loc_005A79B9:   mov var_64, edi
loc_005A79BC:   mov var_74, edi
loc_005A79BF:   mov var_84, edi
loc_005A79C5:   mov var_88, edi
loc_005A79CB:   mov var_A0, edi
  loc_005A79D1: call [004011FCh] ; __vbaStrCopy
  loc_005A79D7: push 0042DF30h
  loc_005A79DC: call [00401168h] ; __vbaNew
loc_005A79E2:   push eax
  loc_005A79E3: push 00668054h
  loc_005A79E8: call [004010B8h] ; __vbaObjSet
loc_005A79EE:   cmp [00668054h], edi
  loc_005A79F4: jnz 005A7A06h
  loc_005A79F6: push 00668054h
  loc_005A79FB: push 0042DF30h
  loc_005A7A00: call [004011E8h] ; __vbaNew2
loc_005A7A06:   mov eax, [0066803Ch]
loc_005A7A0B:   mov esi, [00668054h]
loc_005A7A11:   cmp eax, edi
  loc_005A7A13: jnz 005A7A25h
  loc_005A7A15: push 0066803Ch
  loc_005A7A1A: push 0042DEFCh
  loc_005A7A1F: call [004011E8h] ; __vbaNew2
loc_005A7A25:   push FFFFFFFFh
  loc_005A7A27: push 00000004h
  loc_005A7A29: push 00000002h
loc_005A7A2B:   mov eax, [0066803Ch]
  loc_005A7A30: sub esp, 00000010h
  loc_005A7A33: mov ecx, 00000009h
loc_005A7A38:   mov ebx, esp
loc_005A7A3A:   mov var_74, ecx
loc_005A7A3D:   mov var_6C, eax
  loc_005A7A40: sub esp, 00000010h
loc_005A7A43:   mov [ebx], ecx
loc_005A7A45:   mov ecx, var_70
loc_005A7A48:   mov edx, [0066809Ch]
  loc_005A7A4E: mov var_64, 00000008h
loc_005A7A55:   mov [ebx+00000004h], ecx
loc_005A7A58:   mov ecx, esp
loc_005A7A5A:   mov var_5C, edx
loc_005A7A5D:   mov edi, [esi]
loc_005A7A5F:   mov [ebx+00000008h], eax
loc_005A7A62:   mov eax, var_68
loc_005A7A65:   push esi
loc_005A7A66:   mov [ebx+0000000Ch], eax
loc_005A7A69:   mov eax, var_64
loc_005A7A6C:   mov [ecx], eax
loc_005A7A6E:   mov eax, var_60
loc_005A7A71:   mov [ecx+00000004h], eax
loc_005A7A74:   mov [ecx+00000008h], edx
loc_005A7A77:   mov edx, var_58
loc_005A7A7A:   mov [ecx+0000000Ch], edx
loc_005A7A7D:   Call [edi+000000A0h]
loc_005A7A83:   test eax, eax
loc_005A7A85:   fnclex
  loc_005A7A87: jge 005A7A9Fh
  loc_005A7A89: mov edi, [00401080h] ; __vbaHresultCheckObj
  loc_005A7A8F: push 000000A0h
  loc_005A7A94: push 0042DF44h
loc_005A7A99:   push esi
loc_005A7A9A:   push eax
loc_005A7A9B:   Call edi
  loc_005A7A9D: jmp 005A7AA5h
  loc_005A7A9F: mov edi, [00401080h] ; __vbaHresultCheckObj
loc_005A7AA5:   mov eax, [00668054h]
loc_005A7AAA:   test eax, eax
  loc_005A7AAC: jnz 005A7ABEh
  loc_005A7AAE: push 00668054h
  loc_005A7AB3: push 0042DF30h
  loc_005A7AB8: call [004011E8h] ; __vbaNew2
loc_005A7ABE:   mov esi, [00668054h]
loc_005A7AC4:   push FFFFFFFFh
loc_005A7AC6:   push esi
loc_005A7AC7:   mov eax, [esi]
loc_005A7AC9:   Call [eax+000000A4h]
loc_005A7ACF:   test eax, eax
loc_005A7AD1:   fnclex
  loc_005A7AD3: jge 005A7AE3h
  loc_005A7AD5: push 000000A4h
  loc_005A7ADA: push 0042DF44h
loc_005A7ADF:   push esi
loc_005A7AE0:   push eax
loc_005A7AE1:   Call edi
loc_005A7AE3:   mov eax, [00668054h]
loc_005A7AE8:   test eax, eax
  loc_005A7AEA: jnz 005A7AFCh
  loc_005A7AEC: push 00668054h
  loc_005A7AF1: push 0042DF30h
  loc_005A7AF6: call [004011E8h] ; __vbaNew2
loc_005A7AFC:   mov ecx, [00668054h]
loc_005A7B02:   lea edx, var_A0
loc_005A7B08:   push ecx
loc_005A7B09:   push edx
  loc_005A7B0A: call [004010C8h] ; __vbaObjSetAddref
loc_005A7B10:   mov eax, var_A0
loc_005A7B16:   lea edx, var_88
loc_005A7B1C:   push edx
loc_005A7B1D:   push eax
loc_005A7B1E:   mov ecx, [eax]
loc_005A7B20:   Call [ecx+00000034h]
loc_005A7B23:   test eax, eax
loc_005A7B25:   fnclex
  loc_005A7B27: jge 005A7B3Ah
loc_005A7B29:   mov ecx, var_A0
  loc_005A7B2F: push 00000034h
  loc_005A7B31: push 0042DF44h
loc_005A7B36:   push ecx
loc_005A7B37:   push eax
loc_005A7B38:   Call edi
  loc_005A7B3A: cmp var_88, 0000h
  loc_005A7B42: jz 005A7B5Ah
loc_005A7B44:   mov eax, Me
  loc_005A7B47: mov edx, 00435AC0h ; "JL00000001"
loc_005A7B4C:   lea ecx, [eax+0000006Ch]
  loc_005A7B4F: call [004011FCh] ; __vbaStrCopy
  loc_005A7B55: jmp 005A7D76h
loc_005A7B5A:   mov eax, var_A0
loc_005A7B60:   push eax
loc_005A7B61:   mov ecx, [eax]
loc_005A7B63:   Call [ecx+0000009Ch]
loc_005A7B69:   test eax, eax
loc_005A7B6B:   fnclex
  loc_005A7B6D: jge 005A7B83h
loc_005A7B6F:   mov edx, var_A0
  loc_005A7B75: push 0000009Ch
  loc_005A7B7A: push 0042DF44h
loc_005A7B7F:   push edx
loc_005A7B80:   push eax
loc_005A7B81:   Call edi
loc_005A7B83:   mov eax, var_A0
loc_005A7B89:   lea edx, var_20
loc_005A7B8C:   push edx
loc_005A7B8D:   push eax
loc_005A7B8E:   mov ecx, [eax]
loc_005A7B90:   Call [ecx+00000054h]
loc_005A7B93:   test eax, eax
loc_005A7B95:   fnclex
  loc_005A7B97: jge 005A7BAAh
loc_005A7B99:   mov ecx, var_A0
  loc_005A7B9F: push 00000054h
  loc_005A7BA1: push 0042DF44h
loc_005A7BA6:   push ecx
loc_005A7BA7:   push eax
loc_005A7BA8:   Call edi
loc_005A7BAA:   lea ebx, var_24
loc_005A7BAD:   mov eax, var_20
loc_005A7BB0:   push ebx
  loc_005A7BB1: mov edx, 00000008h
  loc_005A7BB6: sub esp, 00000010h
loc_005A7BB9:   mov var_64, edx
loc_005A7BBC:   mov ebx, esp
  loc_005A7BBE: mov ecx, 00435ADCh ; "no_nota"
loc_005A7BC3:   mov var_5C, ecx
loc_005A7BC6:   mov edi, [eax]
loc_005A7BC8:   mov [ebx], edx
loc_005A7BCA:   mov edx, var_60
loc_005A7BCD:   push eax
loc_005A7BCE:   mov esi, eax
loc_005A7BD0:   mov [ebx+00000004h], edx
loc_005A7BD3:   mov [ebx+00000008h], ecx
loc_005A7BD6:   mov ecx, var_58
loc_005A7BD9:   mov [ebx+0000000Ch], ecx
loc_005A7BDC:   Call [edi+00000028h]
loc_005A7BDF:   test eax, eax
loc_005A7BE1:   fnclex
  loc_005A7BE3: jge 005A7BF4h
  loc_005A7BE5: push 00000028h
  loc_005A7BE7: push 0042DFACh
loc_005A7BEC:   push esi
loc_005A7BED:   push eax
  loc_005A7BEE: call [00401080h] ; __vbaHresultCheckObj
loc_005A7BF4:   mov eax, var_24
loc_005A7BF7:   lea ecx, var_34
loc_005A7BFA:   push ecx
loc_005A7BFB:   push eax
loc_005A7BFC:   mov edx, [eax]
loc_005A7BFE:   mov esi, eax
loc_005A7C00:   Call [edx+00000034h]
loc_005A7C03:   test eax, eax
loc_005A7C05:   fnclex
  loc_005A7C07: jge 005A7C18h
  loc_005A7C09: push 00000034h
  loc_005A7C0B: push 0042DFBCh
loc_005A7C10:   push esi
loc_005A7C11:   push eax
  loc_005A7C12: call [00401080h] ; __vbaHresultCheckObj
  loc_005A7C18: mov edi, [00401034h] ; __vbaStrVarMove
loc_005A7C1E:   lea edx, var_34
loc_005A7C21:   push edx
loc_005A7C22:   Call edi
  loc_005A7C24: mov esi, [0040126Ch] ; __vbaStrMove
loc_005A7C2A:   mov edx, eax
loc_005A7C2C:   lea ecx, var_18
  loc_005A7C2F: call __vbaStrMove
loc_005A7C31:   lea eax, var_24
loc_005A7C34:   lea ecx, var_20
loc_005A7C37:   push eax
loc_005A7C38:   push ecx
  loc_005A7C39: push 00000002h
  loc_005A7C3B: call [0040104Ch] ; __vbaFreeObjList
  loc_005A7C41: mov ebx, [00401020h] ; __vbaFreeVar
  loc_005A7C47: add esp, 0000000Ch
loc_005A7C4A:   lea ecx, var_34
loc_005A7C4D:   Call ebx
loc_005A7C4F:   lea eax, var_64
  loc_005A7C52: push 00000008h
loc_005A7C54:   lea ecx, var_34
loc_005A7C57:   lea edx, var_18
loc_005A7C5A:   push eax
loc_005A7C5B:   push ecx
loc_005A7C5C:   mov var_5C, edx
  loc_005A7C5F: mov var_64, 00004008h
  loc_005A7C66: call [00401274h] ; rtcRightCharVar
loc_005A7C6C:   lea edx, var_34
loc_005A7C6F:   lea eax, var_1C
loc_005A7C72:   push edx
loc_005A7C73:   push eax
  loc_005A7C74: call [004011C0h] ; __vbaStrVarVal
loc_005A7C7A:   push eax
  loc_005A7C7B: call [004012A4h] ; rtcR8ValFromBstr
  loc_005A7C81: sub esp, 00000008h
  loc_005A7C84: fstp real8 ptr [esp]
  loc_005A7C87: call [0040115Ch] ; __vbaStrR8
loc_005A7C8D:   mov edx, eax
loc_005A7C8F:   lea ecx, var_18
  loc_005A7C92: call __vbaStrMove
loc_005A7C94:   lea ecx, var_1C
  loc_005A7C97: call [0040129Ch] ; __vbaFreeStr
loc_005A7C9D:   lea ecx, var_34
loc_005A7CA0:   Call ebx
loc_005A7CA2:   mov ecx, var_18
loc_005A7CA5:   push ecx
  loc_005A7CA6: call [004011E4h] ; __vbaR8Str
  loc_005A7CAC: fadd st0, real8 ptr [004015F0h]
  loc_005A7CB2: sub esp, 00000008h
loc_005A7CB5:   fnstsw ax
  loc_005A7CB7: test al, 0Dh
  loc_005A7CB9: jnz 005A7DE4h
  loc_005A7CBF: fstp real8 ptr [esp]
  loc_005A7CC2: call [0040115Ch] ; __vbaStrR8
loc_005A7CC8:   mov edx, eax
loc_005A7CCA:   lea ecx, var_18
  loc_005A7CCD: call __vbaStrMove
  loc_005A7CCF: mov eax, 00000008h
loc_005A7CD4:   lea edx, var_74
loc_005A7CD7:   lea ecx, var_34
  loc_005A7CDA: mov var_7C, 00435AF0h ; "JL"
loc_005A7CE1:   mov var_84, eax
  loc_005A7CE7: mov var_6C, 0043406Ch ; "00000000"
loc_005A7CEE:   mov var_74, eax
  loc_005A7CF1: call [00401238h] ; __vbaVarDup
loc_005A7CF7:   lea edx, var_18
  loc_005A7CFA: push 00000001h
loc_005A7CFC:   lea eax, var_34
loc_005A7CFF:   mov var_5C, edx
  loc_005A7D02: push 00000001h
loc_005A7D04:   lea ecx, var_64
loc_005A7D07:   push eax
loc_005A7D08:   lea edx, var_44
loc_005A7D0B:   push ecx
loc_005A7D0C:   push edx
  loc_005A7D0D: mov var_64, 00004008h
  loc_005A7D14: call [00401078h] ; rtcVarFromFormatVar
loc_005A7D1A:   lea eax, var_84
loc_005A7D20:   lea ecx, var_44
loc_005A7D23:   push eax
loc_005A7D24:   push ecx
loc_005A7D25:   lea edx, var_54
loc_005A7D28:   push edx
  loc_005A7D29: call [0040122Ch] ; __vbaVarAdd
loc_005A7D2F:   push eax
loc_005A7D30:   Call edi
loc_005A7D32:   mov edx, eax
loc_005A7D34:   lea ecx, var_1C
  loc_005A7D37: call __vbaStrMove
loc_005A7D39:   mov edx, eax
loc_005A7D3B:   mov eax, Me
loc_005A7D3E:   lea ecx, [eax+0000006Ch]
  loc_005A7D41: call [004011FCh] ; __vbaStrCopy
loc_005A7D47:   lea ecx, var_1C
  loc_005A7D4A: call [0040129Ch] ; __vbaFreeStr
loc_005A7D50:   lea ecx, var_54
loc_005A7D53:   lea edx, var_44
loc_005A7D56:   push ecx
loc_005A7D57:   lea eax, var_34
loc_005A7D5A:   push edx
loc_005A7D5B:   push eax
  loc_005A7D5C: push 00000003h
  loc_005A7D5E: call [0040103Ch] ; __vbaFreeVarList
  loc_005A7D64: add esp, 00000010h
loc_005A7D67:   lea ecx, var_A0
  loc_005A7D6D: push 00000000h
loc_005A7D6F:   push ecx
  loc_005A7D70: call [004010C8h] ; __vbaObjSetAddref
loc_005A7D76:   fwait
  loc_005A7D77: push 005A7DC5h
  loc_005A7D7C: jmp 005A7DAFh
loc_005A7D7E:   lea ecx, var_1C
  loc_005A7D81: call [0040129Ch] ; __vbaFreeStr
loc_005A7D87:   lea edx, var_24
loc_005A7D8A:   lea eax, var_20
loc_005A7D8D:   push edx
loc_005A7D8E:   push eax
  loc_005A7D8F: push 00000002h
  loc_005A7D91: call [0040104Ch] ; __vbaFreeObjList
loc_005A7D97:   lea ecx, var_54
loc_005A7D9A:   lea edx, var_44
loc_005A7D9D:   push ecx
loc_005A7D9E:   lea eax, var_34
loc_005A7DA1:   push edx
loc_005A7DA2:   push eax
  loc_005A7DA3: push 00000003h
  loc_005A7DA5: call [0040103Ch] ; __vbaFreeVarList
  loc_005A7DAB: add esp, 0000001Ch
loc_005A7DAE:   ret
loc_005A7DAF:   lea ecx, var_A0
  loc_005A7DB5: call [004012A0h] ; __vbaFreeObj
loc_005A7DBB:   lea ecx, var_18
  loc_005A7DBE: call [0040129Ch] ; __vbaFreeStr
loc_005A7DC4:   ret
loc_005A7DC5:   mov eax, Me
loc_005A7DC8:   push eax
loc_005A7DC9:   mov ecx, [eax]
loc_005A7DCB:   Call [ecx+00000008h]
loc_005A7DCE:   mov eax, var_4
loc_005A7DD1:   mov ecx, var_14
loc_005A7DD4:   pop edi
loc_005A7DD5:   pop esi
loc_005A7DD6:   mov fs: [00000000h] , ecx
loc_005A7DDD:   pop ebx
loc_005A7DDE:   mov esp, ebp
loc_005A7DE0:   pop ebp
  loc_005A7DE1: retn 0004h
End Sub

Public Sub AktifGridJual() '5A7DF0
loc_005A7DF0:   push ebp
loc_005A7DF1:   mov ebp, esp
  loc_005A7DF3: sub esp, 0000000Ch
  loc_005A7DF6: push 00405696h ; __vbaExceptHandler
loc_005A7DFB:   mov eax, fs: [00000000h]
loc_005A7E01:   push eax
loc_005A7E02:   mov fs: [00000000h] , esp
  loc_005A7E09: sub esp, 0000004Ch
loc_005A7E0C:   push ebx
loc_005A7E0D:   push esi
loc_005A7E0E:   push edi
loc_005A7E0F:   mov var_C, esp
  loc_005A7E12: mov var_8, 00402DC0h
  loc_005A7E19: xor edi, edi
loc_005A7E1B:   mov var_4, edi
loc_005A7E1E:   mov esi, Me
loc_005A7E21:   push esi
loc_005A7E22:   mov eax, [esi]
loc_005A7E24:   Call [eax+00000004h]
loc_005A7E27:   mov ecx, [esi]
loc_005A7E29:   push esi
loc_005A7E2A:   mov var_58, edi
loc_005A7E2D:   Call [ecx+00000390h]
loc_005A7E33:   lea edx, var_58
loc_005A7E36:   push eax
loc_005A7E37:   push edx
  loc_005A7E38: call [004010B8h] ; __vbaObjSet
loc_005A7E3E:   mov esi, var_20
  loc_005A7E41: sub esp, 00000010h
loc_005A7E44:   mov edx, esp
  loc_005A7E46: mov ecx, 00000003h
loc_005A7E4B:   mov edi, var_18
  loc_005A7E4E: mov ebx, [00401280h] ; __vbaLateIdSt
loc_005A7E54:   mov [edx], ecx
  loc_005A7E56: mov eax, 00000008h
  loc_005A7E5B: push 00000005h
loc_005A7E5D:   mov [edx+00000004h], esi
loc_005A7E60:   mov [edx+00000008h], eax
loc_005A7E63:   mov eax, var_58
loc_005A7E66:   push eax
loc_005A7E67:   mov [edx+0000000Ch], edi
loc_005A7E6A:   Call ebx
  loc_005A7E6C: sub esp, 00000010h
  loc_005A7E6F: mov ecx, 00000003h
loc_005A7E74:   mov edx, esp
  loc_005A7E76: mov eax, 0000012Ch
  loc_005A7E7B: push 00000021h
loc_005A7E7D:   mov [edx], ecx
loc_005A7E7F:   mov [edx+00000004h], esi
loc_005A7E82:   mov [edx+00000008h], eax
loc_005A7E85:   mov eax, var_58
loc_005A7E88:   push eax
loc_005A7E89:   mov [edx+0000000Ch], edi
loc_005A7E8C:   Call ebx
  loc_005A7E8E: sub esp, 00000010h
  loc_005A7E91: mov ecx, 00000003h
loc_005A7E96:   mov edx, esp
  loc_005A7E98: xor eax, eax
  loc_005A7E9A: push 0000000Bh
loc_005A7E9C:   mov [edx], ecx
loc_005A7E9E:   mov [edx+00000004h], esi
loc_005A7EA1:   mov [edx+00000008h], eax
loc_005A7EA4:   mov eax, var_58
loc_005A7EA7:   push eax
loc_005A7EA8:   mov [edx+0000000Ch], edi
loc_005A7EAB:   Call ebx
  loc_005A7EAD: sub esp, 00000010h
  loc_005A7EB0: mov ecx, 00000003h
loc_005A7EB5:   mov edx, esp
  loc_005A7EB7: xor eax, eax
  loc_005A7EB9: push 0000000Ah
loc_005A7EBB:   mov [edx], ecx
loc_005A7EBD:   mov [edx+00000004h], esi
loc_005A7EC0:   mov [edx+00000008h], eax
loc_005A7EC3:   mov eax, var_58
loc_005A7EC6:   push eax
loc_005A7EC7:   mov [edx+0000000Ch], edi
loc_005A7ECA:   Call ebx
  loc_005A7ECC: sub esp, 00000010h
  loc_005A7ECF: mov ecx, 00000008h
loc_005A7ED4:   mov edx, esp
  loc_005A7ED6: mov eax, 00433298h ; "Kode"
  loc_005A7EDB: push 00000000h
loc_005A7EDD:   mov [edx], ecx
loc_005A7EDF:   mov [edx+00000004h], esi
loc_005A7EE2:   mov [edx+00000008h], eax
loc_005A7EE5:   mov eax, var_58
loc_005A7EE8:   push eax
loc_005A7EE9:   mov [edx+0000000Ch], edi
loc_005A7EEC:   Call ebx
  loc_005A7EEE: sub esp, 00000010h
  loc_005A7EF1: mov ecx, 0000000Bh
loc_005A7EF6:   mov edx, esp
  loc_005A7EF8: or eax, FFFFFFFFh
  loc_005A7EFB: push 0000004Fh
loc_005A7EFD:   mov [edx], ecx
loc_005A7EFF:   mov [edx+00000004h], esi
loc_005A7F02:   mov [edx+00000008h], eax
loc_005A7F05:   mov eax, var_58
loc_005A7F08:   push eax
loc_005A7F09:   mov [edx+0000000Ch], edi
loc_005A7F0C:   Call ebx
  loc_005A7F0E: sub esp, 00000010h
  loc_005A7F11: mov ecx, 00000003h
loc_005A7F16:   mov edx, esp
  loc_005A7F18: xor eax, eax
  loc_005A7F1A: sub esp, 00000010h
loc_005A7F1D:   mov var_44, ecx
loc_005A7F20:   mov [edx], ecx
loc_005A7F22:   mov ecx, esp
  loc_005A7F24: push 00000001h
  loc_005A7F26: push 00000039h
loc_005A7F28:   mov [edx+00000004h], esi
loc_005A7F2B:   mov [edx+00000008h], eax
loc_005A7F2E:   mov eax, var_40
loc_005A7F31:   mov [edx+0000000Ch], edi
loc_005A7F34:   mov edx, var_44
loc_005A7F37:   mov [ecx], edx
loc_005A7F39:   mov edx, var_38
loc_005A7F3C:   mov [ecx+00000004h], eax
  loc_005A7F3F: mov eax, 000007D0h
loc_005A7F44:   mov [ecx+00000008h], eax
loc_005A7F47:   mov eax, var_58
loc_005A7F4A:   push eax
loc_005A7F4B:   mov [ecx+0000000Ch], edx
  loc_005A7F4E: call [0040117Ch] ; __vbaLateIdCallSt
  loc_005A7F54: add esp, 0000001Ch
  loc_005A7F57: mov ecx, 00000003h
loc_005A7F5C:   mov edx, esp
  loc_005A7F5E: mov eax, 00000001h
loc_005A7F63:   mov var_1C, eax
  loc_005A7F66: push 00000034h
loc_005A7F68:   mov [edx], ecx
loc_005A7F6A:   mov [edx+00000004h], esi
loc_005A7F6D:   mov [edx+00000008h], eax
loc_005A7F70:   mov eax, var_58
loc_005A7F73:   push eax
loc_005A7F74:   mov [edx+0000000Ch], edi
loc_005A7F77:   Call ebx
  loc_005A7F79: mov ecx, 00000004h
  loc_005A7F7E: call [00401138h] ; __vbaI2I4
  loc_005A7F84: sub esp, 00000010h
loc_005A7F87:   mov var_1C, ax
loc_005A7F8B:   mov edx, var_1C
loc_005A7F8E:   mov ecx, esp
  loc_005A7F90: mov eax, 00000002h
  loc_005A7F95: push 00000028h
loc_005A7F97:   mov [ecx], eax
loc_005A7F99:   mov eax, var_58
loc_005A7F9C:   push eax
loc_005A7F9D:   mov [ecx+00000004h], esi
loc_005A7FA0:   mov [ecx+00000008h], edx
loc_005A7FA3:   mov [ecx+0000000Ch], edi
loc_005A7FA6:   Call ebx
  loc_005A7FA8: sub esp, 00000010h
  loc_005A7FAB: mov ecx, 00000003h
loc_005A7FB0:   mov edx, esp
  loc_005A7FB2: mov eax, 00000001h
  loc_005A7FB7: push 0000000Bh
loc_005A7FB9:   mov [edx], ecx
loc_005A7FBB:   mov [edx+00000004h], esi
loc_005A7FBE:   mov [edx+00000008h], eax
loc_005A7FC1:   mov eax, var_58
loc_005A7FC4:   push eax
loc_005A7FC5:   mov [edx+0000000Ch], edi
loc_005A7FC8:   Call ebx
  loc_005A7FCA: xor eax, eax
  loc_005A7FCC: sub esp, 00000010h
  loc_005A7FCF: mov ecx, 00000003h
loc_005A7FD4:   mov edx, esp
loc_005A7FD6:   mov [edx], ecx
  loc_005A7FD8: push 0000000Ah
loc_005A7FDA:   mov [edx+00000004h], esi
loc_005A7FDD:   mov [edx+00000008h], eax
loc_005A7FE0:   mov eax, var_58
loc_005A7FE3:   push eax
loc_005A7FE4:   mov [edx+0000000Ch], edi
loc_005A7FE7:   Call ebx
  loc_005A7FE9: sub esp, 00000010h
  loc_005A7FEC: mov ecx, 00000008h
loc_005A7FF1:   mov edx, esp
  loc_005A7FF3: mov eax, 004332B8h ; "Nama Barang"
  loc_005A7FF8: push 00000000h
loc_005A7FFA:   mov [edx], ecx
loc_005A7FFC:   mov [edx+00000004h], esi
loc_005A7FFF:   mov [edx+00000008h], eax
loc_005A8002:   mov eax, var_58
loc_005A8005:   push eax
loc_005A8006:   mov [edx+0000000Ch], edi
loc_005A8009:   Call ebx
  loc_005A800B: sub esp, 00000010h
  loc_005A800E: mov ecx, 0000000Bh
loc_005A8013:   mov edx, esp
  loc_005A8015: or eax, FFFFFFFFh
  loc_005A8018: push 0000004Fh
loc_005A801A:   mov [edx], ecx
loc_005A801C:   mov [edx+00000004h], esi
loc_005A801F:   mov [edx+00000008h], eax
loc_005A8022:   mov eax, var_58
loc_005A8025:   push eax
loc_005A8026:   mov [edx+0000000Ch], edi
loc_005A8029:   Call ebx
  loc_005A802B: sub esp, 00000010h
  loc_005A802E: mov ecx, 00000003h
loc_005A8033:   mov edx, esp
  loc_005A8035: mov eax, 00000001h
  loc_005A803A: sub esp, 00000010h
loc_005A803D:   mov var_44, ecx
loc_005A8040:   mov [edx], ecx
loc_005A8042:   mov ecx, esp
  loc_005A8044: push 00000001h
  loc_005A8046: push 00000039h
loc_005A8048:   mov [edx+00000004h], esi
loc_005A804B:   mov [edx+00000008h], eax
loc_005A804E:   mov eax, var_40
loc_005A8051:   mov [edx+0000000Ch], edi
loc_005A8054:   mov edx, var_44
loc_005A8057:   mov [ecx], edx
loc_005A8059:   mov edx, var_38
loc_005A805C:   mov [ecx+00000004h], eax
  loc_005A805F: mov eax, 00001770h
loc_005A8064:   mov [ecx+00000008h], eax
loc_005A8067:   mov eax, var_58
loc_005A806A:   push eax
loc_005A806B:   mov [ecx+0000000Ch], edx
  loc_005A806E: call [0040117Ch] ; __vbaLateIdCallSt
  loc_005A8074: add esp, 0000001Ch
  loc_005A8077: mov ecx, 00000003h
loc_005A807C:   mov edx, esp
  loc_005A807E: mov eax, 00000001h
loc_005A8083:   mov var_1C, eax
  loc_005A8086: push 00000034h
loc_005A8088:   mov [edx], ecx
loc_005A808A:   mov [edx+00000004h], esi
loc_005A808D:   mov [edx+00000008h], eax
loc_005A8090:   mov eax, var_58
loc_005A8093:   push eax
loc_005A8094:   mov [edx+0000000Ch], edi
loc_005A8097:   Call ebx
  loc_005A8099: mov ecx, 00000004h
  loc_005A809E: call [00401138h] ; __vbaI2I4
  loc_005A80A4: sub esp, 00000010h
loc_005A80A7:   mov var_1C, ax
loc_005A80AB:   mov edx, var_1C
loc_005A80AE:   mov ecx, esp
  loc_005A80B0: mov eax, 00000002h
  loc_005A80B5: push 00000028h
loc_005A80B7:   mov [ecx], eax
loc_005A80B9:   mov [ecx+00000004h], esi
loc_005A80BC:   mov [ecx+00000008h], edx
loc_005A80BF:   mov [ecx+0000000Ch], edi
loc_005A80C2:   mov eax, var_58
loc_005A80C5:   push eax
loc_005A80C6:   Call ebx
  loc_005A80C8: sub esp, 00000010h
  loc_005A80CB: mov ecx, 00000003h
loc_005A80D0:   mov edx, esp
  loc_005A80D2: mov eax, 00000002h
  loc_005A80D7: push 0000000Bh
loc_005A80D9:   mov [edx], ecx
loc_005A80DB:   mov [edx+00000004h], esi
loc_005A80DE:   mov [edx+00000008h], eax
loc_005A80E1:   mov eax, var_58
loc_005A80E4:   push eax
loc_005A80E5:   mov [edx+0000000Ch], edi
loc_005A80E8:   Call ebx
  loc_005A80EA: sub esp, 00000010h
  loc_005A80ED: mov ecx, 00000003h
loc_005A80F2:   mov edx, esp
  loc_005A80F4: xor eax, eax
  loc_005A80F6: push 0000000Ah
loc_005A80F8:   mov [edx], ecx
loc_005A80FA:   mov [edx+00000004h], esi
loc_005A80FD:   mov [edx+00000008h], eax
loc_005A8100:   mov eax, var_58
loc_005A8103:   push eax
loc_005A8104:   mov [edx+0000000Ch], edi
loc_005A8107:   Call ebx
  loc_005A8109: sub esp, 00000010h
  loc_005A810C: mov ecx, 00000008h
loc_005A8111:   mov edx, esp
  loc_005A8113: mov eax, 00435AFCh ; "Harga(Rp)"
  loc_005A8118: push 00000000h
loc_005A811A:   mov [edx], ecx
loc_005A811C:   mov [edx+00000004h], esi
loc_005A811F:   mov [edx+00000008h], eax
loc_005A8122:   mov eax, var_58
loc_005A8125:   push eax
loc_005A8126:   mov [edx+0000000Ch], edi
loc_005A8129:   Call ebx
  loc_005A812B: sub esp, 00000010h
  loc_005A812E: mov ecx, 0000000Bh
loc_005A8133:   mov edx, esp
  loc_005A8135: or eax, FFFFFFFFh
  loc_005A8138: push 0000004Fh
loc_005A813A:   mov [edx], ecx
loc_005A813C:   mov [edx+00000004h], esi
loc_005A813F:   mov [edx+00000008h], eax
loc_005A8142:   mov eax, var_58
loc_005A8145:   push eax
loc_005A8146:   mov [edx+0000000Ch], edi
loc_005A8149:   Call ebx
  loc_005A814B: sub esp, 00000010h
  loc_005A814E: mov ecx, 00000003h
loc_005A8153:   mov edx, esp
  loc_005A8155: mov eax, 00000002h
  loc_005A815A: sub esp, 00000010h
loc_005A815D:   mov var_44, ecx
loc_005A8160:   mov [edx], ecx
loc_005A8162:   mov ecx, esp
  loc_005A8164: push 00000001h
  loc_005A8166: push 00000039h
loc_005A8168:   mov [edx+00000004h], esi
loc_005A816B:   mov [edx+00000008h], eax
loc_005A816E:   mov eax, var_40
loc_005A8171:   mov [edx+0000000Ch], edi
loc_005A8174:   mov edx, var_44
loc_005A8177:   mov [ecx], edx
loc_005A8179:   mov edx, var_38
loc_005A817C:   mov [ecx+00000004h], eax
  loc_005A817F: mov eax, 00000640h
loc_005A8184:   mov [ecx+00000008h], eax
loc_005A8187:   mov eax, var_58
loc_005A818A:   push eax
loc_005A818B:   mov [ecx+0000000Ch], edx
  loc_005A818E: call [0040117Ch] ; __vbaLateIdCallSt
  loc_005A8194: add esp, 0000001Ch
  loc_005A8197: mov eax, 00000001h
loc_005A819C:   mov edx, esp
  loc_005A819E: mov ecx, 00000003h
loc_005A81A3:   mov var_1C, eax
loc_005A81A6:   mov [edx], ecx
loc_005A81A8:   mov [edx+00000004h], esi
  loc_005A81AB: push 00000034h
loc_005A81AD:   mov [edx+00000008h], eax
loc_005A81B0:   mov eax, var_58
loc_005A81B3:   push eax
loc_005A81B4:   mov [edx+0000000Ch], edi
loc_005A81B7:   Call ebx
  loc_005A81B9: mov ecx, 00000004h
  loc_005A81BE: call [00401138h] ; __vbaI2I4
  loc_005A81C4: sub esp, 00000010h
loc_005A81C7:   mov var_1C, ax
loc_005A81CB:   mov edx, var_1C
loc_005A81CE:   mov ecx, esp
  loc_005A81D0: mov eax, 00000002h
  loc_005A81D5: push 00000028h
loc_005A81D7:   mov [ecx], eax
loc_005A81D9:   mov eax, var_58
loc_005A81DC:   push eax
loc_005A81DD:   mov [ecx+00000004h], esi
loc_005A81E0:   mov [ecx+00000008h], edx
loc_005A81E3:   mov [ecx+0000000Ch], edi
loc_005A81E6:   Call ebx
  loc_005A81E8: sub esp, 00000010h
  loc_005A81EB: mov eax, 00000003h
loc_005A81F0:   mov edx, esp
loc_005A81F2:   mov ecx, eax
  loc_005A81F4: push 0000000Bh
loc_005A81F6:   mov [edx], ecx
loc_005A81F8:   mov [edx+00000004h], esi
loc_005A81FB:   mov [edx+00000008h], eax
loc_005A81FE:   mov eax, var_58
loc_005A8201:   push eax
loc_005A8202:   mov [edx+0000000Ch], edi
loc_005A8205:   Call ebx
  loc_005A8207: sub esp, 00000010h
  loc_005A820A: mov ecx, 00000003h
loc_005A820F:   mov edx, esp
  loc_005A8211: xor eax, eax
  loc_005A8213: push 0000000Ah
loc_005A8215:   mov [edx], ecx
loc_005A8217:   mov [edx+00000004h], esi
loc_005A821A:   mov [edx+00000008h], eax
loc_005A821D:   mov eax, var_58
loc_005A8220:   push eax
loc_005A8221:   mov [edx+0000000Ch], edi
loc_005A8224:   Call ebx
  loc_005A8226: sub esp, 00000010h
  loc_005A8229: mov ecx, 00000008h
loc_005A822E:   mov edx, esp
  loc_005A8230: mov eax, 00433328h ; "Disc(" & Chr(37) & ")"
  loc_005A8235: push 00000000h
loc_005A8237:   mov [edx], ecx
loc_005A8239:   mov [edx+00000004h], esi
loc_005A823C:   mov [edx+00000008h], eax
loc_005A823F:   mov eax, var_58
loc_005A8242:   push eax
loc_005A8243:   mov [edx+0000000Ch], edi
loc_005A8246:   Call ebx
  loc_005A8248: sub esp, 00000010h
  loc_005A824B: mov ecx, 0000000Bh
loc_005A8250:   mov edx, esp
  loc_005A8252: or eax, FFFFFFFFh
  loc_005A8255: push 0000004Fh
loc_005A8257:   mov [edx], ecx
loc_005A8259:   mov [edx+00000004h], esi
loc_005A825C:   mov [edx+00000008h], eax
loc_005A825F:   mov eax, var_58
loc_005A8262:   push eax
loc_005A8263:   mov [edx+0000000Ch], edi
loc_005A8266:   Call ebx
  loc_005A8268: sub esp, 00000010h
  loc_005A826B: mov ecx, 00000003h
loc_005A8270:   mov edx, esp
loc_005A8272:   mov eax, ecx
  loc_005A8274: sub esp, 00000010h
loc_005A8277:   mov var_44, ecx
loc_005A827A:   mov [edx], ecx
loc_005A827C:   mov ecx, esp
loc_005A827E:   mov [edx+00000004h], esi
loc_005A8281:   mov [edx+00000008h], eax
loc_005A8284:   mov [edx+0000000Ch], edi
loc_005A8287:   mov edx, eax
loc_005A8289:   mov eax, var_40
loc_005A828C:   mov [ecx], edx
loc_005A828E:   mov edx, var_38
  loc_005A8291: push 00000001h
  loc_005A8293: push 00000039h
loc_005A8295:   mov [ecx+00000004h], eax
  loc_005A8298: mov eax, 000003E8h
loc_005A829D:   mov [ecx+00000008h], eax
loc_005A82A0:   mov eax, var_58
loc_005A82A3:   push eax
loc_005A82A4:   mov [ecx+0000000Ch], edx
  loc_005A82A7: call [0040117Ch] ; __vbaLateIdCallSt
  loc_005A82AD: add esp, 0000001Ch
  loc_005A82B0: mov ecx, 00000003h
loc_005A82B5:   mov edx, esp
  loc_005A82B7: mov eax, 00000001h
loc_005A82BC:   mov var_1C, eax
  loc_005A82BF: push 00000034h
loc_005A82C1:   mov [edx], ecx
loc_005A82C3:   mov [edx+00000004h], esi
loc_005A82C6:   mov [edx+00000008h], eax
loc_005A82C9:   mov eax, var_58
loc_005A82CC:   push eax
loc_005A82CD:   mov [edx+0000000Ch], edi
loc_005A82D0:   Call ebx
  loc_005A82D2: mov ecx, 00000004h
  loc_005A82D7: call [00401138h] ; __vbaI2I4
  loc_005A82DD: sub esp, 00000010h
loc_005A82E0:   mov var_1C, ax
loc_005A82E4:   mov edx, var_1C
loc_005A82E7:   mov ecx, esp
  loc_005A82E9: mov eax, 00000002h
  loc_005A82EE: push 00000028h
loc_005A82F0:   mov [ecx], eax
loc_005A82F2:   mov eax, var_58
loc_005A82F5:   push eax
loc_005A82F6:   mov [ecx+00000004h], esi
loc_005A82F9:   mov [ecx+00000008h], edx
loc_005A82FC:   mov [ecx+0000000Ch], edi
loc_005A82FF:   Call ebx
  loc_005A8301: sub esp, 00000010h
  loc_005A8304: mov ecx, 00000003h
loc_005A8309:   mov edx, esp
  loc_005A830B: mov eax, 00000004h
  loc_005A8310: push 0000000Bh
loc_005A8312:   mov [edx], ecx
loc_005A8314:   mov [edx+00000004h], esi
loc_005A8317:   mov [edx+00000008h], eax
loc_005A831A:   mov eax, var_58
loc_005A831D:   push eax
loc_005A831E:   mov [edx+0000000Ch], edi
loc_005A8321:   Call ebx
  loc_005A8323: sub esp, 00000010h
  loc_005A8326: mov ecx, 00000003h
loc_005A832B:   mov edx, esp
  loc_005A832D: xor eax, eax
  loc_005A832F: push 0000000Ah
loc_005A8331:   mov [edx], ecx
loc_005A8333:   mov [edx+00000004h], esi
loc_005A8336:   mov [edx+00000008h], eax
loc_005A8339:   mov eax, var_58
loc_005A833C:   push eax
loc_005A833D:   mov [edx+0000000Ch], edi
loc_005A8340:   Call ebx
  loc_005A8342: sub esp, 00000010h
  loc_005A8345: mov ecx, 00000008h
loc_005A834A:   mov edx, esp
  loc_005A834C: mov eax, 004340C4h ; "Qty"
  loc_005A8351: push 00000000h
loc_005A8353:   mov [edx], ecx
loc_005A8355:   mov [edx+00000004h], esi
loc_005A8358:   mov [edx+00000008h], eax
loc_005A835B:   mov eax, var_58
loc_005A835E:   push eax
loc_005A835F:   mov [edx+0000000Ch], edi
loc_005A8362:   Call ebx
  loc_005A8364: or eax, FFFFFFFFh
  loc_005A8367: sub esp, 00000010h
  loc_005A836A: mov ecx, 0000000Bh
loc_005A836F:   mov edx, esp
loc_005A8371:   mov [edx], ecx
  loc_005A8373: push 0000004Fh
loc_005A8375:   mov [edx+00000004h], esi
loc_005A8378:   mov [edx+00000008h], eax
loc_005A837B:   mov eax, var_58
loc_005A837E:   push eax
loc_005A837F:   mov [edx+0000000Ch], edi
loc_005A8382:   Call ebx
  loc_005A8384: sub esp, 00000010h
  loc_005A8387: mov ecx, 00000003h
loc_005A838C:   mov edx, esp
  loc_005A838E: mov eax, 00000004h
  loc_005A8393: sub esp, 00000010h
loc_005A8396:   mov var_44, ecx
loc_005A8399:   mov [edx], ecx
loc_005A839B:   mov ecx, esp
  loc_005A839D: push 00000001h
  loc_005A839F: push 00000039h
loc_005A83A1:   mov [edx+00000004h], esi
loc_005A83A4:   mov [edx+00000008h], eax
loc_005A83A7:   mov eax, var_40
loc_005A83AA:   mov [edx+0000000Ch], edi
loc_005A83AD:   mov edx, var_44
loc_005A83B0:   mov [ecx], edx
loc_005A83B2:   mov edx, var_38
loc_005A83B5:   mov [ecx+00000004h], eax
  loc_005A83B8: mov eax, 00000320h
loc_005A83BD:   mov [ecx+00000008h], eax
loc_005A83C0:   mov eax, var_58
loc_005A83C3:   push eax
loc_005A83C4:   mov [ecx+0000000Ch], edx
  loc_005A83C7: call [0040117Ch] ; __vbaLateIdCallSt
  loc_005A83CD: add esp, 0000001Ch
  loc_005A83D0: mov ecx, 00000003h
loc_005A83D5:   mov edx, esp
  loc_005A83D7: mov eax, 00000001h
loc_005A83DC:   mov var_1C, eax
  loc_005A83DF: push 00000034h
loc_005A83E1:   mov [edx], ecx
loc_005A83E3:   mov [edx+00000004h], esi
loc_005A83E6:   mov [edx+00000008h], eax
loc_005A83E9:   mov eax, var_58
loc_005A83EC:   push eax
loc_005A83ED:   mov [edx+0000000Ch], edi
loc_005A83F0:   Call ebx
  loc_005A83F2: mov ecx, 00000004h
  loc_005A83F7: call [00401138h] ; __vbaI2I4
  loc_005A83FD: sub esp, 00000010h
loc_005A8400:   mov var_1C, ax
loc_005A8404:   mov edx, var_1C
loc_005A8407:   mov ecx, esp
  loc_005A8409: mov eax, 00000002h
  loc_005A840E: push 00000028h
loc_005A8410:   mov [ecx], eax
loc_005A8412:   mov eax, var_58
loc_005A8415:   push eax
loc_005A8416:   mov [ecx+00000004h], esi
loc_005A8419:   mov [ecx+00000008h], edx
loc_005A841C:   mov [ecx+0000000Ch], edi
loc_005A841F:   Call ebx
  loc_005A8421: sub esp, 00000010h
  loc_005A8424: mov ecx, 00000003h
loc_005A8429:   mov edx, esp
  loc_005A842B: mov eax, 00000005h
  loc_005A8430: push 0000000Bh
loc_005A8432:   mov [edx], ecx
loc_005A8434:   mov [edx+00000004h], esi
loc_005A8437:   mov [edx+00000008h], eax
loc_005A843A:   mov eax, var_58
loc_005A843D:   push eax
loc_005A843E:   mov [edx+0000000Ch], edi
loc_005A8441:   Call ebx
  loc_005A8443: xor eax, eax
  loc_005A8445: sub esp, 00000010h
loc_005A8448:   mov edx, esp
  loc_005A844A: mov ecx, 00000003h
  loc_005A844F: push 0000000Ah
loc_005A8451:   mov [edx], ecx
loc_005A8453:   mov [edx+00000004h], esi
loc_005A8456:   mov [edx+00000008h], eax
loc_005A8459:   mov [edx+0000000Ch], edi
loc_005A845C:   mov eax, var_58
loc_005A845F:   push eax
loc_005A8460:   Call ebx
  loc_005A8462: sub esp, 00000010h
  loc_005A8465: mov ecx, 00000008h
loc_005A846A:   mov edx, esp
  loc_005A846C: mov eax, 004340D0h ; "Sub Total(Rp)"
  loc_005A8471: push 00000000h
loc_005A8473:   mov [edx], ecx
loc_005A8475:   mov [edx+00000004h], esi
loc_005A8478:   mov [edx+00000008h], eax
loc_005A847B:   mov eax, var_58
loc_005A847E:   push eax
loc_005A847F:   mov [edx+0000000Ch], edi
loc_005A8482:   Call ebx
  loc_005A8484: sub esp, 00000010h
  loc_005A8487: mov ecx, 0000000Bh
loc_005A848C:   mov edx, esp
  loc_005A848E: or eax, FFFFFFFFh
  loc_005A8491: push 0000004Fh
loc_005A8493:   mov [edx], ecx
loc_005A8495:   mov [edx+00000004h], esi
loc_005A8498:   mov [edx+00000008h], eax
loc_005A849B:   mov eax, var_58
loc_005A849E:   push eax
loc_005A849F:   mov [edx+0000000Ch], edi
loc_005A84A2:   Call ebx
  loc_005A84A4: sub esp, 00000010h
  loc_005A84A7: mov ecx, 00000003h
loc_005A84AC:   mov edx, esp
  loc_005A84AE: mov eax, 00000005h
  loc_005A84B3: sub esp, 00000010h
loc_005A84B6:   mov var_44, ecx
loc_005A84B9:   mov [edx], ecx
loc_005A84BB:   mov ecx, esp
  loc_005A84BD: push 00000001h
  loc_005A84BF: push 00000039h
loc_005A84C1:   mov [edx+00000004h], esi
loc_005A84C4:   mov [edx+00000008h], eax
loc_005A84C7:   mov eax, var_40
loc_005A84CA:   mov [edx+0000000Ch], edi
loc_005A84CD:   mov edx, var_44
loc_005A84D0:   mov [ecx], edx
loc_005A84D2:   mov edx, var_38
loc_005A84D5:   mov [ecx+00000004h], eax
  loc_005A84D8: mov eax, 000007D0h
loc_005A84DD:   mov [ecx+00000008h], eax
loc_005A84E0:   mov eax, var_58
loc_005A84E3:   push eax
loc_005A84E4:   mov [ecx+0000000Ch], edx
  loc_005A84E7: call [0040117Ch] ; __vbaLateIdCallSt
  loc_005A84ED: add esp, 0000001Ch
  loc_005A84F0: mov ecx, 00000003h
loc_005A84F5:   mov edx, esp
  loc_005A84F7: mov eax, 00000001h
loc_005A84FC:   mov var_1C, eax
  loc_005A84FF: push 00000034h
loc_005A8501:   mov [edx], ecx
loc_005A8503:   mov [edx+00000004h], esi
loc_005A8506:   mov [edx+00000008h], eax
loc_005A8509:   mov eax, var_58
loc_005A850C:   push eax
loc_005A850D:   mov [edx+0000000Ch], edi
loc_005A8510:   Call ebx
  loc_005A8512: mov ecx, 00000004h
  loc_005A8517: call [00401138h] ; __vbaI2I4
  loc_005A851D: sub esp, 00000010h
loc_005A8520:   mov var_1C, ax
loc_005A8524:   mov edx, var_1C
loc_005A8527:   mov ecx, esp
  loc_005A8529: mov eax, 00000002h
  loc_005A852E: push 00000028h
loc_005A8530:   mov [ecx], eax
loc_005A8532:   mov eax, var_58
loc_005A8535:   push eax
loc_005A8536:   mov [ecx+00000004h], esi
loc_005A8539:   mov [ecx+00000008h], edx
loc_005A853C:   mov [ecx+0000000Ch], edi
loc_005A853F:   Call ebx
  loc_005A8541: mov eax, 00000006h
  loc_005A8546: mov ecx, 00000003h
  loc_005A854B: sub esp, 00000010h
loc_005A854E:   mov edx, esp
  loc_005A8550: push 0000000Bh
loc_005A8552:   mov [edx], ecx
loc_005A8554:   mov [edx+00000004h], esi
loc_005A8557:   mov [edx+00000008h], eax
loc_005A855A:   mov eax, var_58
loc_005A855D:   push eax
loc_005A855E:   mov [edx+0000000Ch], edi
loc_005A8561:   Call ebx
  loc_005A8563: sub esp, 00000010h
  loc_005A8566: mov ecx, 00000003h
loc_005A856B:   mov edx, esp
  loc_005A856D: xor eax, eax
  loc_005A856F: push 0000000Ah
loc_005A8571:   mov [edx], ecx
loc_005A8573:   mov [edx+00000004h], esi
loc_005A8576:   mov [edx+00000008h], eax
loc_005A8579:   mov eax, var_58
loc_005A857C:   push eax
loc_005A857D:   mov [edx+0000000Ch], edi
loc_005A8580:   Call ebx
  loc_005A8582: sub esp, 00000010h
  loc_005A8585: mov ecx, 00000008h
loc_005A858A:   mov edx, esp
  loc_005A858C: mov eax, 00435B14h ; "Harga Beli(Rp)"
  loc_005A8591: push 00000000h
loc_005A8593:   mov [edx], ecx
loc_005A8595:   mov [edx+00000004h], esi
loc_005A8598:   mov [edx+00000008h], eax
loc_005A859B:   mov eax, var_58
loc_005A859E:   push eax
loc_005A859F:   mov [edx+0000000Ch], edi
loc_005A85A2:   Call ebx
  loc_005A85A4: sub esp, 00000010h
  loc_005A85A7: mov ecx, 0000000Bh
loc_005A85AC:   mov edx, esp
  loc_005A85AE: or eax, FFFFFFFFh
  loc_005A85B1: push 0000004Fh
loc_005A85B3:   mov [edx], ecx
loc_005A85B5:   mov [edx+00000004h], esi
loc_005A85B8:   mov [edx+00000008h], eax
loc_005A85BB:   mov eax, var_58
loc_005A85BE:   push eax
loc_005A85BF:   mov [edx+0000000Ch], edi
loc_005A85C2:   Call ebx
  loc_005A85C4: sub esp, 00000010h
  loc_005A85C7: mov ecx, 00000003h
loc_005A85CC:   mov edx, esp
  loc_005A85CE: mov eax, 00000005h
  loc_005A85D3: sub esp, 00000010h
loc_005A85D6:   mov var_44, ecx
loc_005A85D9:   mov [edx], ecx
loc_005A85DB:   mov ecx, esp
  loc_005A85DD: push 00000001h
  loc_005A85DF: push 00000039h
loc_005A85E1:   mov [edx+00000004h], esi
loc_005A85E4:   mov [edx+00000008h], eax
loc_005A85E7:   mov eax, var_40
loc_005A85EA:   mov [edx+0000000Ch], edi
loc_005A85ED:   mov edx, var_44
loc_005A85F0:   mov [ecx], edx
loc_005A85F2:   mov edx, var_38
loc_005A85F5:   mov [ecx+00000004h], eax
  loc_005A85F8: mov eax, 000007D0h
loc_005A85FD:   mov [ecx+00000008h], eax
loc_005A8600:   mov eax, var_58
loc_005A8603:   push eax
loc_005A8604:   mov [ecx+0000000Ch], edx
  loc_005A8607: call [0040117Ch] ; __vbaLateIdCallSt
  loc_005A860D: add esp, 0000001Ch
  loc_005A8610: mov ecx, 00000003h
loc_005A8615:   mov edx, esp
  loc_005A8617: mov eax, 00000001h
loc_005A861C:   mov var_1C, eax
  loc_005A861F: push 00000034h
loc_005A8621:   mov [edx], ecx
loc_005A8623:   mov [edx+00000004h], esi
loc_005A8626:   mov [edx+00000008h], eax
loc_005A8629:   mov eax, var_58
loc_005A862C:   mov [edx+0000000Ch], edi
loc_005A862F:   push eax
loc_005A8630:   Call ebx
  loc_005A8632: mov ecx, 00000004h
  loc_005A8637: call [00401138h] ; __vbaI2I4
  loc_005A863D: sub esp, 00000010h
loc_005A8640:   mov var_1C, ax
loc_005A8644:   mov edx, var_1C
loc_005A8647:   mov ecx, esp
  loc_005A8649: mov eax, 00000002h
  loc_005A864E: push 00000028h
loc_005A8650:   mov [ecx], eax
loc_005A8652:   mov eax, var_58
loc_005A8655:   push eax
loc_005A8656:   mov [ecx+00000004h], esi
loc_005A8659:   mov [ecx+00000008h], edx
loc_005A865C:   mov [ecx+0000000Ch], edi
loc_005A865F:   Call ebx
  loc_005A8661: sub esp, 00000010h
  loc_005A8664: mov ecx, 00000003h
loc_005A8669:   mov edx, esp
  loc_005A866B: mov eax, 00000007h
  loc_005A8670: push 0000000Bh
loc_005A8672:   mov [edx], ecx
loc_005A8674:   mov [edx+00000004h], esi
loc_005A8677:   mov [edx+00000008h], eax
loc_005A867A:   mov eax, var_58
loc_005A867D:   push eax
loc_005A867E:   mov [edx+0000000Ch], edi
loc_005A8681:   Call ebx
  loc_005A8683: sub esp, 00000010h
  loc_005A8686: mov ecx, 00000003h
loc_005A868B:   mov edx, esp
  loc_005A868D: xor eax, eax
  loc_005A868F: push 0000000Ah
loc_005A8691:   mov [edx], ecx
loc_005A8693:   mov [edx+00000004h], esi
loc_005A8696:   mov [edx+00000008h], eax
loc_005A8699:   mov eax, var_58
loc_005A869C:   push eax
loc_005A869D:   mov [edx+0000000Ch], edi
loc_005A86A0:   Call ebx
  loc_005A86A2: sub esp, 00000010h
  loc_005A86A5: mov ecx, 00000008h
loc_005A86AA:   mov edx, esp
  loc_005A86AC: mov eax, 00435B38h ; "Laba(Rp)"
  loc_005A86B1: push 00000000h
loc_005A86B3:   mov [edx], ecx
loc_005A86B5:   mov [edx+00000004h], esi
loc_005A86B8:   mov [edx+00000008h], eax
loc_005A86BB:   mov eax, var_58
loc_005A86BE:   push eax
loc_005A86BF:   mov [edx+0000000Ch], edi
loc_005A86C2:   Call ebx
  loc_005A86C4: sub esp, 00000010h
  loc_005A86C7: mov ecx, 0000000Bh
loc_005A86CC:   mov edx, esp
  loc_005A86CE: or eax, FFFFFFFFh
  loc_005A86D1: push 0000004Fh
loc_005A86D3:   mov [edx], ecx
loc_005A86D5:   mov [edx+00000004h], esi
loc_005A86D8:   mov [edx+00000008h], eax
loc_005A86DB:   mov eax, var_58
loc_005A86DE:   push eax
loc_005A86DF:   mov [edx+0000000Ch], edi
loc_005A86E2:   Call ebx
  loc_005A86E4: sub esp, 00000010h
  loc_005A86E7: mov ecx, 00000003h
loc_005A86EC:   mov edx, esp
  loc_005A86EE: mov eax, 00000005h
  loc_005A86F3: sub esp, 00000010h
loc_005A86F6:   mov var_44, ecx
loc_005A86F9:   mov [edx], ecx
loc_005A86FB:   mov ecx, esp
loc_005A86FD:   mov [edx+00000004h], esi
loc_005A8700:   mov [edx+00000008h], eax
loc_005A8703:   mov eax, var_40
loc_005A8706:   mov [edx+0000000Ch], edi
loc_005A8709:   mov edx, var_44
loc_005A870C:   mov [ecx], edx
loc_005A870E:   mov [ecx+00000004h], eax
  loc_005A8711: mov eax, 000007D0h
loc_005A8716:   mov edx, var_38
loc_005A8719:   mov [ecx+00000008h], eax
loc_005A871C:   mov eax, var_58
  loc_005A871F: push 00000001h
  loc_005A8721: push 00000039h
loc_005A8723:   push eax
loc_005A8724:   mov [ecx+0000000Ch], edx
  loc_005A8727: call [0040117Ch] ; __vbaLateIdCallSt
  loc_005A872D: add esp, 0000001Ch
  loc_005A8730: mov ecx, 00000003h
loc_005A8735:   mov edx, esp
  loc_005A8737: mov eax, 00000001h
loc_005A873C:   mov var_1C, eax
  loc_005A873F: push 00000034h
loc_005A8741:   mov [edx], ecx
loc_005A8743:   mov [edx+00000004h], esi
loc_005A8746:   mov [edx+00000008h], eax
loc_005A8749:   mov eax, var_58
loc_005A874C:   push eax
loc_005A874D:   mov [edx+0000000Ch], edi
loc_005A8750:   Call ebx
  loc_005A8752: mov ecx, 00000004h
  loc_005A8757: call [00401138h] ; __vbaI2I4
  loc_005A875D: sub esp, 00000010h
loc_005A8760:   mov var_1C, ax
loc_005A8764:   mov edx, var_1C
loc_005A8767:   mov ecx, esp
  loc_005A8769: mov eax, 00000002h
  loc_005A876E: push 00000028h
loc_005A8770:   mov [ecx], eax
loc_005A8772:   mov eax, var_58
loc_005A8775:   push eax
loc_005A8776:   mov [ecx+00000004h], esi
loc_005A8779:   mov [ecx+00000008h], edx
loc_005A877C:   mov [ecx+0000000Ch], edi
loc_005A877F:   Call ebx
loc_005A8781:   lea ecx, var_58
  loc_005A8784: push 00000000h
loc_005A8786:   push ecx
  loc_005A8787: call [004010C8h] ; __vbaObjSetAddref
  loc_005A878D: push 005A879Ch
loc_005A8792:   lea ecx, var_58
  loc_005A8795: call [004012A0h] ; __vbaFreeObj
loc_005A879B:   ret
loc_005A879C:   mov eax, Me
loc_005A879F:   push eax
loc_005A87A0:   mov edx, [eax]
loc_005A87A2:   Call [edx+00000008h]
loc_005A87A5:   mov eax, var_4
loc_005A87A8:   mov ecx, var_14
loc_005A87AB:   pop edi
loc_005A87AC:   pop esi
loc_005A87AD:   mov fs: [00000000h] , ecx
loc_005A87B4:   pop ebx
loc_005A87B5:   mov esp, ebp
loc_005A87B7:   pop ebp
  loc_005A87B8: retn 0004h
End Sub

Public Sub DataPelanggan() '5A87C0
loc_005A87C0:   push ebp
loc_005A87C1:   mov ebp, esp
  loc_005A87C3: sub esp, 0000000Ch
  loc_005A87C6: push 00405696h ; __vbaExceptHandler
loc_005A87CB:   mov eax, fs: [00000000h]
loc_005A87D1:   push eax
loc_005A87D2:   mov fs: [00000000h] , esp
  loc_005A87D9: sub esp, 000000D0h
loc_005A87DF:   push ebx
loc_005A87E0:   push esi
loc_005A87E1:   push edi
loc_005A87E2:   mov var_C, esp
  loc_005A87E5: mov var_8, 00402DD0h
  loc_005A87EC: xor edi, edi
loc_005A87EE:   mov var_4, edi
loc_005A87F1:   mov eax, Me
loc_005A87F4:   push eax
loc_005A87F5:   mov ecx, [eax]
loc_005A87F7:   Call [ecx+00000004h]
  loc_005A87FA: mov edx, 00435B50h ; "SELECT * FROM Pelanggan"
  loc_005A87FF: mov ecx, 0066809Ch
loc_005A8804:   mov var_18, edi
loc_005A8807:   mov var_1C, edi
loc_005A880A:   mov var_20, edi
loc_005A880D:   mov var_24, edi
loc_005A8810:   mov var_28, edi
loc_005A8813:   mov var_38, edi
loc_005A8816:   mov var_48, edi
loc_005A8819:   mov var_58, edi
loc_005A881C:   mov var_68, edi
loc_005A881F:   mov var_78, edi
loc_005A8822:   mov var_88, edi
loc_005A8828:   mov var_98, edi
loc_005A882E:   mov var_AC, edi
loc_005A8834:   mov var_B0, edi
loc_005A883A:   mov var_DC, edi
  loc_005A8840: call [004011FCh] ; __vbaStrCopy
  loc_005A8846: push 0042DF30h
  loc_005A884B: call [00401168h] ; __vbaNew
loc_005A8851:   push eax
  loc_005A8852: push 0066805Ch
  loc_005A8857: call [004010B8h] ; __vbaObjSet
loc_005A885D:   cmp [0066805Ch], edi
  loc_005A8863: jnz 005A8875h
  loc_005A8865: push 0066805Ch
  loc_005A886A: push 0042DF30h
  loc_005A886F: call [004011E8h] ; __vbaNew2
loc_005A8875:   mov eax, [0066803Ch]
loc_005A887A:   mov esi, [0066805Ch]
loc_005A8880:   cmp eax, edi
  loc_005A8882: jnz 005A8894h
  loc_005A8884: push 0066803Ch
  loc_005A8889: push 0042DEFCh
  loc_005A888E: call [004011E8h] ; __vbaNew2
loc_005A8894:   push FFFFFFFFh
  loc_005A8896: push 00000004h
  loc_005A8898: push 00000002h
loc_005A889A:   mov eax, [0066803Ch]
  loc_005A889F: sub esp, 00000010h
  loc_005A88A2: mov ecx, 00000009h
loc_005A88A7:   mov ebx, esp
loc_005A88A9:   mov var_88, ecx
loc_005A88AF:   mov var_80, eax
  loc_005A88B2: sub esp, 00000010h
loc_005A88B5:   mov [ebx], ecx
loc_005A88B7:   mov ecx, var_84
loc_005A88BD:   mov edx, [0066809Ch]
  loc_005A88C3: mov var_78, 00000008h
loc_005A88CA:   mov [ebx+00000004h], ecx
loc_005A88CD:   mov ecx, esp
loc_005A88CF:   mov var_70, edx
loc_005A88D2:   mov edi, [esi]
loc_005A88D4:   mov [ebx+00000008h], eax
loc_005A88D7:   mov eax, var_7C
loc_005A88DA:   push esi
loc_005A88DB:   mov [ebx+0000000Ch], eax
loc_005A88DE:   mov eax, var_78
loc_005A88E1:   mov [ecx], eax
loc_005A88E3:   mov eax, var_74
loc_005A88E6:   mov [ecx+00000004h], eax
loc_005A88E9:   mov [ecx+00000008h], edx
loc_005A88EC:   mov edx, var_6C
loc_005A88EF:   mov [ecx+0000000Ch], edx
loc_005A88F2:   Call [edi+000000A0h]
loc_005A88F8:   test eax, eax
loc_005A88FA:   fnclex
  loc_005A88FC: jge 005A8914h
  loc_005A88FE: push 000000A0h
  loc_005A8903: push 0042DF44h
loc_005A8908:   push esi
  loc_005A8909: mov esi, [00401080h] ; __vbaHresultCheckObj
loc_005A890F:   push eax
  loc_005A8910: call __vbaHresultCheckObj
  loc_005A8912: jmp 005A891Ah
  loc_005A8914: mov esi, [00401080h] ; __vbaHresultCheckObj
loc_005A891A:   mov eax, [0066805Ch]
loc_005A891F:   test eax, eax
  loc_005A8921: jnz 005A8933h
  loc_005A8923: push 0066805Ch
  loc_005A8928: push 0042DF30h
  loc_005A892D: call [004011E8h] ; __vbaNew2
loc_005A8933:   mov edi, [0066805Ch]
loc_005A8939:   push FFFFFFFFh
loc_005A893B:   push edi
loc_005A893C:   mov eax, [edi]
loc_005A893E:   Call [eax+000000A4h]
loc_005A8944:   test eax, eax
loc_005A8946:   fnclex
  loc_005A8948: jge 005A8958h
  loc_005A894A: push 000000A4h
  loc_005A894F: push 0042DF44h
loc_005A8954:   push edi
loc_005A8955:   push eax
  loc_005A8956: call __vbaHresultCheckObj
loc_005A8958:   mov eax, [0066805Ch]
loc_005A895D:   test eax, eax
  loc_005A895F: jnz 005A8971h
  loc_005A8961: push 0066805Ch
  loc_005A8966: push 0042DF30h
  loc_005A896B: call [004011E8h] ; __vbaNew2
loc_005A8971:   mov ecx, [0066805Ch]
loc_005A8977:   lea edx, var_DC
loc_005A897D:   push ecx
loc_005A897E:   push edx
  loc_005A897F: call [004010C8h] ; __vbaObjSetAddref
loc_005A8985:   mov eax, var_DC
loc_005A898B:   lea edx, var_AC
loc_005A8991:   push edx
loc_005A8992:   push eax
loc_005A8993:   mov ecx, [eax]
loc_005A8995:   Call [ecx+00000050h]
loc_005A8998:   test eax, eax
loc_005A899A:   fnclex
  loc_005A899C: jge 005A89AFh
loc_005A899E:   mov ecx, var_DC
  loc_005A89A4: push 00000050h
  loc_005A89A6: push 0042DF44h
loc_005A89AB:   push ecx
loc_005A89AC:   push eax
  loc_005A89AD: call __vbaHresultCheckObj
loc_005A89AF:   mov eax, var_DC
loc_005A89B5:   lea ecx, var_B0
loc_005A89BB:   push ecx
loc_005A89BC:   push eax
loc_005A89BD:   mov edx, [eax]
loc_005A89BF:   Call [edx+00000034h]
loc_005A89C2:   test eax, eax
loc_005A89C4:   fnclex
  loc_005A89C6: jge 005A89D9h
loc_005A89C8:   mov edx, var_DC
  loc_005A89CE: push 00000034h
  loc_005A89D0: push 0042DF44h
loc_005A89D5:   push edx
loc_005A89D6:   push eax
  loc_005A89D7: call __vbaHresultCheckObj
  loc_005A89D9: xor eax, eax
loc_005A89DB:   cmp var_B0, ax
loc_005A89E2:   setz al
  loc_005A89E5: xor ecx, ecx
loc_005A89E7:   cmp var_AC, cx
loc_005A89EE:   setz cl
  loc_005A89F1: or eax, ecx
  loc_005A89F3: jnz 005A8A7Ch
  loc_005A89F9: mov esi, [00401238h] ; __vbaVarDup
  loc_005A89FF: mov ecx, 80020004h
loc_005A8A04:   mov var_60, ecx
  loc_005A8A07: mov eax, 0000000Ah
loc_005A8A0C:   mov var_50, ecx
  loc_005A8A0F: mov edi, 00000008h
loc_005A8A14:   lea edx, var_88
loc_005A8A1A:   lea ecx, var_48
loc_005A8A1D:   mov var_68, eax
loc_005A8A20:   mov var_58, eax
  loc_005A8A23: mov var_80, 0042E624h ; "Error"
loc_005A8A2A:   mov var_88, edi
  loc_005A8A30: call __vbaVarDup
loc_005A8A32:   lea edx, var_78
loc_005A8A35:   lea ecx, var_38
  loc_005A8A38: mov var_70, 00435B84h ; "DATA PELANGGAN TIDAK ADA"
loc_005A8A3F:   mov var_78, edi
  loc_005A8A42: call __vbaVarDup
loc_005A8A44:   lea edx, var_68
loc_005A8A47:   lea eax, var_58
loc_005A8A4A:   push edx
loc_005A8A4B:   lea ecx, var_48
loc_005A8A4E:   push eax
loc_005A8A4F:   push ecx
loc_005A8A50:   lea edx, var_38
  loc_005A8A53: push 00000010h
loc_005A8A55:   push edx
  loc_005A8A56: call [004010B4h] ; rtcMsgBox
loc_005A8A5C:   lea eax, var_68
loc_005A8A5F:   lea ecx, var_58
loc_005A8A62:   push eax
loc_005A8A63:   lea edx, var_48
loc_005A8A66:   push ecx
loc_005A8A67:   lea eax, var_38
loc_005A8A6A:   push edx
loc_005A8A6B:   push eax
  loc_005A8A6C: push 00000004h
  loc_005A8A6E: call [0040103Ch] ; __vbaFreeVarList
  loc_005A8A74: add esp, 00000014h
  loc_005A8A77: jmp 005A8D16h
loc_005A8A7C:   mov eax, Me
  loc_005A8A7F: push 00000000h
loc_005A8A81:   push FFFFFDD6h
loc_005A8A86:   push eax
loc_005A8A87:   mov ecx, [eax]
loc_005A8A89:   Call [ecx+0000038Ch]
loc_005A8A8F:   lea edx, var_18
loc_005A8A92:   push eax
loc_005A8A93:   push edx
  loc_005A8A94: call [004010B8h] ; __vbaObjSet
loc_005A8A9A:   push eax
  loc_005A8A9B: call [0040102Ch] ; __vbaLateIdCall
  loc_005A8AA1: add esp, 0000000Ch
loc_005A8AA4:   lea ecx, var_18
  loc_005A8AA7: call [004012A0h] ; __vbaFreeObj
loc_005A8AAD:   mov eax, var_DC
loc_005A8AB3:   lea edx, var_AC
loc_005A8AB9:   push edx
loc_005A8ABA:   push eax
loc_005A8ABB:   mov ecx, [eax]
loc_005A8ABD:   Call [ecx+00000050h]
loc_005A8AC0:   test eax, eax
loc_005A8AC2:   fnclex
  loc_005A8AC4: jge 005A8AD7h
loc_005A8AC6:   mov ecx, var_DC
  loc_005A8ACC: push 00000050h
  loc_005A8ACE: push 0042DF44h
loc_005A8AD3:   push ecx
loc_005A8AD4:   push eax
  loc_005A8AD5: call __vbaVarDup
  loc_005A8AD7: cmp var_AC, 0000h
loc_005A8ADF:   mov eax, var_DC
  loc_005A8AE5: jnz 005A8CF3h
loc_005A8AEB:   mov edx, [eax]
loc_005A8AED:   lea ecx, var_18
loc_005A8AF0:   push ecx
loc_005A8AF1:   push eax
loc_005A8AF2:   Call [edx+00000054h]
loc_005A8AF5:   test eax, eax
loc_005A8AF7:   fnclex
  loc_005A8AF9: jge 005A8B0Ch
loc_005A8AFB:   mov edx, var_DC
  loc_005A8B01: push 00000054h
  loc_005A8B03: push 0042DF44h
loc_005A8B08:   push edx
loc_005A8B09:   push eax
  loc_005A8B0A: call __vbaVarDup
loc_005A8B0C:   lea ebx, var_1C
loc_005A8B0F:   mov eax, var_18
loc_005A8B12:   push ebx
  loc_005A8B13: mov edx, 00000008h
  loc_005A8B18: sub esp, 00000010h
loc_005A8B1B:   mov var_78, edx
loc_005A8B1E:   mov ebx, esp
  loc_005A8B20: mov ecx, 0042F1B8h ; "kode_Pelanggan"
loc_005A8B25:   mov var_70, ecx
loc_005A8B28:   mov edi, [eax]
loc_005A8B2A:   mov [ebx], edx
loc_005A8B2C:   mov edx, var_74
loc_005A8B2F:   push eax
loc_005A8B30:   mov var_B8, eax
loc_005A8B36:   mov [ebx+00000004h], edx
loc_005A8B39:   mov [ebx+00000008h], ecx
loc_005A8B3C:   mov ecx, var_6C
loc_005A8B3F:   mov [ebx+0000000Ch], ecx
loc_005A8B42:   Call [edi+00000028h]
loc_005A8B45:   test eax, eax
loc_005A8B47:   fnclex
  loc_005A8B49: jge 005A8B5Ch
loc_005A8B4B:   mov edx, var_B8
  loc_005A8B51: push 00000028h
  loc_005A8B53: push 0042DFACh
loc_005A8B58:   push edx
loc_005A8B59:   push eax
  loc_005A8B5A: call __vbaVarDup
loc_005A8B5C:   mov eax, var_1C
loc_005A8B5F:   lea edx, var_38
loc_005A8B62:   push edx
loc_005A8B63:   push eax
loc_005A8B64:   mov ecx, [eax]
loc_005A8B66:   mov edi, eax
loc_005A8B68:   Call [ecx+00000034h]
loc_005A8B6B:   test eax, eax
loc_005A8B6D:   fnclex
  loc_005A8B6F: jge 005A8B7Ch
  loc_005A8B71: push 00000034h
  loc_005A8B73: push 0042DFBCh
loc_005A8B78:   push edi
loc_005A8B79:   push eax
  loc_005A8B7A: call __vbaVarDup
loc_005A8B7C:   mov eax, var_DC
loc_005A8B82:   lea edx, var_20
  loc_005A8B85: mov var_80, 004302FCh ; " | "
  loc_005A8B8C: mov var_88, 00000008h
loc_005A8B96:   mov ecx, [eax]
loc_005A8B98:   push edx
loc_005A8B99:   push eax
loc_005A8B9A:   Call [ecx+00000054h]
loc_005A8B9D:   test eax, eax
loc_005A8B9F:   fnclex
  loc_005A8BA1: jge 005A8BB4h
loc_005A8BA3:   mov ecx, var_DC
  loc_005A8BA9: push 00000054h
  loc_005A8BAB: push 0042DF44h
loc_005A8BB0:   push ecx
loc_005A8BB1:   push eax
  loc_005A8BB2: call __vbaVarDup
loc_005A8BB4:   lea ebx, var_24
loc_005A8BB7:   mov eax, var_20
loc_005A8BBA:   push ebx
  loc_005A8BBB: mov edx, 00000008h
  loc_005A8BC0: sub esp, 00000010h
loc_005A8BC3:   mov edi, [eax]
loc_005A8BC5:   mov ebx, esp
  loc_005A8BC7: mov ecx, 0042F358h ; "nama_pelanggan"
loc_005A8BCC:   push eax
loc_005A8BCD:   mov esi, eax
loc_005A8BCF:   mov [ebx], edx
loc_005A8BD1:   mov edx, var_94
loc_005A8BD7:   mov [ebx+00000004h], edx
loc_005A8BDA:   mov [ebx+00000008h], ecx
loc_005A8BDD:   mov ecx, var_8C
loc_005A8BE3:   mov [ebx+0000000Ch], ecx
loc_005A8BE6:   Call [edi+00000028h]
loc_005A8BE9:   test eax, eax
loc_005A8BEB:   fnclex
  loc_005A8BED: jge 005A8C02h
  loc_005A8BEF: push 00000028h
  loc_005A8BF1: push 0042DFACh
loc_005A8BF6:   push esi
  loc_005A8BF7: mov esi, [00401080h] ; __vbaHresultCheckObj
loc_005A8BFD:   push eax
  loc_005A8BFE: call __vbaHresultCheckObj
  loc_005A8C00: jmp 005A8C08h
  loc_005A8C02: mov esi, [00401080h] ; __vbaHresultCheckObj
loc_005A8C08:   mov eax, var_24
loc_005A8C0B:   lea ecx, var_58
loc_005A8C0E:   push ecx
loc_005A8C0F:   push eax
loc_005A8C10:   mov edx, [eax]
loc_005A8C12:   mov edi, eax
loc_005A8C14:   Call [edx+00000034h]
loc_005A8C17:   test eax, eax
loc_005A8C19:   fnclex
  loc_005A8C1B: jge 005A8C28h
  loc_005A8C1D: push 00000034h
  loc_005A8C1F: push 0042DFBCh
loc_005A8C24:   push edi
loc_005A8C25:   push eax
  loc_005A8C26: call __vbaHresultCheckObj
  loc_005A8C28: mov edi, [0040122Ch] ; __vbaVarAdd
loc_005A8C2E:   lea edx, var_38
loc_005A8C31:   lea eax, var_88
loc_005A8C37:   push edx
loc_005A8C38:   lea ecx, var_48
loc_005A8C3B:   push eax
loc_005A8C3C:   push ecx
loc_005A8C3D:   Call edi
loc_005A8C3F:   push eax
loc_005A8C40:   lea edx, var_58
loc_005A8C43:   lea eax, var_68
loc_005A8C46:   push edx
loc_005A8C47:   push eax
loc_005A8C48:   Call edi
loc_005A8C4A:   mov edx, [eax]
  loc_005A8C4C: sub esp, 00000010h
loc_005A8C4F:   mov ecx, esp
  loc_005A8C51: push 00000001h
loc_005A8C53:   push FFFFFDD7h
loc_005A8C58:   mov [ecx], edx
loc_005A8C5A:   mov edx, [eax+00000004h]
loc_005A8C5D:   mov [ecx+00000004h], edx
loc_005A8C60:   mov edx, [eax+00000008h]
loc_005A8C63:   mov eax, [eax+0000000Ch]
loc_005A8C66:   mov [ecx+00000008h], edx
loc_005A8C69:   mov [ecx+0000000Ch], eax
loc_005A8C6C:   mov eax, Me
loc_005A8C6F:   push eax
loc_005A8C70:   mov ecx, [eax]
loc_005A8C72:   Call [ecx+0000038Ch]
loc_005A8C78:   lea edx, var_28
loc_005A8C7B:   push eax
loc_005A8C7C:   push edx
  loc_005A8C7D: call [004010B8h] ; __vbaObjSet
loc_005A8C83:   push eax
  loc_005A8C84: call [0040102Ch] ; __vbaLateIdCall
loc_005A8C8A:   lea eax, var_28
loc_005A8C8D:   lea ecx, var_24
loc_005A8C90:   push eax
loc_005A8C91:   lea edx, var_20
loc_005A8C94:   push ecx
loc_005A8C95:   lea eax, var_1C
loc_005A8C98:   push edx
loc_005A8C99:   lea ecx, var_18
loc_005A8C9C:   push eax
loc_005A8C9D:   push ecx
  loc_005A8C9E: push 00000005h
  loc_005A8CA0: call [0040104Ch] ; __vbaFreeObjList
loc_005A8CA6:   lea edx, var_68
loc_005A8CA9:   lea eax, var_58
loc_005A8CAC:   push edx
loc_005A8CAD:   lea ecx, var_48
loc_005A8CB0:   push eax
loc_005A8CB1:   lea edx, var_38
loc_005A8CB4:   push ecx
loc_005A8CB5:   push edx
  loc_005A8CB6: push 00000004h
  loc_005A8CB8: call [0040103Ch] ; __vbaFreeVarList
loc_005A8CBE:   mov eax, var_DC
  loc_005A8CC4: add esp, 00000048h
loc_005A8CC7:   mov ecx, [eax]
loc_005A8CC9:   push eax
loc_005A8CCA:   Call [ecx+00000090h]
loc_005A8CD0:   test eax, eax
loc_005A8CD2:   fnclex
  loc_005A8CD4: jge 005A8AADh
loc_005A8CDA:   mov edx, var_DC
  loc_005A8CE0: push 00000090h
  loc_005A8CE5: push 0042DF44h
loc_005A8CEA:   push edx
loc_005A8CEB:   push eax
  loc_005A8CEC: call __vbaHresultCheckObj
  loc_005A8CEE: jmp 005A8AADh
loc_005A8CF3:   mov ecx, [eax]
loc_005A8CF5:   push eax
loc_005A8CF6:   Call [ecx+00000098h]
loc_005A8CFC:   test eax, eax
loc_005A8CFE:   fnclex
  loc_005A8D00: jge 005A8D16h
loc_005A8D02:   mov edx, var_DC
  loc_005A8D08: push 00000098h
  loc_005A8D0D: push 0042DF44h
loc_005A8D12:   push edx
loc_005A8D13:   push eax
  loc_005A8D14: call __vbaHresultCheckObj
loc_005A8D16:   lea eax, var_DC
  loc_005A8D1C: push 00000000h
loc_005A8D1E:   push eax
  loc_005A8D1F: call [004010C8h] ; __vbaObjSetAddref
  loc_005A8D25: push 005A8D71h
  loc_005A8D2A: jmp 005A8D64h
loc_005A8D2C:   lea ecx, var_28
loc_005A8D2F:   lea edx, var_24
loc_005A8D32:   push ecx
loc_005A8D33:   lea eax, var_20
loc_005A8D36:   push edx
loc_005A8D37:   lea ecx, var_1C
loc_005A8D3A:   push eax
loc_005A8D3B:   lea edx, var_18
loc_005A8D3E:   push ecx
loc_005A8D3F:   push edx
  loc_005A8D40: push 00000005h
  loc_005A8D42: call [0040104Ch] ; __vbaFreeObjList
loc_005A8D48:   lea eax, var_68
loc_005A8D4B:   lea ecx, var_58
loc_005A8D4E:   push eax
loc_005A8D4F:   lea edx, var_48
loc_005A8D52:   push ecx
loc_005A8D53:   lea eax, var_38
loc_005A8D56:   push edx
loc_005A8D57:   push eax
  loc_005A8D58: push 00000004h
  loc_005A8D5A: call [0040103Ch] ; __vbaFreeVarList
  loc_005A8D60: add esp, 0000002Ch
loc_005A8D63:   ret
loc_005A8D64:   lea ecx, var_DC
  loc_005A8D6A: call [004012A0h] ; __vbaFreeObj
loc_005A8D70:   ret
loc_005A8D71:   mov eax, Me
loc_005A8D74:   push eax
loc_005A8D75:   mov ecx, [eax]
loc_005A8D77:   Call [ecx+00000008h]
loc_005A8D7A:   mov eax, var_4
loc_005A8D7D:   mov ecx, var_14
loc_005A8D80:   pop edi
loc_005A8D81:   pop esi
loc_005A8D82:   mov fs: [00000000h] , ecx
loc_005A8D89:   pop ebx
loc_005A8D8A:   mov esp, ebp
loc_005A8D8C:   pop ebp
  loc_005A8D8D: retn 0004h
End Sub

Public Sub KurangiBarisGrid() '5AEDA0
loc_005AEDA0:   push ebp
loc_005AEDA1:   mov ebp, esp
  loc_005AEDA3: sub esp, 0000000Ch
  loc_005AEDA6: push 00405696h ; __vbaExceptHandler
loc_005AEDAB:   mov eax, fs: [00000000h]
loc_005AEDB1:   push eax
loc_005AEDB2:   mov fs: [00000000h] , esp
  loc_005AEDB9: sub esp, 00000030h
loc_005AEDBC:   push ebx
loc_005AEDBD:   push esi
loc_005AEDBE:   push edi
loc_005AEDBF:   mov var_C, esp
  loc_005AEDC2: mov var_8, 00402F58h
  loc_005AEDC9: xor edi, edi
loc_005AEDCB:   mov var_4, edi
loc_005AEDCE:   mov esi, Me
loc_005AEDD1:   push esi
loc_005AEDD2:   mov eax, [esi]
loc_005AEDD4:   Call [eax+00000004h]
loc_005AEDD7:   mov ax, [esi+00000034h]
loc_005AEDDB:   mov var_18, edi
  loc_005AEDDE: cmp ax, 0002h
loc_005AEDE2:   mov var_28, edi
  loc_005AEDE5: jz 005AEDF3h
  loc_005AEDE7: jl 005AEDF3h
  loc_005AEDE9: cmp ax, 0001h
  loc_005AEDED: jnz 005AEE86h
loc_005AEDF3:   mov ecx, [esi]
loc_005AEDF5:   push edi
  loc_005AEDF6: push 00000044h
loc_005AEDF8:   push esi
  loc_005AEDF9: mov [esi+00000034h], 0001h
loc_005AEDFF:   Call [ecx+00000390h]
  loc_005AEE05: mov edi, [004010B8h] ; __vbaObjSet
loc_005AEE0B:   lea edx, var_18
loc_005AEE0E:   push eax
loc_005AEE0F:   push edx
loc_005AEE10:   Call edi
loc_005AEE12:   push eax
  loc_005AEE13: call [0040102Ch] ; __vbaLateIdCall
  loc_005AEE19: mov ebx, [004012A0h] ; __vbaFreeObj
  loc_005AEE1F: add esp, 0000000Ch
loc_005AEE22:   lea ecx, var_18
loc_005AEE25:   Call ebx
  loc_005AEE27: sub esp, 00000010h
  loc_005AEE2A: mov ecx, 00000003h
loc_005AEE2F:   mov edx, esp
  loc_005AEE31: mov eax, 00000002h
  loc_005AEE36: push 00000004h
loc_005AEE38:   push esi
loc_005AEE39:   mov [edx], ecx
loc_005AEE3B:   mov ecx, var_24
loc_005AEE3E:   mov [edx+00000004h], ecx
loc_005AEE41:   mov ecx, [esi]
loc_005AEE43:   mov [edx+00000008h], eax
loc_005AEE46:   mov eax, var_1C
loc_005AEE49:   mov [edx+0000000Ch], eax
loc_005AEE4C:   Call [ecx+00000390h]
loc_005AEE52:   lea edx, var_18
loc_005AEE55:   push eax
loc_005AEE56:   push edx
loc_005AEE57:   Call edi
loc_005AEE59:   push eax
  loc_005AEE5A: call [00401280h] ; __vbaLateIdSt
loc_005AEE60:   lea ecx, var_18
loc_005AEE63:   Call ebx
loc_005AEE65:   mov eax, [esi]
loc_005AEE67:   push esi
loc_005AEE68:   Call [eax+00000714h]
loc_005AEE6E:   test eax, eax
  loc_005AEE70: jge 005AEE90h
  loc_005AEE72: push 00000714h
  loc_005AEE77: push 00434B78h
loc_005AEE7C:   push esi
loc_005AEE7D:   push eax
  loc_005AEE7E: call [00401080h] ; __vbaHresultCheckObj
  loc_005AEE84: jmp 005AEE90h
  loc_005AEE86: sub ax, 0001h
  loc_005AEE8A: jo 005AEEE0h
loc_005AEE8C:   mov [esi+00000034h], ax
loc_005AEE90:   mov ecx, [esi]
loc_005AEE92:   push esi
loc_005AEE93:   Call [ecx+000006F8h]
loc_005AEE99:   test eax, eax
  loc_005AEE9B: jge 005AEEAFh
  loc_005AEE9D: push 000006F8h
  loc_005AEEA2: push 00434B78h
loc_005AEEA7:   push esi
loc_005AEEA8:   push eax
  loc_005AEEA9: call [00401080h] ; __vbaHresultCheckObj
  loc_005AEEAF: push 005AEEC1h
  loc_005AEEB4: jmp 005AEEC0h
loc_005AEEB6:   lea ecx, var_18
  loc_005AEEB9: call [004012A0h] ; __vbaFreeObj
loc_005AEEBF:   ret
loc_005AEEC0:   ret
loc_005AEEC1:   mov eax, Me
loc_005AEEC4:   push eax
loc_005AEEC5:   mov edx, [eax]
loc_005AEEC7:   Call [edx+00000008h]
loc_005AEECA:   mov eax, var_4
loc_005AEECD:   mov ecx, var_14
loc_005AEED0:   pop edi
loc_005AEED1:   pop esi
loc_005AEED2:   mov fs: [00000000h] , ecx
loc_005AEED9:   pop ebx
loc_005AEEDA:   mov esp, ebp
loc_005AEEDC:   pop ebp
  loc_005AEEDD: retn 0004h
End Sub
