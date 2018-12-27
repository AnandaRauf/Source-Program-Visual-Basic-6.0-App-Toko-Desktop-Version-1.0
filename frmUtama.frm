VERSION 5.00
Begin VB.Form frmUtama 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Program POS Versi 1.0"
   ClientHeight    =   7395
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15120
   Icon            =   "frmUtama.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   7095
      Begin VB.Label Label3 
         BackColor       =   &H80000014&
         Caption         =   "23 December ©2018"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   480
         TabIndex        =   22
         Top             =   2040
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000014&
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   480
         TabIndex        =   21
         Top             =   1320
         Width           =   4335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Software App Toko"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   1095
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   6255
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   5400
      Top             =   5760
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   4920
      TabIndex        =   13
      Top             =   0
      Width           =   3255
      Begin VB.PictureBox CmdPembelian 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   1635
         TabIndex        =   14
         ToolTipText     =   "Transaksi Pembelian"
         Top             =   240
         Width           =   1695
      End
      Begin VB.PictureBox CmdPenjualan 
         Height          =   495
         Left            =   1320
         ScaleHeight     =   435
         ScaleWidth      =   1635
         TabIndex        =   15
         ToolTipText     =   "Transaksi Penjualan"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   8160
      TabIndex        =   10
      Top             =   0
      Width           =   4575
      Begin VB.PictureBox CmdLogOut 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   1635
         TabIndex        =   11
         ToolTipText     =   "Log Out"
         Top             =   240
         Width           =   1695
      End
      Begin VB.PictureBox CmdKeluar 
         Height          =   495
         Left            =   2760
         ScaleHeight     =   435
         ScaleWidth      =   1635
         TabIndex        =   12
         ToolTipText     =   "Keluar"
         Top             =   240
         Width           =   1695
      End
      Begin VB.PictureBox CmdLogin 
         Height          =   495
         Left            =   1200
         ScaleHeight     =   435
         ScaleWidth      =   1995
         TabIndex        =   19
         ToolTipText     =   "Keluar"
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   4815
      Begin VB.PictureBox CmdPelanggan 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   1635
         TabIndex        =   7
         ToolTipText     =   "Data Pelanggan"
         Top             =   240
         Width           =   1695
      End
      Begin VB.PictureBox CmdPemasok 
         Height          =   495
         Left            =   1320
         ScaleHeight     =   435
         ScaleWidth      =   1995
         TabIndex        =   8
         ToolTipText     =   "Data Pemasok"
         Top             =   240
         Width           =   2055
      End
      Begin VB.PictureBox CmdBarang 
         Height          =   495
         Left            =   2880
         ScaleHeight     =   435
         ScaleWidth      =   1635
         TabIndex        =   9
         ToolTipText     =   "Data Barang"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   7320
      TabIndex        =   1
      Top             =   1080
      Width           =   6135
      Begin VB.TextBox txtAlamat 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         MaxLength       =   35
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   960
         Width           =   4680
      End
      Begin VB.TextBox txtNama 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   3
         Top             =   360
         Width           =   4665
      End
      Begin VB.TextBox TxtKota 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   2
         Top             =   1560
         Width           =   4695
      End
      Begin VB.PictureBox txtTelepon 
         Height          =   615
         Left            =   1080
         ScaleHeight     =   555
         ScaleWidth      =   4635
         TabIndex        =   5
         Top             =   2160
         Width           =   4695
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   240
         Picture         =   "frmUtama.frx":0ECA
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   480
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   480
      Top             =   360
   End
   Begin VB.PictureBox StatusBar1 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   15060
      TabIndex        =   0
      Top             =   7020
      Width           =   15120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Software App Toko"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1095
      Index           =   1
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label TxtInfo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   6120
      Width           =   11895
   End
   Begin VB.Menu MuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLogin 
         Caption         =   "Log In"
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "Log Out"
      End
      Begin VB.Menu MnuGantiPass 
         Caption         =   "Ganti Password"
      End
      Begin VB.Menu Bts1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKeluar 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu MuData 
      Caption         =   "&Data"
      Begin VB.Menu mnuPelanggan 
         Caption         =   "Data Pe&langgan"
      End
      Begin VB.Menu mnuPemasok 
         Caption         =   "Data &Pemasok"
      End
      Begin VB.Menu Bts2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSatuan 
         Caption         =   "Satuan Barang"
      End
      Begin VB.Menu mnuGolongan 
         Caption         =   "Data &Golongan"
      End
      Begin VB.Menu mnuJenis 
         Caption         =   "Data &Jenis"
      End
      Begin VB.Menu mnuProduk 
         Caption         =   "Data Pro&duk"
      End
      Begin VB.Menu mnuBarang 
         Caption         =   "Data &Barang"
      End
      Begin VB.Menu Bts3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKontrolStok 
         Caption         =   "Kontrol Stock dan Harga Barang"
      End
   End
   Begin VB.Menu MuTransaksi 
      Caption         =   "&Transaksi"
      Begin VB.Menu mnuPembelian 
         Caption         =   "&Pembelian Barang"
      End
      Begin VB.Menu MnuReturPembelian 
         Caption         =   "Retur Pembelian"
      End
      Begin VB.Menu MnuHutang 
         Caption         =   "Hutang"
      End
      Begin VB.Menu Bts4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPenjualan 
         Caption         =   "Penjualan &Retail"
      End
      Begin VB.Menu MnuRetur 
         Caption         =   "Retur Penjualan"
      End
      Begin VB.Menu MnuPiutang 
         Caption         =   "Piutang"
      End
   End
   Begin VB.Menu MuLaporan 
      Caption         =   "&Laporan"
      Begin VB.Menu mnuLapPemasok 
         Caption         =   "Seluruh Pemasok (Supplier)"
      End
      Begin VB.Menu mnuLapPelanggan 
         Caption         =   "Seluruh Pelanggan (Customer)"
      End
      Begin VB.Menu mnuLapGolongan 
         Caption         =   "Seluruh Golongan"
      End
      Begin VB.Menu mnuLapJenis 
         Caption         =   "Seluruh Jenis Produk"
      End
      Begin VB.Menu mnuLapProduk 
         Caption         =   "Seluruh Produk Barang"
      End
      Begin VB.Menu mnuLapBarang 
         Caption         =   "Seluruh Barang"
      End
      Begin VB.Menu BtsLap1 
         Caption         =   "-"
      End
      Begin VB.Menu sbPembelian 
         Caption         =   "Transaksi Pembelian"
         Begin VB.Menu mnuLapSelBeli 
            Caption         =   "Pembelian Semuanya"
         End
         Begin VB.Menu mnuLapBeliNota 
            Caption         =   "Pembelian Per-Nota"
         End
         Begin VB.Menu mnuLapBeliPeriode 
            Caption         =   "Pembelian Per-Periode"
         End
      End
      Begin VB.Menu MnuLapReturBeli 
         Caption         =   "Retur Pembelian"
         Begin VB.Menu MnuLapReturBeliSemua 
            Caption         =   "Retur Semuanya"
         End
         Begin VB.Menu MnuLapReturBeliNota 
            Caption         =   "Retur Per-Nota"
         End
         Begin VB.Menu MnuLapReturBeliPeriode 
            Caption         =   "Retur Per-Periode"
         End
      End
      Begin VB.Menu BtsLap2 
         Caption         =   "-"
      End
      Begin VB.Menu smPenjualan 
         Caption         =   "Transaksi Penjualan"
         Begin VB.Menu mnuLapJualSemua 
            Caption         =   "Penjualan Semuanya"
         End
         Begin VB.Menu mnuLapJualNota 
            Caption         =   "Penjualan Per Nota Transaksi"
         End
         Begin VB.Menu mnuLapJualPeriode 
            Caption         =   "Penjualan Per-Periode"
         End
      End
      Begin VB.Menu MnuLapReturJual 
         Caption         =   "Retur Penjualan"
         Begin VB.Menu MnuLapReturJualSemua 
            Caption         =   "Retur Semuanya"
         End
         Begin VB.Menu MnuLapReturJualNota 
            Caption         =   "Retur Per Nota"
         End
         Begin VB.Menu MnuLapReturPeriode 
            Caption         =   "Retur Per-Periode"
         End
      End
      Begin VB.Menu BtsLap3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuLapHutang 
         Caption         =   "Laporan Hutang"
      End
      Begin VB.Menu MnuLapPiutang 
         Caption         =   "Laporan Piutang"
      End
      Begin VB.Menu Mnupemisah8 
         Caption         =   "-"
      End
      Begin VB.Menu MnuLapLaba 
         Caption         =   "Laporan Laba"
      End
      Begin VB.Menu BtsLap4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLapPengguna 
         Caption         =   "Seluruh Pengguna"
      End
   End
   Begin VB.Menu MuSetting 
      Caption         =   "&Setting"
      Begin VB.Menu mnuPengguna 
         Caption         =   "Data Pengguna"
      End
      Begin VB.Menu MnuSetProfil 
         Caption         =   "Setting Profil"
      End
      Begin VB.Menu mnuSetServer 
         Caption         =   "Setting Server"
      End
      Begin VB.Menu pemisah 
         Caption         =   "-"
      End
      Begin VB.Menu MnuReset 
         Caption         =   "Reset Data"
      End
      Begin VB.Menu Mnubarcode 
         Caption         =   "Buat Barcode"
      End
   End
   Begin VB.Menu MnuRegister 
      Caption         =   "&Registrasi"
      Begin VB.Menu MnuRegisterProg 
         Caption         =   "Registrasi Program"
      End
      Begin VB.Menu MnuPeraturan 
         Caption         =   "Peraturan"
      End
   End
End
Attribute VB_Name = "frmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MnuRegisterProg_Click() '5A54D0
loc_005A54D0:   push ebp
loc_005A54D1:   mov ebp, esp
  loc_005A54D3: sub esp, 0000000Ch
  loc_005A54D6: push 00405696h ; __vbaExceptHandler
loc_005A54DB:   mov eax, fs: [00000000h]
loc_005A54E1:   push eax
loc_005A54E2:   mov fs: [00000000h] , esp
  loc_005A54E9: sub esp, 00000030h
loc_005A54EC:   push ebx
loc_005A54ED:   push esi
loc_005A54EE:   push edi
loc_005A54EF:   mov var_C, esp
  loc_005A54F2: mov var_8, 00402CC8h
loc_005A54F9:   mov eax, Me
loc_005A54FC:   mov ecx, eax
  loc_005A54FE: and ecx, 00000001h
loc_005A5501:   mov var_4, ecx
  loc_005A5504: and al, FEh
loc_005A5506:   push eax
loc_005A5507:   mov Me, eax
loc_005A550A:   mov edx, [eax]
loc_005A550C:   Call [edx+00000004h]
loc_005A550F:   mov eax, [006682A0h]
loc_005A5514:   test eax, eax
  loc_005A5516: jnz 005A5528h
  loc_005A5518: push 006682A0h
  loc_005A551D: push 00414A94h
  loc_005A5522: call [004011E8h] ; __vbaNew2
  loc_005A5528: sub esp, 00000010h
  loc_005A552B: mov ecx, 0000000Ah
loc_005A5530:   mov ebx, esp
  loc_005A5532: mov eax, 80020004h
  loc_005A5537: sub esp, 00000010h
loc_005A553A:   mov esi, [006682A0h]
loc_005A5540:   mov [ebx], ecx
loc_005A5542:   mov ecx, var_30
loc_005A5545:   mov edi, [esi]
  loc_005A5547: mov edx, 00000001h
loc_005A554C:   mov [ebx+00000004h], ecx
loc_005A554F:   mov ecx, esp
loc_005A5551:   push esi
loc_005A5552:   mov [ebx+00000008h], eax
loc_005A5555:   mov eax, var_28
loc_005A5558:   mov [ebx+0000000Ch], eax
  loc_005A555B: mov eax, 00000002h
loc_005A5560:   mov [ecx], eax
loc_005A5562:   mov eax, var_20
loc_005A5565:   mov [ecx+00000004h], eax
loc_005A5568:   mov [ecx+00000008h], edx
loc_005A556B:   mov edx, var_18
loc_005A556E:   mov [ecx+0000000Ch], edx
loc_005A5571:   Call [edi+000002B0h]
loc_005A5577:   test eax, eax
loc_005A5579:   fnclex
  loc_005A557B: jge 005A558Fh
  loc_005A557D: push 000002B0h
  loc_005A5582: push 00434D2Ch
loc_005A5587:   push esi
loc_005A5588:   push eax
  loc_005A5589: call [00401080h] ; __vbaHresultCheckObj
  loc_005A558F: mov var_4, 00000000h
loc_005A5596:   mov eax, Me
loc_005A5599:   push eax
loc_005A559A:   mov ecx, [eax]
loc_005A559C:   Call [ecx+00000008h]
loc_005A559F:   mov eax, var_4
loc_005A55A2:   mov ecx, var_14
loc_005A55A5:   pop edi
loc_005A55A6:   pop esi
loc_005A55A7:   mov fs: [00000000h] , ecx
loc_005A55AE:   pop ebx
loc_005A55AF:   mov esp, ebp
loc_005A55B1:   pop ebp
  loc_005A55B2: retn 0004h
End Sub

Private Sub MnuLapLaba_Click() '5A34D0
loc_005A34D0:   push ebp
loc_005A34D1:   mov ebp, esp
  loc_005A34D3: sub esp, 0000000Ch
  loc_005A34D6: push 00405696h ; __vbaExceptHandler
loc_005A34DB:   mov eax, fs: [00000000h]
loc_005A34E1:   push eax
loc_005A34E2:   mov fs: [00000000h] , esp
  loc_005A34E9: sub esp, 00000030h
loc_005A34EC:   push ebx
loc_005A34ED:   push esi
loc_005A34EE:   push edi
loc_005A34EF:   mov var_C, esp
  loc_005A34F2: mov var_8, 00402BA8h
loc_005A34F9:   mov eax, Me
loc_005A34FC:   mov ecx, eax
  loc_005A34FE: and ecx, 00000001h
loc_005A3501:   mov var_4, ecx
  loc_005A3504: and al, FEh
loc_005A3506:   push eax
loc_005A3507:   mov Me, eax
loc_005A350A:   mov edx, [eax]
loc_005A350C:   Call [edx+00000004h]
loc_005A350F:   mov eax, [00668300h]
loc_005A3514:   test eax, eax
  loc_005A3516: jnz 005A3528h
  loc_005A3518: push 00668300h
  loc_005A351D: push 004100CCh
  loc_005A3522: call [004011E8h] ; __vbaNew2
  loc_005A3528: sub esp, 00000010h
  loc_005A352B: mov ecx, 0000000Ah
loc_005A3530:   mov ebx, esp
  loc_005A3532: mov eax, 80020004h
  loc_005A3537: sub esp, 00000010h
loc_005A353A:   mov esi, [00668300h]
loc_005A3540:   mov [ebx], ecx
loc_005A3542:   mov ecx, var_30
loc_005A3545:   mov edi, [esi]
  loc_005A3547: mov edx, 00000001h
loc_005A354C:   mov [ebx+00000004h], ecx
loc_005A354F:   mov ecx, esp
loc_005A3551:   push esi
loc_005A3552:   mov [ebx+00000008h], eax
loc_005A3555:   mov eax, var_28
loc_005A3558:   mov [ebx+0000000Ch], eax
  loc_005A355B: mov eax, 00000002h
loc_005A3560:   mov [ecx], eax
loc_005A3562:   mov eax, var_20
loc_005A3565:   mov [ecx+00000004h], eax
loc_005A3568:   mov [ecx+00000008h], edx
loc_005A356B:   mov edx, var_18
loc_005A356E:   mov [ecx+0000000Ch], edx
loc_005A3571:   Call [edi+000002B0h]
loc_005A3577:   test eax, eax
loc_005A3579:   fnclex
  loc_005A357B: jge 005A358Fh
  loc_005A357D: push 000002B0h
  loc_005A3582: push 00435290h
loc_005A3587:   push esi
loc_005A3588:   push eax
  loc_005A3589: call [00401080h] ; __vbaHresultCheckObj
  loc_005A358F: mov var_4, 00000000h
loc_005A3596:   mov eax, Me
loc_005A3599:   push eax
loc_005A359A:   mov ecx, [eax]
loc_005A359C:   Call [ecx+00000008h]
loc_005A359F:   mov eax, var_4
loc_005A35A2:   mov ecx, var_14
loc_005A35A5:   pop edi
loc_005A35A6:   pop esi
loc_005A35A7:   mov fs: [00000000h] , ecx
loc_005A35AE:   pop ebx
loc_005A35AF:   mov esp, ebp
loc_005A35B1:   pop ebp
  loc_005A35B2: retn 0004h
End Sub

Private Sub mnuLapProduk_Click() '5A4E50
loc_005A4E50:   push ebp
loc_005A4E51:   mov ebp, esp
  loc_005A4E53: sub esp, 0000000Ch
  loc_005A4E56: push 00405696h ; __vbaExceptHandler
loc_005A4E5B:   mov eax, fs: [00000000h]
loc_005A4E61:   push eax
loc_005A4E62:   mov fs: [00000000h] , esp
  loc_005A4E69: sub esp, 00000030h
loc_005A4E6C:   push ebx
loc_005A4E6D:   push esi
loc_005A4E6E:   push edi
loc_005A4E6F:   mov var_C, esp
  loc_005A4E72: mov var_8, 00402C88h
loc_005A4E79:   mov eax, Me
loc_005A4E7C:   mov ecx, eax
  loc_005A4E7E: and ecx, 00000001h
loc_005A4E81:   mov var_4, ecx
  loc_005A4E84: and al, FEh
loc_005A4E86:   push eax
loc_005A4E87:   mov Me, eax
loc_005A4E8A:   mov edx, [eax]
loc_005A4E8C:   Call [edx+00000004h]
loc_005A4E8F:   mov eax, [006683A0h]
loc_005A4E94:   test eax, eax
  loc_005A4E96: jnz 005A4EA8h
  loc_005A4E98: push 006683A0h
  loc_005A4E9D: push 0040B2E0h
  loc_005A4EA2: call [004011E8h] ; __vbaNew2
  loc_005A4EA8: sub esp, 00000010h
  loc_005A4EAB: mov ecx, 0000000Ah
loc_005A4EB0:   mov ebx, esp
  loc_005A4EB2: mov eax, 80020004h
  loc_005A4EB7: sub esp, 00000010h
loc_005A4EBA:   mov esi, [006683A0h]
loc_005A4EC0:   mov [ebx], ecx
loc_005A4EC2:   mov ecx, var_30
loc_005A4EC5:   mov edi, [esi]
  loc_005A4EC7: mov edx, 00000001h
loc_005A4ECC:   mov [ebx+00000004h], ecx
loc_005A4ECF:   mov ecx, esp
loc_005A4ED1:   push esi
loc_005A4ED2:   mov [ebx+00000008h], eax
loc_005A4ED5:   mov eax, var_28
loc_005A4ED8:   mov [ebx+0000000Ch], eax
  loc_005A4EDB: mov eax, 00000002h
loc_005A4EE0:   mov [ecx], eax
loc_005A4EE2:   mov eax, var_20
loc_005A4EE5:   mov [ecx+00000004h], eax
loc_005A4EE8:   mov [ecx+00000008h], edx
loc_005A4EEB:   mov edx, var_18
loc_005A4EEE:   mov [ecx+0000000Ch], edx
loc_005A4EF1:   Call [edi+000002B0h]
loc_005A4EF7:   test eax, eax
loc_005A4EF9:   fnclex
  loc_005A4EFB: jge 005A4F0Fh
  loc_005A4EFD: push 000002B0h
  loc_005A4F02: push 004355E4h
loc_005A4F07:   push esi
loc_005A4F08:   push eax
  loc_005A4F09: call [00401080h] ; __vbaHresultCheckObj
  loc_005A4F0F: mov var_4, 00000000h
loc_005A4F16:   mov eax, Me
loc_005A4F19:   push eax
loc_005A4F1A:   mov ecx, [eax]
loc_005A4F1C:   Call [ecx+00000008h]
loc_005A4F1F:   mov eax, var_4
loc_005A4F22:   mov ecx, var_14
loc_005A4F25:   pop edi
loc_005A4F26:   pop esi
loc_005A4F27:   mov fs: [00000000h] , ecx
loc_005A4F2E:   pop ebx
loc_005A4F2F:   mov esp, ebp
loc_005A4F31:   pop ebp
  loc_005A4F32: retn 0004h
End Sub

Private Sub MnuLapReturJualNota_Click() '5A39E0
loc_005A39E0:   push ebp
loc_005A39E1:   mov ebp, esp
  loc_005A39E3: sub esp, 0000000Ch
  loc_005A39E6: push 00405696h ; __vbaExceptHandler
loc_005A39EB:   mov eax, fs: [00000000h]
loc_005A39F1:   push eax
loc_005A39F2:   mov fs: [00000000h] , esp
  loc_005A39F9: sub esp, 00000030h
loc_005A39FC:   push ebx
loc_005A39FD:   push esi
loc_005A39FE:   push edi
loc_005A39FF:   mov var_C, esp
  loc_005A3A02: mov var_8, 00402BE0h
loc_005A3A09:   mov eax, Me
loc_005A3A0C:   mov ecx, eax
  loc_005A3A0E: and ecx, 00000001h
loc_005A3A11:   mov var_4, ecx
  loc_005A3A14: and al, FEh
loc_005A3A16:   push eax
loc_005A3A17:   mov Me, eax
loc_005A3A1A:   mov edx, [eax]
loc_005A3A1C:   Call [edx+00000004h]
loc_005A3A1F:   mov eax, [00668350h]
loc_005A3A24:   test eax, eax
  loc_005A3A26: jnz 005A3A38h
  loc_005A3A28: push 00668350h
  loc_005A3A2D: push 0040C9A4h
  loc_005A3A32: call [004011E8h] ; __vbaNew2
  loc_005A3A38: sub esp, 00000010h
  loc_005A3A3B: mov ecx, 0000000Ah
loc_005A3A40:   mov ebx, esp
  loc_005A3A42: mov eax, 80020004h
  loc_005A3A47: sub esp, 00000010h
loc_005A3A4A:   mov esi, [00668350h]
loc_005A3A50:   mov [ebx], ecx
loc_005A3A52:   mov ecx, var_30
loc_005A3A55:   mov edi, [esi]
  loc_005A3A57: mov edx, 00000001h
loc_005A3A5C:   mov [ebx+00000004h], ecx
loc_005A3A5F:   mov ecx, esp
loc_005A3A61:   push esi
loc_005A3A62:   mov [ebx+00000008h], eax
loc_005A3A65:   mov eax, var_28
loc_005A3A68:   mov [ebx+0000000Ch], eax
  loc_005A3A6B: mov eax, 00000002h
loc_005A3A70:   mov [ecx], eax
loc_005A3A72:   mov eax, var_20
loc_005A3A75:   mov [ecx+00000004h], eax
loc_005A3A78:   mov [ecx+00000008h], edx
loc_005A3A7B:   mov edx, var_18
loc_005A3A7E:   mov [ecx+0000000Ch], edx
loc_005A3A81:   Call [edi+000002B0h]
loc_005A3A87:   test eax, eax
loc_005A3A89:   fnclex
  loc_005A3A8B: jge 005A3A9Fh
  loc_005A3A8D: push 000002B0h
  loc_005A3A92: push 004354BCh
loc_005A3A97:   push esi
loc_005A3A98:   push eax
  loc_005A3A99: call [00401080h] ; __vbaHresultCheckObj
  loc_005A3A9F: mov var_4, 00000000h
loc_005A3AA6:   mov eax, Me
loc_005A3AA9:   push eax
loc_005A3AAA:   mov ecx, [eax]
loc_005A3AAC:   Call [ecx+00000008h]
loc_005A3AAF:   mov eax, var_4
loc_005A3AB2:   mov ecx, var_14
loc_005A3AB5:   pop edi
loc_005A3AB6:   pop esi
loc_005A3AB7:   mov fs: [00000000h] , ecx
loc_005A3ABE:   pop ebx
loc_005A3ABF:   mov esp, ebp
loc_005A3AC1:   pop ebp
  loc_005A3AC2: retn 0004h
End Sub

Private Sub MnuLapPiutang_Click() '5A36A0
loc_005A36A0:   push ebp
loc_005A36A1:   mov ebp, esp
  loc_005A36A3: sub esp, 0000000Ch
  loc_005A36A6: push 00405696h ; __vbaExceptHandler
loc_005A36AB:   mov eax, fs: [00000000h]
loc_005A36B1:   push eax
loc_005A36B2:   mov fs: [00000000h] , esp
  loc_005A36B9: sub esp, 00000030h
loc_005A36BC:   push ebx
loc_005A36BD:   push esi
loc_005A36BE:   push edi
loc_005A36BF:   mov var_C, esp
  loc_005A36C2: mov var_8, 00402BC0h
loc_005A36C9:   mov eax, Me
loc_005A36CC:   mov ecx, eax
  loc_005A36CE: and ecx, 00000001h
loc_005A36D1:   mov var_4, ecx
  loc_005A36D4: and al, FEh
loc_005A36D6:   push eax
loc_005A36D7:   mov Me, eax
loc_005A36DA:   mov edx, [eax]
loc_005A36DC:   Call [edx+00000004h]
loc_005A36DF:   mov eax, [00668314h]
loc_005A36E4:   test eax, eax
  loc_005A36E6: jnz 005A36F8h
  loc_005A36E8: push 00668314h
  loc_005A36ED: push 0041A2C8h
  loc_005A36F2: call [004011E8h] ; __vbaNew2
  loc_005A36F8: sub esp, 00000010h
  loc_005A36FB: mov ecx, 0000000Ah
loc_005A3700:   mov ebx, esp
  loc_005A3702: mov eax, 80020004h
  loc_005A3707: sub esp, 00000010h
loc_005A370A:   mov esi, [00668314h]
loc_005A3710:   mov [ebx], ecx
loc_005A3712:   mov ecx, var_30
loc_005A3715:   mov edi, [esi]
  loc_005A3717: mov edx, 00000001h
loc_005A371C:   mov [ebx+00000004h], ecx
loc_005A371F:   mov ecx, esp
loc_005A3721:   push esi
loc_005A3722:   mov [ebx+00000008h], eax
loc_005A3725:   mov eax, var_28
loc_005A3728:   mov [ebx+0000000Ch], eax
  loc_005A372B: mov eax, 00000002h
loc_005A3730:   mov [ecx], eax
loc_005A3732:   mov eax, var_20
loc_005A3735:   mov [ecx+00000004h], eax
loc_005A3738:   mov [ecx+00000008h], edx
loc_005A373B:   mov edx, var_18
loc_005A373E:   mov [ecx+0000000Ch], edx
loc_005A3741:   Call [edi+000002B0h]
loc_005A3747:   test eax, eax
loc_005A3749:   fnclex
  loc_005A374B: jge 005A375Fh
  loc_005A374D: push 000002B0h
  loc_005A3752: push 00435340h
loc_005A3757:   push esi
loc_005A3758:   push eax
  loc_005A3759: call [00401080h] ; __vbaHresultCheckObj
  loc_005A375F: mov var_4, 00000000h
loc_005A3766:   mov eax, Me
loc_005A3769:   push eax
loc_005A376A:   mov ecx, [eax]
loc_005A376C:   Call [ecx+00000008h]
loc_005A376F:   mov eax, var_4
loc_005A3772:   mov ecx, var_14
loc_005A3775:   pop edi
loc_005A3776:   pop esi
loc_005A3777:   mov fs: [00000000h] , ecx
loc_005A377E:   pop ebx
loc_005A377F:   mov esp, ebp
loc_005A3781:   pop ebp
  loc_005A3782: retn 0004h
End Sub

Private Sub mnuGolongan_Click() '5A4580
loc_005A4580:   push ebp
loc_005A4581:   mov ebp, esp
  loc_005A4583: sub esp, 0000000Ch
  loc_005A4586: push 00405696h ; __vbaExceptHandler
loc_005A458B:   mov eax, fs: [00000000h]
loc_005A4591:   push eax
loc_005A4592:   mov fs: [00000000h] , esp
  loc_005A4599: sub esp, 00000030h
loc_005A459C:   push ebx
loc_005A459D:   push esi
loc_005A459E:   push edi
loc_005A459F:   mov var_C, esp
  loc_005A45A2: mov var_8, 00402C30h
loc_005A45A9:   mov eax, Me
loc_005A45AC:   mov ecx, eax
  loc_005A45AE: and ecx, 00000001h
loc_005A45B1:   mov var_4, ecx
  loc_005A45B4: and al, FEh
loc_005A45B6:   push eax
loc_005A45B7:   mov Me, eax
loc_005A45BA:   mov edx, [eax]
loc_005A45BC:   Call [edx+00000004h]
loc_005A45BF:   mov eax, [0066818Ch]
loc_005A45C4:   test eax, eax
  loc_005A45C6: jnz 005A45D8h
  loc_005A45C8: push 0066818Ch
  loc_005A45CD: push 00412CA4h
  loc_005A45D2: call [004011E8h] ; __vbaNew2
  loc_005A45D8: sub esp, 00000010h
  loc_005A45DB: mov ecx, 0000000Ah
loc_005A45E0:   mov ebx, esp
  loc_005A45E2: mov eax, 80020004h
  loc_005A45E7: sub esp, 00000010h
loc_005A45EA:   mov esi, [0066818Ch]
loc_005A45F0:   mov [ebx], ecx
loc_005A45F2:   mov ecx, var_30
loc_005A45F5:   mov edi, [esi]
  loc_005A45F7: mov edx, 00000001h
loc_005A45FC:   mov [ebx+00000004h], ecx
loc_005A45FF:   mov ecx, esp
loc_005A4601:   push esi
loc_005A4602:   mov [ebx+00000008h], eax
loc_005A4605:   mov eax, var_28
loc_005A4608:   mov [ebx+0000000Ch], eax
  loc_005A460B: mov eax, 00000002h
loc_005A4610:   mov [ecx], eax
loc_005A4612:   mov eax, var_20
loc_005A4615:   mov [ecx+00000004h], eax
loc_005A4618:   mov [ecx+00000008h], edx
loc_005A461B:   mov edx, var_18
loc_005A461E:   mov [ecx+0000000Ch], edx
loc_005A4621:   Call [edi+000002B0h]
loc_005A4627:   test eax, eax
loc_005A4629:   fnclex
  loc_005A462B: jge 005A463Fh
  loc_005A462D: push 000002B0h
  loc_005A4632: push 0042FB74h
loc_005A4637:   push esi
loc_005A4638:   push eax
  loc_005A4639: call [00401080h] ; __vbaHresultCheckObj
  loc_005A463F: mov var_4, 00000000h
loc_005A4646:   mov eax, Me
loc_005A4649:   push eax
loc_005A464A:   mov ecx, [eax]
loc_005A464C:   Call [ecx+00000008h]
loc_005A464F:   mov eax, var_4
loc_005A4652:   mov ecx, var_14
loc_005A4655:   pop edi
loc_005A4656:   pop esi
loc_005A4657:   mov fs: [00000000h] , ecx
loc_005A465E:   pop ebx
loc_005A465F:   mov esp, ebp
loc_005A4661:   pop ebp
  loc_005A4662: retn 0004h
End Sub

Private Sub MnuLapReturBeliNota_Click() '5A3790
loc_005A3790:   push ebp
loc_005A3791:   mov ebp, esp
  loc_005A3793: sub esp, 0000000Ch
  loc_005A3796: push 00405696h ; __vbaExceptHandler
loc_005A379B:   mov eax, fs: [00000000h]
loc_005A37A1:   push eax
loc_005A37A2:   mov fs: [00000000h] , esp
  loc_005A37A9: sub esp, 00000030h
loc_005A37AC:   push ebx
loc_005A37AD:   push esi
loc_005A37AE:   push edi
loc_005A37AF:   mov var_C, esp
  loc_005A37B2: mov var_8, 00402BC8h
loc_005A37B9:   mov eax, Me
loc_005A37BC:   mov ecx, eax
  loc_005A37BE: and ecx, 00000001h
loc_005A37C1:   mov var_4, ecx
  loc_005A37C4: and al, FEh
loc_005A37C6:   push eax
loc_005A37C7:   mov Me, eax
loc_005A37CA:   mov edx, [eax]
loc_005A37CC:   Call [edx+00000004h]
loc_005A37CF:   mov eax, [00668328h]
loc_005A37D4:   test eax, eax
  loc_005A37D6: jnz 005A37E8h
  loc_005A37D8: push 00668328h
  loc_005A37DD: push 0040CFB8h
  loc_005A37E2: call [004011E8h] ; __vbaNew2
  loc_005A37E8: sub esp, 00000010h
  loc_005A37EB: mov ecx, 0000000Ah
loc_005A37F0:   mov ebx, esp
  loc_005A37F2: mov eax, 80020004h
  loc_005A37F7: sub esp, 00000010h
loc_005A37FA:   mov esi, [00668328h]
loc_005A3800:   mov [ebx], ecx
loc_005A3802:   mov ecx, var_30
loc_005A3805:   mov edi, [esi]
  loc_005A3807: mov edx, 00000001h
loc_005A380C:   mov [ebx+00000004h], ecx
loc_005A380F:   mov ecx, esp
loc_005A3811:   push esi
loc_005A3812:   mov [ebx+00000008h], eax
loc_005A3815:   mov eax, var_28
loc_005A3818:   mov [ebx+0000000Ch], eax
  loc_005A381B: mov eax, 00000002h
loc_005A3820:   mov [ecx], eax
loc_005A3822:   mov eax, var_20
loc_005A3825:   mov [ecx+00000004h], eax
loc_005A3828:   mov [ecx+00000008h], edx
loc_005A382B:   mov edx, var_18
loc_005A382E:   mov [ecx+0000000Ch], edx
loc_005A3831:   Call [edi+000002B0h]
loc_005A3837:   test eax, eax
loc_005A3839:   fnclex
  loc_005A383B: jge 005A384Fh
  loc_005A383D: push 000002B0h
  loc_005A3842: push 004353C0h
loc_005A3847:   push esi
loc_005A3848:   push eax
  loc_005A3849: call [00401080h] ; __vbaHresultCheckObj
  loc_005A384F: mov var_4, 00000000h
loc_005A3856:   mov eax, Me
loc_005A3859:   push eax
loc_005A385A:   mov ecx, [eax]
loc_005A385C:   Call [ecx+00000008h]
loc_005A385F:   mov eax, var_4
loc_005A3862:   mov ecx, var_14
loc_005A3865:   pop edi
loc_005A3866:   pop esi
loc_005A3867:   mov fs: [00000000h] , ecx
loc_005A386E:   pop ebx
loc_005A386F:   mov esp, ebp
loc_005A3871:   pop ebp
  loc_005A3872: retn 0004h
End Sub

Private Sub mnuLapBeliPeriode_Click() '5A5110
loc_005A5110:   push ebp
loc_005A5111:   mov ebp, esp
  loc_005A5113: sub esp, 0000000Ch
  loc_005A5116: push 00405696h ; __vbaExceptHandler
loc_005A511B:   mov eax, fs: [00000000h]
loc_005A5121:   push eax
loc_005A5122:   mov fs: [00000000h] , esp
  loc_005A5129: sub esp, 00000030h
loc_005A512C:   push ebx
loc_005A512D:   push esi
loc_005A512E:   push edi
loc_005A512F:   mov var_C, esp
  loc_005A5132: mov var_8, 00402CA8h
loc_005A5139:   mov eax, Me
loc_005A513C:   mov ecx, eax
  loc_005A513E: and ecx, 00000001h
loc_005A5141:   mov var_4, ecx
  loc_005A5144: and al, FEh
loc_005A5146:   push eax
loc_005A5147:   mov Me, eax
loc_005A514A:   mov edx, [eax]
loc_005A514C:   Call [edx+00000004h]
loc_005A514F:   mov eax, [006683C8h]
loc_005A5154:   test eax, eax
  loc_005A5156: jnz 005A5168h
  loc_005A5158: push 006683C8h
  loc_005A515D: push 0040F09Ch
  loc_005A5162: call [004011E8h] ; __vbaNew2
  loc_005A5168: sub esp, 00000010h
  loc_005A516B: mov ecx, 0000000Ah
loc_005A5170:   mov ebx, esp
  loc_005A5172: mov eax, 80020004h
  loc_005A5177: sub esp, 00000010h
loc_005A517A:   mov esi, [006683C8h]
loc_005A5180:   mov [ebx], ecx
loc_005A5182:   mov ecx, var_30
loc_005A5185:   mov edi, [esi]
  loc_005A5187: mov edx, 00000001h
loc_005A518C:   mov [ebx+00000004h], ecx
loc_005A518F:   mov ecx, esp
loc_005A5191:   push esi
loc_005A5192:   mov [ebx+00000008h], eax
loc_005A5195:   mov eax, var_28
loc_005A5198:   mov [ebx+0000000Ch], eax
  loc_005A519B: mov eax, 00000002h
loc_005A51A0:   mov [ecx], eax
loc_005A51A2:   mov eax, var_20
loc_005A51A5:   mov [ecx+00000004h], eax
loc_005A51A8:   mov [ecx+00000008h], edx
loc_005A51AB:   mov edx, var_18
loc_005A51AE:   mov [ecx+0000000Ch], edx
loc_005A51B1:   Call [edi+000002B0h]
loc_005A51B7:   test eax, eax
loc_005A51B9:   fnclex
  loc_005A51BB: jge 005A51CFh
  loc_005A51BD: push 000002B0h
  loc_005A51C2: push 004356B8h
loc_005A51C7:   push esi
loc_005A51C8:   push eax
  loc_005A51C9: call [00401080h] ; __vbaHresultCheckObj
  loc_005A51CF: mov var_4, 00000000h
loc_005A51D6:   mov eax, Me
loc_005A51D9:   push eax
loc_005A51DA:   mov ecx, [eax]
loc_005A51DC:   Call [ecx+00000008h]
loc_005A51DF:   mov eax, var_4
loc_005A51E2:   mov ecx, var_14
loc_005A51E5:   pop edi
loc_005A51E6:   pop esi
loc_005A51E7:   mov fs: [00000000h] , ecx
loc_005A51EE:   pop ebx
loc_005A51EF:   mov esp, ebp
loc_005A51F1:   pop ebp
  loc_005A51F2: retn 0004h
End Sub

Private Sub MnuLapReturPeriode_Click() '5A3B40
loc_005A3B40:   push ebp
loc_005A3B41:   mov ebp, esp
  loc_005A3B43: sub esp, 0000000Ch
  loc_005A3B46: push 00405696h ; __vbaExceptHandler
loc_005A3B4B:   mov eax, fs: [00000000h]
loc_005A3B51:   push eax
loc_005A3B52:   mov fs: [00000000h] , esp
  loc_005A3B59: sub esp, 00000030h
loc_005A3B5C:   push ebx
loc_005A3B5D:   push esi
loc_005A3B5E:   push edi
loc_005A3B5F:   mov var_C, esp
  loc_005A3B62: mov var_8, 00402BF0h
loc_005A3B69:   mov eax, Me
loc_005A3B6C:   mov ecx, eax
  loc_005A3B6E: and ecx, 00000001h
loc_005A3B71:   mov var_4, ecx
  loc_005A3B74: and al, FEh
loc_005A3B76:   push eax
loc_005A3B77:   mov Me, eax
loc_005A3B7A:   mov edx, [eax]
loc_005A3B7C:   Call [edx+00000004h]
loc_005A3B7F:   mov eax, [00668364h]
loc_005A3B84:   test eax, eax
  loc_005A3B86: jnz 005A3B98h
  loc_005A3B88: push 00668364h
  loc_005A3B8D: push 0041124Ch
  loc_005A3B92: call [004011E8h] ; __vbaNew2
  loc_005A3B98: sub esp, 00000010h
  loc_005A3B9B: mov ecx, 0000000Ah
loc_005A3BA0:   mov ebx, esp
  loc_005A3BA2: mov eax, 80020004h
  loc_005A3BA7: sub esp, 00000010h
loc_005A3BAA:   mov esi, [00668364h]
loc_005A3BB0:   mov [ebx], ecx
loc_005A3BB2:   mov ecx, var_30
loc_005A3BB5:   mov edi, [esi]
  loc_005A3BB7: mov edx, 00000001h
loc_005A3BBC:   mov [ebx+00000004h], ecx
loc_005A3BBF:   mov ecx, esp
loc_005A3BC1:   push esi
loc_005A3BC2:   mov [ebx+00000008h], eax
loc_005A3BC5:   mov eax, var_28
loc_005A3BC8:   mov [ebx+0000000Ch], eax
  loc_005A3BCB: mov eax, 00000002h
loc_005A3BD0:   mov [ecx], eax
loc_005A3BD2:   mov eax, var_20
loc_005A3BD5:   mov [ecx+00000004h], eax
loc_005A3BD8:   mov [ecx+00000008h], edx
loc_005A3BDB:   mov edx, var_18
loc_005A3BDE:   mov [ecx+0000000Ch], edx
loc_005A3BE1:   Call [edi+000002B0h]
loc_005A3BE7:   test eax, eax
loc_005A3BE9:   fnclex
  loc_005A3BEB: jge 005A3BFFh
  loc_005A3BED: push 000002B0h
  loc_005A3BF2: push 004354FCh
loc_005A3BF7:   push esi
loc_005A3BF8:   push eax
  loc_005A3BF9: call [00401080h] ; __vbaHresultCheckObj
  loc_005A3BFF: mov var_4, 00000000h
loc_005A3C06:   mov eax, Me
loc_005A3C09:   push eax
loc_005A3C0A:   mov ecx, [eax]
loc_005A3C0C:   Call [ecx+00000008h]
loc_005A3C0F:   mov eax, var_4
loc_005A3C12:   mov ecx, var_14
loc_005A3C15:   pop edi
loc_005A3C16:   pop esi
loc_005A3C17:   mov fs: [00000000h] , ecx
loc_005A3C1E:   pop ebx
loc_005A3C1F:   mov esp, ebp
loc_005A3C21:   pop ebp
  loc_005A3C22: retn 0004h
End Sub

Private Sub mnuLapJualSemua_Click() '5A3460
loc_005A3460:   push ebp
loc_005A3461:   mov ebp, esp
  loc_005A3463: sub esp, 0000000Ch
  loc_005A3466: push 00405696h ; __vbaExceptHandler
loc_005A346B:   mov eax, fs: [00000000h]
loc_005A3471:   push eax
loc_005A3472:   mov fs: [00000000h] , esp
  loc_005A3479: sub esp, 00000008h
loc_005A347C:   push ebx
loc_005A347D:   push esi
loc_005A347E:   push edi
loc_005A347F:   mov var_C, esp
  loc_005A3482: mov var_8, 00402BA0h
loc_005A3489:   mov eax, Me
loc_005A348C:   mov ecx, eax
  loc_005A348E: and ecx, 00000001h
loc_005A3491:   mov var_4, ecx
  loc_005A3494: and al, FEh
loc_005A3496:   push eax
loc_005A3497:   mov Me, eax
loc_005A349A:   mov edx, [eax]
loc_005A349C:   Call [edx+00000004h]
  loc_005A349F: call 0060FD20h
  loc_005A34A4: mov var_4, 00000000h
loc_005A34AB:   mov eax, Me
loc_005A34AE:   push eax
loc_005A34AF:   mov ecx, [eax]
loc_005A34B1:   Call [ecx+00000008h]
loc_005A34B4:   mov eax, var_4
loc_005A34B7:   mov ecx, var_14
loc_005A34BA:   pop edi
loc_005A34BB:   pop esi
loc_005A34BC:   mov fs: [00000000h] , ecx
loc_005A34C3:   pop ebx
loc_005A34C4:   mov esp, ebp
loc_005A34C6:   pop ebp
  loc_005A34C7: retn 0004h
End Sub

Private Sub MnuSetProfil_Click() '5A5980
loc_005A5980:   push ebp
loc_005A5981:   mov ebp, esp
  loc_005A5983: sub esp, 0000000Ch
  loc_005A5986: push 00405696h ; __vbaExceptHandler
loc_005A598B:   mov eax, fs: [00000000h]
loc_005A5991:   push eax
loc_005A5992:   mov fs: [00000000h] , esp
  loc_005A5999: sub esp, 00000030h
loc_005A599C:   push ebx
loc_005A599D:   push esi
loc_005A599E:   push edi
loc_005A599F:   mov var_C, esp
  loc_005A59A2: mov var_8, 00402CF0h
loc_005A59A9:   mov eax, Me
loc_005A59AC:   mov ecx, eax
  loc_005A59AE: and ecx, 00000001h
loc_005A59B1:   mov var_4, ecx
  loc_005A59B4: and al, FEh
loc_005A59B6:   push eax
loc_005A59B7:   mov Me, eax
loc_005A59BA:   mov edx, [eax]
loc_005A59BC:   Call [edx+00000004h]
loc_005A59BF:   mov eax, [00668440h]
loc_005A59C4:   test eax, eax
  loc_005A59C6: jnz 005A59D8h
  loc_005A59C8: push 00668440h
  loc_005A59CD: push 0041946Ch
  loc_005A59D2: call [004011E8h] ; __vbaNew2
  loc_005A59D8: sub esp, 00000010h
  loc_005A59DB: mov ecx, 0000000Ah
loc_005A59E0:   mov ebx, esp
  loc_005A59E2: mov eax, 80020004h
  loc_005A59E7: sub esp, 00000010h
loc_005A59EA:   mov esi, [00668440h]
loc_005A59F0:   mov [ebx], ecx
loc_005A59F2:   mov ecx, var_30
loc_005A59F5:   mov edi, [esi]
  loc_005A59F7: mov edx, 00000001h
loc_005A59FC:   mov [ebx+00000004h], ecx
loc_005A59FF:   mov ecx, esp
loc_005A5A01:   push esi
loc_005A5A02:   mov [ebx+00000008h], eax
loc_005A5A05:   mov eax, var_28
loc_005A5A08:   mov [ebx+0000000Ch], eax
  loc_005A5A0B: mov eax, 00000002h
loc_005A5A10:   mov [ecx], eax
loc_005A5A12:   mov eax, var_20
loc_005A5A15:   mov [ecx+00000004h], eax
loc_005A5A18:   mov [ecx+00000008h], edx
loc_005A5A1B:   mov edx, var_18
loc_005A5A1E:   mov [ecx+0000000Ch], edx
loc_005A5A21:   Call [edi+000002B0h]
loc_005A5A27:   test eax, eax
loc_005A5A29:   fnclex
  loc_005A5A2B: jge 005A5A3Fh
  loc_005A5A2D: push 000002B0h
  loc_005A5A32: push 0043598Ch
loc_005A5A37:   push esi
loc_005A5A38:   push eax
  loc_005A5A39: call [00401080h] ; __vbaHresultCheckObj
  loc_005A5A3F: mov var_4, 00000000h
loc_005A5A46:   mov eax, Me
loc_005A5A49:   push eax
loc_005A5A4A:   mov ecx, [eax]
loc_005A5A4C:   Call [ecx+00000008h]
loc_005A5A4F:   mov eax, var_4
loc_005A5A52:   mov ecx, var_14
loc_005A5A55:   pop edi
loc_005A5A56:   pop esi
loc_005A5A57:   mov fs: [00000000h] , ecx
loc_005A5A5E:   pop ebx
loc_005A5A5F:   mov esp, ebp
loc_005A5A61:   pop ebp
  loc_005A5A62: retn 0004h
End Sub

Private Sub mnuSetServer_Click() '5A5A70
loc_005A5A70:   push ebp
loc_005A5A71:   mov ebp, esp
  loc_005A5A73: sub esp, 0000000Ch
  loc_005A5A76: push 00405696h ; __vbaExceptHandler
loc_005A5A7B:   mov eax, fs: [00000000h]
loc_005A5A81:   push eax
loc_005A5A82:   mov fs: [00000000h] , esp
  loc_005A5A89: sub esp, 00000030h
loc_005A5A8C:   push ebx
loc_005A5A8D:   push esi
loc_005A5A8E:   push edi
loc_005A5A8F:   mov var_C, esp
  loc_005A5A92: mov var_8, 00402CF8h
loc_005A5A99:   mov eax, Me
loc_005A5A9C:   mov ecx, eax
  loc_005A5A9E: and ecx, 00000001h
loc_005A5AA1:   mov var_4, ecx
  loc_005A5AA4: and al, FEh
loc_005A5AA6:   push eax
loc_005A5AA7:   mov Me, eax
loc_005A5AAA:   mov edx, [eax]
loc_005A5AAC:   Call [edx+00000004h]
loc_005A5AAF:   mov eax, [00668010h]
loc_005A5AB4:   test eax, eax
  loc_005A5AB6: jnz 005A5AC8h
  loc_005A5AB8: push 00668010h
  loc_005A5ABD: push 0040E2ECh
  loc_005A5AC2: call [004011E8h] ; __vbaNew2
  loc_005A5AC8: sub esp, 00000010h
  loc_005A5ACB: mov ecx, 0000000Ah
loc_005A5AD0:   mov ebx, esp
  loc_005A5AD2: mov eax, 80020004h
  loc_005A5AD7: sub esp, 00000010h
loc_005A5ADA:   mov esi, [00668010h]
loc_005A5AE0:   mov [ebx], ecx
loc_005A5AE2:   mov ecx, var_30
loc_005A5AE5:   mov edi, [esi]
  loc_005A5AE7: mov edx, 00000001h
loc_005A5AEC:   mov [ebx+00000004h], ecx
loc_005A5AEF:   mov ecx, esp
loc_005A5AF1:   push esi
loc_005A5AF2:   mov [ebx+00000008h], eax
loc_005A5AF5:   mov eax, var_28
loc_005A5AF8:   mov [ebx+0000000Ch], eax
  loc_005A5AFB: mov eax, 00000002h
loc_005A5B00:   mov [ecx], eax
loc_005A5B02:   mov eax, var_20
loc_005A5B05:   mov [ecx+00000004h], eax
loc_005A5B08:   mov [ecx+00000008h], edx
loc_005A5B0B:   mov edx, var_18
loc_005A5B0E:   mov [ecx+0000000Ch], edx
loc_005A5B11:   Call [edi+000002B0h]
loc_005A5B17:   test eax, eax
loc_005A5B19:   fnclex
  loc_005A5B1B: jge 005A5B2Fh
  loc_005A5B1D: push 000002B0h
  loc_005A5B22: push 0042DD20h
loc_005A5B27:   push esi
loc_005A5B28:   push eax
  loc_005A5B29: call [00401080h] ; __vbaHresultCheckObj
  loc_005A5B2F: mov var_4, 00000000h
loc_005A5B36:   mov eax, Me
loc_005A5B39:   push eax
loc_005A5B3A:   mov ecx, [eax]
loc_005A5B3C:   Call [ecx+00000008h]
loc_005A5B3F:   mov eax, var_4
loc_005A5B42:   mov ecx, var_14
loc_005A5B45:   pop edi
loc_005A5B46:   pop esi
loc_005A5B47:   mov fs: [00000000h] , ecx
loc_005A5B4E:   pop ebx
loc_005A5B4F:   mov esp, ebp
loc_005A5B51:   pop ebp
  loc_005A5B52: retn 0004h
End Sub

Private Sub mnuPengguna_Click() '5A53E0
loc_005A53E0:   push ebp
loc_005A53E1:   mov ebp, esp
  loc_005A53E3: sub esp, 0000000Ch
  loc_005A53E6: push 00405696h ; __vbaExceptHandler
loc_005A53EB:   mov eax, fs: [00000000h]
loc_005A53F1:   push eax
loc_005A53F2:   mov fs: [00000000h] , esp
  loc_005A53F9: sub esp, 00000030h
loc_005A53FC:   push ebx
loc_005A53FD:   push esi
loc_005A53FE:   push edi
loc_005A53FF:   mov var_C, esp
  loc_005A5402: mov var_8, 00402CC0h
loc_005A5409:   mov eax, Me
loc_005A540C:   mov ecx, eax
  loc_005A540E: and ecx, 00000001h
loc_005A5411:   mov var_4, ecx
  loc_005A5414: and al, FEh
loc_005A5416:   push eax
loc_005A5417:   mov Me, eax
loc_005A541A:   mov edx, [eax]
loc_005A541C:   Call [edx+00000004h]
loc_005A541F:   mov eax, [00668150h]
loc_005A5424:   test eax, eax
  loc_005A5426: jnz 005A5438h
  loc_005A5428: push 00668150h
  loc_005A542D: push 00415594h
  loc_005A5432: call [004011E8h] ; __vbaNew2
  loc_005A5438: sub esp, 00000010h
  loc_005A543B: mov ecx, 0000000Ah
loc_005A5440:   mov ebx, esp
  loc_005A5442: mov eax, 80020004h
  loc_005A5447: sub esp, 00000010h
loc_005A544A:   mov esi, [00668150h]
loc_005A5450:   mov [ebx], ecx
loc_005A5452:   mov ecx, var_30
loc_005A5455:   mov edi, [esi]
  loc_005A5457: mov edx, 00000001h
loc_005A545C:   mov [ebx+00000004h], ecx
loc_005A545F:   mov ecx, esp
loc_005A5461:   push esi
loc_005A5462:   mov [ebx+00000008h], eax
loc_005A5465:   mov eax, var_28
loc_005A5468:   mov [ebx+0000000Ch], eax
  loc_005A546B: mov eax, 00000002h
loc_005A5470:   mov [ecx], eax
loc_005A5472:   mov eax, var_20
loc_005A5475:   mov [ecx+00000004h], eax
loc_005A5478:   mov [ecx+00000008h], edx
loc_005A547B:   mov edx, var_18
loc_005A547E:   mov [ecx+0000000Ch], edx
loc_005A5481:   Call [edi+000002B0h]
loc_005A5487:   test eax, eax
loc_005A5489:   fnclex
  loc_005A548B: jge 005A549Fh
  loc_005A548D: push 000002B0h
  loc_005A5492: push 0042E6ECh
loc_005A5497:   push esi
loc_005A5498:   push eax
  loc_005A5499: call [00401080h] ; __vbaHresultCheckObj
  loc_005A549F: mov var_4, 00000000h
loc_005A54A6:   mov eax, Me
loc_005A54A9:   push eax
loc_005A54AA:   mov ecx, [eax]
loc_005A54AC:   Call [ecx+00000008h]
loc_005A54AF:   mov eax, var_4
loc_005A54B2:   mov ecx, var_14
loc_005A54B5:   pop edi
loc_005A54B6:   pop esi
loc_005A54B7:   mov fs: [00000000h] , ecx
loc_005A54BE:   pop ebx
loc_005A54BF:   mov esp, ebp
loc_005A54C1:   pop ebp
  loc_005A54C2: retn 0004h
End Sub

Private Sub MnuReset_Click() '5A55C0
loc_005A55C0:   push ebp
loc_005A55C1:   mov ebp, esp
  loc_005A55C3: sub esp, 0000000Ch
  loc_005A55C6: push 00405696h ; __vbaExceptHandler
loc_005A55CB:   mov eax, fs: [00000000h]
loc_005A55D1:   push eax
loc_005A55D2:   mov fs: [00000000h] , esp
  loc_005A55D9: sub esp, 00000030h
loc_005A55DC:   push ebx
loc_005A55DD:   push esi
loc_005A55DE:   push edi
loc_005A55DF:   mov var_C, esp
  loc_005A55E2: mov var_8, 00402CD0h
loc_005A55E9:   mov eax, Me
loc_005A55EC:   mov ecx, eax
  loc_005A55EE: and ecx, 00000001h
loc_005A55F1:   mov var_4, ecx
  loc_005A55F4: and al, FEh
loc_005A55F6:   push eax
loc_005A55F7:   mov Me, eax
loc_005A55FA:   mov edx, [eax]
loc_005A55FC:   Call [edx+00000004h]
loc_005A55FF:   mov eax, [00668404h]
loc_005A5604:   test eax, eax
  loc_005A5606: jnz 005A5618h
  loc_005A5608: push 00668404h
  loc_005A560D: push 0040DC1Ch
  loc_005A5612: call [004011E8h] ; __vbaNew2
  loc_005A5618: sub esp, 00000010h
  loc_005A561B: mov ecx, 0000000Ah
loc_005A5620:   mov ebx, esp
  loc_005A5622: mov eax, 80020004h
  loc_005A5627: sub esp, 00000010h
loc_005A562A:   mov esi, [00668404h]
loc_005A5630:   mov [ebx], ecx
loc_005A5632:   mov ecx, var_30
loc_005A5635:   mov edi, [esi]
  loc_005A5637: mov edx, 00000001h
loc_005A563C:   mov [ebx+00000004h], ecx
loc_005A563F:   mov ecx, esp
loc_005A5641:   push esi
loc_005A5642:   mov [ebx+00000008h], eax
loc_005A5645:   mov eax, var_28
loc_005A5648:   mov [ebx+0000000Ch], eax
  loc_005A564B: mov eax, 00000002h
loc_005A5650:   mov [ecx], eax
loc_005A5652:   mov eax, var_20
loc_005A5655:   mov [ecx+00000004h], eax
loc_005A5658:   mov [ecx+00000008h], edx
loc_005A565B:   mov edx, var_18
loc_005A565E:   mov [ecx+0000000Ch], edx
loc_005A5661:   Call [edi+000002B0h]
loc_005A5667:   test eax, eax
loc_005A5669:   fnclex
  loc_005A566B: jge 005A567Fh
  loc_005A566D: push 000002B0h
  loc_005A5672: push 0043578Ch
loc_005A5677:   push esi
loc_005A5678:   push eax
  loc_005A5679: call [00401080h] ; __vbaHresultCheckObj
  loc_005A567F: mov var_4, 00000000h
loc_005A5686:   mov eax, Me
loc_005A5689:   push eax
loc_005A568A:   mov ecx, [eax]
loc_005A568C:   Call [ecx+00000008h]
loc_005A568F:   mov eax, var_4
loc_005A5692:   mov ecx, var_14
loc_005A5695:   pop edi
loc_005A5696:   pop esi
loc_005A5697:   mov fs: [00000000h] , ecx
loc_005A569E:   pop ebx
loc_005A569F:   mov esp, ebp
loc_005A56A1:   pop ebp
  loc_005A56A2: retn 0004h
End Sub

Private Sub Mnubarcode_Click() '5A2F60
loc_005A2F60:   push ebp
loc_005A2F61:   mov ebp, esp
  loc_005A2F63: sub esp, 0000000Ch
  loc_005A2F66: push 00405696h ; __vbaExceptHandler
loc_005A2F6B:   mov eax, fs: [00000000h]
loc_005A2F71:   push eax
loc_005A2F72:   mov fs: [00000000h] , esp
  loc_005A2F79: sub esp, 00000034h
loc_005A2F7C:   push ebx
loc_005A2F7D:   push esi
loc_005A2F7E:   push edi
loc_005A2F7F:   mov var_C, esp
  loc_005A2F82: mov var_8, 00402B70h
loc_005A2F89:   mov eax, Me
loc_005A2F8C:   mov ecx, eax
  loc_005A2F8E: and ecx, 00000001h
loc_005A2F91:   mov var_4, ecx
  loc_005A2F94: and al, FEh
loc_005A2F96:   push eax
loc_005A2F97:   mov Me, eax
loc_005A2F9A:   mov edx, [eax]
loc_005A2F9C:   Call [edx+00000004h]
loc_005A2F9F:   mov eax, [0066A318h]
  loc_005A2FA4: xor edi, edi
loc_005A2FA6:   cmp eax, edi
loc_005A2FA8:   mov var_1C, edi
loc_005A2FAB:   mov var_20, edi
loc_005A2FAE:   mov var_30, edi
  loc_005A2FB1: jnz 005A2FC3h
  loc_005A2FB3: push 0066A318h
  loc_005A2FB8: push 0042E390h
  loc_005A2FBD: call [004011E8h] ; __vbaNew2
loc_005A2FC3:   mov esi, [0066A318h]
loc_005A2FC9:   lea ecx, var_20
loc_005A2FCC:   push ecx
loc_005A2FCD:   push esi
loc_005A2FCE:   mov eax, [esi]
loc_005A2FD0:   Call [eax+00000014h]
loc_005A2FD3:   cmp eax, edi
loc_005A2FD5:   fnclex
  loc_005A2FD7: jge 005A2FE8h
  loc_005A2FD9: push 00000014h
  loc_005A2FDB: push 0042E380h
loc_005A2FE0:   push esi
loc_005A2FE1:   push eax
  loc_005A2FE2: call [00401080h] ; __vbaHresultCheckObj
loc_005A2FE8:   mov eax, var_20
loc_005A2FEB:   lea ecx, var_1C
loc_005A2FEE:   push ecx
loc_005A2FEF:   push eax
loc_005A2FF0:   mov edx, [eax]
loc_005A2FF2:   mov esi, eax
loc_005A2FF4:   Call [edx+00000050h]
loc_005A2FF7:   cmp eax, edi
loc_005A2FF9:   fnclex
  loc_005A2FFB: jge 005A300Ch
  loc_005A2FFD: push 00000050h
  loc_005A2FFF: push 0042E528h
loc_005A3004:   push esi
loc_005A3005:   push eax
  loc_005A3006: call [00401080h] ; __vbaHresultCheckObj
loc_005A300C:   mov edx, var_1C
loc_005A300F:   push edx
  loc_005A3010: push 00435018h ; "\barcode.exe"
  loc_005A3015: call [00401070h] ; __vbaStrCat
loc_005A301B:   mov var_28, eax
loc_005A301E:   lea eax, var_30
  loc_005A3021: push 00000001h
loc_005A3023:   push eax
  loc_005A3024: mov var_30, 00000008h
  loc_005A302B: call [0040116Ch] ; rtcShell
  loc_005A3031: call [00401244h] ; __vbaFpI2
loc_005A3037:   lea ecx, var_1C
  loc_005A303A: call [0040129Ch] ; __vbaFreeStr
loc_005A3040:   lea ecx, var_20
  loc_005A3043: call [004012A0h] ; __vbaFreeObj
loc_005A3049:   lea ecx, var_30
  loc_005A304C: call [00401020h] ; __vbaFreeVar
loc_005A3052:   mov var_4, edi
loc_005A3055:   fwait
  loc_005A3056: push 005A307Ah
  loc_005A305B: jmp 005A3079h
loc_005A305D:   lea ecx, var_1C
  loc_005A3060: call [0040129Ch] ; __vbaFreeStr
loc_005A3066:   lea ecx, var_20
  loc_005A3069: call [004012A0h] ; __vbaFreeObj
loc_005A306F:   lea ecx, var_30
  loc_005A3072: call [00401020h] ; __vbaFreeVar
loc_005A3078:   ret
loc_005A3079:   ret
loc_005A307A:   mov eax, Me
loc_005A307D:   push eax
loc_005A307E:   mov ecx, [eax]
loc_005A3080:   Call [ecx+00000008h]
loc_005A3083:   mov eax, var_4
loc_005A3086:   mov ecx, var_14
loc_005A3089:   pop edi
loc_005A308A:   pop esi
loc_005A308B:   mov fs: [00000000h] , ecx
loc_005A3092:   pop ebx
loc_005A3093:   mov esp, ebp
loc_005A3095:   pop ebp
  loc_005A3096: retn 0004h
End Sub

Private Sub MnuReturPembelian_Click() '5A57A0
loc_005A57A0:   push ebp
loc_005A57A1:   mov ebp, esp
  loc_005A57A3: sub esp, 0000000Ch
  loc_005A57A6: push 00405696h ; __vbaExceptHandler
loc_005A57AB:   mov eax, fs: [00000000h]
loc_005A57B1:   push eax
loc_005A57B2:   mov fs: [00000000h] , esp
  loc_005A57B9: sub esp, 00000030h
loc_005A57BC:   push ebx
loc_005A57BD:   push esi
loc_005A57BE:   push edi
loc_005A57BF:   mov var_C, esp
  loc_005A57C2: mov var_8, 00402CE0h
loc_005A57C9:   mov eax, Me
loc_005A57CC:   mov ecx, eax
  loc_005A57CE: and ecx, 00000001h
loc_005A57D1:   mov var_4, ecx
  loc_005A57D4: and al, FEh
loc_005A57D6:   push eax
loc_005A57D7:   mov Me, eax
loc_005A57DA:   mov edx, [eax]
loc_005A57DC:   Call [edx+00000004h]
loc_005A57DF:   mov eax, [0066842Ch]
loc_005A57E4:   test eax, eax
  loc_005A57E6: jnz 005A57F8h
  loc_005A57E8: push 0066842Ch
  loc_005A57ED: push 00422F54h
  loc_005A57F2: call [004011E8h] ; __vbaNew2
  loc_005A57F8: sub esp, 00000010h
  loc_005A57FB: mov ecx, 0000000Ah
loc_005A5800:   mov ebx, esp
  loc_005A5802: mov eax, 80020004h
  loc_005A5807: sub esp, 00000010h
loc_005A580A:   mov esi, [0066842Ch]
loc_005A5810:   mov [ebx], ecx
loc_005A5812:   mov ecx, var_30
loc_005A5815:   mov edi, [esi]
  loc_005A5817: mov edx, 00000001h
loc_005A581C:   mov [ebx+00000004h], ecx
loc_005A581F:   mov ecx, esp
loc_005A5821:   push esi
loc_005A5822:   mov [ebx+00000008h], eax
loc_005A5825:   mov eax, var_28
loc_005A5828:   mov [ebx+0000000Ch], eax
  loc_005A582B: mov eax, 00000002h
loc_005A5830:   mov [ecx], eax
loc_005A5832:   mov eax, var_20
loc_005A5835:   mov [ecx+00000004h], eax
loc_005A5838:   mov [ecx+00000008h], edx
loc_005A583B:   mov edx, var_18
loc_005A583E:   mov [ecx+0000000Ch], edx
loc_005A5841:   Call [edi+000002B0h]
loc_005A5847:   test eax, eax
loc_005A5849:   fnclex
  loc_005A584B: jge 005A585Fh
  loc_005A584D: push 000002B0h
  loc_005A5852: push 004358F4h
loc_005A5857:   push esi
loc_005A5858:   push eax
  loc_005A5859: call [00401080h] ; __vbaHresultCheckObj
  loc_005A585F: mov var_4, 00000000h
loc_005A5866:   mov eax, Me
loc_005A5869:   push eax
loc_005A586A:   mov ecx, [eax]
loc_005A586C:   Call [ecx+00000008h]
loc_005A586F:   mov eax, var_4
loc_005A5872:   mov ecx, var_14
loc_005A5875:   pop edi
loc_005A5876:   pop esi
loc_005A5877:   mov fs: [00000000h] , ecx
loc_005A587E:   pop ebx
loc_005A587F:   mov esp, ebp
loc_005A5881:   pop ebp
  loc_005A5882: retn 0004h
End Sub

Private Sub MnuHutang_Click() '5A3190
loc_005A3190:   push ebp
loc_005A3191:   mov ebp, esp
  loc_005A3193: sub esp, 0000000Ch
  loc_005A3196: push 00405696h ; __vbaExceptHandler
loc_005A319B:   mov eax, fs: [00000000h]
loc_005A31A1:   push eax
loc_005A31A2:   mov fs: [00000000h] , esp
  loc_005A31A9: sub esp, 00000030h
loc_005A31AC:   push ebx
loc_005A31AD:   push esi
loc_005A31AE:   push edi
loc_005A31AF:   mov var_C, esp
  loc_005A31B2: mov var_8, 00402B88h
loc_005A31B9:   mov eax, Me
loc_005A31BC:   mov ecx, eax
  loc_005A31BE: and ecx, 00000001h
loc_005A31C1:   mov var_4, ecx
  loc_005A31C4: and al, FEh
loc_005A31C6:   push eax
loc_005A31C7:   mov Me, eax
loc_005A31CA:   mov edx, [eax]
loc_005A31CC:   Call [edx+00000004h]
loc_005A31CF:   mov eax, [006682C8h]
loc_005A31D4:   test eax, eax
  loc_005A31D6: jnz 005A31E8h
  loc_005A31D8: push 006682C8h
  loc_005A31DD: push 0041D790h
  loc_005A31E2: call [004011E8h] ; __vbaNew2
  loc_005A31E8: sub esp, 00000010h
  loc_005A31EB: mov ecx, 0000000Ah
loc_005A31F0:   mov ebx, esp
  loc_005A31F2: mov eax, 80020004h
  loc_005A31F7: sub esp, 00000010h
loc_005A31FA:   mov esi, [006682C8h]
loc_005A3200:   mov [ebx], ecx
loc_005A3202:   mov ecx, var_30
loc_005A3205:   mov edi, [esi]
  loc_005A3207: mov edx, 00000001h
loc_005A320C:   mov [ebx+00000004h], ecx
loc_005A320F:   mov ecx, esp
loc_005A3211:   push esi
loc_005A3212:   mov [ebx+00000008h], eax
loc_005A3215:   mov eax, var_28
loc_005A3218:   mov [ebx+0000000Ch], eax
  loc_005A321B: mov eax, 00000002h
loc_005A3220:   mov [ecx], eax
loc_005A3222:   mov eax, var_20
loc_005A3225:   mov [ecx+00000004h], eax
loc_005A3228:   mov [ecx+00000008h], edx
loc_005A322B:   mov edx, var_18
loc_005A322E:   mov [ecx+0000000Ch], edx
loc_005A3231:   Call [edi+000002B0h]
loc_005A3237:   test eax, eax
loc_005A3239:   fnclex
  loc_005A323B: jge 005A324Fh
  loc_005A323D: push 000002B0h
  loc_005A3242: push 004350D0h
loc_005A3247:   push esi
loc_005A3248:   push eax
  loc_005A3249: call [00401080h] ; __vbaHresultCheckObj
  loc_005A324F: mov var_4, 00000000h
loc_005A3256:   mov eax, Me
loc_005A3259:   push eax
loc_005A325A:   mov ecx, [eax]
loc_005A325C:   Call [ecx+00000008h]
loc_005A325F:   mov eax, var_4
loc_005A3262:   mov ecx, var_14
loc_005A3265:   pop edi
loc_005A3266:   pop esi
loc_005A3267:   mov fs: [00000000h] , ecx
loc_005A326E:   pop ebx
loc_005A326F:   mov esp, ebp
loc_005A3271:   pop ebp
  loc_005A3272: retn 0004h
End Sub

Private Sub MnuGantiPass_Click() '5A30A0
loc_005A30A0:   push ebp
loc_005A30A1:   mov ebp, esp
  loc_005A30A3: sub esp, 0000000Ch
  loc_005A30A6: push 00405696h ; __vbaExceptHandler
loc_005A30AB:   mov eax, fs: [00000000h]
loc_005A30B1:   push eax
loc_005A30B2:   mov fs: [00000000h] , esp
  loc_005A30B9: sub esp, 00000030h
loc_005A30BC:   push ebx
loc_005A30BD:   push esi
loc_005A30BE:   push edi
loc_005A30BF:   mov var_C, esp
  loc_005A30C2: mov var_8, 00402B80h
loc_005A30C9:   mov eax, Me
loc_005A30CC:   mov ecx, eax
  loc_005A30CE: and ecx, 00000001h
loc_005A30D1:   mov var_4, ecx
  loc_005A30D4: and al, FEh
loc_005A30D6:   push eax
loc_005A30D7:   mov Me, eax
loc_005A30DA:   mov edx, [eax]
loc_005A30DC:   Call [edx+00000004h]
loc_005A30DF:   mov eax, [006682B4h]
loc_005A30E4:   test eax, eax
  loc_005A30E6: jnz 005A30F8h
  loc_005A30E8: push 006682B4h
  loc_005A30ED: push 0040D5CCh
  loc_005A30F2: call [004011E8h] ; __vbaNew2
  loc_005A30F8: sub esp, 00000010h
  loc_005A30FB: mov ecx, 0000000Ah
loc_005A3100:   mov ebx, esp
  loc_005A3102: mov eax, 80020004h
  loc_005A3107: sub esp, 00000010h
loc_005A310A:   mov esi, [006682B4h]
loc_005A3110:   mov [ebx], ecx
loc_005A3112:   mov ecx, var_30
loc_005A3115:   mov edi, [esi]
  loc_005A3117: mov edx, 00000001h
loc_005A311C:   mov [ebx+00000004h], ecx
loc_005A311F:   mov ecx, esp
loc_005A3121:   push esi
loc_005A3122:   mov [ebx+00000008h], eax
loc_005A3125:   mov eax, var_28
loc_005A3128:   mov [ebx+0000000Ch], eax
  loc_005A312B: mov eax, 00000002h
loc_005A3130:   mov [ecx], eax
loc_005A3132:   mov eax, var_20
loc_005A3135:   mov [ecx+00000004h], eax
loc_005A3138:   mov [ecx+00000008h], edx
loc_005A313B:   mov edx, var_18
loc_005A313E:   mov [ecx+0000000Ch], edx
loc_005A3141:   Call [edi+000002B0h]
loc_005A3147:   test eax, eax
loc_005A3149:   fnclex
  loc_005A314B: jge 005A315Fh
  loc_005A314D: push 000002B0h
  loc_005A3152: push 00435034h
loc_005A3157:   push esi
loc_005A3158:   push eax
  loc_005A3159: call [00401080h] ; __vbaHresultCheckObj
  loc_005A315F: mov var_4, 00000000h
loc_005A3166:   mov eax, Me
loc_005A3169:   push eax
loc_005A316A:   mov ecx, [eax]
loc_005A316C:   Call [ecx+00000008h]
loc_005A316F:   mov eax, var_4
loc_005A3172:   mov ecx, var_14
loc_005A3175:   pop edi
loc_005A3176:   pop esi
loc_005A3177:   mov fs: [00000000h] , ecx
loc_005A317E:   pop ebx
loc_005A317F:   mov esp, ebp
loc_005A3181:   pop ebp
  loc_005A3182: retn 0004h
End Sub

Private Sub mnuLapBarang_Click() '5A4F40
loc_005A4F40:   push ebp
loc_005A4F41:   mov ebp, esp
  loc_005A4F43: sub esp, 0000000Ch
  loc_005A4F46: push 00405696h ; __vbaExceptHandler
loc_005A4F4B:   mov eax, fs: [00000000h]
loc_005A4F51:   push eax
loc_005A4F52:   mov fs: [00000000h] , esp
  loc_005A4F59: sub esp, 00000008h
loc_005A4F5C:   push ebx
loc_005A4F5D:   push esi
loc_005A4F5E:   push edi
loc_005A4F5F:   mov var_C, esp
  loc_005A4F62: mov var_8, 00402C90h
loc_005A4F69:   mov eax, Me
loc_005A4F6C:   mov ecx, eax
  loc_005A4F6E: and ecx, 00000001h
loc_005A4F71:   mov var_4, ecx
  loc_005A4F74: and al, FEh
loc_005A4F76:   push eax
loc_005A4F77:   mov Me, eax
loc_005A4F7A:   mov edx, [eax]
loc_005A4F7C:   Call [edx+00000004h]
  loc_005A4F7F: call 0060E200h
  loc_005A4F84: mov var_4, 00000000h
loc_005A4F8B:   mov eax, Me
loc_005A4F8E:   push eax
loc_005A4F8F:   mov ecx, [eax]
loc_005A4F91:   Call [ecx+00000008h]
loc_005A4F94:   mov eax, var_4
loc_005A4F97:   mov ecx, var_14
loc_005A4F9A:   pop edi
loc_005A4F9B:   pop esi
loc_005A4F9C:   mov fs: [00000000h] , ecx
loc_005A4FA3:   pop ebx
loc_005A4FA4:   mov esp, ebp
loc_005A4FA6:   pop ebp
  loc_005A4FA7: retn 0004h
End Sub

Private Sub mnuLapPengguna_Click() '5A3630
loc_005A3630:   push ebp
loc_005A3631:   mov ebp, esp
  loc_005A3633: sub esp, 0000000Ch
  loc_005A3636: push 00405696h ; __vbaExceptHandler
loc_005A363B:   mov eax, fs: [00000000h]
loc_005A3641:   push eax
loc_005A3642:   mov fs: [00000000h] , esp
  loc_005A3649: sub esp, 00000008h
loc_005A364C:   push ebx
loc_005A364D:   push esi
loc_005A364E:   push edi
loc_005A364F:   mov var_C, esp
  loc_005A3652: mov var_8, 00402BB8h
loc_005A3659:   mov eax, Me
loc_005A365C:   mov ecx, eax
  loc_005A365E: and ecx, 00000001h
loc_005A3661:   mov var_4, ecx
  loc_005A3664: and al, FEh
loc_005A3666:   push eax
loc_005A3667:   mov Me, eax
loc_005A366A:   mov edx, [eax]
loc_005A366C:   Call [edx+00000004h]
  loc_005A366F: call 00609EC0h
  loc_005A3674: mov var_4, 00000000h
loc_005A367B:   mov eax, Me
loc_005A367E:   push eax
loc_005A367F:   mov ecx, [eax]
loc_005A3681:   Call [ecx+00000008h]
loc_005A3684:   mov eax, var_4
loc_005A3687:   mov ecx, var_14
loc_005A368A:   pop edi
loc_005A368B:   pop esi
loc_005A368C:   mov fs: [00000000h] , ecx
loc_005A3693:   pop ebx
loc_005A3694:   mov esp, ebp
loc_005A3696:   pop ebp
  loc_005A3697: retn 0004h
End Sub

Private Sub CmdPemasok_UnknownEvent_10() '5A26C0
loc_005A26C0:   push ebp
loc_005A26C1:   mov ebp, esp
  loc_005A26C3: sub esp, 0000000Ch
  loc_005A26C6: push 00405696h ; __vbaExceptHandler
loc_005A26CB:   mov eax, fs: [00000000h]
loc_005A26D1:   push eax
loc_005A26D2:   mov fs: [00000000h] , esp
  loc_005A26D9: sub esp, 00000030h
loc_005A26DC:   push ebx
loc_005A26DD:   push esi
loc_005A26DE:   push edi
loc_005A26DF:   mov var_C, esp
  loc_005A26E2: mov var_8, 00402B28h
loc_005A26E9:   mov eax, Me
loc_005A26EC:   mov ecx, eax
  loc_005A26EE: and ecx, 00000001h
loc_005A26F1:   mov var_4, ecx
  loc_005A26F4: and al, FEh
loc_005A26F6:   push eax
loc_005A26F7:   mov Me, eax
loc_005A26FA:   mov edx, [eax]
loc_005A26FC:   Call [edx+00000004h]
loc_005A26FF:   mov eax, [00668178h]
loc_005A2704:   test eax, eax
  loc_005A2706: jnz 005A2718h
  loc_005A2708: push 00668178h
  loc_005A270D: push 00416CD8h
  loc_005A2712: call [004011E8h] ; __vbaNew2
  loc_005A2718: sub esp, 00000010h
  loc_005A271B: mov ecx, 0000000Ah
loc_005A2720:   mov ebx, esp
  loc_005A2722: mov eax, 80020004h
  loc_005A2727: sub esp, 00000010h
loc_005A272A:   mov esi, [00668178h]
loc_005A2730:   mov [ebx], ecx
loc_005A2732:   mov ecx, var_30
loc_005A2735:   mov edi, [esi]
  loc_005A2737: mov edx, 00000001h
loc_005A273C:   mov [ebx+00000004h], ecx
loc_005A273F:   mov ecx, esp
loc_005A2741:   push esi
loc_005A2742:   mov [ebx+00000008h], eax
loc_005A2745:   mov eax, var_28
loc_005A2748:   mov [ebx+0000000Ch], eax
  loc_005A274B: mov eax, 00000002h
loc_005A2750:   mov [ecx], eax
loc_005A2752:   mov eax, var_20
loc_005A2755:   mov [ecx+00000004h], eax
loc_005A2758:   mov [ecx+00000008h], edx
loc_005A275B:   mov edx, var_18
loc_005A275E:   mov [ecx+0000000Ch], edx
loc_005A2761:   Call [edi+000002B0h]
loc_005A2767:   test eax, eax
loc_005A2769:   fnclex
  loc_005A276B: jge 005A277Fh
  loc_005A276D: push 000002B0h
  loc_005A2772: push 0042F6DCh ; "ø'³:©†éC·"
loc_005A2777:   push esi
loc_005A2778:   push eax
  loc_005A2779: call [00401080h] ; __vbaHresultCheckObj
  loc_005A277F: mov var_4, 00000000h
loc_005A2786:   mov eax, Me
loc_005A2789:   push eax
loc_005A278A:   mov ecx, [eax]
loc_005A278C:   Call [ecx+00000008h]
loc_005A278F:   mov eax, var_4
loc_005A2792:   mov ecx, var_14
loc_005A2795:   pop edi
loc_005A2796:   pop esi
loc_005A2797:   mov fs: [00000000h] , ecx
loc_005A279E:   pop ebx
loc_005A279F:   mov esp, ebp
loc_005A27A1:   pop ebp
  loc_005A27A2: retn 0004h
End Sub

Private Sub CmdBarang_UnknownEvent_10() '5A1D70
loc_005A1D70:   push ebp
loc_005A1D71:   mov ebp, esp
  loc_005A1D73: sub esp, 0000000Ch
  loc_005A1D76: push 00405696h ; __vbaExceptHandler
loc_005A1D7B:   mov eax, fs: [00000000h]
loc_005A1D81:   push eax
loc_005A1D82:   mov fs: [00000000h] , esp
  loc_005A1D89: sub esp, 00000030h
loc_005A1D8C:   push ebx
loc_005A1D8D:   push esi
loc_005A1D8E:   push edi
loc_005A1D8F:   mov var_C, esp
  loc_005A1D92: mov var_8, 00402AF0h
loc_005A1D99:   mov eax, Me
loc_005A1D9C:   mov ecx, eax
  loc_005A1D9E: and ecx, 00000001h
loc_005A1DA1:   mov var_4, ecx
  loc_005A1DA4: and al, FEh
loc_005A1DA6:   push eax
loc_005A1DA7:   mov Me, eax
loc_005A1DAA:   mov edx, [eax]
loc_005A1DAC:   Call [edx+00000004h]
loc_005A1DAF:   mov eax, [006681DCh]
loc_005A1DB4:   test eax, eax
  loc_005A1DB6: jnz 005A1DC8h
  loc_005A1DB8: push 006681DCh
  loc_005A1DBD: push 0041EC78h
  loc_005A1DC2: call [004011E8h] ; __vbaNew2
  loc_005A1DC8: sub esp, 00000010h
  loc_005A1DCB: mov ecx, 0000000Ah
loc_005A1DD0:   mov ebx, esp
  loc_005A1DD2: mov eax, 80020004h
  loc_005A1DD7: sub esp, 00000010h
loc_005A1DDA:   mov esi, [006681DCh]
loc_005A1DE0:   mov [ebx], ecx
loc_005A1DE2:   mov ecx, var_30
loc_005A1DE5:   mov edi, [esi]
  loc_005A1DE7: mov edx, 00000001h
loc_005A1DEC:   mov [ebx+00000004h], ecx
loc_005A1DEF:   mov ecx, esp
loc_005A1DF1:   push esi
loc_005A1DF2:   mov [ebx+00000008h], eax
loc_005A1DF5:   mov eax, var_28
loc_005A1DF8:   mov [ebx+0000000Ch], eax
  loc_005A1DFB: mov eax, 00000002h
loc_005A1E00:   mov [ecx], eax
loc_005A1E02:   mov eax, var_20
loc_005A1E05:   mov [ecx+00000004h], eax
loc_005A1E08:   mov [ecx+00000008h], edx
loc_005A1E0B:   mov edx, var_18
loc_005A1E0E:   mov [ecx+0000000Ch], edx
loc_005A1E11:   Call [edi+000002B0h]
loc_005A1E17:   test eax, eax
loc_005A1E19:   fnclex
  loc_005A1E1B: jge 005A1E2Fh
  loc_005A1E1D: push 000002B0h
  loc_005A1E22: push 00431BBCh
loc_005A1E27:   push esi
loc_005A1E28:   push eax
  loc_005A1E29: call [00401080h] ; __vbaHresultCheckObj
  loc_005A1E2F: mov var_4, 00000000h
loc_005A1E36:   mov eax, Me
loc_005A1E39:   push eax
loc_005A1E3A:   mov ecx, [eax]
loc_005A1E3C:   Call [ecx+00000008h]
loc_005A1E3F:   mov eax, var_4
loc_005A1E42:   mov ecx, var_14
loc_005A1E45:   pop edi
loc_005A1E46:   pop esi
loc_005A1E47:   mov fs: [00000000h] , ecx
loc_005A1E4E:   pop ebx
loc_005A1E4F:   mov esp, ebp
loc_005A1E51:   pop ebp
  loc_005A1E52: retn 0004h
End Sub

Private Sub MnuLapReturBeliSemua_Click() '5A3970
loc_005A3970:   push ebp
loc_005A3971:   mov ebp, esp
  loc_005A3973: sub esp, 0000000Ch
  loc_005A3976: push 00405696h ; __vbaExceptHandler
loc_005A397B:   mov eax, fs: [00000000h]
loc_005A3981:   push eax
loc_005A3982:   mov fs: [00000000h] , esp
  loc_005A3989: sub esp, 00000008h
loc_005A398C:   push ebx
loc_005A398D:   push esi
loc_005A398E:   push edi
loc_005A398F:   mov var_C, esp
  loc_005A3992: mov var_8, 00402BD8h
loc_005A3999:   mov eax, Me
loc_005A399C:   mov ecx, eax
  loc_005A399E: and ecx, 00000001h
loc_005A39A1:   mov var_4, ecx
  loc_005A39A4: and al, FEh
loc_005A39A6:   push eax
loc_005A39A7:   mov Me, eax
loc_005A39AA:   mov edx, [eax]
loc_005A39AC:   Call [edx+00000004h]
  loc_005A39AF: call 00610AA0h
  loc_005A39B4: mov var_4, 00000000h
loc_005A39BB:   mov eax, Me
loc_005A39BE:   push eax
loc_005A39BF:   mov ecx, [eax]
loc_005A39C1:   Call [ecx+00000008h]
loc_005A39C4:   mov eax, var_4
loc_005A39C7:   mov ecx, var_14
loc_005A39CA:   pop edi
loc_005A39CB:   pop esi
loc_005A39CC:   mov fs: [00000000h] , ecx
loc_005A39D3:   pop ebx
loc_005A39D4:   mov esp, ebp
loc_005A39D6:   pop ebp
  loc_005A39D7: retn 0004h
End Sub

Private Sub mnuLapGolongan_Click() '5A4D70
loc_005A4D70:   push ebp
loc_005A4D71:   mov ebp, esp
  loc_005A4D73: sub esp, 0000000Ch
  loc_005A4D76: push 00405696h ; __vbaExceptHandler
loc_005A4D7B:   mov eax, fs: [00000000h]
loc_005A4D81:   push eax
loc_005A4D82:   mov fs: [00000000h] , esp
  loc_005A4D89: sub esp, 00000008h
loc_005A4D8C:   push ebx
loc_005A4D8D:   push esi
loc_005A4D8E:   push edi
loc_005A4D8F:   mov var_C, esp
  loc_005A4D92: mov var_8, 00402C78h
loc_005A4D99:   mov eax, Me
loc_005A4D9C:   mov ecx, eax
  loc_005A4D9E: and ecx, 00000001h
loc_005A4DA1:   mov var_4, ecx
  loc_005A4DA4: and al, FEh
loc_005A4DA6:   push eax
loc_005A4DA7:   mov Me, eax
loc_005A4DAA:   mov edx, [eax]
loc_005A4DAC:   Call [edx+00000004h]
  loc_005A4DAF: call 0060C740h
  loc_005A4DB4: mov var_4, 00000000h
loc_005A4DBB:   mov eax, Me
loc_005A4DBE:   push eax
loc_005A4DBF:   mov ecx, [eax]
loc_005A4DC1:   Call [ecx+00000008h]
loc_005A4DC4:   mov eax, var_4
loc_005A4DC7:   mov ecx, var_14
loc_005A4DCA:   pop edi
loc_005A4DCB:   pop esi
loc_005A4DCC:   mov fs: [00000000h] , ecx
loc_005A4DD3:   pop ebx
loc_005A4DD4:   mov esp, ebp
loc_005A4DD6:   pop ebp
  loc_005A4DD7: retn 0004h
End Sub

Private Sub MnuPiutang_Click() '5A4850
loc_005A4850:   push ebp
loc_005A4851:   mov ebp, esp
  loc_005A4853: sub esp, 0000000Ch
  loc_005A4856: push 00405696h ; __vbaExceptHandler
loc_005A485B:   mov eax, fs: [00000000h]
loc_005A4861:   push eax
loc_005A4862:   mov fs: [00000000h] , esp
  loc_005A4869: sub esp, 00000030h
loc_005A486C:   push ebx
loc_005A486D:   push esi
loc_005A486E:   push edi
loc_005A486F:   mov var_C, esp
  loc_005A4872: mov var_8, 00402C48h
loc_005A4879:   mov eax, Me
loc_005A487C:   mov ecx, eax
  loc_005A487E: and ecx, 00000001h
loc_005A4881:   mov var_4, ecx
  loc_005A4884: and al, FEh
loc_005A4886:   push eax
loc_005A4887:   mov Me, eax
loc_005A488A:   mov edx, [eax]
loc_005A488C:   Call [edx+00000004h]
loc_005A488F:   mov eax, [0066838Ch]
loc_005A4894:   test eax, eax
  loc_005A4896: jnz 005A48A8h
  loc_005A4898: push 0066838Ch
  loc_005A489D: push 0041C3D0h
  loc_005A48A2: call [004011E8h] ; __vbaNew2
  loc_005A48A8: sub esp, 00000010h
  loc_005A48AB: mov ecx, 0000000Ah
loc_005A48B0:   mov ebx, esp
  loc_005A48B2: mov eax, 80020004h
  loc_005A48B7: sub esp, 00000010h
loc_005A48BA:   mov esi, [0066838Ch]
loc_005A48C0:   mov [ebx], ecx
loc_005A48C2:   mov ecx, var_30
loc_005A48C5:   mov edi, [esi]
  loc_005A48C7: mov edx, 00000001h
loc_005A48CC:   mov [ebx+00000004h], ecx
loc_005A48CF:   mov ecx, esp
loc_005A48D1:   push esi
loc_005A48D2:   mov [ebx+00000008h], eax
loc_005A48D5:   mov eax, var_28
loc_005A48D8:   mov [ebx+0000000Ch], eax
  loc_005A48DB: mov eax, 00000002h
loc_005A48E0:   mov [ecx], eax
loc_005A48E2:   mov eax, var_20
loc_005A48E5:   mov [ecx+00000004h], eax
loc_005A48E8:   mov [ecx+00000008h], edx
loc_005A48EB:   mov edx, var_18
loc_005A48EE:   mov [ecx+0000000Ch], edx
loc_005A48F1:   Call [edi+000002B0h]
loc_005A48F7:   test eax, eax
loc_005A48F9:   fnclex
  loc_005A48FB: jge 005A490Fh
  loc_005A48FD: push 000002B0h
  loc_005A4902: push 0043557Ch
loc_005A4907:   push esi
loc_005A4908:   push eax
  loc_005A4909: call [00401080h] ; __vbaHresultCheckObj
  loc_005A490F: mov var_4, 00000000h
loc_005A4916:   mov eax, Me
loc_005A4919:   push eax
loc_005A491A:   mov ecx, [eax]
loc_005A491C:   Call [ecx+00000008h]
loc_005A491F:   mov eax, var_4
loc_005A4922:   mov ecx, var_14
loc_005A4925:   pop edi
loc_005A4926:   pop esi
loc_005A4927:   mov fs: [00000000h] , ecx
loc_005A492E:   pop ebx
loc_005A492F:   mov esp, ebp
loc_005A4931:   pop ebp
  loc_005A4932: retn 0004h
End Sub

Private Sub mnuLapJenis_Click() '5A4DE0
loc_005A4DE0:   push ebp
loc_005A4DE1:   mov ebp, esp
  loc_005A4DE3: sub esp, 0000000Ch
  loc_005A4DE6: push 00405696h ; __vbaExceptHandler
loc_005A4DEB:   mov eax, fs: [00000000h]
loc_005A4DF1:   push eax
loc_005A4DF2:   mov fs: [00000000h] , esp
  loc_005A4DF9: sub esp, 00000008h
loc_005A4DFC:   push ebx
loc_005A4DFD:   push esi
loc_005A4DFE:   push edi
loc_005A4DFF:   mov var_C, esp
  loc_005A4E02: mov var_8, 00402C80h
loc_005A4E09:   mov eax, Me
loc_005A4E0C:   mov ecx, eax
  loc_005A4E0E: and ecx, 00000001h
loc_005A4E11:   mov var_4, ecx
  loc_005A4E14: and al, FEh
loc_005A4E16:   push eax
loc_005A4E17:   mov Me, eax
loc_005A4E1A:   mov edx, [eax]
loc_005A4E1C:   Call [edx+00000004h]
  loc_005A4E1F: call 0060D4C0h
  loc_005A4E24: mov var_4, 00000000h
loc_005A4E2B:   mov eax, Me
loc_005A4E2E:   push eax
loc_005A4E2F:   mov ecx, [eax]
loc_005A4E31:   Call [ecx+00000008h]
loc_005A4E34:   mov eax, var_4
loc_005A4E37:   mov ecx, var_14
loc_005A4E3A:   pop edi
loc_005A4E3B:   pop esi
loc_005A4E3C:   mov fs: [00000000h] , ecx
loc_005A4E43:   pop ebx
loc_005A4E44:   mov esp, ebp
loc_005A4E46:   pop ebp
  loc_005A4E47: retn 0004h
End Sub

Private Sub CmdPembelian_UnknownEvent_10() '5A27B0
loc_005A27B0:   push ebp
loc_005A27B1:   mov ebp, esp
  loc_005A27B3: sub esp, 0000000Ch
  loc_005A27B6: push 00405696h ; __vbaExceptHandler
loc_005A27BB:   mov eax, fs: [00000000h]
loc_005A27C1:   push eax
loc_005A27C2:   mov fs: [00000000h] , esp
  loc_005A27C9: sub esp, 00000030h
loc_005A27CC:   push ebx
loc_005A27CD:   push esi
loc_005A27CE:   push edi
loc_005A27CF:   mov var_C, esp
  loc_005A27D2: mov var_8, 00402B30h
loc_005A27D9:   mov eax, Me
loc_005A27DC:   mov ecx, eax
  loc_005A27DE: and ecx, 00000001h
loc_005A27E1:   mov var_4, ecx
  loc_005A27E4: and al, FEh
loc_005A27E6:   push eax
loc_005A27E7:   mov Me, eax
loc_005A27EA:   mov edx, [eax]
loc_005A27EC:   Call [edx+00000004h]
loc_005A27EF:   mov eax, [0066823Ch]
loc_005A27F4:   test eax, eax
  loc_005A27F6: jnz 005A2808h
  loc_005A27F8: push 0066823Ch
  loc_005A27FD: push 00427B78h
  loc_005A2802: call [004011E8h] ; __vbaNew2
  loc_005A2808: sub esp, 00000010h
  loc_005A280B: mov ecx, 0000000Ah
loc_005A2810:   mov ebx, esp
  loc_005A2812: mov eax, 80020004h
  loc_005A2817: sub esp, 00000010h
loc_005A281A:   mov esi, [0066823Ch]
loc_005A2820:   mov [ebx], ecx
loc_005A2822:   mov ecx, var_30
loc_005A2825:   mov edi, [esi]
  loc_005A2827: mov edx, 00000001h
loc_005A282C:   mov [ebx+00000004h], ecx
loc_005A282F:   mov ecx, esp
loc_005A2831:   push esi
loc_005A2832:   mov [ebx+00000008h], eax
loc_005A2835:   mov eax, var_28
loc_005A2838:   mov [ebx+0000000Ch], eax
  loc_005A283B: mov eax, 00000002h
loc_005A2840:   mov [ecx], eax
loc_005A2842:   mov eax, var_20
loc_005A2845:   mov [ecx+00000004h], eax
loc_005A2848:   mov [ecx+00000008h], edx
loc_005A284B:   mov edx, var_18
loc_005A284E:   mov [ecx+0000000Ch], edx
loc_005A2851:   Call [edi+000002B0h]
loc_005A2857:   test eax, eax
loc_005A2859:   fnclex
  loc_005A285B: jge 005A286Fh
  loc_005A285D: push 000002B0h
  loc_005A2862: push 00433DA8h
loc_005A2867:   push esi
loc_005A2868:   push eax
  loc_005A2869: call [00401080h] ; __vbaHresultCheckObj
  loc_005A286F: mov var_4, 00000000h
loc_005A2876:   mov eax, Me
loc_005A2879:   push eax
loc_005A287A:   mov ecx, [eax]
loc_005A287C:   Call [ecx+00000008h]
loc_005A287F:   mov eax, var_4
loc_005A2882:   mov ecx, var_14
loc_005A2885:   pop edi
loc_005A2886:   pop esi
loc_005A2887:   mov fs: [00000000h] , ecx
loc_005A288E:   pop ebx
loc_005A288F:   mov esp, ebp
loc_005A2891:   pop ebp
  loc_005A2892: retn 0004h
End Sub

Private Sub CmdLogOut_UnknownEvent_10() '5A2020
loc_005A2020:   push ebp
loc_005A2021:   mov ebp, esp
  loc_005A2023: sub esp, 0000000Ch
  loc_005A2026: push 00405696h ; __vbaExceptHandler
loc_005A202B:   mov eax, fs: [00000000h]
loc_005A2031:   push eax
loc_005A2032:   mov fs: [00000000h] , esp
  loc_005A2039: sub esp, 0000003Ch
loc_005A203C:   push ebx
loc_005A203D:   push esi
loc_005A203E:   push edi
loc_005A203F:   mov var_C, esp
  loc_005A2042: mov var_8, 00402B10h
loc_005A2049:   mov esi, Me
loc_005A204C:   mov eax, esi
  loc_005A204E: and eax, 00000001h
loc_005A2051:   mov var_4, eax
  loc_005A2054: and esi, FFFFFFFEh
loc_005A2057:   push esi
loc_005A2058:   mov Me, esi
loc_005A205B:   mov ecx, [esi]
loc_005A205D:   Call [ecx+00000004h]
loc_005A2060:   mov eax, [006682A0h]
  loc_005A2065: xor ecx, ecx
loc_005A2067:   cmp eax, ecx
loc_005A2069:   mov var_18, ecx
loc_005A206C:   mov var_1C, ecx
  loc_005A206F: jnz 005A2086h
  loc_005A2071: push 006682A0h
  loc_005A2076: push 00414A94h
  loc_005A207B: call [004011E8h] ; __vbaNew2
loc_005A2081:   mov eax, [006682A0h]
loc_005A2086:   mov edx, [eax]
loc_005A2088:   push eax
loc_005A2089:   Call [edx+00000308h]
  loc_005A208F: mov ebx, [004010B8h] ; __vbaObjSet
loc_005A2095:   push eax
loc_005A2096:   lea eax, var_1C
loc_005A2099:   push eax
loc_005A209A:   Call ebx
loc_005A209C:   mov edi, eax
loc_005A209E:   lea edx, var_18
loc_005A20A1:   push edx
loc_005A20A2:   push edi
loc_005A20A3:   mov ecx, [edi]
loc_005A20A5:   Call [ecx+000000A0h]
loc_005A20AB:   test eax, eax
loc_005A20AD:   fnclex
  loc_005A20AF: jge 005A20C3h
  loc_005A20B1: push 000000A0h
  loc_005A20B6: push 0042DFCCh
loc_005A20BB:   push edi
loc_005A20BC:   push eax
  loc_005A20BD: call [00401080h] ; __vbaHresultCheckObj
loc_005A20C3:   mov eax, var_18
loc_005A20C6:   push eax
  loc_005A20C7: push 00434FA8h ; "UNREGISTERED"
  loc_005A20CC: call [0040112Ch] ; __vbaStrCmp
loc_005A20D2:   neg eax
loc_005A20D4:   sbb eax, eax
loc_005A20D6:   lea ecx, var_18
loc_005A20D9:   inc eax
loc_005A20DA:   neg eax
loc_005A20DC:   mov var_48, eax
  loc_005A20DF: call [0040129Ch] ; __vbaFreeStr
  loc_005A20E5: mov edi, [004012A0h] ; __vbaFreeObj
loc_005A20EB:   lea ecx, var_1C
loc_005A20EE:   Call edi
  loc_005A20F0: cmp var_48, 0000h
  loc_005A20F5: jz 005A21A3h
loc_005A20FB:   mov ecx, [esi]
loc_005A20FD:   push esi
loc_005A20FE:   Call [ecx+00000358h]
loc_005A2104:   lea edx, var_1C
loc_005A2107:   push eax
loc_005A2108:   push edx
loc_005A2109:   Call ebx
loc_005A210B:   mov ecx, [eax]
  loc_005A210D: push 00000000h
loc_005A210F:   push eax
loc_005A2110:   mov var_40, eax
loc_005A2113:   Call [ecx+00000074h]
loc_005A2116:   test eax, eax
loc_005A2118:   fnclex
  loc_005A211A: jge 005A212Eh
loc_005A211C:   mov edx, var_40
  loc_005A211F: push 00000074h
  loc_005A2121: push 00434274h
loc_005A2126:   push edx
loc_005A2127:   push eax
  loc_005A2128: call [00401080h] ; __vbaHresultCheckObj
loc_005A212E:   lea ecx, var_1C
loc_005A2131:   Call edi
loc_005A2133:   mov eax, [esi]
loc_005A2135:   push esi
loc_005A2136:   Call [eax+0000043Ch]
loc_005A213C:   lea ecx, var_1C
loc_005A213F:   push eax
loc_005A2140:   push ecx
loc_005A2141:   Call ebx
loc_005A2143:   mov edx, [eax]
  loc_005A2145: push 00000000h
loc_005A2147:   push eax
loc_005A2148:   mov var_40, eax
loc_005A214B:   Call [edx+00000074h]
loc_005A214E:   test eax, eax
loc_005A2150:   fnclex
  loc_005A2152: jge 005A2166h
loc_005A2154:   mov ecx, var_40
  loc_005A2157: push 00000074h
  loc_005A2159: push 00434274h
loc_005A215E:   push ecx
loc_005A215F:   push eax
  loc_005A2160: call [00401080h] ; __vbaHresultCheckObj
loc_005A2166:   lea ecx, var_1C
loc_005A2169:   Call edi
loc_005A216B:   mov edx, [esi]
loc_005A216D:   push esi
loc_005A216E:   Call [edx+000003A4h]
loc_005A2174:   push eax
loc_005A2175:   lea eax, var_1C
loc_005A2178:   push eax
loc_005A2179:   Call ebx
loc_005A217B:   mov ecx, [eax]
  loc_005A217D: push 00000000h
loc_005A217F:   push eax
loc_005A2180:   mov var_40, eax
loc_005A2183:   Call [ecx+00000074h]
loc_005A2186:   test eax, eax
loc_005A2188:   fnclex
  loc_005A218A: jge 005A219Eh
loc_005A218C:   mov edx, var_40
  loc_005A218F: push 00000074h
  loc_005A2191: push 00434274h
loc_005A2196:   push edx
loc_005A2197:   push eax
  loc_005A2198: call [00401080h] ; __vbaHresultCheckObj
loc_005A219E:   lea ecx, var_1C
loc_005A21A1:   Call edi
loc_005A21A3:   mov eax, [esi]
loc_005A21A5:   push esi
loc_005A21A6:   Call [eax+000003A4h]
loc_005A21AC:   lea ecx, var_1C
loc_005A21AF:   push eax
loc_005A21B0:   push ecx
loc_005A21B1:   Call ebx
loc_005A21B3:   mov edx, [eax]
  loc_005A21B5: push 00000000h
loc_005A21B7:   push eax
loc_005A21B8:   mov var_40, eax
loc_005A21BB:   Call [edx+00000074h]
loc_005A21BE:   test eax, eax
loc_005A21C0:   fnclex
  loc_005A21C2: jge 005A21D6h
loc_005A21C4:   mov ecx, var_40
  loc_005A21C7: push 00000074h
  loc_005A21C9: push 00434274h
loc_005A21CE:   push ecx
loc_005A21CF:   push eax
  loc_005A21D0: call [00401080h] ; __vbaHresultCheckObj
loc_005A21D6:   lea ecx, var_1C
loc_005A21D9:   Call edi
loc_005A21DB:   mov edx, [esi]
loc_005A21DD:   push esi
loc_005A21DE:   Call [edx+00000358h]
loc_005A21E4:   push eax
loc_005A21E5:   lea eax, var_1C
loc_005A21E8:   push eax
loc_005A21E9:   Call ebx
loc_005A21EB:   mov ecx, [eax]
  loc_005A21ED: push 00000000h
loc_005A21EF:   push eax
loc_005A21F0:   mov var_40, eax
loc_005A21F3:   Call [ecx+00000074h]
loc_005A21F6:   test eax, eax
loc_005A21F8:   fnclex
  loc_005A21FA: jge 005A220Eh
loc_005A21FC:   mov edx, var_40
  loc_005A21FF: push 00000074h
  loc_005A2201: push 00434274h
loc_005A2206:   push edx
loc_005A2207:   push eax
  loc_005A2208: call [00401080h] ; __vbaHresultCheckObj
loc_005A220E:   lea ecx, var_1C
loc_005A2211:   Call edi
loc_005A2213:   mov eax, [esi]
loc_005A2215:   push esi
loc_005A2216:   Call [eax+00000384h]
loc_005A221C:   lea ecx, var_1C
loc_005A221F:   push eax
loc_005A2220:   push ecx
loc_005A2221:   Call ebx
loc_005A2223:   mov edx, [eax]
  loc_005A2225: push 00000000h
loc_005A2227:   push eax
loc_005A2228:   mov var_40, eax
loc_005A222B:   Call [edx+00000074h]
loc_005A222E:   test eax, eax
loc_005A2230:   fnclex
  loc_005A2232: jge 005A2246h
loc_005A2234:   mov ecx, var_40
  loc_005A2237: push 00000074h
  loc_005A2239: push 00434274h
loc_005A223E:   push ecx
loc_005A223F:   push eax
  loc_005A2240: call [00401080h] ; __vbaHresultCheckObj
loc_005A2246:   lea ecx, var_1C
loc_005A2249:   Call edi
loc_005A224B:   mov edx, [esi]
loc_005A224D:   push esi
loc_005A224E:   Call [edx+000003A4h]
loc_005A2254:   push eax
loc_005A2255:   lea eax, var_1C
loc_005A2258:   push eax
loc_005A2259:   Call ebx
loc_005A225B:   mov ecx, [eax]
  loc_005A225D: push 00000000h
loc_005A225F:   push eax
loc_005A2260:   mov var_40, eax
loc_005A2263:   Call [ecx+00000074h]
loc_005A2266:   test eax, eax
loc_005A2268:   fnclex
  loc_005A226A: jge 005A227Eh
loc_005A226C:   mov edx, var_40
  loc_005A226F: push 00000074h
  loc_005A2271: push 00434274h
loc_005A2276:   push edx
loc_005A2277:   push eax
  loc_005A2278: call [00401080h] ; __vbaHresultCheckObj
loc_005A227E:   lea ecx, var_1C
loc_005A2281:   Call edi
loc_005A2283:   mov eax, [esi]
loc_005A2285:   push esi
loc_005A2286:   Call [eax+00000344h]
loc_005A228C:   lea ecx, var_1C
loc_005A228F:   push eax
loc_005A2290:   push ecx
loc_005A2291:   Call ebx
loc_005A2293:   mov edx, [eax]
loc_005A2295:   push FFFFFFFFh
loc_005A2297:   push eax
loc_005A2298:   mov var_40, eax
loc_005A229B:   Call [edx+00000074h]
loc_005A229E:   test eax, eax
loc_005A22A0:   fnclex
  loc_005A22A2: jge 005A22B6h
loc_005A22A4:   mov ecx, var_40
  loc_005A22A7: push 00000074h
  loc_005A22A9: push 00434274h
loc_005A22AE:   push ecx
loc_005A22AF:   push eax
  loc_005A22B0: call [00401080h] ; __vbaHresultCheckObj
loc_005A22B6:   lea ecx, var_1C
loc_005A22B9:   Call edi
loc_005A22BB:   mov edx, [esi]
loc_005A22BD:   push esi
loc_005A22BE:   Call [edx+00000348h]
loc_005A22C4:   push eax
loc_005A22C5:   lea eax, var_1C
loc_005A22C8:   push eax
loc_005A22C9:   Call ebx
loc_005A22CB:   mov ecx, [eax]
  loc_005A22CD: push 00000000h
loc_005A22CF:   push eax
loc_005A22D0:   mov var_40, eax
loc_005A22D3:   Call [ecx+00000074h]
loc_005A22D6:   test eax, eax
loc_005A22D8:   fnclex
  loc_005A22DA: jge 005A22EEh
loc_005A22DC:   mov edx, var_40
  loc_005A22DF: push 00000074h
  loc_005A22E1: push 00434274h
loc_005A22E6:   push edx
loc_005A22E7:   push eax
  loc_005A22E8: call [00401080h] ; __vbaHresultCheckObj
loc_005A22EE:   lea ecx, var_1C
loc_005A22F1:   Call edi
loc_005A22F3:   mov eax, [esi]
loc_005A22F5:   push esi
loc_005A22F6:   Call [eax+0000034Ch]
loc_005A22FC:   lea ecx, var_1C
loc_005A22FF:   push eax
loc_005A2300:   push ecx
loc_005A2301:   Call ebx
loc_005A2303:   mov edx, [eax]
  loc_005A2305: push 00000000h
loc_005A2307:   push eax
loc_005A2308:   mov var_40, eax
loc_005A230B:   Call [edx+00000074h]
loc_005A230E:   test eax, eax
loc_005A2310:   fnclex
  loc_005A2312: jge 005A2326h
loc_005A2314:   mov ecx, var_40
  loc_005A2317: push 00000074h
  loc_005A2319: push 00434274h
loc_005A231E:   push ecx
loc_005A231F:   push eax
  loc_005A2320: call [00401080h] ; __vbaHresultCheckObj
loc_005A2326:   lea ecx, var_1C
loc_005A2329:   Call edi
loc_005A232B:   mov edx, [esi]
loc_005A232D:   push esi
loc_005A232E:   Call [edx+00000424h]
loc_005A2334:   push eax
loc_005A2335:   lea eax, var_1C
loc_005A2338:   push eax
loc_005A2339:   Call ebx
loc_005A233B:   mov ecx, [eax]
  loc_005A233D: push 00000000h
loc_005A233F:   push eax
loc_005A2340:   mov var_40, eax
loc_005A2343:   Call [ecx+00000074h]
loc_005A2346:   test eax, eax
loc_005A2348:   fnclex
  loc_005A234A: jge 005A235Eh
loc_005A234C:   mov edx, var_40
  loc_005A234F: push 00000074h
  loc_005A2351: push 00434274h
loc_005A2356:   push edx
loc_005A2357:   push eax
  loc_005A2358: call [00401080h] ; __vbaHresultCheckObj
loc_005A235E:   lea ecx, var_1C
loc_005A2361:   Call edi
loc_005A2363:   mov eax, [esi]
loc_005A2365:   push esi
loc_005A2366:   Call [eax+00000440h]
loc_005A236C:   lea ecx, var_1C
loc_005A236F:   push eax
loc_005A2370:   push ecx
loc_005A2371:   Call ebx
loc_005A2373:   mov edx, [eax]
  loc_005A2375: push 00000000h
loc_005A2377:   push eax
loc_005A2378:   mov var_40, eax
loc_005A237B:   Call [edx+00000074h]
loc_005A237E:   test eax, eax
loc_005A2380:   fnclex
  loc_005A2382: jge 005A2396h
loc_005A2384:   mov ecx, var_40
  loc_005A2387: push 00000074h
  loc_005A2389: push 00434274h
loc_005A238E:   push ecx
loc_005A238F:   push eax
  loc_005A2390: call [00401080h] ; __vbaHresultCheckObj
loc_005A2396:   lea ecx, var_1C
loc_005A2399:   Call edi
  loc_005A239B: sub esp, 00000010h
  loc_005A239E: mov ecx, 0000000Bh
loc_005A23A3:   mov edx, esp
  loc_005A23A5: xor eax, eax
  loc_005A23A7: push 8001000Dh
loc_005A23AC:   push esi
loc_005A23AD:   mov [edx], ecx
loc_005A23AF:   mov ecx, var_28
loc_005A23B2:   mov [edx+00000004h], ecx
loc_005A23B5:   mov ecx, [esi]
loc_005A23B7:   mov [edx+00000008h], eax
loc_005A23BA:   mov eax, var_20
loc_005A23BD:   mov [edx+0000000Ch], eax
loc_005A23C0:   Call [ecx+00000460h]
loc_005A23C6:   lea edx, var_1C
loc_005A23C9:   push eax
loc_005A23CA:   push edx
loc_005A23CB:   Call ebx
loc_005A23CD:   push eax
  loc_005A23CE: call [00401280h] ; __vbaLateIdSt
loc_005A23D4:   lea ecx, var_1C
loc_005A23D7:   Call edi
  loc_005A23D9: sub esp, 00000010h
  loc_005A23DC: mov ecx, 0000000Bh
loc_005A23E1:   mov edx, esp
  loc_005A23E3: xor eax, eax
  loc_005A23E5: push 8001000Dh
loc_005A23EA:   push esi
loc_005A23EB:   mov [edx], ecx
loc_005A23ED:   mov ecx, var_28
loc_005A23F0:   mov [edx+00000004h], ecx
loc_005A23F3:   mov ecx, [esi]
loc_005A23F5:   mov [edx+00000008h], eax
loc_005A23F8:   mov eax, var_20
loc_005A23FB:   mov [edx+0000000Ch], eax
loc_005A23FE:   Call [ecx+00000464h]
loc_005A2404:   lea edx, var_1C
loc_005A2407:   push eax
loc_005A2408:   push edx
loc_005A2409:   Call ebx
loc_005A240B:   push eax
  loc_005A240C: call [00401280h] ; __vbaLateIdSt
loc_005A2412:   lea ecx, var_1C
loc_005A2415:   Call edi
  loc_005A2417: sub esp, 00000010h
  loc_005A241A: mov ecx, 0000000Bh
loc_005A241F:   mov edx, esp
  loc_005A2421: xor eax, eax
  loc_005A2423: push 8001000Dh
loc_005A2428:   push esi
loc_005A2429:   mov [edx], ecx
loc_005A242B:   mov ecx, var_28
loc_005A242E:   mov [edx+00000004h], ecx
loc_005A2431:   mov ecx, [esi]
loc_005A2433:   mov [edx+00000008h], eax
loc_005A2436:   mov eax, var_20
loc_005A2439:   mov [edx+0000000Ch], eax
loc_005A243C:   Call [ecx+00000468h]
loc_005A2442:   lea edx, var_1C
loc_005A2445:   push eax
loc_005A2446:   push edx
loc_005A2447:   Call ebx
loc_005A2449:   push eax
  loc_005A244A: call [00401280h] ; __vbaLateIdSt
loc_005A2450:   lea ecx, var_1C
loc_005A2453:   Call edi
  loc_005A2455: xor eax, eax
  loc_005A2457: sub esp, 00000010h
loc_005A245A:   mov edx, esp
  loc_005A245C: mov ecx, 0000000Bh
  loc_005A2461: push 8001000Dh
loc_005A2466:   push esi
loc_005A2467:   mov [edx], ecx
loc_005A2469:   mov ecx, var_28
loc_005A246C:   mov [edx+00000004h], ecx
loc_005A246F:   mov ecx, [esi]
loc_005A2471:   mov [edx+00000008h], eax
loc_005A2474:   mov eax, var_20
loc_005A2477:   mov [edx+0000000Ch], eax
loc_005A247A:   Call [ecx+0000044Ch]
loc_005A2480:   lea edx, var_1C
loc_005A2483:   push eax
loc_005A2484:   push edx
loc_005A2485:   Call ebx
loc_005A2487:   push eax
  loc_005A2488: call [00401280h] ; __vbaLateIdSt
loc_005A248E:   lea ecx, var_1C
loc_005A2491:   Call edi
  loc_005A2493: sub esp, 00000010h
  loc_005A2496: mov ecx, 0000000Bh
loc_005A249B:   mov edx, esp
  loc_005A249D: xor eax, eax
  loc_005A249F: push 8001000Dh
loc_005A24A4:   push esi
loc_005A24A5:   mov [edx], ecx
loc_005A24A7:   mov ecx, var_28
loc_005A24AA:   mov [edx+00000004h], ecx
loc_005A24AD:   mov ecx, [esi]
loc_005A24AF:   mov [edx+00000008h], eax
loc_005A24B2:   mov eax, var_20
loc_005A24B5:   mov [edx+0000000Ch], eax
loc_005A24B8:   Call [ecx+00000450h]
loc_005A24BE:   lea edx, var_1C
loc_005A24C1:   push eax
loc_005A24C2:   push edx
loc_005A24C3:   Call ebx
loc_005A24C5:   push eax
  loc_005A24C6: call [00401280h] ; __vbaLateIdSt
loc_005A24CC:   lea ecx, var_1C
loc_005A24CF:   Call edi
  loc_005A24D1: sub esp, 00000010h
  loc_005A24D4: mov ecx, 0000000Bh
loc_005A24D9:   mov edx, esp
  loc_005A24DB: xor eax, eax
  loc_005A24DD: push 8001000Dh
loc_005A24E2:   push esi
loc_005A24E3:   mov [edx], ecx
loc_005A24E5:   mov ecx, var_28
loc_005A24E8:   mov [edx+00000004h], ecx
loc_005A24EB:   mov ecx, [esi]
loc_005A24ED:   mov [edx+00000008h], eax
loc_005A24F0:   mov eax, var_20
loc_005A24F3:   mov [edx+0000000Ch], eax
loc_005A24F6:   Call [ecx+00000454h]
loc_005A24FC:   lea edx, var_1C
loc_005A24FF:   push eax
loc_005A2500:   push edx
loc_005A2501:   Call ebx
loc_005A2503:   push eax
  loc_005A2504: call [00401280h] ; __vbaLateIdSt
loc_005A250A:   lea ecx, var_1C
loc_005A250D:   Call edi
  loc_005A250F: sub esp, 00000010h
  loc_005A2512: mov ecx, 0000000Bh
loc_005A2517:   mov edx, esp
  loc_005A2519: or eax, FFFFFFFFh
  loc_005A251C: push 8001000Dh
loc_005A2521:   push esi
loc_005A2522:   mov [edx], ecx
loc_005A2524:   mov ecx, var_28
loc_005A2527:   mov [edx+00000004h], ecx
loc_005A252A:   mov ecx, [esi]
loc_005A252C:   mov [edx+00000008h], eax
loc_005A252F:   mov eax, var_20
loc_005A2532:   mov [edx+0000000Ch], eax
loc_005A2535:   Call [ecx+0000045Ch]
loc_005A253B:   lea edx, var_1C
loc_005A253E:   push eax
loc_005A253F:   push edx
loc_005A2540:   Call ebx
loc_005A2542:   push eax
  loc_005A2543: call [00401280h] ; __vbaLateIdSt
loc_005A2549:   lea ecx, var_1C
loc_005A254C:   Call edi
  loc_005A254E: or eax, FFFFFFFFh
  loc_005A2551: sub esp, 00000010h
loc_005A2554:   mov edx, esp
  loc_005A2556: mov ecx, 0000000Bh
loc_005A255B:   mov [edx], ecx
loc_005A255D:   mov ecx, var_28
loc_005A2560:   mov [edx+00000004h], ecx
loc_005A2563:   mov ecx, [esi]
  loc_005A2565: push 8001000Dh
loc_005A256A:   push esi
loc_005A256B:   mov [edx+00000008h], eax
loc_005A256E:   mov eax, var_20
loc_005A2571:   mov [edx+0000000Ch], eax
loc_005A2574:   Call [ecx+00000458h]
loc_005A257A:   lea edx, var_1C
loc_005A257D:   push eax
loc_005A257E:   push edx
loc_005A257F:   Call ebx
loc_005A2581:   push eax
  loc_005A2582: call [00401280h] ; __vbaLateIdSt
loc_005A2588:   lea ecx, var_1C
loc_005A258B:   Call edi
  loc_005A258D: mov var_4, 00000000h
  loc_005A2594: push 005A25AFh
  loc_005A2599: jmp 005A25AEh
loc_005A259B:   lea ecx, var_18
  loc_005A259E: call [0040129Ch] ; __vbaFreeStr
loc_005A25A4:   lea ecx, var_1C
  loc_005A25A7: call [004012A0h] ; __vbaFreeObj
loc_005A25AD:   ret
loc_005A25AE:   ret
loc_005A25AF:   mov eax, Me
loc_005A25B2:   push eax
loc_005A25B3:   mov ecx, [eax]
loc_005A25B5:   Call [ecx+00000008h]
loc_005A25B8:   mov eax, var_4
loc_005A25BB:   mov ecx, var_14
loc_005A25BE:   pop edi
loc_005A25BF:   pop esi
loc_005A25C0:   mov fs: [00000000h] , ecx
loc_005A25C7:   pop ebx
loc_005A25C8:   mov esp, ebp
loc_005A25CA:   pop ebp
  loc_005A25CB: retn 0004h
End Sub

Private Sub mnuPenjualan_Click() '5A4C10
loc_005A4C10:   push ebp
loc_005A4C11:   mov ebp, esp
  loc_005A4C13: sub esp, 0000000Ch
  loc_005A4C16: push 00405696h ; __vbaExceptHandler
loc_005A4C1B:   mov eax, fs: [00000000h]
loc_005A4C21:   push eax
loc_005A4C22:   mov fs: [00000000h] , esp
  loc_005A4C29: sub esp, 00000030h
loc_005A4C2C:   push ebx
loc_005A4C2D:   push esi
loc_005A4C2E:   push edi
loc_005A4C2F:   mov var_C, esp
  loc_005A4C32: mov var_8, 00402C68h
loc_005A4C39:   mov eax, Me
loc_005A4C3C:   mov ecx, eax
  loc_005A4C3E: and ecx, 00000001h
loc_005A4C41:   mov var_4, ecx
  loc_005A4C44: and al, FEh
loc_005A4C46:   push eax
loc_005A4C47:   mov Me, eax
loc_005A4C4A:   mov edx, [eax]
loc_005A4C4C:   Call [edx+00000004h]
loc_005A4C4F:   mov eax, [00668278h]
loc_005A4C54:   test eax, eax
  loc_005A4C56: jnz 005A4C68h
  loc_005A4C58: push 00668278h
  loc_005A4C5D: push 00424728h
  loc_005A4C62: call [004011E8h] ; __vbaNew2
  loc_005A4C68: sub esp, 00000010h
  loc_005A4C6B: mov ecx, 0000000Ah
loc_005A4C70:   mov ebx, esp
  loc_005A4C72: mov eax, 80020004h
  loc_005A4C77: sub esp, 00000010h
loc_005A4C7A:   mov esi, [00668278h]
loc_005A4C80:   mov [ebx], ecx
loc_005A4C82:   mov ecx, var_30
loc_005A4C85:   mov edi, [esi]
  loc_005A4C87: mov edx, 00000001h
loc_005A4C8C:   mov [ebx+00000004h], ecx
loc_005A4C8F:   mov ecx, esp
loc_005A4C91:   push esi
loc_005A4C92:   mov [ebx+00000008h], eax
loc_005A4C95:   mov eax, var_28
loc_005A4C98:   mov [ebx+0000000Ch], eax
  loc_005A4C9B: mov eax, 00000002h
loc_005A4CA0:   mov [ecx], eax
loc_005A4CA2:   mov eax, var_20
loc_005A4CA5:   mov [ecx+00000004h], eax
loc_005A4CA8:   mov [ecx+00000008h], edx
loc_005A4CAB:   mov edx, var_18
loc_005A4CAE:   mov [ecx+0000000Ch], edx
loc_005A4CB1:   Call [edi+000002B0h]
loc_005A4CB7:   test eax, eax
loc_005A4CB9:   fnclex
  loc_005A4CBB: jge 005A4CCFh
  loc_005A4CBD: push 000002B0h
  loc_005A4CC2: push 00434B48h
loc_005A4CC7:   push esi
loc_005A4CC8:   push eax
  loc_005A4CC9: call [00401080h] ; __vbaHresultCheckObj
  loc_005A4CCF: mov var_4, 00000000h
loc_005A4CD6:   mov eax, Me
loc_005A4CD9:   push eax
loc_005A4CDA:   mov ecx, [eax]
loc_005A4CDC:   Call [ecx+00000008h]
loc_005A4CDF:   mov eax, var_4
loc_005A4CE2:   mov ecx, var_14
loc_005A4CE5:   pop edi
loc_005A4CE6:   pop esi
loc_005A4CE7:   mov fs: [00000000h] , ecx
loc_005A4CEE:   pop ebx
loc_005A4CEF:   mov esp, ebp
loc_005A4CF1:   pop ebp
  loc_005A4CF2: retn 0004h
End Sub

Private Sub cmdKeluar_UnknownEvent_10() '5A1E60
loc_005A1E60:   push ebp
loc_005A1E61:   mov ebp, esp
  loc_005A1E63: sub esp, 0000000Ch
  loc_005A1E66: push 00405696h ; __vbaExceptHandler
loc_005A1E6B:   mov eax, fs: [00000000h]
loc_005A1E71:   push eax
loc_005A1E72:   mov fs: [00000000h] , esp
  loc_005A1E79: sub esp, 00000018h
loc_005A1E7C:   push ebx
loc_005A1E7D:   push esi
loc_005A1E7E:   push edi
loc_005A1E7F:   mov var_C, esp
  loc_005A1E82: mov var_8, 00402AF8h
loc_005A1E89:   mov edi, Me
loc_005A1E8C:   mov eax, edi
  loc_005A1E8E: and eax, 00000001h
loc_005A1E91:   mov var_4, eax
  loc_005A1E94: and edi, FFFFFFFEh
loc_005A1E97:   push edi
loc_005A1E98:   mov Me, edi
loc_005A1E9B:   mov ecx, [edi]
loc_005A1E9D:   Call [ecx+00000004h]
loc_005A1EA0:   mov eax, [0066A318h]
  loc_005A1EA5: xor ebx, ebx
loc_005A1EA7:   cmp eax, ebx
loc_005A1EA9:   mov var_18, ebx
  loc_005A1EAC: jnz 005A1EBEh
  loc_005A1EAE: push 0066A318h
  loc_005A1EB3: push 0042E390h
  loc_005A1EB8: call [004011E8h] ; __vbaNew2
loc_005A1EBE:   mov esi, [0066A318h]
loc_005A1EC4:   lea eax, var_18
loc_005A1EC7:   push edi
loc_005A1EC8:   push eax
loc_005A1EC9:   mov edx, [esi]
loc_005A1ECB:   mov var_2C, edx
  loc_005A1ECE: call [004010C8h] ; __vbaObjSetAddref
loc_005A1ED4:   mov ecx, var_2C
loc_005A1ED7:   push eax
loc_005A1ED8:   push esi
loc_005A1ED9:   Call [ecx+00000010h]
loc_005A1EDC:   cmp eax, ebx
loc_005A1EDE:   fnclex
  loc_005A1EE0: jge 005A1EF1h
  loc_005A1EE2: push 00000010h
  loc_005A1EE4: push 0042E380h
loc_005A1EE9:   push esi
loc_005A1EEA:   push eax
  loc_005A1EEB: call [00401080h] ; __vbaHresultCheckObj
loc_005A1EF1:   lea ecx, var_18
  loc_005A1EF4: call [004012A0h] ; __vbaFreeObj
loc_005A1EFA:   mov var_4, ebx
  loc_005A1EFD: push 005A1F0Fh
  loc_005A1F02: jmp 005A1F0Eh
loc_005A1F04:   lea ecx, var_18
  loc_005A1F07: call [004012A0h] ; __vbaFreeObj
loc_005A1F0D:   ret
loc_005A1F0E:   ret
loc_005A1F0F:   mov eax, Me
loc_005A1F12:   push eax
loc_005A1F13:   mov edx, [eax]
loc_005A1F15:   Call [edx+00000008h]
loc_005A1F18:   mov eax, var_4
loc_005A1F1B:   mov ecx, var_14
loc_005A1F1E:   pop edi
loc_005A1F1F:   pop esi
loc_005A1F20:   mov fs: [00000000h] , ecx
loc_005A1F27:   pop ebx
loc_005A1F28:   mov esp, ebp
loc_005A1F2A:   pop ebp
  loc_005A1F2B: retn 0004h
End Sub

Private Sub cmdLogin_UnknownEvent_10() '5A1F30
loc_005A1F30:   push ebp
loc_005A1F31:   mov ebp, esp
  loc_005A1F33: sub esp, 0000000Ch
  loc_005A1F36: push 00405696h ; __vbaExceptHandler
loc_005A1F3B:   mov eax, fs: [00000000h]
loc_005A1F41:   push eax
loc_005A1F42:   mov fs: [00000000h] , esp
  loc_005A1F49: sub esp, 00000030h
loc_005A1F4C:   push ebx
loc_005A1F4D:   push esi
loc_005A1F4E:   push edi
loc_005A1F4F:   mov var_C, esp
  loc_005A1F52: mov var_8, 00402B08h
loc_005A1F59:   mov eax, Me
loc_005A1F5C:   mov ecx, eax
  loc_005A1F5E: and ecx, 00000001h
loc_005A1F61:   mov var_4, ecx
  loc_005A1F64: and al, FEh
loc_005A1F66:   push eax
loc_005A1F67:   mov Me, eax
loc_005A1F6A:   mov edx, [eax]
loc_005A1F6C:   Call [edx+00000004h]
loc_005A1F6F:   mov eax, [0066828Ch]
loc_005A1F74:   test eax, eax
  loc_005A1F76: jnz 005A1F88h
  loc_005A1F78: push 0066828Ch
  loc_005A1F7D: push 0040E9C0h
  loc_005A1F82: call [004011E8h] ; __vbaNew2
  loc_005A1F88: sub esp, 00000010h
  loc_005A1F8B: mov ecx, 0000000Ah
loc_005A1F90:   mov ebx, esp
  loc_005A1F92: mov eax, 80020004h
  loc_005A1F97: sub esp, 00000010h
loc_005A1F9A:   mov esi, [0066828Ch]
loc_005A1FA0:   mov [ebx], ecx
loc_005A1FA2:   mov ecx, var_30
loc_005A1FA5:   mov edi, [esi]
  loc_005A1FA7: mov edx, 00000001h
loc_005A1FAC:   mov [ebx+00000004h], ecx
loc_005A1FAF:   mov ecx, esp
loc_005A1FB1:   push esi
loc_005A1FB2:   mov [ebx+00000008h], eax
loc_005A1FB5:   mov eax, var_28
loc_005A1FB8:   mov [ebx+0000000Ch], eax
  loc_005A1FBB: mov eax, 00000002h
loc_005A1FC0:   mov [ecx], eax
loc_005A1FC2:   mov eax, var_20
loc_005A1FC5:   mov [ecx+00000004h], eax
loc_005A1FC8:   mov [ecx+00000008h], edx
loc_005A1FCB:   mov edx, var_18
loc_005A1FCE:   mov [ecx+0000000Ch], edx
loc_005A1FD1:   Call [edi+000002B0h]
loc_005A1FD7:   test eax, eax
loc_005A1FD9:   fnclex
  loc_005A1FDB: jge 005A1FEFh
  loc_005A1FDD: push 000002B0h
  loc_005A1FE2: push 00434CC4h
loc_005A1FE7:   push esi
loc_005A1FE8:   push eax
  loc_005A1FE9: call [00401080h] ; __vbaHresultCheckObj
  loc_005A1FEF: mov var_4, 00000000h
loc_005A1FF6:   mov eax, Me
loc_005A1FF9:   push eax
loc_005A1FFA:   mov ecx, [eax]
loc_005A1FFC:   Call [ecx+00000008h]
loc_005A1FFF:   mov eax, var_4
loc_005A2002:   mov ecx, var_14
loc_005A2005:   pop edi
loc_005A2006:   pop esi
loc_005A2007:   mov fs: [00000000h] , ecx
loc_005A200E:   pop ebx
loc_005A200F:   mov esp, ebp
loc_005A2011:   pop ebp
  loc_005A2012: retn 0004h
End Sub

Private Sub CmdPenjualan_UnknownEvent_10() '5A28A0
loc_005A28A0:   push ebp
loc_005A28A1:   mov ebp, esp
  loc_005A28A3: sub esp, 0000000Ch
  loc_005A28A6: push 00405696h ; __vbaExceptHandler
loc_005A28AB:   mov eax, fs: [00000000h]
loc_005A28B1:   push eax
loc_005A28B2:   mov fs: [00000000h] , esp
  loc_005A28B9: sub esp, 00000030h
loc_005A28BC:   push ebx
loc_005A28BD:   push esi
loc_005A28BE:   push edi
loc_005A28BF:   mov var_C, esp
  loc_005A28C2: mov var_8, 00402B38h
loc_005A28C9:   mov eax, Me
loc_005A28CC:   mov ecx, eax
  loc_005A28CE: and ecx, 00000001h
loc_005A28D1:   mov var_4, ecx
  loc_005A28D4: and al, FEh
loc_005A28D6:   push eax
loc_005A28D7:   mov Me, eax
loc_005A28DA:   mov edx, [eax]
loc_005A28DC:   Call [edx+00000004h]
loc_005A28DF:   mov eax, [00668278h]
loc_005A28E4:   test eax, eax
  loc_005A28E6: jnz 005A28F8h
  loc_005A28E8: push 00668278h
  loc_005A28ED: push 00424728h
  loc_005A28F2: call [004011E8h] ; __vbaNew2
  loc_005A28F8: sub esp, 00000010h
  loc_005A28FB: mov ecx, 0000000Ah
loc_005A2900:   mov ebx, esp
  loc_005A2902: mov eax, 80020004h
  loc_005A2907: sub esp, 00000010h
loc_005A290A:   mov esi, [00668278h]
loc_005A2910:   mov [ebx], ecx
loc_005A2912:   mov ecx, var_30
loc_005A2915:   mov edi, [esi]
  loc_005A2917: mov edx, 00000001h
loc_005A291C:   mov [ebx+00000004h], ecx
loc_005A291F:   mov ecx, esp
loc_005A2921:   push esi
loc_005A2922:   mov [ebx+00000008h], eax
loc_005A2925:   mov eax, var_28
loc_005A2928:   mov [ebx+0000000Ch], eax
  loc_005A292B: mov eax, 00000002h
loc_005A2930:   mov [ecx], eax
loc_005A2932:   mov eax, var_20
loc_005A2935:   mov [ecx+00000004h], eax
loc_005A2938:   mov [ecx+00000008h], edx
loc_005A293B:   mov edx, var_18
loc_005A293E:   mov [ecx+0000000Ch], edx
loc_005A2941:   Call [edi+000002B0h]
loc_005A2947:   test eax, eax
loc_005A2949:   fnclex
  loc_005A294B: jge 005A295Fh
  loc_005A294D: push 000002B0h
  loc_005A2952: push 00434B48h
loc_005A2957:   push esi
loc_005A2958:   push eax
  loc_005A2959: call [00401080h] ; __vbaHresultCheckObj
  loc_005A295F: mov var_4, 00000000h
loc_005A2966:   mov eax, Me
loc_005A2969:   push eax
loc_005A296A:   mov ecx, [eax]
loc_005A296C:   Call [ecx+00000008h]
loc_005A296F:   mov eax, var_4
loc_005A2972:   mov ecx, var_14
loc_005A2975:   pop edi
loc_005A2976:   pop esi
loc_005A2977:   mov fs: [00000000h] , ecx
loc_005A297E:   pop ebx
loc_005A297F:   mov esp, ebp
loc_005A2981:   pop ebp
  loc_005A2982: retn 0004h
End Sub

Private Sub CmdPelanggan_UnknownEvent_10() '5A25D0
loc_005A25D0:   push ebp
loc_005A25D1:   mov ebp, esp
  loc_005A25D3: sub esp, 0000000Ch
  loc_005A25D6: push 00405696h ; __vbaExceptHandler
loc_005A25DB:   mov eax, fs: [00000000h]
loc_005A25E1:   push eax
loc_005A25E2:   mov fs: [00000000h] , esp
  loc_005A25E9: sub esp, 00000030h
loc_005A25EC:   push ebx
loc_005A25ED:   push esi
loc_005A25EE:   push edi
loc_005A25EF:   mov var_C, esp
  loc_005A25F2: mov var_8, 00402B20h
loc_005A25F9:   mov eax, Me
loc_005A25FC:   mov ecx, eax
  loc_005A25FE: and ecx, 00000001h
loc_005A2601:   mov var_4, ecx
  loc_005A2604: and al, FEh
loc_005A2606:   push eax
loc_005A2607:   mov Me, eax
loc_005A260A:   mov edx, [eax]
loc_005A260C:   Call [edx+00000004h]
loc_005A260F:   mov eax, [00668164h]
loc_005A2614:   test eax, eax
  loc_005A2616: jnz 005A2628h
  loc_005A2618: push 00668164h
  loc_005A261D: push 00413F9Ch
  loc_005A2622: call [004011E8h] ; __vbaNew2
  loc_005A2628: sub esp, 00000010h
  loc_005A262B: mov ecx, 0000000Ah
loc_005A2630:   mov ebx, esp
  loc_005A2632: mov eax, 80020004h
  loc_005A2637: sub esp, 00000010h
loc_005A263A:   mov esi, [00668164h]
loc_005A2640:   mov [ebx], ecx
loc_005A2642:   mov ecx, var_30
loc_005A2645:   mov edi, [esi]
  loc_005A2647: mov edx, 00000001h
loc_005A264C:   mov [ebx+00000004h], ecx
loc_005A264F:   mov ecx, esp
loc_005A2651:   push esi
loc_005A2652:   mov [ebx+00000008h], eax
loc_005A2655:   mov eax, var_28
loc_005A2658:   mov [ebx+0000000Ch], eax
  loc_005A265B: mov eax, 00000002h
loc_005A2660:   mov [ecx], eax
loc_005A2662:   mov eax, var_20
loc_005A2665:   mov [ecx+00000004h], eax
loc_005A2668:   mov [ecx+00000008h], edx
loc_005A266B:   mov edx, var_18
loc_005A266E:   mov [ecx+0000000Ch], edx
loc_005A2671:   Call [edi+000002B0h]
loc_005A2677:   test eax, eax
loc_005A2679:   fnclex
  loc_005A267B: jge 005A268Fh
  loc_005A267D: push 000002B0h
  loc_005A2682: push 0042F03Ch
loc_005A2687:   push esi
loc_005A2688:   push eax
  loc_005A2689: call [00401080h] ; __vbaHresultCheckObj
  loc_005A268F: mov var_4, 00000000h
loc_005A2696:   mov eax, Me
loc_005A2699:   push eax
loc_005A269A:   mov ecx, [eax]
loc_005A269C:   Call [ecx+00000008h]
loc_005A269F:   mov eax, var_4
loc_005A26A2:   mov ecx, var_14
loc_005A26A5:   pop edi
loc_005A26A6:   pop esi
loc_005A26A7:   mov fs: [00000000h] , ecx
loc_005A26AE:   pop ebx
loc_005A26AF:   mov esp, ebp
loc_005A26B1:   pop ebp
  loc_005A26B2: retn 0004h
End Sub

Private Sub MnuRetur_Click() '5A56B0
loc_005A56B0:   push ebp
loc_005A56B1:   mov ebp, esp
  loc_005A56B3: sub esp, 0000000Ch
  loc_005A56B6: push 00405696h ; __vbaExceptHandler
loc_005A56BB:   mov eax, fs: [00000000h]
loc_005A56C1:   push eax
loc_005A56C2:   mov fs: [00000000h] , esp
  loc_005A56C9: sub esp, 00000030h
loc_005A56CC:   push ebx
loc_005A56CD:   push esi
loc_005A56CE:   push edi
loc_005A56CF:   mov var_C, esp
  loc_005A56D2: mov var_8, 00402CD8h
loc_005A56D9:   mov eax, Me
loc_005A56DC:   mov ecx, eax
  loc_005A56DE: and ecx, 00000001h
loc_005A56E1:   mov var_4, ecx
  loc_005A56E4: and al, FEh
loc_005A56E6:   push eax
loc_005A56E7:   mov Me, eax
loc_005A56EA:   mov edx, [eax]
loc_005A56EC:   Call [edx+00000004h]
loc_005A56EF:   mov eax, [00668418h]
loc_005A56F4:   test eax, eax
  loc_005A56F6: jnz 005A5708h
  loc_005A56F8: push 00668418h
  loc_005A56FD: push 00420198h
  loc_005A5702: call [004011E8h] ; __vbaNew2
  loc_005A5708: sub esp, 00000010h
  loc_005A570B: mov ecx, 0000000Ah
loc_005A5710:   mov ebx, esp
  loc_005A5712: mov eax, 80020004h
  loc_005A5717: sub esp, 00000010h
loc_005A571A:   mov esi, [00668418h]
loc_005A5720:   mov [ebx], ecx
loc_005A5722:   mov ecx, var_30
loc_005A5725:   mov edi, [esi]
  loc_005A5727: mov edx, 00000001h
loc_005A572C:   mov [ebx+00000004h], ecx
loc_005A572F:   mov ecx, esp
loc_005A5731:   push esi
loc_005A5732:   mov [ebx+00000008h], eax
loc_005A5735:   mov eax, var_28
loc_005A5738:   mov [ebx+0000000Ch], eax
  loc_005A573B: mov eax, 00000002h
loc_005A5740:   mov [ecx], eax
loc_005A5742:   mov eax, var_20
loc_005A5745:   mov [ecx+00000004h], eax
loc_005A5748:   mov [ecx+00000008h], edx
loc_005A574B:   mov edx, var_18
loc_005A574E:   mov [ecx+0000000Ch], edx
loc_005A5751:   Call [edi+000002B0h]
loc_005A5757:   test eax, eax
loc_005A5759:   fnclex
  loc_005A575B: jge 005A576Fh
  loc_005A575D: push 000002B0h
  loc_005A5762: push 004357F8h
loc_005A5767:   push esi
loc_005A5768:   push eax
  loc_005A5769: call [00401080h] ; __vbaHresultCheckObj
  loc_005A576F: mov var_4, 00000000h
loc_005A5776:   mov eax, Me
loc_005A5779:   push eax
loc_005A577A:   mov ecx, [eax]
loc_005A577C:   Call [ecx+00000008h]
loc_005A577F:   mov eax, var_4
loc_005A5782:   mov ecx, var_14
loc_005A5785:   pop edi
loc_005A5786:   pop esi
loc_005A5787:   mov fs: [00000000h] , ecx
loc_005A578E:   pop ebx
loc_005A578F:   mov esp, ebp
loc_005A5791:   pop ebp
  loc_005A5792: retn 0004h
End Sub

Private Sub mnuLapSelBeli_Click() '5A4FB0
loc_005A4FB0:   push ebp
loc_005A4FB1:   mov ebp, esp
  loc_005A4FB3: sub esp, 0000000Ch
  loc_005A4FB6: push 00405696h ; __vbaExceptHandler
loc_005A4FBB:   mov eax, fs: [00000000h]
loc_005A4FC1:   push eax
loc_005A4FC2:   mov fs: [00000000h] , esp
  loc_005A4FC9: sub esp, 00000008h
loc_005A4FCC:   push ebx
loc_005A4FCD:   push esi
loc_005A4FCE:   push edi
loc_005A4FCF:   mov var_C, esp
  loc_005A4FD2: mov var_8, 00402C98h
loc_005A4FD9:   mov eax, Me
loc_005A4FDC:   mov ecx, eax
  loc_005A4FDE: and ecx, 00000001h
loc_005A4FE1:   mov var_4, ecx
  loc_005A4FE4: and al, FEh
loc_005A4FE6:   push eax
loc_005A4FE7:   mov Me, eax
loc_005A4FEA:   mov edx, [eax]
loc_005A4FEC:   Call [edx+00000004h]
  loc_005A4FEF: call 0060EF90h
  loc_005A4FF4: mov var_4, 00000000h
loc_005A4FFB:   mov eax, Me
loc_005A4FFE:   push eax
loc_005A4FFF:   mov ecx, [eax]
loc_005A5001:   Call [ecx+00000008h]
loc_005A5004:   mov eax, var_4
loc_005A5007:   mov ecx, var_14
loc_005A500A:   pop edi
loc_005A500B:   pop esi
loc_005A500C:   mov fs: [00000000h] , ecx
loc_005A5013:   pop ebx
loc_005A5014:   mov esp, ebp
loc_005A5016:   pop ebp
  loc_005A5017: retn 0004h
End Sub

Private Sub mnuBarang_Click() '5A4A30
loc_005A4A30:   push ebp
loc_005A4A31:   mov ebp, esp
  loc_005A4A33: sub esp, 0000000Ch
  loc_005A4A36: push 00405696h ; __vbaExceptHandler
loc_005A4A3B:   mov eax, fs: [00000000h]
loc_005A4A41:   push eax
loc_005A4A42:   mov fs: [00000000h] , esp
  loc_005A4A49: sub esp, 00000030h
loc_005A4A4C:   push ebx
loc_005A4A4D:   push esi
loc_005A4A4E:   push edi
loc_005A4A4F:   mov var_C, esp
  loc_005A4A52: mov var_8, 00402C58h
loc_005A4A59:   mov eax, Me
loc_005A4A5C:   mov ecx, eax
  loc_005A4A5E: and ecx, 00000001h
loc_005A4A61:   mov var_4, ecx
  loc_005A4A64: and al, FEh
loc_005A4A66:   push eax
loc_005A4A67:   mov Me, eax
loc_005A4A6A:   mov edx, [eax]
loc_005A4A6C:   Call [edx+00000004h]
loc_005A4A6F:   mov eax, [006681DCh]
loc_005A4A74:   test eax, eax
  loc_005A4A76: jnz 005A4A88h
  loc_005A4A78: push 006681DCh
  loc_005A4A7D: push 0041EC78h
  loc_005A4A82: call [004011E8h] ; __vbaNew2
  loc_005A4A88: sub esp, 00000010h
  loc_005A4A8B: mov ecx, 0000000Ah
loc_005A4A90:   mov ebx, esp
  loc_005A4A92: mov eax, 80020004h
  loc_005A4A97: sub esp, 00000010h
loc_005A4A9A:   mov esi, [006681DCh]
loc_005A4AA0:   mov [ebx], ecx
loc_005A4AA2:   mov ecx, var_30
loc_005A4AA5:   mov edi, [esi]
  loc_005A4AA7: mov edx, 00000001h
loc_005A4AAC:   mov [ebx+00000004h], ecx
loc_005A4AAF:   mov ecx, esp
loc_005A4AB1:   push esi
loc_005A4AB2:   mov [ebx+00000008h], eax
loc_005A4AB5:   mov eax, var_28
loc_005A4AB8:   mov [ebx+0000000Ch], eax
  loc_005A4ABB: mov eax, 00000002h
loc_005A4AC0:   mov [ecx], eax
loc_005A4AC2:   mov eax, var_20
loc_005A4AC5:   mov [ecx+00000004h], eax
loc_005A4AC8:   mov [ecx+00000008h], edx
loc_005A4ACB:   mov edx, var_18
loc_005A4ACE:   mov [ecx+0000000Ch], edx
loc_005A4AD1:   Call [edi+000002B0h]
loc_005A4AD7:   test eax, eax
loc_005A4AD9:   fnclex
  loc_005A4ADB: jge 005A4AEFh
  loc_005A4ADD: push 000002B0h
  loc_005A4AE2: push 00431BBCh
loc_005A4AE7:   push esi
loc_005A4AE8:   push eax
  loc_005A4AE9: call [00401080h] ; __vbaHresultCheckObj
  loc_005A4AEF: mov var_4, 00000000h
loc_005A4AF6:   mov eax, Me
loc_005A4AF9:   push eax
loc_005A4AFA:   mov ecx, [eax]
loc_005A4AFC:   Call [ecx+00000008h]
loc_005A4AFF:   mov eax, var_4
loc_005A4B02:   mov ecx, var_14
loc_005A4B05:   pop edi
loc_005A4B06:   pop esi
loc_005A4B07:   mov fs: [00000000h] , ecx
loc_005A4B0E:   pop ebx
loc_005A4B0F:   mov esp, ebp
loc_005A4B11:   pop ebp
  loc_005A4B12: retn 0004h
End Sub

Private Sub MnuLapHutang_Click() '5A3370
loc_005A3370:   push ebp
loc_005A3371:   mov ebp, esp
  loc_005A3373: sub esp, 0000000Ch
  loc_005A3376: push 00405696h ; __vbaExceptHandler
loc_005A337B:   mov eax, fs: [00000000h]
loc_005A3381:   push eax
loc_005A3382:   mov fs: [00000000h] , esp
  loc_005A3389: sub esp, 00000030h
loc_005A338C:   push ebx
loc_005A338D:   push esi
loc_005A338E:   push edi
loc_005A338F:   mov var_C, esp
  loc_005A3392: mov var_8, 00402B98h
loc_005A3399:   mov eax, Me
loc_005A339C:   mov ecx, eax
  loc_005A339E: and ecx, 00000001h
loc_005A33A1:   mov var_4, ecx
  loc_005A33A4: and al, FEh
loc_005A33A6:   push eax
loc_005A33A7:   mov Me, eax
loc_005A33AA:   mov edx, [eax]
loc_005A33AC:   Call [edx+00000004h]
loc_005A33AF:   mov eax, [006682DCh]
loc_005A33B4:   test eax, eax
  loc_005A33B6: jnz 005A33C8h
  loc_005A33B8: push 006682DCh
  loc_005A33BD: push 0041B2B8h
  loc_005A33C2: call [004011E8h] ; __vbaNew2
  loc_005A33C8: sub esp, 00000010h
  loc_005A33CB: mov ecx, 0000000Ah
loc_005A33D0:   mov ebx, esp
  loc_005A33D2: mov eax, 80020004h
  loc_005A33D7: sub esp, 00000010h
loc_005A33DA:   mov esi, [006682DCh]
loc_005A33E0:   mov [ebx], ecx
loc_005A33E2:   mov ecx, var_30
loc_005A33E5:   mov edi, [esi]
  loc_005A33E7: mov edx, 00000001h
loc_005A33EC:   mov [ebx+00000004h], ecx
loc_005A33EF:   mov ecx, esp
loc_005A33F1:   push esi
loc_005A33F2:   mov [ebx+00000008h], eax
loc_005A33F5:   mov eax, var_28
loc_005A33F8:   mov [ebx+0000000Ch], eax
  loc_005A33FB: mov eax, 00000002h
loc_005A3400:   mov [ecx], eax
loc_005A3402:   mov eax, var_20
loc_005A3405:   mov [ecx+00000004h], eax
loc_005A3408:   mov [ecx+00000008h], edx
loc_005A340B:   mov edx, var_18
loc_005A340E:   mov [ecx+0000000Ch], edx
loc_005A3411:   Call [edi+000002B0h]
loc_005A3417:   test eax, eax
loc_005A3419:   fnclex
  loc_005A341B: jge 005A342Fh
  loc_005A341D: push 000002B0h
  loc_005A3422: push 004351F4h
loc_005A3427:   push esi
loc_005A3428:   push eax
  loc_005A3429: call [00401080h] ; __vbaHresultCheckObj
  loc_005A342F: mov var_4, 00000000h
loc_005A3436:   mov eax, Me
loc_005A3439:   push eax
loc_005A343A:   mov ecx, [eax]
loc_005A343C:   Call [ecx+00000008h]
loc_005A343F:   mov eax, var_4
loc_005A3442:   mov ecx, var_14
loc_005A3445:   pop edi
loc_005A3446:   pop esi
loc_005A3447:   mov fs: [00000000h] , ecx
loc_005A344E:   pop ebx
loc_005A344F:   mov esp, ebp
loc_005A3451:   pop ebp
  loc_005A3452: retn 0004h
End Sub

Private Sub mnuPembelian_Click() '5A4B20
loc_005A4B20:   push ebp
loc_005A4B21:   mov ebp, esp
  loc_005A4B23: sub esp, 0000000Ch
  loc_005A4B26: push 00405696h ; __vbaExceptHandler
loc_005A4B2B:   mov eax, fs: [00000000h]
loc_005A4B31:   push eax
loc_005A4B32:   mov fs: [00000000h] , esp
  loc_005A4B39: sub esp, 00000030h
loc_005A4B3C:   push ebx
loc_005A4B3D:   push esi
loc_005A4B3E:   push edi
loc_005A4B3F:   mov var_C, esp
  loc_005A4B42: mov var_8, 00402C60h
loc_005A4B49:   mov eax, Me
loc_005A4B4C:   mov ecx, eax
  loc_005A4B4E: and ecx, 00000001h
loc_005A4B51:   mov var_4, ecx
  loc_005A4B54: and al, FEh
loc_005A4B56:   push eax
loc_005A4B57:   mov Me, eax
loc_005A4B5A:   mov edx, [eax]
loc_005A4B5C:   Call [edx+00000004h]
loc_005A4B5F:   mov eax, [0066823Ch]
loc_005A4B64:   test eax, eax
  loc_005A4B66: jnz 005A4B78h
  loc_005A4B68: push 0066823Ch
  loc_005A4B6D: push 00427B78h
  loc_005A4B72: call [004011E8h] ; __vbaNew2
  loc_005A4B78: sub esp, 00000010h
  loc_005A4B7B: mov ecx, 0000000Ah
loc_005A4B80:   mov ebx, esp
  loc_005A4B82: mov eax, 80020004h
  loc_005A4B87: sub esp, 00000010h
loc_005A4B8A:   mov esi, [0066823Ch]
loc_005A4B90:   mov [ebx], ecx
loc_005A4B92:   mov ecx, var_30
loc_005A4B95:   mov edi, [esi]
  loc_005A4B97: mov edx, 00000001h
loc_005A4B9C:   mov [ebx+00000004h], ecx
loc_005A4B9F:   mov ecx, esp
loc_005A4BA1:   push esi
loc_005A4BA2:   mov [ebx+00000008h], eax
loc_005A4BA5:   mov eax, var_28
loc_005A4BA8:   mov [ebx+0000000Ch], eax
  loc_005A4BAB: mov eax, 00000002h
loc_005A4BB0:   mov [ecx], eax
loc_005A4BB2:   mov eax, var_20
loc_005A4BB5:   mov [ecx+00000004h], eax
loc_005A4BB8:   mov [ecx+00000008h], edx
loc_005A4BBB:   mov edx, var_18
loc_005A4BBE:   mov [ecx+0000000Ch], edx
loc_005A4BC1:   Call [edi+000002B0h]
loc_005A4BC7:   test eax, eax
loc_005A4BC9:   fnclex
  loc_005A4BCB: jge 005A4BDFh
  loc_005A4BCD: push 000002B0h
  loc_005A4BD2: push 00433DA8h
loc_005A4BD7:   push esi
loc_005A4BD8:   push eax
  loc_005A4BD9: call [00401080h] ; __vbaHresultCheckObj
  loc_005A4BDF: mov var_4, 00000000h
loc_005A4BE6:   mov eax, Me
loc_005A4BE9:   push eax
loc_005A4BEA:   mov ecx, [eax]
loc_005A4BEC:   Call [ecx+00000008h]
loc_005A4BEF:   mov eax, var_4
loc_005A4BF2:   mov ecx, var_14
loc_005A4BF5:   pop edi
loc_005A4BF6:   pop esi
loc_005A4BF7:   mov fs: [00000000h] , ecx
loc_005A4BFE:   pop ebx
loc_005A4BFF:   mov esp, ebp
loc_005A4C01:   pop ebp
  loc_005A4C02: retn 0004h
End Sub

Private Sub mnuJenis_Click() '5A4670
loc_005A4670:   push ebp
loc_005A4671:   mov ebp, esp
  loc_005A4673: sub esp, 0000000Ch
  loc_005A4676: push 00405696h ; __vbaExceptHandler
loc_005A467B:   mov eax, fs: [00000000h]
loc_005A4681:   push eax
loc_005A4682:   mov fs: [00000000h] , esp
  loc_005A4689: sub esp, 00000030h
loc_005A468C:   push ebx
loc_005A468D:   push esi
loc_005A468E:   push edi
loc_005A468F:   mov var_C, esp
  loc_005A4692: mov var_8, 00402C38h
loc_005A4699:   mov eax, Me
loc_005A469C:   mov ecx, eax
  loc_005A469E: and ecx, 00000001h
loc_005A46A1:   mov var_4, ecx
  loc_005A46A4: and al, FEh
loc_005A46A6:   push eax
loc_005A46A7:   mov Me, eax
loc_005A46AA:   mov edx, [eax]
loc_005A46AC:   Call [edx+00000004h]
loc_005A46AF:   mov eax, [006681A0h]
loc_005A46B4:   test eax, eax
  loc_005A46B6: jnz 005A46C8h
  loc_005A46B8: push 006681A0h
  loc_005A46BD: push 0041611Ch
  loc_005A46C2: call [004011E8h] ; __vbaNew2
  loc_005A46C8: sub esp, 00000010h
  loc_005A46CB: mov ecx, 0000000Ah
loc_005A46D0:   mov ebx, esp
  loc_005A46D2: mov eax, 80020004h
  loc_005A46D7: sub esp, 00000010h
loc_005A46DA:   mov esi, [006681A0h]
loc_005A46E0:   mov [ebx], ecx
loc_005A46E2:   mov ecx, var_30
loc_005A46E5:   mov edi, [esi]
  loc_005A46E7: mov edx, 00000001h
loc_005A46EC:   mov [ebx+00000004h], ecx
loc_005A46EF:   mov ecx, esp
loc_005A46F1:   push esi
loc_005A46F2:   mov [ebx+00000008h], eax
loc_005A46F5:   mov eax, var_28
loc_005A46F8:   mov [ebx+0000000Ch], eax
  loc_005A46FB: mov eax, 00000002h
loc_005A4700:   mov [ecx], eax
loc_005A4702:   mov eax, var_20
loc_005A4705:   mov [ecx+00000004h], eax
loc_005A4708:   mov [ecx+00000008h], edx
loc_005A470B:   mov edx, var_18
loc_005A470E:   mov [ecx+0000000Ch], edx
loc_005A4711:   Call [edi+000002B0h]
loc_005A4717:   test eax, eax
loc_005A4719:   fnclex
  loc_005A471B: jge 005A472Fh
  loc_005A471D: push 000002B0h
  loc_005A4722: push 004300F8h
loc_005A4727:   push esi
loc_005A4728:   push eax
  loc_005A4729: call [00401080h] ; __vbaHresultCheckObj
  loc_005A472F: mov var_4, 00000000h
loc_005A4736:   mov eax, Me
loc_005A4739:   push eax
loc_005A473A:   mov ecx, [eax]
loc_005A473C:   Call [ecx+00000008h]
loc_005A473F:   mov eax, var_4
loc_005A4742:   mov ecx, var_14
loc_005A4745:   pop edi
loc_005A4746:   pop esi
loc_005A4747:   mov fs: [00000000h] , ecx
loc_005A474E:   pop ebx
loc_005A474F:   mov esp, ebp
loc_005A4751:   pop ebp
  loc_005A4752: retn 0004h
End Sub

Private Sub mnuLogin_Click() '5A3C30
loc_005A3C30:   push ebp
loc_005A3C31:   mov ebp, esp
  loc_005A3C33: sub esp, 0000000Ch
  loc_005A3C36: push 00405696h ; __vbaExceptHandler
loc_005A3C3B:   mov eax, fs: [00000000h]
loc_005A3C41:   push eax
loc_005A3C42:   mov fs: [00000000h] , esp
  loc_005A3C49: sub esp, 00000030h
loc_005A3C4C:   push ebx
loc_005A3C4D:   push esi
loc_005A3C4E:   push edi
loc_005A3C4F:   mov var_C, esp
  loc_005A3C52: mov var_8, 00402BF8h
loc_005A3C59:   mov eax, Me
loc_005A3C5C:   mov ecx, eax
  loc_005A3C5E: and ecx, 00000001h
loc_005A3C61:   mov var_4, ecx
  loc_005A3C64: and al, FEh
loc_005A3C66:   push eax
loc_005A3C67:   mov Me, eax
loc_005A3C6A:   mov edx, [eax]
loc_005A3C6C:   Call [edx+00000004h]
loc_005A3C6F:   mov eax, [0066828Ch]
loc_005A3C74:   test eax, eax
  loc_005A3C76: jnz 005A3C88h
  loc_005A3C78: push 0066828Ch
  loc_005A3C7D: push 0040E9C0h
  loc_005A3C82: call [004011E8h] ; __vbaNew2
  loc_005A3C88: sub esp, 00000010h
  loc_005A3C8B: mov ecx, 0000000Ah
loc_005A3C90:   mov ebx, esp
  loc_005A3C92: mov eax, 80020004h
  loc_005A3C97: sub esp, 00000010h
loc_005A3C9A:   mov esi, [0066828Ch]
loc_005A3CA0:   mov [ebx], ecx
loc_005A3CA2:   mov ecx, var_30
loc_005A3CA5:   mov edi, [esi]
  loc_005A3CA7: mov edx, 00000001h
loc_005A3CAC:   mov [ebx+00000004h], ecx
loc_005A3CAF:   mov ecx, esp
loc_005A3CB1:   push esi
loc_005A3CB2:   mov [ebx+00000008h], eax
loc_005A3CB5:   mov eax, var_28
loc_005A3CB8:   mov [ebx+0000000Ch], eax
  loc_005A3CBB: mov eax, 00000002h
loc_005A3CC0:   mov [ecx], eax
loc_005A3CC2:   mov eax, var_20
loc_005A3CC5:   mov [ecx+00000004h], eax
loc_005A3CC8:   mov [ecx+00000008h], edx
loc_005A3CCB:   mov edx, var_18
loc_005A3CCE:   mov [ecx+0000000Ch], edx
loc_005A3CD1:   Call [edi+000002B0h]
loc_005A3CD7:   test eax, eax
loc_005A3CD9:   fnclex
  loc_005A3CDB: jge 005A3CEFh
  loc_005A3CDD: push 000002B0h
  loc_005A3CE2: push 00434CC4h
loc_005A3CE7:   push esi
loc_005A3CE8:   push eax
  loc_005A3CE9: call [00401080h] ; __vbaHresultCheckObj
  loc_005A3CEF: mov var_4, 00000000h
loc_005A3CF6:   mov eax, Me
loc_005A3CF9:   push eax
loc_005A3CFA:   mov ecx, [eax]
loc_005A3CFC:   Call [ecx+00000008h]
loc_005A3CFF:   mov eax, var_4
loc_005A3D02:   mov ecx, var_14
loc_005A3D05:   pop edi
loc_005A3D06:   pop esi
loc_005A3D07:   mov fs: [00000000h] , ecx
loc_005A3D0E:   pop ebx
loc_005A3D0F:   mov esp, ebp
loc_005A3D11:   pop ebp
  loc_005A3D12: retn 0004h
End Sub

Private Sub MnuLapReturBeliPeriode_Click() '5A3880
loc_005A3880:   push ebp
loc_005A3881:   mov ebp, esp
  loc_005A3883: sub esp, 0000000Ch
  loc_005A3886: push 00405696h ; __vbaExceptHandler
loc_005A388B:   mov eax, fs: [00000000h]
loc_005A3891:   push eax
loc_005A3892:   mov fs: [00000000h] , esp
  loc_005A3899: sub esp, 00000030h
loc_005A389C:   push ebx
loc_005A389D:   push esi
loc_005A389E:   push edi
loc_005A389F:   mov var_C, esp
  loc_005A38A2: mov var_8, 00402BD0h
loc_005A38A9:   mov eax, Me
loc_005A38AC:   mov ecx, eax
  loc_005A38AE: and ecx, 00000001h
loc_005A38B1:   mov var_4, ecx
  loc_005A38B4: and al, FEh
loc_005A38B6:   push eax
loc_005A38B7:   mov Me, eax
loc_005A38BA:   mov edx, [eax]
loc_005A38BC:   Call [edx+00000004h]
loc_005A38BF:   mov eax, [0066833Ch]
loc_005A38C4:   test eax, eax
  loc_005A38C6: jnz 005A38D8h
  loc_005A38C8: push 0066833Ch
  loc_005A38CD: push 0041098Ch
  loc_005A38D2: call [004011E8h] ; __vbaNew2
  loc_005A38D8: sub esp, 00000010h
  loc_005A38DB: mov ecx, 0000000Ah
loc_005A38E0:   mov ebx, esp
  loc_005A38E2: mov eax, 80020004h
  loc_005A38E7: sub esp, 00000010h
loc_005A38EA:   mov esi, [0066833Ch]
loc_005A38F0:   mov [ebx], ecx
loc_005A38F2:   mov ecx, var_30
loc_005A38F5:   mov edi, [esi]
  loc_005A38F7: mov edx, 00000001h
loc_005A38FC:   mov [ebx+00000004h], ecx
loc_005A38FF:   mov ecx, esp
loc_005A3901:   push esi
loc_005A3902:   mov [ebx+00000008h], eax
loc_005A3905:   mov eax, var_28
loc_005A3908:   mov [ebx+0000000Ch], eax
  loc_005A390B: mov eax, 00000002h
loc_005A3910:   mov [ecx], eax
loc_005A3912:   mov eax, var_20
loc_005A3915:   mov [ecx+00000004h], eax
loc_005A3918:   mov [ecx+00000008h], edx
loc_005A391B:   mov edx, var_18
loc_005A391E:   mov [ecx+0000000Ch], edx
loc_005A3921:   Call [edi+000002B0h]
loc_005A3927:   test eax, eax
loc_005A3929:   fnclex
  loc_005A392B: jge 005A393Fh
  loc_005A392D: push 000002B0h
  loc_005A3932: push 00435478h
loc_005A3937:   push esi
loc_005A3938:   push eax
  loc_005A3939: call [00401080h] ; __vbaHresultCheckObj
  loc_005A393F: mov var_4, 00000000h
loc_005A3946:   mov eax, Me
loc_005A3949:   push eax
loc_005A394A:   mov ecx, [eax]
loc_005A394C:   Call [ecx+00000008h]
loc_005A394F:   mov eax, var_4
loc_005A3952:   mov ecx, var_14
loc_005A3955:   pop edi
loc_005A3956:   pop esi
loc_005A3957:   mov fs: [00000000h] , ecx
loc_005A395E:   pop ebx
loc_005A395F:   mov esp, ebp
loc_005A3961:   pop ebp
  loc_005A3962: retn 0004h
End Sub

Private Sub mnuKeluar_Click()
loc_005A42D0:   push ebp
loc_005A42D1:   mov ebp, esp
  loc_005A42D3: sub esp, 0000000Ch
  loc_005A42D6: push 00405696h ; __vbaExceptHandler
loc_005A42DB:   mov eax, fs: [00000000h]
loc_005A42E1:   push eax
loc_005A42E2:   mov fs: [00000000h] , esp
  loc_005A42E9: sub esp, 00000018h
loc_005A42EC:   push ebx
loc_005A42ED:   push esi
loc_005A42EE:   push edi
loc_005A42EF:   mov var_C, esp
  loc_005A42F2: mov var_8, 00402C10h
loc_005A42F9:   mov edi, Me
loc_005A42FC:   mov eax, edi
  loc_005A42FE: and eax, 00000001h
loc_005A4301:   mov var_4, eax
  loc_005A4304: and edi, FFFFFFFEh
loc_005A4307:   push edi
loc_005A4308:   mov Me, edi
loc_005A430B:   mov ecx, [edi]
loc_005A430D:   Call [ecx+00000004h]
loc_005A4310:   mov eax, [0066A318h]
  loc_005A4315: xor ebx, ebx
loc_005A4317:   cmp eax, ebx
loc_005A4319:   mov var_18, ebx
  loc_005A431C: jnz 005A432Eh
  loc_005A431E: push 0066A318h
  loc_005A4323: push 0042E390h
  loc_005A4328: call [004011E8h] ; __vbaNew2
loc_005A432E:   mov esi, [0066A318h]
loc_005A4334:   lea eax, var_18
loc_005A4337:   push edi
loc_005A4338:   push eax
loc_005A4339:   mov edx, [esi]
loc_005A433B:   mov var_2C, edx
  loc_005A433E: call [004010C8h] ; __vbaObjSetAddref
loc_005A4344:   mov ecx, var_2C
loc_005A4347:   push eax
loc_005A4348:   push esi
loc_005A4349:   Call [ecx+00000010h]
loc_005A434C:   cmp eax, ebx
loc_005A434E:   fnclex
  loc_005A4350: jge 005A4361h
  loc_005A4352: push 00000010h
  loc_005A4354: push 0042E380h
loc_005A4359:   push esi
loc_005A435A:   push eax
  loc_005A435B: call [00401080h] ; __vbaHresultCheckObj
loc_005A4361:   lea ecx, var_18
  loc_005A4364: call [004012A0h] ; __vbaFreeObj
loc_005A436A:   mov var_4, ebx
  loc_005A436D: push 005A437Fh
  loc_005A4372: jmp 005A437Eh
loc_005A4374:   lea ecx, var_18
  loc_005A4377: call [004012A0h] ; __vbaFreeObj
loc_005A437D:   ret
loc_005A437E:   ret
loc_005A437F:   mov eax, Me
loc_005A4382:   push eax
loc_005A4383:   mov edx, [eax]
loc_005A4385:   Call [edx+00000008h]
loc_005A4388:   mov eax, var_4
loc_005A438B:   mov ecx, var_14
loc_005A438E:   pop edi
loc_005A438F:   pop esi
loc_005A4390:   mov fs: [00000000h] , ecx
loc_005A4397:   pop ebx
loc_005A4398:   mov esp, ebp
loc_005A439A:   pop ebp
  loc_005A439B: retn 0004h
  
End Sub

Private Sub mnuProduk_Click() '5A4940
loc_005A4940:   push ebp
loc_005A4941:   mov ebp, esp
  loc_005A4943: sub esp, 0000000Ch
  loc_005A4946: push 00405696h ; __vbaExceptHandler
loc_005A494B:   mov eax, fs: [00000000h]
loc_005A4951:   push eax
loc_005A4952:   mov fs: [00000000h] , esp
  loc_005A4959: sub esp, 00000030h
loc_005A495C:   push ebx
loc_005A495D:   push esi
loc_005A495E:   push edi
loc_005A495F:   mov var_C, esp
  loc_005A4962: mov var_8, 00402C50h
loc_005A4969:   mov eax, Me
loc_005A496C:   mov ecx, eax
  loc_005A496E: and ecx, 00000001h
loc_005A4971:   mov var_4, ecx
  loc_005A4974: and al, FEh
loc_005A4976:   push eax
loc_005A4977:   mov Me, eax
loc_005A497A:   mov edx, [eax]
loc_005A497C:   Call [edx+00000004h]
loc_005A497F:   mov eax, [006681C8h]
loc_005A4984:   test eax, eax
  loc_005A4986: jnz 005A4998h
  loc_005A4988: push 006681C8h
  loc_005A498D: push 0041790Ch
  loc_005A4992: call [004011E8h] ; __vbaNew2
  loc_005A4998: sub esp, 00000010h
  loc_005A499B: mov ecx, 0000000Ah
loc_005A49A0:   mov ebx, esp
  loc_005A49A2: mov eax, 80020004h
  loc_005A49A7: sub esp, 00000010h
loc_005A49AA:   mov esi, [006681C8h]
loc_005A49B0:   mov [ebx], ecx
loc_005A49B2:   mov ecx, var_30
loc_005A49B5:   mov edi, [esi]
  loc_005A49B7: mov edx, 00000001h
loc_005A49BC:   mov [ebx+00000004h], ecx
loc_005A49BF:   mov ecx, esp
loc_005A49C1:   push esi
loc_005A49C2:   mov [ebx+00000008h], eax
loc_005A49C5:   mov eax, var_28
loc_005A49C8:   mov [ebx+0000000Ch], eax
  loc_005A49CB: mov eax, 00000002h
loc_005A49D0:   mov [ecx], eax
loc_005A49D2:   mov eax, var_20
loc_005A49D5:   mov [ecx+00000004h], eax
loc_005A49D8:   mov [ecx+00000008h], edx
loc_005A49DB:   mov edx, var_18
loc_005A49DE:   mov [ecx+0000000Ch], edx
loc_005A49E1:   Call [edi+000002B0h]
loc_005A49E7:   test eax, eax
loc_005A49E9:   fnclex
  loc_005A49EB: jge 005A49FFh
  loc_005A49ED: push 000002B0h
  loc_005A49F2: push 00430E58h
loc_005A49F7:   push esi
loc_005A49F8:   push eax
  loc_005A49F9: call [00401080h] ; __vbaHresultCheckObj
  loc_005A49FF: mov var_4, 00000000h
loc_005A4A06:   mov eax, Me
loc_005A4A09:   push eax
loc_005A4A0A:   mov ecx, [eax]
loc_005A4A0C:   Call [ecx+00000008h]
loc_005A4A0F:   mov eax, var_4
loc_005A4A12:   mov ecx, var_14
loc_005A4A15:   pop edi
loc_005A4A16:   pop esi
loc_005A4A17:   mov fs: [00000000h] , ecx
loc_005A4A1E:   pop ebx
loc_005A4A1F:   mov esp, ebp
loc_005A4A21:   pop ebp
  loc_005A4A22: retn 0004h
End Sub

Private Sub MnuPeraturan_Click() '5A4760
loc_005A4760:   push ebp
loc_005A4761:   mov ebp, esp
  loc_005A4763: sub esp, 0000000Ch
  loc_005A4766: push 00405696h ; __vbaExceptHandler
loc_005A476B:   mov eax, fs: [00000000h]
loc_005A4771:   push eax
loc_005A4772:   mov fs: [00000000h] , esp
  loc_005A4779: sub esp, 00000030h
loc_005A477C:   push ebx
loc_005A477D:   push esi
loc_005A477E:   push edi
loc_005A477F:   mov var_C, esp
  loc_005A4782: mov var_8, 00402C40h
loc_005A4789:   mov eax, Me
loc_005A478C:   mov ecx, eax
  loc_005A478E: and ecx, 00000001h
loc_005A4791:   mov var_4, ecx
  loc_005A4794: and al, FEh
loc_005A4796:   push eax
loc_005A4797:   mov Me, eax
loc_005A479A:   mov edx, [eax]
loc_005A479C:   Call [edx+00000004h]
loc_005A479F:   mov eax, [00668378h]
loc_005A47A4:   test eax, eax
  loc_005A47A6: jnz 005A47B8h
  loc_005A47A8: push 00668378h
  loc_005A47AD: push 0040BDE8h
  loc_005A47B2: call [004011E8h] ; __vbaNew2
  loc_005A47B8: sub esp, 00000010h
  loc_005A47BB: mov ecx, 0000000Ah
loc_005A47C0:   mov ebx, esp
  loc_005A47C2: mov eax, 80020004h
  loc_005A47C7: sub esp, 00000010h
loc_005A47CA:   mov esi, [00668378h]
loc_005A47D0:   mov [ebx], ecx
loc_005A47D2:   mov ecx, var_30
loc_005A47D5:   mov edi, [esi]
  loc_005A47D7: mov edx, 00000001h
loc_005A47DC:   mov [ebx+00000004h], ecx
loc_005A47DF:   mov ecx, esp
loc_005A47E1:   push esi
loc_005A47E2:   mov [ebx+00000008h], eax
loc_005A47E5:   mov eax, var_28
loc_005A47E8:   mov [ebx+0000000Ch], eax
  loc_005A47EB: mov eax, 00000002h
loc_005A47F0:   mov [ecx], eax
loc_005A47F2:   mov eax, var_20
loc_005A47F5:   mov [ecx+00000004h], eax
loc_005A47F8:   mov [ecx+00000008h], edx
loc_005A47FB:   mov edx, var_18
loc_005A47FE:   mov [ecx+0000000Ch], edx
loc_005A4801:   Call [edi+000002B0h]
loc_005A4807:   test eax, eax
loc_005A4809:   fnclex
  loc_005A480B: jge 005A481Fh
  loc_005A480D: push 000002B0h
  loc_005A4812: push 0043553Ch
loc_005A4817:   push esi
loc_005A4818:   push eax
  loc_005A4819: call [00401080h] ; __vbaHresultCheckObj
  loc_005A481F: mov var_4, 00000000h
loc_005A4826:   mov eax, Me
loc_005A4829:   push eax
loc_005A482A:   mov ecx, [eax]
loc_005A482C:   Call [ecx+00000008h]
loc_005A482F:   mov eax, var_4
loc_005A4832:   mov ecx, var_14
loc_005A4835:   pop edi
loc_005A4836:   pop esi
loc_005A4837:   mov fs: [00000000h] , ecx
loc_005A483E:   pop ebx
loc_005A483F:   mov esp, ebp
loc_005A4841:   pop ebp
  loc_005A4842: retn 0004h
End Sub

Private Sub mnuLapJualNota_Click() '5A5200
loc_005A5200:   push ebp
loc_005A5201:   mov ebp, esp
  loc_005A5203: sub esp, 0000000Ch
  loc_005A5206: push 00405696h ; __vbaExceptHandler
loc_005A520B:   mov eax, fs: [00000000h]
loc_005A5211:   push eax
loc_005A5212:   mov fs: [00000000h] , esp
  loc_005A5219: sub esp, 00000030h
loc_005A521C:   push ebx
loc_005A521D:   push esi
loc_005A521E:   push edi
loc_005A521F:   mov var_C, esp
  loc_005A5222: mov var_8, 00402CB0h
loc_005A5229:   mov eax, Me
loc_005A522C:   mov ecx, eax
  loc_005A522E: and ecx, 00000001h
loc_005A5231:   mov var_4, ecx
  loc_005A5234: and al, FEh
loc_005A5236:   push eax
loc_005A5237:   mov Me, eax
loc_005A523A:   mov edx, [eax]
loc_005A523C:   Call [edx+00000004h]
loc_005A523F:   mov eax, [006683DCh]
loc_005A5244:   test eax, eax
  loc_005A5246: jnz 005A5258h
  loc_005A5248: push 006683DCh
  loc_005A524D: push 0040C390h
  loc_005A5252: call [004011E8h] ; __vbaNew2
  loc_005A5258: sub esp, 00000010h
  loc_005A525B: mov ecx, 0000000Ah
loc_005A5260:   mov ebx, esp
  loc_005A5262: mov eax, 80020004h
  loc_005A5267: sub esp, 00000010h
loc_005A526A:   mov esi, [006683DCh]
loc_005A5270:   mov [ebx], ecx
loc_005A5272:   mov ecx, var_30
loc_005A5275:   mov edi, [esi]
  loc_005A5277: mov edx, 00000001h
loc_005A527C:   mov [ebx+00000004h], ecx
loc_005A527F:   mov ecx, esp
loc_005A5281:   push esi
loc_005A5282:   mov [ebx+00000008h], eax
loc_005A5285:   mov eax, var_28
loc_005A5288:   mov [ebx+0000000Ch], eax
  loc_005A528B: mov eax, 00000002h
loc_005A5290:   mov [ecx], eax
loc_005A5292:   mov eax, var_20
loc_005A5295:   mov [ecx+00000004h], eax
loc_005A5298:   mov [ecx+00000008h], edx
loc_005A529B:   mov edx, var_18
loc_005A529E:   mov [ecx+0000000Ch], edx
loc_005A52A1:   Call [edi+000002B0h]
loc_005A52A7:   test eax, eax
loc_005A52A9:   fnclex
  loc_005A52AB: jge 005A52BFh
  loc_005A52AD: push 000002B0h
  loc_005A52B2: push 0043570Ch
loc_005A52B7:   push esi
loc_005A52B8:   push eax
  loc_005A52B9: call [00401080h] ; __vbaHresultCheckObj
  loc_005A52BF: mov var_4, 00000000h
loc_005A52C6:   mov eax, Me
loc_005A52C9:   push eax
loc_005A52CA:   mov ecx, [eax]
loc_005A52CC:   Call [ecx+00000008h]
loc_005A52CF:   mov eax, var_4
loc_005A52D2:   mov ecx, var_14
loc_005A52D5:   pop edi
loc_005A52D6:   pop esi
loc_005A52D7:   mov fs: [00000000h] , ecx
loc_005A52DE:   pop ebx
loc_005A52DF:   mov esp, ebp
loc_005A52E1:   pop ebp
  loc_005A52E2: retn 0004h
End Sub

Private Sub mnuLogout_Click() '5A3D20
loc_005A3D20:   push ebp
loc_005A3D21:   mov ebp, esp
  loc_005A3D23: sub esp, 0000000Ch
  loc_005A3D26: push 00405696h ; __vbaExceptHandler
loc_005A3D2B:   mov eax, fs: [00000000h]
loc_005A3D31:   push eax
loc_005A3D32:   mov fs: [00000000h] , esp
  loc_005A3D39: sub esp, 0000003Ch
loc_005A3D3C:   push ebx
loc_005A3D3D:   push esi
loc_005A3D3E:   push edi
loc_005A3D3F:   mov var_C, esp
  loc_005A3D42: mov var_8, 00402C00h
loc_005A3D49:   mov esi, Me
loc_005A3D4C:   mov eax, esi
  loc_005A3D4E: and eax, 00000001h
loc_005A3D51:   mov var_4, eax
  loc_005A3D54: and esi, FFFFFFFEh
loc_005A3D57:   push esi
loc_005A3D58:   mov Me, esi
loc_005A3D5B:   mov ecx, [esi]
loc_005A3D5D:   Call [ecx+00000004h]
loc_005A3D60:   mov eax, [006682A0h]
  loc_005A3D65: xor ecx, ecx
loc_005A3D67:   cmp eax, ecx
loc_005A3D69:   mov var_18, ecx
loc_005A3D6C:   mov var_1C, ecx
  loc_005A3D6F: jnz 005A3D86h
  loc_005A3D71: push 006682A0h
  loc_005A3D76: push 00414A94h
  loc_005A3D7B: call [004011E8h] ; __vbaNew2
loc_005A3D81:   mov eax, [006682A0h]
loc_005A3D86:   mov edx, [eax]
loc_005A3D88:   push eax
loc_005A3D89:   Call [edx+00000308h]
  loc_005A3D8F: mov ebx, [004010B8h] ; __vbaObjSet
loc_005A3D95:   push eax
loc_005A3D96:   lea eax, var_1C
loc_005A3D99:   push eax
loc_005A3D9A:   Call ebx
loc_005A3D9C:   mov edi, eax
loc_005A3D9E:   lea edx, var_18
loc_005A3DA1:   push edx
loc_005A3DA2:   push edi
loc_005A3DA3:   mov ecx, [edi]
loc_005A3DA5:   Call [ecx+000000A0h]
loc_005A3DAB:   test eax, eax
loc_005A3DAD:   fnclex
  loc_005A3DAF: jge 005A3DC3h
  loc_005A3DB1: push 000000A0h
  loc_005A3DB6: push 0042DFCCh
loc_005A3DBB:   push edi
loc_005A3DBC:   push eax
  loc_005A3DBD: call [00401080h] ; __vbaHresultCheckObj
loc_005A3DC3:   mov eax, var_18
loc_005A3DC6:   push eax
  loc_005A3DC7: push 00434FA8h ; "UNREGISTERED"
  loc_005A3DCC: call [0040112Ch] ; __vbaStrCmp
loc_005A3DD2:   neg eax
loc_005A3DD4:   sbb eax, eax
loc_005A3DD6:   lea ecx, var_18
loc_005A3DD9:   inc eax
loc_005A3DDA:   neg eax
loc_005A3DDC:   mov var_48, eax
  loc_005A3DDF: call [0040129Ch] ; __vbaFreeStr
  loc_005A3DE5: mov edi, [004012A0h] ; __vbaFreeObj
loc_005A3DEB:   lea ecx, var_1C
loc_005A3DEE:   Call edi
  loc_005A3DF0: cmp var_48, 0000h
  loc_005A3DF5: jz 005A3EA3h
loc_005A3DFB:   mov ecx, [esi]
loc_005A3DFD:   push esi
loc_005A3DFE:   Call [ecx+00000358h]
loc_005A3E04:   lea edx, var_1C
loc_005A3E07:   push eax
loc_005A3E08:   push edx
loc_005A3E09:   Call ebx
loc_005A3E0B:   mov ecx, [eax]
  loc_005A3E0D: push 00000000h
loc_005A3E0F:   push eax
loc_005A3E10:   mov var_40, eax
loc_005A3E13:   Call [ecx+00000074h]
loc_005A3E16:   test eax, eax
loc_005A3E18:   fnclex
  loc_005A3E1A: jge 005A3E2Eh
loc_005A3E1C:   mov edx, var_40
  loc_005A3E1F: push 00000074h
  loc_005A3E21: push 00434274h
loc_005A3E26:   push edx
loc_005A3E27:   push eax
  loc_005A3E28: call [00401080h] ; __vbaHresultCheckObj
loc_005A3E2E:   lea ecx, var_1C
loc_005A3E31:   Call edi
loc_005A3E33:   mov eax, [esi]
loc_005A3E35:   push esi
loc_005A3E36:   Call [eax+0000043Ch]
loc_005A3E3C:   lea ecx, var_1C
loc_005A3E3F:   push eax
loc_005A3E40:   push ecx
loc_005A3E41:   Call ebx
loc_005A3E43:   mov edx, [eax]
  loc_005A3E45: push 00000000h
loc_005A3E47:   push eax
loc_005A3E48:   mov var_40, eax
loc_005A3E4B:   Call [edx+00000074h]
loc_005A3E4E:   test eax, eax
loc_005A3E50:   fnclex
  loc_005A3E52: jge 005A3E66h
loc_005A3E54:   mov ecx, var_40
  loc_005A3E57: push 00000074h
  loc_005A3E59: push 00434274h
loc_005A3E5E:   push ecx
loc_005A3E5F:   push eax
  loc_005A3E60: call [00401080h] ; __vbaHresultCheckObj
loc_005A3E66:   lea ecx, var_1C
loc_005A3E69:   Call edi
loc_005A3E6B:   mov edx, [esi]
loc_005A3E6D:   push esi
loc_005A3E6E:   Call [edx+000003A4h]
loc_005A3E74:   push eax
loc_005A3E75:   lea eax, var_1C
loc_005A3E78:   push eax
loc_005A3E79:   Call ebx
loc_005A3E7B:   mov ecx, [eax]
  loc_005A3E7D: push 00000000h
loc_005A3E7F:   push eax
loc_005A3E80:   mov var_40, eax
loc_005A3E83:   Call [ecx+00000074h]
loc_005A3E86:   test eax, eax
loc_005A3E88:   fnclex
  loc_005A3E8A: jge 005A3E9Eh
loc_005A3E8C:   mov edx, var_40
  loc_005A3E8F: push 00000074h
  loc_005A3E91: push 00434274h
loc_005A3E96:   push edx
loc_005A3E97:   push eax
  loc_005A3E98: call [00401080h] ; __vbaHresultCheckObj
loc_005A3E9E:   lea ecx, var_1C
loc_005A3EA1:   Call edi
loc_005A3EA3:   mov eax, [esi]
loc_005A3EA5:   push esi
loc_005A3EA6:   Call [eax+000003A4h]
loc_005A3EAC:   lea ecx, var_1C
loc_005A3EAF:   push eax
loc_005A3EB0:   push ecx
loc_005A3EB1:   Call ebx
loc_005A3EB3:   mov edx, [eax]
  loc_005A3EB5: push 00000000h
loc_005A3EB7:   push eax
loc_005A3EB8:   mov var_40, eax
loc_005A3EBB:   Call [edx+00000074h]
loc_005A3EBE:   test eax, eax
loc_005A3EC0:   fnclex
  loc_005A3EC2: jge 005A3ED6h
loc_005A3EC4:   mov ecx, var_40
  loc_005A3EC7: push 00000074h
  loc_005A3EC9: push 00434274h
loc_005A3ECE:   push ecx
loc_005A3ECF:   push eax
  loc_005A3ED0: call [00401080h] ; __vbaHresultCheckObj
loc_005A3ED6:   lea ecx, var_1C
loc_005A3ED9:   Call edi
loc_005A3EDB:   mov edx, [esi]
loc_005A3EDD:   push esi
loc_005A3EDE:   Call [edx+00000358h]
loc_005A3EE4:   push eax
loc_005A3EE5:   lea eax, var_1C
loc_005A3EE8:   push eax
loc_005A3EE9:   Call ebx
loc_005A3EEB:   mov ecx, [eax]
  loc_005A3EED: push 00000000h
loc_005A3EEF:   push eax
loc_005A3EF0:   mov var_40, eax
loc_005A3EF3:   Call [ecx+00000074h]
loc_005A3EF6:   test eax, eax
loc_005A3EF8:   fnclex
  loc_005A3EFA: jge 005A3F0Eh
loc_005A3EFC:   mov edx, var_40
  loc_005A3EFF: push 00000074h
  loc_005A3F01: push 00434274h
loc_005A3F06:   push edx
loc_005A3F07:   push eax
  loc_005A3F08: call [00401080h] ; __vbaHresultCheckObj
loc_005A3F0E:   lea ecx, var_1C
loc_005A3F11:   Call edi
loc_005A3F13:   mov eax, [esi]
loc_005A3F15:   push esi
loc_005A3F16:   Call [eax+00000384h]
loc_005A3F1C:   lea ecx, var_1C
loc_005A3F1F:   push eax
loc_005A3F20:   push ecx
loc_005A3F21:   Call ebx
loc_005A3F23:   mov edx, [eax]
  loc_005A3F25: push 00000000h
loc_005A3F27:   push eax
loc_005A3F28:   mov var_40, eax
loc_005A3F2B:   Call [edx+00000074h]
loc_005A3F2E:   test eax, eax
loc_005A3F30:   fnclex
  loc_005A3F32: jge 005A3F46h
loc_005A3F34:   mov ecx, var_40
  loc_005A3F37: push 00000074h
  loc_005A3F39: push 00434274h
loc_005A3F3E:   push ecx
loc_005A3F3F:   push eax
  loc_005A3F40: call [00401080h] ; __vbaHresultCheckObj
loc_005A3F46:   lea ecx, var_1C
loc_005A3F49:   Call edi
loc_005A3F4B:   mov edx, [esi]
loc_005A3F4D:   push esi
loc_005A3F4E:   Call [edx+000003A4h]
loc_005A3F54:   push eax
loc_005A3F55:   lea eax, var_1C
loc_005A3F58:   push eax
loc_005A3F59:   Call ebx
loc_005A3F5B:   mov ecx, [eax]
  loc_005A3F5D: push 00000000h
loc_005A3F5F:   push eax
loc_005A3F60:   mov var_40, eax
loc_005A3F63:   Call [ecx+00000074h]
loc_005A3F66:   test eax, eax
loc_005A3F68:   fnclex
  loc_005A3F6A: jge 005A3F7Eh
loc_005A3F6C:   mov edx, var_40
  loc_005A3F6F: push 00000074h
  loc_005A3F71: push 00434274h
loc_005A3F76:   push edx
loc_005A3F77:   push eax
  loc_005A3F78: call [00401080h] ; __vbaHresultCheckObj
loc_005A3F7E:   lea ecx, var_1C
loc_005A3F81:   Call edi
loc_005A3F83:   mov eax, [esi]
loc_005A3F85:   push esi
loc_005A3F86:   Call [eax+00000344h]
loc_005A3F8C:   lea ecx, var_1C
loc_005A3F8F:   push eax
loc_005A3F90:   push ecx
loc_005A3F91:   Call ebx
loc_005A3F93:   mov edx, [eax]
loc_005A3F95:   push FFFFFFFFh
loc_005A3F97:   push eax
loc_005A3F98:   mov var_40, eax
loc_005A3F9B:   Call [edx+00000074h]
loc_005A3F9E:   test eax, eax
loc_005A3FA0:   fnclex
  loc_005A3FA2: jge 005A3FB6h
loc_005A3FA4:   mov ecx, var_40
  loc_005A3FA7: push 00000074h
  loc_005A3FA9: push 00434274h
loc_005A3FAE:   push ecx
loc_005A3FAF:   push eax
  loc_005A3FB0: call [00401080h] ; __vbaHresultCheckObj
loc_005A3FB6:   lea ecx, var_1C
loc_005A3FB9:   Call edi
loc_005A3FBB:   mov edx, [esi]
loc_005A3FBD:   push esi
loc_005A3FBE:   Call [edx+00000348h]
loc_005A3FC4:   push eax
loc_005A3FC5:   lea eax, var_1C
loc_005A3FC8:   push eax
loc_005A3FC9:   Call ebx
loc_005A3FCB:   mov ecx, [eax]
  loc_005A3FCD: push 00000000h
loc_005A3FCF:   push eax
loc_005A3FD0:   mov var_40, eax
loc_005A3FD3:   Call [ecx+00000074h]
loc_005A3FD6:   test eax, eax
loc_005A3FD8:   fnclex
  loc_005A3FDA: jge 005A3FEEh
loc_005A3FDC:   mov edx, var_40
  loc_005A3FDF: push 00000074h
  loc_005A3FE1: push 00434274h
loc_005A3FE6:   push edx
loc_005A3FE7:   push eax
  loc_005A3FE8: call [00401080h] ; __vbaHresultCheckObj
loc_005A3FEE:   lea ecx, var_1C
loc_005A3FF1:   Call edi
loc_005A3FF3:   mov eax, [esi]
loc_005A3FF5:   push esi
loc_005A3FF6:   Call [eax+0000034Ch]
loc_005A3FFC:   lea ecx, var_1C
loc_005A3FFF:   push eax
loc_005A4000:   push ecx
loc_005A4001:   Call ebx
loc_005A4003:   mov edx, [eax]
  loc_005A4005: push 00000000h
loc_005A4007:   push eax
loc_005A4008:   mov var_40, eax
loc_005A400B:   Call [edx+00000074h]
loc_005A400E:   test eax, eax
loc_005A4010:   fnclex
  loc_005A4012: jge 005A4026h
loc_005A4014:   mov ecx, var_40
  loc_005A4017: push 00000074h
  loc_005A4019: push 00434274h
loc_005A401E:   push ecx
loc_005A401F:   push eax
  loc_005A4020: call [00401080h] ; __vbaHresultCheckObj
loc_005A4026:   lea ecx, var_1C
loc_005A4029:   Call edi
loc_005A402B:   mov edx, [esi]
loc_005A402D:   push esi
loc_005A402E:   Call [edx+00000424h]
loc_005A4034:   push eax
loc_005A4035:   lea eax, var_1C
loc_005A4038:   push eax
loc_005A4039:   Call ebx
loc_005A403B:   mov ecx, [eax]
  loc_005A403D: push 00000000h
loc_005A403F:   push eax
loc_005A4040:   mov var_40, eax
loc_005A4043:   Call [ecx+00000074h]
loc_005A4046:   test eax, eax
loc_005A4048:   fnclex
  loc_005A404A: jge 005A405Eh
loc_005A404C:   mov edx, var_40
  loc_005A404F: push 00000074h
  loc_005A4051: push 00434274h
loc_005A4056:   push edx
loc_005A4057:   push eax
  loc_005A4058: call [00401080h] ; __vbaHresultCheckObj
loc_005A405E:   lea ecx, var_1C
loc_005A4061:   Call edi
loc_005A4063:   mov eax, [esi]
loc_005A4065:   push esi
loc_005A4066:   Call [eax+00000440h]
loc_005A406C:   lea ecx, var_1C
loc_005A406F:   push eax
loc_005A4070:   push ecx
loc_005A4071:   Call ebx
loc_005A4073:   mov edx, [eax]
  loc_005A4075: push 00000000h
loc_005A4077:   push eax
loc_005A4078:   mov var_40, eax
loc_005A407B:   Call [edx+00000074h]
loc_005A407E:   test eax, eax
loc_005A4080:   fnclex
  loc_005A4082: jge 005A4096h
loc_005A4084:   mov ecx, var_40
  loc_005A4087: push 00000074h
  loc_005A4089: push 00434274h
loc_005A408E:   push ecx
loc_005A408F:   push eax
  loc_005A4090: call [00401080h] ; __vbaHresultCheckObj
loc_005A4096:   lea ecx, var_1C
loc_005A4099:   Call edi
  loc_005A409B: sub esp, 00000010h
  loc_005A409E: mov ecx, 0000000Bh
loc_005A40A3:   mov edx, esp
  loc_005A40A5: xor eax, eax
  loc_005A40A7: push 8001000Dh
loc_005A40AC:   push esi
loc_005A40AD:   mov [edx], ecx
loc_005A40AF:   mov ecx, var_28
loc_005A40B2:   mov [edx+00000004h], ecx
loc_005A40B5:   mov ecx, [esi]
loc_005A40B7:   mov [edx+00000008h], eax
loc_005A40BA:   mov eax, var_20
loc_005A40BD:   mov [edx+0000000Ch], eax
loc_005A40C0:   Call [ecx+00000460h]
loc_005A40C6:   lea edx, var_1C
loc_005A40C9:   push eax
loc_005A40CA:   push edx
loc_005A40CB:   Call ebx
loc_005A40CD:   push eax
  loc_005A40CE: call [00401280h] ; __vbaLateIdSt
loc_005A40D4:   lea ecx, var_1C
loc_005A40D7:   Call edi
  loc_005A40D9: sub esp, 00000010h
  loc_005A40DC: mov ecx, 0000000Bh
loc_005A40E1:   mov edx, esp
  loc_005A40E3: xor eax, eax
  loc_005A40E5: push 8001000Dh
loc_005A40EA:   push esi
loc_005A40EB:   mov [edx], ecx
loc_005A40ED:   mov ecx, var_28
loc_005A40F0:   mov [edx+00000004h], ecx
loc_005A40F3:   mov ecx, [esi]
loc_005A40F5:   mov [edx+00000008h], eax
loc_005A40F8:   mov eax, var_20
loc_005A40FB:   mov [edx+0000000Ch], eax
loc_005A40FE:   Call [ecx+00000464h]
loc_005A4104:   lea edx, var_1C
loc_005A4107:   push eax
loc_005A4108:   push edx
loc_005A4109:   Call ebx
loc_005A410B:   push eax
  loc_005A410C: call [00401280h] ; __vbaLateIdSt
loc_005A4112:   lea ecx, var_1C
loc_005A4115:   Call edi
  loc_005A4117: sub esp, 00000010h
  loc_005A411A: mov ecx, 0000000Bh
loc_005A411F:   mov edx, esp
  loc_005A4121: xor eax, eax
  loc_005A4123: push 8001000Dh
loc_005A4128:   push esi
loc_005A4129:   mov [edx], ecx
loc_005A412B:   mov ecx, var_28
loc_005A412E:   mov [edx+00000004h], ecx
loc_005A4131:   mov ecx, [esi]
loc_005A4133:   mov [edx+00000008h], eax
loc_005A4136:   mov eax, var_20
loc_005A4139:   mov [edx+0000000Ch], eax
loc_005A413C:   Call [ecx+00000468h]
loc_005A4142:   lea edx, var_1C
loc_005A4145:   push eax
loc_005A4146:   push edx
loc_005A4147:   Call ebx
loc_005A4149:   push eax
  loc_005A414A: call [00401280h] ; __vbaLateIdSt
loc_005A4150:   lea ecx, var_1C
loc_005A4153:   Call edi
  loc_005A4155: xor eax, eax
  loc_005A4157: sub esp, 00000010h
loc_005A415A:   mov edx, esp
  loc_005A415C: mov ecx, 0000000Bh
  loc_005A4161: push 8001000Dh
loc_005A4166:   push esi
loc_005A4167:   mov [edx], ecx
loc_005A4169:   mov ecx, var_28
loc_005A416C:   mov [edx+00000004h], ecx
loc_005A416F:   mov ecx, [esi]
loc_005A4171:   mov [edx+00000008h], eax
loc_005A4174:   mov eax, var_20
loc_005A4177:   mov [edx+0000000Ch], eax
loc_005A417A:   Call [ecx+0000044Ch]
loc_005A4180:   lea edx, var_1C
loc_005A4183:   push eax
loc_005A4184:   push edx
loc_005A4185:   Call ebx
loc_005A4187:   push eax
  loc_005A4188: call [00401280h] ; __vbaLateIdSt
loc_005A418E:   lea ecx, var_1C
loc_005A4191:   Call edi
  loc_005A4193: sub esp, 00000010h
  loc_005A4196: mov ecx, 0000000Bh
loc_005A419B:   mov edx, esp
  loc_005A419D: xor eax, eax
  loc_005A419F: push 8001000Dh
loc_005A41A4:   push esi
loc_005A41A5:   mov [edx], ecx
loc_005A41A7:   mov ecx, var_28
loc_005A41AA:   mov [edx+00000004h], ecx
loc_005A41AD:   mov ecx, [esi]
loc_005A41AF:   mov [edx+00000008h], eax
loc_005A41B2:   mov eax, var_20
loc_005A41B5:   mov [edx+0000000Ch], eax
loc_005A41B8:   Call [ecx+00000450h]
loc_005A41BE:   lea edx, var_1C
loc_005A41C1:   push eax
loc_005A41C2:   push edx
loc_005A41C3:   Call ebx
loc_005A41C5:   push eax
  loc_005A41C6: call [00401280h] ; __vbaLateIdSt
loc_005A41CC:   lea ecx, var_1C
loc_005A41CF:   Call edi
  loc_005A41D1: sub esp, 00000010h
  loc_005A41D4: mov ecx, 0000000Bh
loc_005A41D9:   mov edx, esp
  loc_005A41DB: xor eax, eax
  loc_005A41DD: push 8001000Dh
loc_005A41E2:   push esi
loc_005A41E3:   mov [edx], ecx
loc_005A41E5:   mov ecx, var_28
loc_005A41E8:   mov [edx+00000004h], ecx
loc_005A41EB:   mov ecx, [esi]
loc_005A41ED:   mov [edx+00000008h], eax
loc_005A41F0:   mov eax, var_20
loc_005A41F3:   mov [edx+0000000Ch], eax
loc_005A41F6:   Call [ecx+00000454h]
loc_005A41FC:   lea edx, var_1C
loc_005A41FF:   push eax
loc_005A4200:   push edx
loc_005A4201:   Call ebx
loc_005A4203:   push eax
  loc_005A4204: call [00401280h] ; __vbaLateIdSt
loc_005A420A:   lea ecx, var_1C
loc_005A420D:   Call edi
  loc_005A420F: sub esp, 00000010h
  loc_005A4212: mov ecx, 0000000Bh
loc_005A4217:   mov edx, esp
  loc_005A4219: or eax, FFFFFFFFh
  loc_005A421C: push 8001000Dh
loc_005A4221:   push esi
loc_005A4222:   mov [edx], ecx
loc_005A4224:   mov ecx, var_28
loc_005A4227:   mov [edx+00000004h], ecx
loc_005A422A:   mov ecx, [esi]
loc_005A422C:   mov [edx+00000008h], eax
loc_005A422F:   mov eax, var_20
loc_005A4232:   mov [edx+0000000Ch], eax
loc_005A4235:   Call [ecx+0000045Ch]
loc_005A423B:   lea edx, var_1C
loc_005A423E:   push eax
loc_005A423F:   push edx
loc_005A4240:   Call ebx
loc_005A4242:   push eax
  loc_005A4243: call [00401280h] ; __vbaLateIdSt
loc_005A4249:   lea ecx, var_1C
loc_005A424C:   Call edi
  loc_005A424E: or eax, FFFFFFFFh
  loc_005A4251: sub esp, 00000010h
loc_005A4254:   mov edx, esp
  loc_005A4256: mov ecx, 0000000Bh
loc_005A425B:   mov [edx], ecx
loc_005A425D:   mov ecx, var_28
loc_005A4260:   mov [edx+00000004h], ecx
loc_005A4263:   mov ecx, [esi]
  loc_005A4265: push 8001000Dh
loc_005A426A:   push esi
loc_005A426B:   mov [edx+00000008h], eax
loc_005A426E:   mov eax, var_20
loc_005A4271:   mov [edx+0000000Ch], eax
loc_005A4274:   Call [ecx+00000458h]
loc_005A427A:   lea edx, var_1C
loc_005A427D:   push eax
loc_005A427E:   push edx
loc_005A427F:   Call ebx
loc_005A4281:   push eax
  loc_005A4282: call [00401280h] ; __vbaLateIdSt
loc_005A4288:   lea ecx, var_1C
loc_005A428B:   Call edi
  loc_005A428D: mov var_4, 00000000h
  loc_005A4294: push 005A42AFh
  loc_005A4299: jmp 005A42AEh
loc_005A429B:   lea ecx, var_18
  loc_005A429E: call [0040129Ch] ; __vbaFreeStr
loc_005A42A4:   lea ecx, var_1C
  loc_005A42A7: call [004012A0h] ; __vbaFreeObj
loc_005A42AD:   ret
loc_005A42AE:   ret
loc_005A42AF:   mov eax, Me
loc_005A42B2:   push eax
loc_005A42B3:   mov ecx, [eax]
loc_005A42B5:   Call [ecx+00000008h]
loc_005A42B8:   mov eax, var_4
loc_005A42BB:   mov ecx, var_14
loc_005A42BE:   pop edi
loc_005A42BF:   pop esi
loc_005A42C0:   mov fs: [00000000h] , ecx
loc_005A42C7:   pop ebx
loc_005A42C8:   mov esp, ebp
loc_005A42CA:   pop ebp
  loc_005A42CB: retn 0004h
End Sub

Private Sub mnuLapPelanggan_Click() '5A35C0
loc_005A35C0:   push ebp
loc_005A35C1:   mov ebp, esp
  loc_005A35C3: sub esp, 0000000Ch
  loc_005A35C6: push 00405696h ; __vbaExceptHandler
loc_005A35CB:   mov eax, fs: [00000000h]
loc_005A35D1:   push eax
loc_005A35D2:   mov fs: [00000000h] , esp
  loc_005A35D9: sub esp, 00000008h
loc_005A35DC:   push ebx
loc_005A35DD:   push esi
loc_005A35DE:   push edi
loc_005A35DF:   mov var_C, esp
  loc_005A35E2: mov var_8, 00402BB0h
loc_005A35E9:   mov eax, Me
loc_005A35EC:   mov ecx, eax
  loc_005A35EE: and ecx, 00000001h
loc_005A35F1:   mov var_4, ecx
  loc_005A35F4: and al, FEh
loc_005A35F6:   push eax
loc_005A35F7:   mov Me, eax
loc_005A35FA:   mov edx, [eax]
loc_005A35FC:   Call [edx+00000004h]
  loc_005A35FF: call 0060B9C0h
  loc_005A3604: mov var_4, 00000000h
loc_005A360B:   mov eax, Me
loc_005A360E:   push eax
loc_005A360F:   mov ecx, [eax]
loc_005A3611:   Call [ecx+00000008h]
loc_005A3614:   mov eax, var_4
loc_005A3617:   mov ecx, var_14
loc_005A361A:   pop edi
loc_005A361B:   pop esi
loc_005A361C:   mov fs: [00000000h] , ecx
loc_005A3623:   pop ebx
loc_005A3624:   mov esp, ebp
loc_005A3626:   pop ebp
  loc_005A3627: retn 0004h
End Sub

Private Sub Timer1_Timer() '5A5B60
loc_005A5B60:   push ebp
loc_005A5B61:   mov ebp, esp
  loc_005A5B63: sub esp, 0000000Ch
  loc_005A5B66: push 00405696h ; __vbaExceptHandler
loc_005A5B6B:   mov eax, fs: [00000000h]
loc_005A5B71:   push eax
loc_005A5B72:   mov fs: [00000000h] , esp
  loc_005A5B79: sub esp, 00000098h
loc_005A5B7F:   push ebx
loc_005A5B80:   push esi
loc_005A5B81:   push edi
loc_005A5B82:   mov var_C, esp
  loc_005A5B85: mov var_8, 00402D00h
loc_005A5B8C:   mov esi, Me
loc_005A5B8F:   mov eax, esi
  loc_005A5B91: and eax, 00000001h
loc_005A5B94:   mov var_4, eax
  loc_005A5B97: and esi, FFFFFFFEh
loc_005A5B9A:   push esi
loc_005A5B9B:   mov Me, esi
loc_005A5B9E:   mov ecx, [esi]
loc_005A5BA0:   Call [ecx+00000004h]
loc_005A5BA3:   lea edx, var_34
  loc_005A5BA6: xor edi, edi
loc_005A5BA8:   push edx
loc_005A5BA9:   mov var_18, edi
loc_005A5BAC:   mov var_1C, edi
loc_005A5BAF:   mov var_20, edi
loc_005A5BB2:   mov var_24, edi
loc_005A5BB5:   mov var_34, edi
loc_005A5BB8:   mov var_44, edi
loc_005A5BBB:   mov var_54, edi
loc_005A5BBE:   mov var_64, edi
loc_005A5BC1:   mov var_74, edi
loc_005A5BC4:   mov var_84, edi
  loc_005A5BCA: call [00401234h] ; rtcGetTimeVar
loc_005A5BD0:   mov eax, [esi]
  loc_005A5BD2: push 00434FF4h
loc_005A5BD7:   push edi
  loc_005A5BD8: push 00000003h
loc_005A5BDA:   push esi
loc_005A5BDB:   Call [eax+00000470h]
  loc_005A5BE1: mov esi, [004010B8h] ; __vbaObjSet
loc_005A5BE7:   lea ecx, var_1C
loc_005A5BEA:   push eax
loc_005A5BEB:   push ecx
  loc_005A5BEC: call __vbaObjSet
loc_005A5BEE:   lea edx, var_64
loc_005A5BF1:   push eax
loc_005A5BF2:   push edx
  loc_005A5BF3: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005A5BF9: add esp, 00000010h
loc_005A5BFC:   push eax
  loc_005A5BFD: call [00401144h] ; __vbaCastObjVar
loc_005A5C03:   push eax
loc_005A5C04:   lea eax, var_20
loc_005A5C07:   push eax
  loc_005A5C08: call __vbaObjSet
loc_005A5C0A:   mov esi, eax
loc_005A5C0C:   lea edx, var_24
loc_005A5C0F:   lea eax, var_74
  loc_005A5C12: mov var_6C, 00000003h
  loc_005A5C19: mov var_74, 00000002h
loc_005A5C20:   mov ecx, [esi]
loc_005A5C22:   push edx
loc_005A5C23:   push eax
loc_005A5C24:   push esi
loc_005A5C25:   Call [ecx+00000034h]
loc_005A5C28:   cmp eax, edi
loc_005A5C2A:   fnclex
  loc_005A5C2C: jge 005A5C3Dh
  loc_005A5C2E: push 00000034h
  loc_005A5C30: push 00434FF4h
loc_005A5C35:   push esi
loc_005A5C36:   push eax
  loc_005A5C37: call [00401080h] ; __vbaHresultCheckObj
loc_005A5C3D:   mov esi, var_24
loc_005A5C40:   lea edx, var_84
loc_005A5C46:   lea ecx, var_44
  loc_005A5C49: mov var_7C, 004359ECh ; "hh:mm:ss AM/PM"
  loc_005A5C50: mov var_84, 00000008h
  loc_005A5C5A: call [00401238h] ; __vbaVarDup
  loc_005A5C60: push 00000001h
loc_005A5C62:   lea ecx, var_44
  loc_005A5C65: push 00000001h
loc_005A5C67:   lea edx, var_34
loc_005A5C6A:   push ecx
loc_005A5C6B:   lea eax, var_54
loc_005A5C6E:   push edx
loc_005A5C6F:   push eax
  loc_005A5C70: call [00401078h] ; rtcVarFromFormatVar
loc_005A5C76:   mov ebx, [esi]
loc_005A5C78:   lea ecx, var_54
loc_005A5C7B:   lea edx, var_18
loc_005A5C7E:   push ecx
loc_005A5C7F:   push edx
  loc_005A5C80: call [004011C0h] ; __vbaStrVarVal
loc_005A5C86:   push eax
loc_005A5C87:   push esi
loc_005A5C88:   Call [ebx+00000080h]
loc_005A5C8E:   cmp eax, edi
loc_005A5C90:   fnclex
  loc_005A5C92: jge 005A5CA6h
  loc_005A5C94: push 00000080h
  loc_005A5C99: push 00435004h
loc_005A5C9E:   push esi
loc_005A5C9F:   push eax
  loc_005A5CA0: call [00401080h] ; __vbaHresultCheckObj
loc_005A5CA6:   lea ecx, var_18
  loc_005A5CA9: call [0040129Ch] ; __vbaFreeStr
loc_005A5CAF:   lea eax, var_24
loc_005A5CB2:   lea ecx, var_20
loc_005A5CB5:   push eax
loc_005A5CB6:   lea edx, var_1C
loc_005A5CB9:   push ecx
loc_005A5CBA:   push edx
  loc_005A5CBB: push 00000003h
  loc_005A5CBD: call [0040104Ch] ; __vbaFreeObjList
loc_005A5CC3:   lea eax, var_54
loc_005A5CC6:   lea ecx, var_74
loc_005A5CC9:   push eax
loc_005A5CCA:   lea edx, var_64
loc_005A5CCD:   push ecx
loc_005A5CCE:   lea eax, var_44
loc_005A5CD1:   push edx
loc_005A5CD2:   lea ecx, var_34
loc_005A5CD5:   push eax
loc_005A5CD6:   push ecx
  loc_005A5CD7: push 00000005h
  loc_005A5CD9: call [0040103Ch] ; __vbaFreeVarList
  loc_005A5CDF: add esp, 00000028h
loc_005A5CE2:   mov var_4, edi
  loc_005A5CE5: push 005A5D2Ah
  loc_005A5CEA: jmp 005A5D29h
loc_005A5CEC:   lea ecx, var_18
  loc_005A5CEF: call [0040129Ch] ; __vbaFreeStr
loc_005A5CF5:   lea edx, var_24
loc_005A5CF8:   lea eax, var_20
loc_005A5CFB:   push edx
loc_005A5CFC:   lea ecx, var_1C
loc_005A5CFF:   push eax
loc_005A5D00:   push ecx
  loc_005A5D01: push 00000003h
  loc_005A5D03: call [0040104Ch] ; __vbaFreeObjList
loc_005A5D09:   lea edx, var_74
loc_005A5D0C:   lea eax, var_64
loc_005A5D0F:   push edx
loc_005A5D10:   lea ecx, var_54
loc_005A5D13:   push eax
loc_005A5D14:   lea edx, var_44
loc_005A5D17:   push ecx
loc_005A5D18:   lea eax, var_34
loc_005A5D1B:   push edx
loc_005A5D1C:   push eax
  loc_005A5D1D: push 00000005h
  loc_005A5D1F: call [0040103Ch] ; __vbaFreeVarList
  loc_005A5D25: add esp, 00000028h
loc_005A5D28:   ret
loc_005A5D29:   ret
loc_005A5D2A:   mov eax, Me
loc_005A5D2D:   push eax
loc_005A5D2E:   mov ecx, [eax]
loc_005A5D30:   Call [ecx+00000008h]
loc_005A5D33:   mov eax, var_4
loc_005A5D36:   mov ecx, var_14
loc_005A5D39:   pop edi
loc_005A5D3A:   pop esi
loc_005A5D3B:   mov fs: [00000000h] , ecx
loc_005A5D42:   pop ebx
loc_005A5D43:   mov esp, ebp
loc_005A5D45:   pop ebp
  loc_005A5D46: retn 0004h
End Sub

Private Sub Form_Load() '5A2AA0
loc_005A2AA0:   push ebp
loc_005A2AA1:   mov ebp, esp
  loc_005A2AA3: sub esp, 0000000Ch
  loc_005A2AA6: push 00405696h ; __vbaExceptHandler
loc_005A2AAB:   mov eax, fs: [00000000h]
loc_005A2AB1:   push eax
loc_005A2AB2:   mov fs: [00000000h] , esp
  loc_005A2AB9: sub esp, 0000009Ch
loc_005A2ABF:   push ebx
loc_005A2AC0:   push esi
loc_005A2AC1:   push edi
loc_005A2AC2:   mov var_C, esp
  loc_005A2AC5: mov var_8, 00402B50h
loc_005A2ACC:   mov edi, Me
loc_005A2ACF:   mov eax, edi
  loc_005A2AD1: and eax, 00000001h
loc_005A2AD4:   mov var_4, eax
  loc_005A2AD7: and edi, FFFFFFFEh
loc_005A2ADA:   push edi
loc_005A2ADB:   mov Me, edi
loc_005A2ADE:   mov ecx, [edi]
loc_005A2AE0:   Call [ecx+00000004h]
loc_005A2AE3:   lea edx, var_34
  loc_005A2AE6: xor ebx, ebx
loc_005A2AE8:   push edx
loc_005A2AE9:   mov var_18, ebx
loc_005A2AEC:   mov var_1C, ebx
loc_005A2AEF:   mov var_20, ebx
loc_005A2AF2:   mov var_24, ebx
loc_005A2AF5:   mov var_34, ebx
loc_005A2AF8:   mov var_44, ebx
loc_005A2AFB:   mov var_54, ebx
loc_005A2AFE:   mov var_64, ebx
loc_005A2B01:   mov var_74, ebx
loc_005A2B04:   mov var_84, ebx
  loc_005A2B0A: call [00401220h] ; rtcGetDateVar
loc_005A2B10:   mov eax, [edi]
  loc_005A2B12: push 00434FF4h
loc_005A2B17:   push ebx
  loc_005A2B18: push 00000003h
loc_005A2B1A:   push edi
loc_005A2B1B:   Call [eax+00000470h]
  loc_005A2B21: mov esi, [004010B8h] ; __vbaObjSet
loc_005A2B27:   lea ecx, var_1C
loc_005A2B2A:   push eax
loc_005A2B2B:   push ecx
  loc_005A2B2C: call __vbaObjSet
loc_005A2B2E:   lea edx, var_64
loc_005A2B31:   push eax
loc_005A2B32:   push edx
  loc_005A2B33: call [0040114Ch] ; __vbaLateIdCallLd
  loc_005A2B39: add esp, 00000010h
loc_005A2B3C:   push eax
  loc_005A2B3D: call [00401144h] ; __vbaCastObjVar
loc_005A2B43:   push eax
loc_005A2B44:   lea eax, var_20
loc_005A2B47:   push eax
  loc_005A2B48: call __vbaObjSet
loc_005A2B4A:   mov esi, eax
loc_005A2B4C:   lea edx, var_24
loc_005A2B4F:   lea eax, var_74
  loc_005A2B52: mov var_6C, 00000004h
  loc_005A2B59: mov var_74, 00000002h
loc_005A2B60:   mov ecx, [esi]
loc_005A2B62:   push edx
loc_005A2B63:   push eax
loc_005A2B64:   push esi
loc_005A2B65:   Call [ecx+00000034h]
loc_005A2B68:   cmp eax, ebx
loc_005A2B6A:   fnclex
  loc_005A2B6C: jge 005A2B7Dh
  loc_005A2B6E: push 00000034h
  loc_005A2B70: push 00434FF4h
loc_005A2B75:   push esi
loc_005A2B76:   push eax
  loc_005A2B77: call [00401080h] ; __vbaHresultCheckObj
loc_005A2B7D:   mov esi, var_24
loc_005A2B80:   lea edx, var_84
loc_005A2B86:   lea ecx, var_44
loc_005A2B89:   mov var_A0, esi
  loc_005A2B8F: mov var_7C, 00434FD0h ; "dddd, dd-mm-yyyy"
  loc_005A2B96: mov var_84, 00000008h
  loc_005A2BA0: call [00401238h] ; __vbaVarDup
  loc_005A2BA6: push 00000001h
loc_005A2BA8:   lea ecx, var_44
  loc_005A2BAB: push 00000001h
loc_005A2BAD:   lea edx, var_34
loc_005A2BB0:   push ecx
loc_005A2BB1:   lea eax, var_54
loc_005A2BB4:   push edx
loc_005A2BB5:   push eax
  loc_005A2BB6: call [00401078h] ; rtcVarFromFormatVar
loc_005A2BBC:   mov esi, [esi]
loc_005A2BBE:   lea ecx, var_54
loc_005A2BC1:   lea edx, var_18
loc_005A2BC4:   push ecx
loc_005A2BC5:   push edx
  loc_005A2BC6: call [004011C0h] ; __vbaStrVarVal
loc_005A2BCC:   mov var_B0, esi
loc_005A2BD2:   mov esi, var_A0
loc_005A2BD8:   push eax
loc_005A2BD9:   mov eax, var_B0
loc_005A2BDF:   push esi
loc_005A2BE0:   Call [eax+00000080h]
loc_005A2BE6:   cmp eax, ebx
loc_005A2BE8:   fnclex
  loc_005A2BEA: jge 005A2BFEh
  loc_005A2BEC: push 00000080h
  loc_005A2BF1: push 00435004h
loc_005A2BF6:   push esi
loc_005A2BF7:   push eax
  loc_005A2BF8: call [00401080h] ; __vbaHresultCheckObj
loc_005A2BFE:   lea ecx, var_18
  loc_005A2C01: call [0040129Ch] ; __vbaFreeStr
loc_005A2C07:   lea ecx, var_24
loc_005A2C0A:   lea edx, var_20
loc_005A2C0D:   push ecx
loc_005A2C0E:   lea eax, var_1C
loc_005A2C11:   push edx
loc_005A2C12:   push eax
  loc_005A2C13: push 00000003h
  loc_005A2C15: call [0040104Ch] ; __vbaFreeObjList
loc_005A2C1B:   lea ecx, var_54
loc_005A2C1E:   lea edx, var_74
loc_005A2C21:   push ecx
loc_005A2C22:   lea eax, var_64
loc_005A2C25:   push edx
loc_005A2C26:   lea ecx, var_44
loc_005A2C29:   push eax
loc_005A2C2A:   lea edx, var_34
loc_005A2C2D:   push ecx
loc_005A2C2E:   push edx
  loc_005A2C2F: push 00000005h
  loc_005A2C31: call [0040103Ch] ; __vbaFreeVarList
loc_005A2C37:   mov eax, [006682A0h]
  loc_005A2C3C: add esp, 00000028h
loc_005A2C3F:   cmp eax, ebx
  loc_005A2C41: jnz 005A2C58h
  loc_005A2C43: push 006682A0h
  loc_005A2C48: push 00414A94h
  loc_005A2C4D: call [004011E8h] ; __vbaNew2
loc_005A2C53:   mov eax, [006682A0h]
loc_005A2C58:   mov ecx, [eax]
loc_005A2C5A:   push eax
loc_005A2C5B:   Call [ecx+00000308h]
loc_005A2C61:   lea edx, var_1C
loc_005A2C64:   push eax
loc_005A2C65:   push edx
  loc_005A2C66: call [004010B8h] ; __vbaObjSet
loc_005A2C6C:   mov esi, eax
loc_005A2C6E:   lea ecx, var_18
loc_005A2C71:   push ecx
loc_005A2C72:   push esi
loc_005A2C73:   mov eax, [esi]
loc_005A2C75:   Call [eax+000000A0h]
loc_005A2C7B:   cmp eax, ebx
loc_005A2C7D:   fnclex
  loc_005A2C7F: jge 005A2C93h
  loc_005A2C81: push 000000A0h
  loc_005A2C86: push 0042DFCCh
loc_005A2C8B:   push esi
loc_005A2C8C:   push eax
  loc_005A2C8D: call [00401080h] ; __vbaHresultCheckObj
loc_005A2C93:   mov edx, var_18
loc_005A2C96:   push edx
  loc_005A2C97: push 00434FA8h ; "UNREGISTERED"
  loc_005A2C9C: call [0040112Ch] ; __vbaStrCmp
loc_005A2CA2:   mov esi, eax
loc_005A2CA4:   lea ecx, var_18
loc_005A2CA7:   neg esi
loc_005A2CA9:   sbb esi, esi
loc_005A2CAB:   inc esi
loc_005A2CAC:   neg esi
  loc_005A2CAE: call [0040129Ch] ; __vbaFreeStr
loc_005A2CB4:   lea ecx, var_1C
  loc_005A2CB7: call [004012A0h] ; __vbaFreeObj
loc_005A2CBD:   cmp si, bx
  loc_005A2CC0: jz 005A2CFDh
loc_005A2CC2:   mov eax, [edi]
loc_005A2CC4:   push edi
loc_005A2CC5:   Call [eax+00000358h]
loc_005A2CCB:   lea ecx, var_1C
loc_005A2CCE:   push eax
loc_005A2CCF:   push ecx
  loc_005A2CD0: call [004010B8h] ; __vbaObjSet
loc_005A2CD6:   mov esi, eax
loc_005A2CD8:   push ebx
loc_005A2CD9:   push esi
loc_005A2CDA:   mov edx, [esi]
loc_005A2CDC:   Call [edx+00000074h]
loc_005A2CDF:   cmp eax, ebx
loc_005A2CE1:   fnclex
  loc_005A2CE3: jge 005A2CF4h
  loc_005A2CE5: push 00000074h
  loc_005A2CE7: push 00434274h
loc_005A2CEC:   push esi
loc_005A2CED:   push eax
  loc_005A2CEE: call [00401080h] ; __vbaHresultCheckObj
loc_005A2CF4:   lea ecx, var_1C
  loc_005A2CF7: call [004012A0h] ; __vbaFreeObj
loc_005A2CFD:   mov var_4, ebx
  loc_005A2D00: push 005A2D45h
  loc_005A2D05: jmp 005A2D44h
loc_005A2D07:   lea ecx, var_18
  loc_005A2D0A: call [0040129Ch] ; __vbaFreeStr
loc_005A2D10:   lea eax, var_24
loc_005A2D13:   lea ecx, var_20
loc_005A2D16:   push eax
loc_005A2D17:   lea edx, var_1C
loc_005A2D1A:   push ecx
loc_005A2D1B:   push edx
  loc_005A2D1C: push 00000003h
  loc_005A2D1E: call [0040104Ch] ; __vbaFreeObjList
loc_005A2D24:   lea eax, var_74
loc_005A2D27:   lea ecx, var_64
loc_005A2D2A:   push eax
loc_005A2D2B:   lea edx, var_54
loc_005A2D2E:   push ecx
loc_005A2D2F:   lea eax, var_44
loc_005A2D32:   push edx
loc_005A2D33:   lea ecx, var_34
loc_005A2D36:   push eax
loc_005A2D37:   push ecx
  loc_005A2D38: push 00000005h
  loc_005A2D3A: call [0040103Ch] ; __vbaFreeVarList
  loc_005A2D40: add esp, 00000028h
loc_005A2D43:   ret
loc_005A2D44:   ret
loc_005A2D45:   mov eax, Me
loc_005A2D48:   push eax
loc_005A2D49:   mov edx, [eax]
loc_005A2D4B:   Call [edx+00000008h]
loc_005A2D4E:   mov eax, var_4
loc_005A2D51:   mov ecx, var_14
loc_005A2D54:   pop edi
loc_005A2D55:   pop esi
loc_005A2D56:   mov fs: [00000000h] , ecx
loc_005A2D5D:   pop ebx
loc_005A2D5E:   mov esp, ebp
loc_005A2D60:   pop ebp
  loc_005A2D61: retn 0004h
End Sub

Private Sub Form_Unload(Cancel As Integer) '5A2D70
loc_005A2D70:   push ebp
loc_005A2D71:   mov ebp, esp
  loc_005A2D73: sub esp, 0000000Ch
  loc_005A2D76: push 00405696h ; __vbaExceptHandler
loc_005A2D7B:   mov eax, fs: [00000000h]
loc_005A2D81:   push eax
loc_005A2D82:   mov fs: [00000000h] , esp
  loc_005A2D89: sub esp, 0000008Ch
loc_005A2D8F:   push ebx
loc_005A2D90:   push esi
loc_005A2D91:   push edi
loc_005A2D92:   mov var_C, esp
  loc_005A2D95: mov var_8, 00402B60h
loc_005A2D9C:   mov eax, Me
loc_005A2D9F:   mov ecx, eax
  loc_005A2DA1: and ecx, 00000001h
loc_005A2DA4:   mov var_4, ecx
  loc_005A2DA7: and al, FEh
loc_005A2DA9:   push eax
loc_005A2DAA:   mov Me, eax
loc_005A2DAD:   mov edx, [eax]
loc_005A2DAF:   Call [edx+00000004h]
  loc_005A2DB2: mov edi, [00401238h] ; __vbaVarDup
  loc_005A2DB8: mov ecx, 80020004h
  loc_005A2DBD: xor esi, esi
loc_005A2DBF:   mov var_4C, ecx
  loc_005A2DC2: mov eax, 0000000Ah
loc_005A2DC7:   mov var_3C, ecx
  loc_005A2DCA: mov ebx, 00000008h
loc_005A2DCF:   mov var_44, esi
loc_005A2DD2:   mov var_54, esi
loc_005A2DD5:   mov var_74, esi
loc_005A2DD8:   lea edx, var_74
loc_005A2DDB:   lea ecx, var_34
loc_005A2DDE:   mov var_24, esi
loc_005A2DE1:   mov var_34, esi
loc_005A2DE4:   mov var_64, esi
loc_005A2DE7:   mov var_54, eax
loc_005A2DEA:   mov var_44, eax
  loc_005A2DED: mov var_6C, 0042ED7Ch ; "Konfirmasi"
loc_005A2DF4:   mov var_74, ebx
loc_005A2DF7:   Call edi
loc_005A2DF9:   lea edx, var_64
loc_005A2DFC:   lea ecx, var_24
  loc_005A2DFF: mov var_5C, 00434874h ; "APAKAH KAMU INGIN MENUTUP APLIKASI INI..?"
loc_005A2E06:   mov var_64, ebx
loc_005A2E09:   Call edi
loc_005A2E0B:   lea eax, var_54
loc_005A2E0E:   lea ecx, var_44
loc_005A2E11:   push eax
loc_005A2E12:   lea edx, var_34
loc_005A2E15:   push ecx
loc_005A2E16:   push edx
loc_005A2E17:   lea eax, var_24
  loc_005A2E1A: push 00000024h
loc_005A2E1C:   push eax
  loc_005A2E1D: call [004010B4h] ; rtcMsgBox
  loc_005A2E23: xor ecx, ecx
  loc_005A2E25: cmp eax, 00000007h
loc_005A2E28:   setz cl
loc_005A2E2B:   neg ecx
loc_005A2E2D:   lea edx, var_54
loc_005A2E30:   mov edi, ecx
loc_005A2E32:   lea eax, var_44
loc_005A2E35:   push edx
loc_005A2E36:   lea ecx, var_34
loc_005A2E39:   push eax
loc_005A2E3A:   lea edx, var_24
loc_005A2E3D:   push ecx
loc_005A2E3E:   push edx
  loc_005A2E3F: push 00000004h
  loc_005A2E41: call [0040103Ch] ; __vbaFreeVarList
  loc_005A2E47: add esp, 00000014h
loc_005A2E4A:   cmp di, si
  loc_005A2E4D: jz 005A2E59h
loc_005A2E4F:   mov eax, Cancel
  loc_005A2E52: mov [eax], 0001h
  loc_005A2E57: jmp 005A2E5Fh
  loc_005A2E59: call [00401038h] ; __vbaEnd
loc_005A2E5F:   mov var_4, esi
  loc_005A2E62: push 005A2E86h
  loc_005A2E67: jmp 005A2E85h
loc_005A2E69:   lea ecx, var_54
loc_005A2E6C:   lea edx, var_44
loc_005A2E6F:   push ecx
loc_005A2E70:   lea eax, var_34
loc_005A2E73:   push edx
loc_005A2E74:   lea ecx, var_24
loc_005A2E77:   push eax
loc_005A2E78:   push ecx
  loc_005A2E79: push 00000004h
  loc_005A2E7B: call [0040103Ch] ; __vbaFreeVarList
  loc_005A2E81: add esp, 00000014h
loc_005A2E84:   ret
loc_005A2E85:   ret
loc_005A2E86:   mov eax, Me
loc_005A2E89:   push eax
loc_005A2E8A:   mov edx, [eax]
loc_005A2E8C:   Call [edx+00000008h]
loc_005A2E8F:   mov eax, var_4
loc_005A2E92:   mov ecx, var_14
loc_005A2E95:   pop edi
loc_005A2E96:   pop esi
loc_005A2E97:   mov fs: [00000000h] , ecx
loc_005A2E9E:   pop ebx
loc_005A2E9F:   mov esp, ebp
loc_005A2EA1:   pop ebp
  loc_005A2EA2: retn 0008h
End Sub

Private Sub Form_Activate() '5A2990
loc_005A2990:   push ebp
loc_005A2991:   mov ebp, esp
  loc_005A2993: sub esp, 0000000Ch
  loc_005A2996: push 00405696h ; __vbaExceptHandler
loc_005A299B:   mov eax, fs: [00000000h]
loc_005A29A1:   push eax
loc_005A29A2:   mov fs: [00000000h] , esp
  loc_005A29A9: sub esp, 0000001Ch
loc_005A29AC:   push ebx
loc_005A29AD:   push esi
loc_005A29AE:   push edi
loc_005A29AF:   mov var_C, esp
  loc_005A29B2: mov var_8, 00402B40h
loc_005A29B9:   mov esi, Me
loc_005A29BC:   mov eax, esi
  loc_005A29BE: and eax, 00000001h
loc_005A29C1:   mov var_4, eax
  loc_005A29C4: and esi, FFFFFFFEh
loc_005A29C7:   push esi
loc_005A29C8:   mov Me, esi
loc_005A29CB:   mov ecx, [esi]
loc_005A29CD:   Call [ecx+00000004h]
loc_005A29D0:   mov edx, [esi]
  loc_005A29D2: xor ebx, ebx
loc_005A29D4:   push esi
loc_005A29D5:   mov var_18, ebx
loc_005A29D8:   mov var_1C, ebx
loc_005A29DB:   mov var_20, ebx
loc_005A29DE:   Call [edx+0000033Ch]
loc_005A29E4:   push eax
loc_005A29E5:   lea eax, var_20
loc_005A29E8:   push eax
  loc_005A29E9: call [004010B8h] ; __vbaObjSet
loc_005A29EF:   mov edi, eax
loc_005A29F1:   lea edx, var_18
loc_005A29F4:   push edx
loc_005A29F5:   push edi
loc_005A29F6:   mov ecx, [edi]
loc_005A29F8:   Call [ecx+00000050h]
loc_005A29FB:   cmp eax, ebx
loc_005A29FD:   fnclex
  loc_005A29FF: jge 005A2A10h
  loc_005A2A01: push 00000050h
  loc_005A2A03: push 00433F80h
loc_005A2A08:   push edi
loc_005A2A09:   push eax
  loc_005A2A0A: call [00401080h] ; __vbaHresultCheckObj
loc_005A2A10:   mov eax, var_18
  loc_005A2A13: push 00434ABCh ; "  "
loc_005A2A18:   push eax
  loc_005A2A19: call [00401070h] ; __vbaStrCat
loc_005A2A1F:   mov edx, eax
loc_005A2A21:   lea ecx, var_1C
  loc_005A2A24: call [0040126Ch] ; __vbaStrMove
loc_005A2A2A:   mov edx, eax
loc_005A2A2C:   lea ecx, [esi+00000034h]
  loc_005A2A2F: call [004011FCh] ; __vbaStrCopy
loc_005A2A35:   lea ecx, var_1C
loc_005A2A38:   lea edx, var_18
loc_005A2A3B:   push ecx
loc_005A2A3C:   push edx
  loc_005A2A3D: push 00000002h
  loc_005A2A3F: call [00401204h] ; __vbaFreeStrList
  loc_005A2A45: add esp, 0000000Ch
loc_005A2A48:   lea ecx, var_20
  loc_005A2A4B: call [004012A0h] ; __vbaFreeObj
loc_005A2A51:   mov var_4, ebx
  loc_005A2A54: push 005A2A79h
  loc_005A2A59: jmp 005A2A78h
loc_005A2A5B:   lea eax, var_1C
loc_005A2A5E:   lea ecx, var_18
loc_005A2A61:   push eax
loc_005A2A62:   push ecx
  loc_005A2A63: push 00000002h
  loc_005A2A65: call [00401204h] ; __vbaFreeStrList
  loc_005A2A6B: add esp, 0000000Ch
loc_005A2A6E:   lea ecx, var_20
  loc_005A2A71: call [004012A0h] ; __vbaFreeObj
loc_005A2A77:   ret
loc_005A2A78:   ret
loc_005A2A79:   mov eax, Me
loc_005A2A7C:   push eax
loc_005A2A7D:   mov edx, [eax]
loc_005A2A7F:   Call [edx+00000008h]
loc_005A2A82:   mov eax, var_4
loc_005A2A85:   mov ecx, var_14
loc_005A2A88:   pop edi
loc_005A2A89:   pop esi
loc_005A2A8A:   mov fs: [00000000h] , ecx
loc_005A2A91:   pop ebx
loc_005A2A92:   mov esp, ebp
loc_005A2A94:   pop ebp
  loc_005A2A95: retn 0004h
End Sub

Private Sub mnuLapJualPeriode_Click() '5A52F0
loc_005A52F0:   push ebp
loc_005A52F1:   mov ebp, esp
  loc_005A52F3: sub esp, 0000000Ch
  loc_005A52F6: push 00405696h ; __vbaExceptHandler
loc_005A52FB:   mov eax, fs: [00000000h]
loc_005A5301:   push eax
loc_005A5302:   mov fs: [00000000h] , esp
  loc_005A5309: sub esp, 00000030h
loc_005A530C:   push ebx
loc_005A530D:   push esi
loc_005A530E:   push edi
loc_005A530F:   mov var_C, esp
  loc_005A5312: mov var_8, 00402CB8h
loc_005A5319:   mov eax, Me
loc_005A531C:   mov ecx, eax
  loc_005A531E: and ecx, 00000001h
loc_005A5321:   mov var_4, ecx
  loc_005A5324: and al, FEh
loc_005A5326:   push eax
loc_005A5327:   mov Me, eax
loc_005A532A:   mov edx, [eax]
loc_005A532C:   Call [edx+00000004h]
loc_005A532F:   mov eax, [006683F0h]
loc_005A5334:   test eax, eax
  loc_005A5336: jnz 005A5348h
  loc_005A5338: push 006683F0h
  loc_005A533D: push 00411B0Ch
  loc_005A5342: call [004011E8h] ; __vbaNew2
  loc_005A5348: sub esp, 00000010h
  loc_005A534B: mov ecx, 0000000Ah
loc_005A5350:   mov ebx, esp
  loc_005A5352: mov eax, 80020004h
  loc_005A5357: sub esp, 00000010h
loc_005A535A:   mov esi, [006683F0h]
loc_005A5360:   mov [ebx], ecx
loc_005A5362:   mov ecx, var_30
loc_005A5365:   mov edi, [esi]
  loc_005A5367: mov edx, 00000001h
loc_005A536C:   mov [ebx+00000004h], ecx
loc_005A536F:   mov ecx, esp
loc_005A5371:   push esi
loc_005A5372:   mov [ebx+00000008h], eax
loc_005A5375:   mov eax, var_28
loc_005A5378:   mov [ebx+0000000Ch], eax
  loc_005A537B: mov eax, 00000002h
loc_005A5380:   mov [ecx], eax
loc_005A5382:   mov eax, var_20
loc_005A5385:   mov [ecx+00000004h], eax
loc_005A5388:   mov [ecx+00000008h], edx
loc_005A538B:   mov edx, var_18
loc_005A538E:   mov [ecx+0000000Ch], edx
loc_005A5391:   Call [edi+000002B0h]
loc_005A5397:   test eax, eax
loc_005A5399:   fnclex
  loc_005A539B: jge 005A53AFh
  loc_005A539D: push 000002B0h
  loc_005A53A2: push 0043574Ch
loc_005A53A7:   push esi
loc_005A53A8:   push eax
  loc_005A53A9: call [00401080h] ; __vbaHresultCheckObj
  loc_005A53AF: mov var_4, 00000000h
loc_005A53B6:   mov eax, Me
loc_005A53B9:   push eax
loc_005A53BA:   mov ecx, [eax]
loc_005A53BC:   Call [ecx+00000008h]
loc_005A53BF:   mov eax, var_4
loc_005A53C2:   mov ecx, var_14
loc_005A53C5:   pop edi
loc_005A53C6:   pop esi
loc_005A53C7:   mov fs: [00000000h] , ecx
loc_005A53CE:   pop ebx
loc_005A53CF:   mov esp, ebp
loc_005A53D1:   pop ebp
  loc_005A53D2: retn 0004h
End Sub

Private Sub mnuLapPemasok_Click() '5A4D00
loc_005A4D00:   push ebp
loc_005A4D01:   mov ebp, esp
  loc_005A4D03: sub esp, 0000000Ch
  loc_005A4D06: push 00405696h ; __vbaExceptHandler
loc_005A4D0B:   mov eax, fs: [00000000h]
loc_005A4D11:   push eax
loc_005A4D12:   mov fs: [00000000h] , esp
  loc_005A4D19: sub esp, 00000008h
loc_005A4D1C:   push ebx
loc_005A4D1D:   push esi
loc_005A4D1E:   push edi
loc_005A4D1F:   mov var_C, esp
  loc_005A4D22: mov var_8, 00402C70h
loc_005A4D29:   mov eax, Me
loc_005A4D2C:   mov ecx, eax
  loc_005A4D2E: and ecx, 00000001h
loc_005A4D31:   mov var_4, ecx
  loc_005A4D34: and al, FEh
loc_005A4D36:   push eax
loc_005A4D37:   mov Me, eax
loc_005A4D3A:   mov edx, [eax]
loc_005A4D3C:   Call [edx+00000004h]
  loc_005A4D3F: call 0060AC40h
  loc_005A4D44: mov var_4, 00000000h
loc_005A4D4B:   mov eax, Me
loc_005A4D4E:   push eax
loc_005A4D4F:   mov ecx, [eax]
loc_005A4D51:   Call [ecx+00000008h]
loc_005A4D54:   mov eax, var_4
loc_005A4D57:   mov ecx, var_14
loc_005A4D5A:   pop edi
loc_005A4D5B:   pop esi
loc_005A4D5C:   mov fs: [00000000h] , ecx
loc_005A4D63:   pop ebx
loc_005A4D64:   mov esp, ebp
loc_005A4D66:   pop ebp
  loc_005A4D67: retn 0004h
End Sub

Private Sub mnuKontrolStok_Click() '5A3280
loc_005A3280:   push ebp
loc_005A3281:   mov ebp, esp
  loc_005A3283: sub esp, 0000000Ch
  loc_005A3286: push 00405696h ; __vbaExceptHandler
loc_005A328B:   mov eax, fs: [00000000h]
loc_005A3291:   push eax
loc_005A3292:   mov fs: [00000000h] , esp
  loc_005A3299: sub esp, 00000030h
loc_005A329C:   push ebx
loc_005A329D:   push esi
loc_005A329E:   push edi
loc_005A329F:   mov var_C, esp
  loc_005A32A2: mov var_8, 00402B90h
loc_005A32A9:   mov eax, Me
loc_005A32AC:   mov ecx, eax
  loc_005A32AE: and ecx, 00000001h
loc_005A32B1:   mov var_4, ecx
  loc_005A32B4: and al, FEh
loc_005A32B6:   push eax
loc_005A32B7:   mov Me, eax
loc_005A32BA:   mov edx, [eax]
loc_005A32BC:   Call [edx+00000004h]
loc_005A32BF:   mov eax, [00668204h]
loc_005A32C4:   test eax, eax
  loc_005A32C6: jnz 005A32D8h
  loc_005A32C8: push 00668204h
  loc_005A32CD: push 0040AE1Ch
  loc_005A32D2: call [004011E8h] ; __vbaNew2
  loc_005A32D8: sub esp, 00000010h
  loc_005A32DB: mov ecx, 0000000Ah
loc_005A32E0:   mov ebx, esp
  loc_005A32E2: mov eax, 80020004h
  loc_005A32E7: sub esp, 00000010h
loc_005A32EA:   mov esi, [00668204h]
loc_005A32F0:   mov [ebx], ecx
loc_005A32F2:   mov ecx, var_30
loc_005A32F5:   mov edi, [esi]
  loc_005A32F7: mov edx, 00000001h
loc_005A32FC:   mov [ebx+00000004h], ecx
loc_005A32FF:   mov ecx, esp
loc_005A3301:   push esi
loc_005A3302:   mov [ebx+00000008h], eax
loc_005A3305:   mov eax, var_28
loc_005A3308:   mov [ebx+0000000Ch], eax
  loc_005A330B: mov eax, 00000002h
loc_005A3310:   mov [ecx], eax
loc_005A3312:   mov eax, var_20
loc_005A3315:   mov [ecx+00000004h], eax
loc_005A3318:   mov [ecx+00000008h], edx
loc_005A331B:   mov edx, var_18
loc_005A331E:   mov [ecx+0000000Ch], edx
loc_005A3321:   Call [edi+000002B0h]
loc_005A3327:   test eax, eax
loc_005A3329:   fnclex
  loc_005A332B: jge 005A333Fh
  loc_005A332D: push 000002B0h
  loc_005A3332: push 00432FF0h
loc_005A3337:   push esi
loc_005A3338:   push eax
  loc_005A3339: call [00401080h] ; __vbaHresultCheckObj
  loc_005A333F: mov var_4, 00000000h
loc_005A3346:   mov eax, Me
loc_005A3349:   push eax
loc_005A334A:   mov ecx, [eax]
loc_005A334C:   Call [ecx+00000008h]
loc_005A334F:   mov eax, var_4
loc_005A3352:   mov ecx, var_14
loc_005A3355:   pop edi
loc_005A3356:   pop esi
loc_005A3357:   mov fs: [00000000h] , ecx
loc_005A335E:   pop ebx
loc_005A335F:   mov esp, ebp
loc_005A3361:   pop ebp
  loc_005A3362: retn 0004h
End Sub

Private Sub mnuLapBeliNota_Click() '5A5020
loc_005A5020:   push ebp
loc_005A5021:   mov ebp, esp
  loc_005A5023: sub esp, 0000000Ch
  loc_005A5026: push 00405696h ; __vbaExceptHandler
loc_005A502B:   mov eax, fs: [00000000h]
loc_005A5031:   push eax
loc_005A5032:   mov fs: [00000000h] , esp
  loc_005A5039: sub esp, 00000030h
loc_005A503C:   push ebx
loc_005A503D:   push esi
loc_005A503E:   push edi
loc_005A503F:   mov var_C, esp
  loc_005A5042: mov var_8, 00402CA0h
loc_005A5049:   mov eax, Me
loc_005A504C:   mov ecx, eax
  loc_005A504E: and ecx, 00000001h
loc_005A5051:   mov var_4, ecx
  loc_005A5054: and al, FEh
loc_005A5056:   push eax
loc_005A5057:   mov Me, eax
loc_005A505A:   mov edx, [eax]
loc_005A505C:   Call [edx+00000004h]
loc_005A505F:   mov eax, [006683B4h]
loc_005A5064:   test eax, eax
  loc_005A5066: jnz 005A5078h
  loc_005A5068: push 006683B4h
  loc_005A506D: push 0040B848h
  loc_005A5072: call [004011E8h] ; __vbaNew2
  loc_005A5078: sub esp, 00000010h
  loc_005A507B: mov ecx, 0000000Ah
loc_005A5080:   mov ebx, esp
  loc_005A5082: mov eax, 80020004h
  loc_005A5087: sub esp, 00000010h
loc_005A508A:   mov esi, [006683B4h]
loc_005A5090:   mov [ebx], ecx
loc_005A5092:   mov ecx, var_30
loc_005A5095:   mov edi, [esi]
  loc_005A5097: mov edx, 00000001h
loc_005A509C:   mov [ebx+00000004h], ecx
loc_005A509F:   mov ecx, esp
loc_005A50A1:   push esi
loc_005A50A2:   mov [ebx+00000008h], eax
loc_005A50A5:   mov eax, var_28
loc_005A50A8:   mov [ebx+0000000Ch], eax
  loc_005A50AB: mov eax, 00000002h
loc_005A50B0:   mov [ecx], eax
loc_005A50B2:   mov eax, var_20
loc_005A50B5:   mov [ecx+00000004h], eax
loc_005A50B8:   mov [ecx+00000008h], edx
loc_005A50BB:   mov edx, var_18
loc_005A50BE:   mov [ecx+0000000Ch], edx
loc_005A50C1:   Call [edi+000002B0h]
loc_005A50C7:   test eax, eax
loc_005A50C9:   fnclex
  loc_005A50CB: jge 005A50DFh
  loc_005A50CD: push 000002B0h
  loc_005A50D2: push 0043566Ch
loc_005A50D7:   push esi
loc_005A50D8:   push eax
  loc_005A50D9: call [00401080h] ; __vbaHresultCheckObj
  loc_005A50DF: mov var_4, 00000000h
loc_005A50E6:   mov eax, Me
loc_005A50E9:   push eax
loc_005A50EA:   mov ecx, [eax]
loc_005A50EC:   Call [ecx+00000008h]
loc_005A50EF:   mov eax, var_4
loc_005A50F2:   mov ecx, var_14
loc_005A50F5:   pop edi
loc_005A50F6:   pop esi
loc_005A50F7:   mov fs: [00000000h] , ecx
loc_005A50FE:   pop ebx
loc_005A50FF:   mov esp, ebp
loc_005A5101:   pop ebp
  loc_005A5102: retn 0004h
End Sub

Private Sub mnuSatuan_Click() '5A5890
loc_005A5890:   push ebp
loc_005A5891:   mov ebp, esp
  loc_005A5893: sub esp, 0000000Ch
  loc_005A5896: push 00405696h ; __vbaExceptHandler
loc_005A589B:   mov eax, fs: [00000000h]
loc_005A58A1:   push eax
loc_005A58A2:   mov fs: [00000000h] , esp
  loc_005A58A9: sub esp, 00000030h
loc_005A58AC:   push ebx
loc_005A58AD:   push esi
loc_005A58AE:   push edi
loc_005A58AF:   mov var_C, esp
  loc_005A58B2: mov var_8, 00402CE8h
loc_005A58B9:   mov eax, Me
loc_005A58BC:   mov ecx, eax
  loc_005A58BE: and ecx, 00000001h
loc_005A58C1:   mov var_4, ecx
  loc_005A58C4: and al, FEh
loc_005A58C6:   push eax
loc_005A58C7:   mov Me, eax
loc_005A58CA:   mov edx, [eax]
loc_005A58CC:   Call [edx+00000004h]
loc_005A58CF:   mov eax, [006681B4h]
loc_005A58D4:   test eax, eax
  loc_005A58D6: jnz 005A58E8h
  loc_005A58D8: push 006681B4h
  loc_005A58DD: push 004123CCh
  loc_005A58E2: call [004011E8h] ; __vbaNew2
  loc_005A58E8: sub esp, 00000010h
  loc_005A58EB: mov ecx, 0000000Ah
loc_005A58F0:   mov ebx, esp
  loc_005A58F2: mov eax, 80020004h
  loc_005A58F7: sub esp, 00000010h
loc_005A58FA:   mov esi, [006681B4h]
loc_005A5900:   mov [ebx], ecx
loc_005A5902:   mov ecx, var_30
loc_005A5905:   mov edi, [esi]
  loc_005A5907: mov edx, 00000001h
loc_005A590C:   mov [ebx+00000004h], ecx
loc_005A590F:   mov ecx, esp
loc_005A5911:   push esi
loc_005A5912:   mov [ebx+00000008h], eax
loc_005A5915:   mov eax, var_28
loc_005A5918:   mov [ebx+0000000Ch], eax
  loc_005A591B: mov eax, 00000002h
loc_005A5920:   mov [ecx], eax
loc_005A5922:   mov eax, var_20
loc_005A5925:   mov [ecx+00000004h], eax
loc_005A5928:   mov [ecx+00000008h], edx
loc_005A592B:   mov edx, var_18
loc_005A592E:   mov [ecx+0000000Ch], edx
loc_005A5931:   Call [edi+000002B0h]
loc_005A5937:   test eax, eax
loc_005A5939:   fnclex
  loc_005A593B: jge 005A594Fh
  loc_005A593D: push 000002B0h
  loc_005A5942: push 00430490h
loc_005A5947:   push esi
loc_005A5948:   push eax
  loc_005A5949: call [00401080h] ; __vbaHresultCheckObj
  loc_005A594F: mov var_4, 00000000h
loc_005A5956:   mov eax, Me
loc_005A5959:   push eax
loc_005A595A:   mov ecx, [eax]
loc_005A595C:   Call [ecx+00000008h]
loc_005A595F:   mov eax, var_4
loc_005A5962:   mov ecx, var_14
loc_005A5965:   pop edi
loc_005A5966:   pop esi
loc_005A5967:   mov fs: [00000000h] , ecx
loc_005A596E:   pop ebx
loc_005A596F:   mov esp, ebp
loc_005A5971:   pop ebp
  loc_005A5972: retn 0004h
End Sub

Private Sub MnuLapReturJualSemua_Click() '5A3AD0
loc_005A3AD0:   push ebp
loc_005A3AD1:   mov ebp, esp
  loc_005A3AD3: sub esp, 0000000Ch
  loc_005A3AD6: push 00405696h ; __vbaExceptHandler
loc_005A3ADB:   mov eax, fs: [00000000h]
loc_005A3AE1:   push eax
loc_005A3AE2:   mov fs: [00000000h] , esp
  loc_005A3AE9: sub esp, 00000008h
loc_005A3AEC:   push ebx
loc_005A3AED:   push esi
loc_005A3AEE:   push edi
loc_005A3AEF:   mov var_C, esp
  loc_005A3AF2: mov var_8, 00402BE8h
loc_005A3AF9:   mov eax, Me
loc_005A3AFC:   mov ecx, eax
  loc_005A3AFE: and ecx, 00000001h
loc_005A3B01:   mov var_4, ecx
  loc_005A3B04: and al, FEh
loc_005A3B06:   push eax
loc_005A3B07:   mov Me, eax
loc_005A3B0A:   mov edx, [eax]
loc_005A3B0C:   Call [edx+00000004h]
  loc_005A3B0F: call 00611830h
  loc_005A3B14: mov var_4, 00000000h
loc_005A3B1B:   mov eax, Me
loc_005A3B1E:   push eax
loc_005A3B1F:   mov ecx, [eax]
loc_005A3B21:   Call [ecx+00000008h]
loc_005A3B24:   mov eax, var_4
loc_005A3B27:   mov ecx, var_14
loc_005A3B2A:   pop edi
loc_005A3B2B:   pop esi
loc_005A3B2C:   mov fs: [00000000h] , ecx
loc_005A3B33:   pop ebx
loc_005A3B34:   mov esp, ebp
loc_005A3B36:   pop ebp
  loc_005A3B37: retn 0004h
End Sub

Private Sub Timer3_Timer() '5A5D50
loc_005A5D50:   push ebp
loc_005A5D51:   mov ebp, esp
  loc_005A5D53: sub esp, 0000000Ch
  loc_005A5D56: push 00405696h ; __vbaExceptHandler
loc_005A5D5B:   mov eax, fs: [00000000h]
loc_005A5D61:   push eax
loc_005A5D62:   mov fs: [00000000h] , esp
  loc_005A5D69: sub esp, 0000001Ch
loc_005A5D6C:   push ebx
loc_005A5D6D:   push esi
loc_005A5D6E:   push edi
loc_005A5D6F:   mov var_C, esp
  loc_005A5D72: mov var_8, 00402D10h
loc_005A5D79:   mov esi, Me
loc_005A5D7C:   mov eax, esi
  loc_005A5D7E: and eax, 00000001h
loc_005A5D81:   mov var_4, eax
  loc_005A5D84: and esi, FFFFFFFEh
loc_005A5D87:   push esi
loc_005A5D88:   mov Me, esi
loc_005A5D8B:   mov ecx, [esi]
loc_005A5D8D:   Call [ecx+00000004h]
loc_005A5D90:   mov edx, [esi]
  loc_005A5D92: xor eax, eax
loc_005A5D94:   push esi
loc_005A5D95:   mov var_18, eax
loc_005A5D98:   mov var_1C, eax
loc_005A5D9B:   Call [edx+0000033Ch]
  loc_005A5DA1: mov ebx, [004010B8h] ; __vbaObjSet
loc_005A5DA7:   push eax
loc_005A5DA8:   lea eax, var_18
loc_005A5DAB:   push eax
loc_005A5DAC:   Call ebx
loc_005A5DAE:   mov edi, eax
loc_005A5DB0:   lea edx, var_1C
loc_005A5DB3:   push edx
loc_005A5DB4:   push edi
loc_005A5DB5:   mov ecx, [edi]
loc_005A5DB7:   Call [ecx+00000090h]
loc_005A5DBD:   test eax, eax
loc_005A5DBF:   fnclex
  loc_005A5DC1: jge 005A5DD5h
  loc_005A5DC3: push 00000090h
  loc_005A5DC8: push 00433F80h
loc_005A5DCD:   push edi
loc_005A5DCE:   push eax
  loc_005A5DCF: call [00401080h] ; __vbaHresultCheckObj
  loc_005A5DD5: xor eax, eax
loc_005A5DD7:   cmp var_1C, FFFFFFh
loc_005A5DDC:   lea ecx, var_18
loc_005A5DDF:   setz al
loc_005A5DE2:   neg eax
loc_005A5DE4:   mov edi, eax
  loc_005A5DE6: call [004012A0h] ; __vbaFreeObj
loc_005A5DEC:   mov ecx, [esi]
loc_005A5DEE:   push esi
loc_005A5DEF:   test di, di
  loc_005A5DF2: jz 005A5E16h
loc_005A5DF4:   Call [ecx+0000033Ch]
loc_005A5DFA:   lea edx, var_18
loc_005A5DFD:   push eax
loc_005A5DFE:   push edx
loc_005A5DFF:   Call ebx
loc_005A5E01:   mov esi, eax
  loc_005A5E03: push 00000000h
loc_005A5E05:   push esi
loc_005A5E06:   mov eax, [esi]
loc_005A5E08:   Call [eax+00000094h]
loc_005A5E0E:   test eax, eax
loc_005A5E10:   fnclex
  loc_005A5E12: jge 005A5E48h
  loc_005A5E14: jmp 005A5E36h
loc_005A5E16:   Call [ecx+0000033Ch]
loc_005A5E1C:   lea edx, var_18
loc_005A5E1F:   push eax
loc_005A5E20:   push edx
loc_005A5E21:   Call ebx
loc_005A5E23:   mov esi, eax
loc_005A5E25:   push FFFFFFFFh
loc_005A5E27:   push esi
loc_005A5E28:   mov eax, [esi]
loc_005A5E2A:   Call [eax+00000094h]
loc_005A5E30:   test eax, eax
loc_005A5E32:   fnclex
  loc_005A5E34: jge 005A5E48h
  loc_005A5E36: push 00000094h
  loc_005A5E3B: push 00433F80h
loc_005A5E40:   push esi
loc_005A5E41:   push eax
  loc_005A5E42: call [00401080h] ; __vbaHresultCheckObj
loc_005A5E48:   lea ecx, var_18
  loc_005A5E4B: call [004012A0h] ; __vbaFreeObj
  loc_005A5E51: mov var_4, 00000000h
  loc_005A5E58: push 005A5E6Ah
  loc_005A5E5D: jmp 005A5E69h
loc_005A5E5F:   lea ecx, var_18
  loc_005A5E62: call [004012A0h] ; __vbaFreeObj
loc_005A5E68:   ret
loc_005A5E69:   ret
loc_005A5E6A:   mov eax, Me
loc_005A5E6D:   push eax
loc_005A5E6E:   mov ecx, [eax]
loc_005A5E70:   Call [ecx+00000008h]
loc_005A5E73:   mov eax, var_4
loc_005A5E76:   mov ecx, var_14
loc_005A5E79:   pop edi
loc_005A5E7A:   pop esi
loc_005A5E7B:   mov fs: [00000000h] , ecx
loc_005A5E82:   pop ebx
loc_005A5E83:   mov esp, ebp
loc_005A5E85:   pop ebp
  loc_005A5E86: retn 0004h
End Sub

Private Sub mnuPelanggan_Click() '5A43A0
loc_005A43A0:   push ebp
loc_005A43A1:   mov ebp, esp
  loc_005A43A3: sub esp, 0000000Ch
  loc_005A43A6: push 00405696h ; __vbaExceptHandler
loc_005A43AB:   mov eax, fs: [00000000h]
loc_005A43B1:   push eax
loc_005A43B2:   mov fs: [00000000h] , esp
  loc_005A43B9: sub esp, 00000030h
loc_005A43BC:   push ebx
loc_005A43BD:   push esi
loc_005A43BE:   push edi
loc_005A43BF:   mov var_C, esp
  loc_005A43C2: mov var_8, 00402C20h
loc_005A43C9:   mov eax, Me
loc_005A43CC:   mov ecx, eax
  loc_005A43CE: and ecx, 00000001h
loc_005A43D1:   mov var_4, ecx
  loc_005A43D4: and al, FEh
loc_005A43D6:   push eax
loc_005A43D7:   mov Me, eax
loc_005A43DA:   mov edx, [eax]
loc_005A43DC:   Call [edx+00000004h]
loc_005A43DF:   mov eax, [00668164h]
loc_005A43E4:   test eax, eax
  loc_005A43E6: jnz 005A43F8h
  loc_005A43E8: push 00668164h
  loc_005A43ED: push 00413F9Ch
  loc_005A43F2: call [004011E8h] ; __vbaNew2
  loc_005A43F8: sub esp, 00000010h
  loc_005A43FB: mov ecx, 0000000Ah
loc_005A4400:   mov ebx, esp
  loc_005A4402: mov eax, 80020004h
  loc_005A4407: sub esp, 00000010h
loc_005A440A:   mov esi, [00668164h]
loc_005A4410:   mov [ebx], ecx
loc_005A4412:   mov ecx, var_30
loc_005A4415:   mov edi, [esi]
  loc_005A4417: mov edx, 00000001h
loc_005A441C:   mov [ebx+00000004h], ecx
loc_005A441F:   mov ecx, esp
loc_005A4421:   push esi
loc_005A4422:   mov [ebx+00000008h], eax
loc_005A4425:   mov eax, var_28
loc_005A4428:   mov [ebx+0000000Ch], eax
  loc_005A442B: mov eax, 00000002h
loc_005A4430:   mov [ecx], eax
loc_005A4432:   mov eax, var_20
loc_005A4435:   mov [ecx+00000004h], eax
loc_005A4438:   mov [ecx+00000008h], edx
loc_005A443B:   mov edx, var_18
loc_005A443E:   mov [ecx+0000000Ch], edx
loc_005A4441:   Call [edi+000002B0h]
loc_005A4447:   test eax, eax
loc_005A4449:   fnclex
  loc_005A444B: jge 005A445Fh
  loc_005A444D: push 000002B0h
  loc_005A4452: push 0042F03Ch
loc_005A4457:   push esi
loc_005A4458:   push eax
  loc_005A4459: call [00401080h] ; __vbaHresultCheckObj
  loc_005A445F: mov var_4, 00000000h
loc_005A4466:   mov eax, Me
loc_005A4469:   push eax
loc_005A446A:   mov ecx, [eax]
loc_005A446C:   Call [ecx+00000008h]
loc_005A446F:   mov eax, var_4
loc_005A4472:   mov ecx, var_14
loc_005A4475:   pop edi
loc_005A4476:   pop esi
loc_005A4477:   mov fs: [00000000h] , ecx
loc_005A447E:   pop ebx
loc_005A447F:   mov esp, ebp
loc_005A4481:   pop ebp
  loc_005A4482: retn 0004h
End Sub

Private Sub mnuPemasok_Click() '5A4490
loc_005A4490:   push ebp
loc_005A4491:   mov ebp, esp
  loc_005A4493: sub esp, 0000000Ch
  loc_005A4496: push 00405696h ; __vbaExceptHandler
loc_005A449B:   mov eax, fs: [00000000h]
loc_005A44A1:   push eax
loc_005A44A2:   mov fs: [00000000h] , esp
  loc_005A44A9: sub esp, 00000030h
loc_005A44AC:   push ebx
loc_005A44AD:   push esi
loc_005A44AE:   push edi
loc_005A44AF:   mov var_C, esp
  loc_005A44B2: mov var_8, 00402C28h
loc_005A44B9:   mov eax, Me
loc_005A44BC:   mov ecx, eax
  loc_005A44BE: and ecx, 00000001h
loc_005A44C1:   mov var_4, ecx
  loc_005A44C4: and al, FEh
loc_005A44C6:   push eax
loc_005A44C7:   mov Me, eax
loc_005A44CA:   mov edx, [eax]
loc_005A44CC:   Call [edx+00000004h]
loc_005A44CF:   mov eax, [00668178h]
loc_005A44D4:   test eax, eax
  loc_005A44D6: jnz 005A44E8h
  loc_005A44D8: push 00668178h
  loc_005A44DD: push 00416CD8h
  loc_005A44E2: call [004011E8h] ; __vbaNew2
  loc_005A44E8: sub esp, 00000010h
  loc_005A44EB: mov ecx, 0000000Ah
loc_005A44F0:   mov ebx, esp
  loc_005A44F2: mov eax, 80020004h
  loc_005A44F7: sub esp, 00000010h
loc_005A44FA:   mov esi, [00668178h]
loc_005A4500:   mov [ebx], ecx
loc_005A4502:   mov ecx, var_30
loc_005A4505:   mov edi, [esi]
  loc_005A4507: mov edx, 00000001h
loc_005A450C:   mov [ebx+00000004h], ecx
loc_005A450F:   mov ecx, esp
loc_005A4511:   push esi
loc_005A4512:   mov [ebx+00000008h], eax
loc_005A4515:   mov eax, var_28
loc_005A4518:   mov [ebx+0000000Ch], eax
  loc_005A451B: mov eax, 00000002h
loc_005A4520:   mov [ecx], eax
loc_005A4522:   mov eax, var_20
loc_005A4525:   mov [ecx+00000004h], eax
loc_005A4528:   mov [ecx+00000008h], edx
loc_005A452B:   mov edx, var_18
loc_005A452E:   mov [ecx+0000000Ch], edx
loc_005A4531:   Call [edi+000002B0h]
loc_005A4537:   test eax, eax
loc_005A4539:   fnclex
  loc_005A453B: jge 005A454Fh
  loc_005A453D: push 000002B0h
  loc_005A4542: push 0042F6DCh ; "ø'³:©†éC·"
loc_005A4547:   push esi
loc_005A4548:   push eax
  loc_005A4549: call [00401080h] ; __vbaHresultCheckObj
  loc_005A454F: mov var_4, 00000000h
loc_005A4556:   mov eax, Me
loc_005A4559:   push eax
loc_005A455A:   mov ecx, [eax]
loc_005A455C:   Call [ecx+00000008h]
loc_005A455F:   mov eax, var_4
loc_005A4562:   mov ecx, var_14
loc_005A4565:   pop edi
loc_005A4566:   pop esi
loc_005A4567:   mov fs: [00000000h] , ecx
loc_005A456E:   pop ebx
loc_005A456F:   mov esp, ebp
loc_005A4571:   pop ebp
  loc_005A4572: retn 0004h
End Sub

Private Function Proc_16_60_5A2EB0(arg_C) '5A2EB0
loc_005A2EB0:   mov eax, [00668164h]
  loc_005A2EB5: sub esp, 00000010h
loc_005A2EB8:   test eax, eax
loc_005A2EBA:   push ebx
loc_005A2EBB:   push ebp
loc_005A2EBC:   push esi
loc_005A2EBD:   push edi
  loc_005A2EBE: jnz 005A2ED0h
  loc_005A2EC0: push 00668164h
  loc_005A2EC5: push 00413F9Ch
  loc_005A2ECA: call [004011E8h] ; __vbaNew2
  loc_005A2ED0: sub esp, 00000010h
  loc_005A2ED3: mov ecx, 0000000Ah
loc_005A2ED8:   mov ebp, esp
  loc_005A2EDA: mov eax, 80020004h
  loc_005A2EDF: sub esp, 00000010h
loc_005A2EE2:   mov esi, [00668164h]
loc_005A2EE8:   mov [ebp], ecx
loc_005A2EEB:   mov ecx, arg_2C
  loc_005A2EEF: mov edi, 00000002h
loc_005A2EF4:   mov ebx, [esi]
loc_005A2EF6:   mov Proc_16_60_5A2EB0, ecx
loc_005A2EF9:   mov ecx, esp
  loc_005A2EFB: mov edx, 00000001h
loc_005A2F00:   push esi
loc_005A2F01:   mov Me, eax
loc_005A2F04:   mov eax, arg_38
loc_005A2F08:   mov arg_C, eax
loc_005A2F0B:   mov eax, arg_30
loc_005A2F0F:   mov [ecx], edi
loc_005A2F11:   mov [ecx+00000004h], eax
loc_005A2F14:   mov [ecx+00000008h], edx
loc_005A2F17:   mov edx, arg_38
loc_005A2F1B:   mov [ecx+0000000Ch], edx
loc_005A2F1E:   Call [ebx+000002B0h]
loc_005A2F24:   test eax, eax
loc_005A2F26:   fnclex
  loc_005A2F28: jge 005A2F3Ch
  loc_005A2F2A: push 000002B0h
  loc_005A2F2F: push 0042F03Ch
loc_005A2F34:   push esi
loc_005A2F35:   push eax
  loc_005A2F36: call [00401080h] ; __vbaHresultCheckObj
loc_005A2F3C:   pop edi
loc_005A2F3D:   pop esi
loc_005A2F3E:   pop ebp
  loc_005A2F3F: xor eax, eax
loc_005A2F41:   pop ebx
  loc_005A2F42: add esp, 00000010h
  loc_005A2F45: retn 0004h
End Function

Private Function Proc_16_61_5A2F50()
  loc_005A2F50: xor eax, eax
  loc_005A2F52: retn 0004h
End Sub

